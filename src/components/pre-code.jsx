import React, { useState, useEffect, useCallback } from "react";
import * as mammoth from "mammoth";
import * as Diff from "diff";

const styleMap = [
  "table => table",
  "tr => tr",
  "td => td",
  "th => th",
  "p[style-name='Normal'] => p",
  "p[style-name='Heading 1'] => h1",
  "p[style-name='Heading 2'] => h2",
  "ul => ul",
  "ol => ol",
  "li => li",
  "b => strong",
  "i => em",
  "u => u",
  "p[style-name='Title'] => h2:fresh",
  "p[style-name='Subtitle'] => h3:fresh",
  // Thêm các style khác nếu cần
  // Bổ sung cho danh sách đánh số
  "p[style-name='List Number'] => ol > li:fresh",
  "p[style-name='List Number 2'] => ol > li:fresh",
  "p[style-name='List Number 3'] => ol > li:fresh",
  "p[style-name='List Paragraph'] => ol > li:fresh",
];

const DocxToHtml = () => {
  const [originalHtmlContent, setOriginalHtmlContent] = useState("");
  const [editedHtmlContent, setEditedHtmlContent] = useState("");
  const [diffHtmlContent, setDiffHtmlContent] = useState("");
  const [error, setError] = useState("");
  const [originalFile, setOriginalFile] = useState(null);
  const [editedFile, setEditedFile] = useState(null);
  const [showDiff, setShowDiff] = useState(false);
  const [processingMode, setProcessingMode] = useState("auto"); // auto, simple, complex

  const handleOriginalFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (
      file.type !==
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      setError("Vui lòng chọn file DOCX hợp lệ!");
      return;
    }
    setOriginalFile(file);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap });
      setOriginalHtmlContent(result.value);
      setError("");
    } catch (err) {
      setError("Lỗi khi đọc file DOCX gốc: " + err.message);
    }
  };

  const handleEditedFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (
      file.type !==
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      setError("Vui lòng chọn file DOCX hợp lệ!");
      return;
    }
    setEditedFile(file);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap });
      setEditedHtmlContent(result.value);
      setError("");
    } catch (err) {
      setError("Lỗi khi đọc file DOCX đã sửa: " + err.message);
    }
  };

  const generateDiff = useCallback(() => {
    try {
      // Create a new HTML document that preserves formatting but highlights changes
      const parser = new DOMParser();
      const originalDoc = parser.parseFromString(
        originalHtmlContent,
        "text/html"
      );
      const diffDoc = parser.parseFromString(editedHtmlContent, "text/html");

      // Function to handle paragraphs and other non-table elements
      const processNonTableText = () => {
        const originalParagraphs = Array.from(
          originalDoc.querySelectorAll(
            "p, h1, h2, h3, h4, h5, h6, li, div:not(:has(table))"
          )
        ).filter((el) => !el.closest("table"));

        const editedParagraphs = Array.from(
          diffDoc.querySelectorAll(
            "p, h1, h2, h3, h4, h5, h6, li, div:not(:has(table))"
          )
        ).filter((el) => !el.closest("table"));

        const minParaCount = Math.min(
          originalParagraphs.length,
          editedParagraphs.length
        );
        for (let i = 0; i < minParaCount; i++) {
          const origPara = originalParagraphs[i];
          const editedPara = editedParagraphs[i];
          // So sánh từng dòng HTML (giữ style)
          const origLines = origPara.innerHTML.split(/<br\s*\/?>(?![^<]*>)/gi);
          const editedLines =
            editedPara.innerHTML.split(/<br\s*\/?>(?![^<]*>)/gi);
          const lineDiffs = Diff.diffLines(
            origLines.join("<br>"),
            editedLines.join("<br>")
          );
          let newParaHtml = "";
          lineDiffs.forEach((part) => {
            const lines = part.value.split(/<br\s*\/?>(?![^<]*>)/gi);
            lines.forEach((line, idx) => {
              if (line === "" && idx === lines.length - 1) return;
              if (part.added) {
                newParaHtml += `<span style="background-color: #CCFFCC; color: green; font-weight: bold; display:inline-block; width:100%">${line}</span><br>`;
              } else if (part.removed) {
                newParaHtml += `<span style="background-color: #FFCCCC; color: red; text-decoration: line-through; display:inline-block; width:100%">${line}</span><br>`;
              } else {
                newParaHtml += `${line}<br>`;
              }
            });
          });
          // Xóa <br> cuối cùng nếu có
          if (newParaHtml.endsWith("<br>"))
            newParaHtml = newParaHtml.slice(0, -4);
          editedPara.innerHTML = newParaHtml;
        }
      };

      // Function to detect merged cells in a table
      const hasTableMergedCells = (table) => {
        const cells = table.querySelectorAll("td, th");
        for (const cell of cells) {
          if (cell.hasAttribute("rowspan") || cell.hasAttribute("colspan")) {
            return true;
          }
        }
        return false;
      };

      // Function to check if a table has section headers (like A., B., etc.)
      const hasTableSections = (table) => {
        const rows = table.querySelectorAll("tr");
        let foundSectionHeader = false;

        for (const row of rows) {
          const firstCell = row.querySelector("td, th");
          if (firstCell) {
            const text = firstCell.textContent.trim();
            // Check for pattern like "A.", "B.", "I.", "II." etc.
            if (/^[A-Z]\.|\s[A-Z]\.|\s[IVX]+\./.test(text)) {
              foundSectionHeader = true;
              break;
            }
          }
        }

        return foundSectionHeader;
      };

      // Function to split a table into sections based on headers
      const splitTableIntoSections = (table) => {
        const rows = Array.from(table.querySelectorAll("tr"));
        const sections = [];
        let currentSection = [];
        let sectionHeader = "";

        rows.forEach((row, index) => {
          const firstCell = row.querySelector("td, th");
          if (firstCell) {
            const text = firstCell.textContent.trim();
            // Check if this is a new section header
            if (index > 0 && /^[A-Z]\.|\s[A-Z]\.|\s[IVX]+\./.test(text)) {
              if (currentSection.length > 0) {
                sections.push({
                  header: sectionHeader,
                  rows: [...currentSection],
                });
                currentSection = [];
              }
              sectionHeader = text;
            }
            currentSection.push(row);
          }
        });

        // Add the last section
        if (currentSection.length > 0) {
          sections.push({
            header: sectionHeader,
            rows: [...currentSection],
          });
        }

        return sections;
      };

      // Function to count matching cells content in tables
      const countMatchingCellContent = (table1, table2) => {
        const cells1 = Array.from(table1.querySelectorAll("td, th"));
        const cells2 = Array.from(table2.querySelectorAll("td, th"));

        let matchCount = 0;
        const minCellCount = Math.min(cells1.length, cells2.length);

        for (let i = 0; i < minCellCount; i++) {
          if (cells1[i].textContent.trim() === cells2[i].textContent.trim()) {
            matchCount++;
          }
        }

        return {
          matchCount,
          total: minCellCount,
          matchRatio: minCellCount > 0 ? matchCount / minCellCount : 0,
        };
      };

      // Function to process tables - AUTO DETECT which mode to use
      const processTables = () => {
        const originalTables = originalDoc.querySelectorAll("table");
        const editedTables = diffDoc.querySelectorAll("table");
        const minTableCount = Math.min(
          originalTables.length,
          editedTables.length
        );

        for (let tableIndex = 0; tableIndex < minTableCount; tableIndex++) {
          const originalTable = originalTables[tableIndex];
          const editedTable = editedTables[tableIndex];

          // AUTO DETECT: Check which processing mode to use
          let tableMode = processingMode;

          if (processingMode === "auto") {
            // First check if table has section headers (like A., B., etc.)
            if (
              hasTableSections(originalTable) ||
              hasTableSections(editedTable)
            ) {
              // Process each section separately as simple tables
              processTableWithSections(originalTable, editedTable);
              continue; // Skip standard processing for this table pair
            }
            // If table has merged cells, use "complex" mode
            else if (
              hasTableMergedCells(originalTable) ||
              hasTableMergedCells(editedTable)
            ) {
              tableMode = "complex";
            } else {
              // Check if table structures are similar by comparing matching content
              const matchInfo = countMatchingCellContent(
                originalTable,
                editedTable
              );
              // If tables have very different content, use complex mode, otherwise use simple
              tableMode = matchInfo.matchRatio < 0.5 ? "complex" : "simple";
            }
          }

          // Process the table based on the determined mode
          if (tableMode === "complex") {
            processComplexTable(originalTable, editedTable);
          } else {
            processSimpleTable(originalTable, editedTable);
          }
        }

        // Handle extra tables in edited document
        for (let i = minTableCount; i < editedTables.length; i++) {
          const editedTable = editedTables[i];
          markTableAsAdded(editedTable);
        }
      };

      // Function to process tables with section headers
      const processTableWithSections = (originalTable, editedTable) => {
        // Split tables into sections
        const originalSections = splitTableIntoSections(originalTable);
        const editedSections = splitTableIntoSections(editedTable);

        // Process each section separately
        const minSectionCount = Math.min(
          originalSections.length,
          editedSections.length
        );

        // Copy overall table attributes to preserve structure
        copyAttributes(originalTable, editedTable);

        // Process matching sections
        for (let i = 0; i < minSectionCount; i++) {
          const origSection = originalSections[i];
          const editSection = editedSections[i];

          // Check if both sections have rows
          if (origSection.rows.length === 0 || editSection.rows.length === 0) {
            continue; // Skip if either section has no rows
          }

          // For each row in this section, compare them as simple table rows
          const minRowCount = Math.min(
            origSection.rows.length,
            editSection.rows.length
          );

          for (let j = 0; j < minRowCount; j++) {
            const origRow = origSection.rows[j];
            const editRow = editSection.rows[j];

            compareTableRows(origRow, editRow);
          }

          // Handle extra rows in edited section - ensure they're properly marked as added
          for (let j = minRowCount; j < editSection.rows.length; j++) {
            markRowAsAdded(editSection.rows[j]);
          }
        }

        // Handle completely new rows that don't belong to any section
        const allOriginalRows = new Set(
          Array.from(originalTable.querySelectorAll("tr"))
        );
        const allEditedRows = Array.from(editedTable.querySelectorAll("tr"));

        // Any row in edited table that doesn't have a matching row in original should be marked as added
        allEditedRows.forEach((editedRow) => {
          // Check if this row was processed in any section
          let wasProcessed = false;

          for (const section of editedSections) {
            if (section.rows.includes(editedRow)) {
              wasProcessed = true;
              break;
            }
          }

          // If the row wasn't processed and doesn't have a match in the original table, mark it as added
          if (!wasProcessed && !allOriginalRows.has(editedRow)) {
            markRowAsAdded(editedRow);
          }
        });

        // Handle extra sections in edited table
        for (let i = minSectionCount; i < editedSections.length; i++) {
          const editSection = editedSections[i];
          editSection.rows.forEach((row) => {
            markRowAsAdded(row);
          });
        }
      };

      // COMPLEX MODE: For tables with merged cells or complex structure
      const processComplexTable = (originalTable, editedTable) => {
        // Copy table attributes to preserve structure
        copyAttributes(originalTable, editedTable);

        const originalRows = originalTable.querySelectorAll("tr");
        const editedRows = editedTable.querySelectorAll("tr");

        // Track all processed rows to find unprocessed ones later
        const processedEditedRows = new Set();

        // Create a map of row indices to row elements for both tables
        const originalRowMap = new Map();
        const editedRowMap = new Map();

        // Track rows that don't have numeric indices
        const originalNonIndexedRows = [];
        const editedNonIndexedRows = [];

        // Extract row indices from the first column of each row
        originalRows.forEach((row) => {
          const firstCell = row.querySelector("td, th");
          if (firstCell) {
            const rowIndex = firstCell.textContent.trim();
            // Only use as index if it's a number or looks like a numbered item
            if (/^\d+\.?$|^\d+\)$|^\d+$/.test(rowIndex)) {
              // Extract just the number part
              const numericIndex = rowIndex.replace(/[^\d]/g, "");
              originalRowMap.set(numericIndex, row);
            } else {
              // Store rows without numeric indices separately
              originalNonIndexedRows.push(row);
            }
          }
        });

        editedRows.forEach((row) => {
          const firstCell = row.querySelector("td, th");
          if (firstCell) {
            const rowIndex = firstCell.textContent.trim();
            if (/^\d+\.?$|^\d+\)$|^\d+$/.test(rowIndex)) {
              const numericIndex = rowIndex.replace(/[^\d]/g, "");
              editedRowMap.set(numericIndex, row);
            } else {
              // Store rows without numeric indices separately
              editedNonIndexedRows.push(row);
            }
          }
        });

        // Process rows that exist in both tables, matched by index
        for (const [index, origRow] of originalRowMap.entries()) {
          const editedRow = editedRowMap.get(index);
          if (editedRow) {
            // Compare rows and preserve structure
            compareAndPreserveTableRowStructure(origRow, editedRow);
            // Mark as processed
            processedEditedRows.add(editedRow);
            // Remove processed rows to track what's left
            originalRowMap.delete(index);
            editedRowMap.delete(index);
          }
        }

        // Handle non-indexed rows by comparing content similarity
        const matchedNonIndexedRows = new Set();
        compareNonIndexedRows(
          originalNonIndexedRows,
          editedNonIndexedRows,
          matchedNonIndexedRows
        );

        // Add matched non-indexed rows to the processed set
        matchedNonIndexedRows.forEach((row) => {
          processedEditedRows.add(row);
        });

        // Mark any remaining rows in editedRowMap as added (completely new rows with indices)
        for (const [, editedRow] of editedRowMap.entries()) {
          markRowAsAdded(editedRow);
          processedEditedRows.add(editedRow);
        }

        // Mark any remaining non-indexed rows that weren't matched as added
        editedNonIndexedRows.forEach((row) => {
          if (!matchedNonIndexedRows.has(row)) {
            markRowAsAdded(row);
            processedEditedRows.add(row);
          }
        });

        // Final check: Find any rows that weren't processed by any method and mark them as added
        Array.from(editedRows).forEach((row) => {
          if (!processedEditedRows.has(row)) {
            markRowAsAdded(row);
          }
        });

        // Special header check: Look for header-like rows with new parenthetical content
        Array.from(editedRows).forEach((row, index) => {
          // Skip if we don't have a corresponding original row to compare with
          if (index >= originalRows.length) return;

          const origRow = originalRows[index];
          const isHeaderRow =
            row.closest("thead") ||
            row.parentElement.firstChild === row ||
            index === 0;

          if (isHeaderRow) {
            const editedCells = row.querySelectorAll("td, th");
            const origCells = origRow.querySelectorAll("td, th");

            const minCellCount = Math.min(editedCells.length, origCells.length);

            for (let i = 0; i < minCellCount; i++) {
              const editedCell = editedCells[i];
              const origCell = origCells[i];

              const origContent = origCell.textContent.trim();
              const editedContent = editedCell.textContent.trim();

              // Look for added parenthetical content like "(Đồng)"
              if (
                editedContent.includes("(") &&
                (!origContent.includes("(") || origContent === "")
              ) {
                // Try to highlight just the parenthetical part
                const parentheticalRegex = /(\([^)]+\))/g;
                if (parentheticalRegex.test(editedCell.innerHTML)) {
                  // Replace all occurrences of parenthetical text with highlighted version
                  editedCell.innerHTML = editedCell.innerHTML.replace(
                    parentheticalRegex,
                    '<span style="background-color: #CCFFCC; color: green; font-weight: bold;">$1</span>'
                  );
                }
              }
            }
          }
        });
      };

      // SIMPLE MODE: For simple tables without merged cells
      const processSimpleTable = (originalTable, editedTable) => {
        const originalRows = Array.from(originalTable.querySelectorAll("tr"));
        const editedRows = Array.from(editedTable.querySelectorAll("tr"));

        // Compare rows by position
        const minRowCount = Math.min(originalRows.length, editedRows.length);

        for (let rowIndex = 0; rowIndex < minRowCount; rowIndex++) {
          const origRow = originalRows[rowIndex];
          const editedRow = editedRows[rowIndex];

          compareTableRows(origRow, editedRow);
        }

        // Handle extra rows in edited table - make sure to highlight them clearly
        for (
          let rowIndex = minRowCount;
          rowIndex < editedRows.length;
          rowIndex++
        ) {
          const editedRow = editedRows[rowIndex];
          // Double-check that we mark all cells in each added row
          const cells = editedRow.querySelectorAll("td, th");

          cells.forEach((cell) => {
            markCellAsAdded(cell);
          });

          // Also mark the entire row for good measure
          markRowAsAdded(editedRow);
        }
      };

      // Function to compare table rows for simple mode
      const compareTableRows = (origRow, editedRow) => {
        if (!origRow || !editedRow) return;

        const origCells = Array.from(origRow.querySelectorAll("td, th"));
        const editedCells = Array.from(editedRow.querySelectorAll("td, th"));

        // Compare cells by position
        const minCellCount = Math.min(origCells.length, editedCells.length);

        for (let cellIndex = 0; cellIndex < minCellCount; cellIndex++) {
          const origCell = origCells[cellIndex];
          const editedCell = editedCells[cellIndex];

          // Compare cell content
          compareCellContent(origCell, editedCell);
        }

        // Handle extra cells in edited row (cells that exist in edited but not in original)
        if (editedCells.length > origCells.length) {
          for (
            let cellIndex = minCellCount;
            cellIndex < editedCells.length;
            cellIndex++
          ) {
            const editedCell = editedCells[cellIndex];
            markCellAsAdded(editedCell);
          }
        }

        // Check header-type cells for new content even if original cells exist
        // This is to catch cases like adding "(Đồng)" to header cells
        if (minCellCount > 0) {
          // Specifically look for header or first row cells that might have new content
          const isHeaderRow =
            origRow.closest("thead") ||
            editedRow.closest("thead") ||
            origRow.parentElement.firstChild === origRow ||
            editedRow.parentElement.firstChild === editedRow;

          if (isHeaderRow) {
            for (let cellIndex = 0; cellIndex < minCellCount; cellIndex++) {
              const origCell = origCells[cellIndex];
              const editedCell = editedCells[cellIndex];

              const origContent = origCell.textContent.trim();
              const editedContent = editedCell.textContent.trim();

              // If content was added to an otherwise empty or minimal cell
              if (
                editedContent &&
                (origContent === "" || !origContent.includes(editedContent))
              ) {
                // Check specifically for the case where text like "(Đồng)" was added
                if (
                  editedContent.includes("(") &&
                  editedContent.includes(")")
                ) {
                  markCellAsAdded(editedCell);
                }
              }
            }
          }
        }
      };

      // CẢI TIẾN: Hàm so sánh và giữ nguyên cấu trúc hàng của bảng
      const compareAndPreserveTableRowStructure = (origRow, editedRow) => {
        if (!origRow || !editedRow) return;

        // 1. BẢO TOÀN CẤU TRÚC: Copy tất cả thuộc tính quan trọng từ hàng gốc
        copyAttributes(origRow, editedRow);

        // 2. Lấy tất cả ô từ cả hai hàng
        const originalCells = origRow.querySelectorAll("td, th");
        const editedCells = editedRow.querySelectorAll("td, th");

        // 3. Chỉ so sánh các ô có cùng vị trí trong hàng
        const minCellCount = Math.min(originalCells.length, editedCells.length);

        // Check if this is a header row (for special handling of header content)
        const isHeaderRow =
          origRow.closest("thead") ||
          editedRow.closest("thead") ||
          origRow.parentElement.firstChild === origRow ||
          editedRow.parentElement.firstChild === editedRow;

        for (let i = 0; i < minCellCount; i++) {
          const origCell = originalCells[i];
          const editedCell = editedCells[i];

          // 4. BẢO TOÀN CẤU TRÚC: Copy tất cả thuộc tính quan trọng từ ô gốc sang ô mới
          copyAttributes(origCell, editedCell);

          // 5. So sánh nội dung hai ô và highlight sự khác biệt
          compareAndHighlightCellContent(origCell, editedCell);

          // 6. Special check for header cells with parenthetical content like "(Đồng)"
          if (isHeaderRow) {
            const origContent = origCell.textContent.trim();
            const editedContent = editedCell.textContent.trim();

            // If content was added to an otherwise empty or minimal cell
            if (
              editedContent &&
              (origContent === "" ||
                (!origContent.includes("(") && editedContent.includes("(")))
            ) {
              // Check specifically for the case where text like "(Đồng)" was added
              if (editedContent.includes("(") && editedContent.includes(")")) {
                // Try to highlight just the parenthetical part
                const parentheticalRegex = /(\([^)]+\))/g;
                if (parentheticalRegex.test(editedCell.innerHTML)) {
                  // Replace all occurrences of parenthetical text with highlighted version
                  editedCell.innerHTML = editedCell.innerHTML.replace(
                    parentheticalRegex,
                    '<span style="background-color: #CCFFCC; color: green; font-weight: bold;">$1</span>'
                  );
                }
              }
            }
          }
        }

        // 7. Handle extra cells in edited row
        if (editedCells.length > originalCells.length) {
          for (let i = minCellCount; i < editedCells.length; i++) {
            markCellAsAdded(editedCells[i]);
          }
        }
      };

      // Hàm copy thuộc tính từ phần tử nguồn sang phần tử đích
      const copyAttributes = (source, target) => {
        if (!source || !target) return;

        // Đặc biệt chú ý đến các thuộc tính rowspan và colspan
        if (source.hasAttribute("rowspan")) {
          target.setAttribute("rowspan", source.getAttribute("rowspan"));
        }

        if (source.hasAttribute("colspan")) {
          target.setAttribute("colspan", source.getAttribute("colspan"));
        }

        // Giữ nguyên các thuộc tính class để bảo toàn style
        if (source.hasAttribute("class")) {
          target.setAttribute("class", source.getAttribute("class"));
        }

        // Giữ nguyên thuộc tính style
        if (source.hasAttribute("style")) {
          const sourceStyle = source.getAttribute("style");
          const targetStyle = target.getAttribute("style") || "";

          // Đảm bảo không ghi đè style hiện có
          if (!targetStyle.includes(sourceStyle)) {
            target.setAttribute("style", targetStyle + "; " + sourceStyle);
          }
        }
      };

      // Hàm so sánh và highlight sự khác biệt trong nội dung ô
      const compareAndHighlightCellContent = (origCell, editedCell) => {
        if (!origCell || !editedCell) return;

        // Lấy nội dung HTML của hai ô
        const origHtml = origCell.innerHTML;
        const editedHtml = editedCell.innerHTML;

        // Nếu nội dung giống nhau, không cần xử lý gì thêm
        if (origHtml === editedHtml) return;

        // So sánh nội dung dùng diff
        const origLines = origHtml.split(/<br\s*\/?>(?![^<]*>)/gi);
        const editedLines = editedHtml.split(/<br\s*\/?>(?![^<]*>)/gi);

        const cellDiff = Diff.diffLines(
          origLines.join("<br>"),
          editedLines.join("<br>")
        );

        // Tạo HTML mới với highlight sự khác biệt
        let newCellHtml = "";

        cellDiff.forEach((part) => {
          const lines = part.value.split(/<br\s*\/?>(?![^<]*>)/gi);

          lines.forEach((line, idx) => {
            // Bỏ qua dòng trống ở cuối
            if (line === "" && idx === lines.length - 1) return;

            if (part.added) {
              // Nội dung được thêm mới
              newCellHtml += `<span style="background-color: #CCFFCC; color: green; font-weight: bold; display:inline-block; width:100%">${line}</span><br>`;
            } else if (part.removed) {
              // Nội dung bị xóa
              newCellHtml += `<span style="background-color: #FFCCCC; color: red; text-decoration: line-through; display:inline-block; width:100%">${line}</span><br>`;
            } else {
              // Nội dung không thay đổi
              newCellHtml += `${line}<br>`;
            }
          });
        });

        // Xóa <br> cuối cùng nếu có
        if (newCellHtml.endsWith("<br>")) {
          newCellHtml = newCellHtml.slice(0, -4);
        }

        // Cập nhật nội dung ô
        editedCell.innerHTML = newCellHtml;
      };

      // Helper function to compare non-indexed rows by content similarity
      const compareNonIndexedRows = (
        originalRows,
        editedRows,
        matchedRows = new Set()
      ) => {
        if (originalRows.length === 0 || editedRows.length === 0) return;

        // For each original row, find the most similar edited row
        const processedEditedRows = new Set();

        originalRows.forEach((origRow) => {
          let bestMatch = null;
          let bestMatchIndex = -1;
          let highestSimilarity = -1;

          // Calculate a simple similarity score based on first cell content
          const origFirstCell = origRow.querySelector("td, th");
          if (!origFirstCell) return;

          // Try to find the best matching row in edited rows
          editedRows.forEach((editedRow, idx) => {
            if (processedEditedRows.has(idx)) return; // Skip already processed rows

            const editedFirstCell = editedRow.querySelector("td, th");
            if (!editedFirstCell) return;

            // Simple similarity: check if first cells have similar content
            const origContent = origFirstCell.textContent.trim();
            const editedContent = editedFirstCell.textContent.trim();

            // Calculate similarity (can be improved with more sophisticated algorithms)
            // Here we use a simple string similarity
            const similarity = calculateStringSimilarity(
              origContent,
              editedContent
            );

            if (similarity > highestSimilarity) {
              highestSimilarity = similarity;
              bestMatch = editedRow;
              bestMatchIndex = idx;
            }
          });

          // If we found a good match, compare the rows
          if (bestMatch && highestSimilarity > 0.5) {
            // Threshold can be adjusted
            compareAndPreserveTableRowStructure(origRow, bestMatch);
            processedEditedRows.add(bestMatchIndex);
            if (matchedRows) matchedRows.add(bestMatch);
          }
        });
      };

      // Simple string similarity function (Levenshtein distance based)
      const calculateStringSimilarity = (str1, str2) => {
        if (!str1 || !str2) return 0;
        if (str1 === str2) return 1.0;

        // Simple implementation of string similarity
        const len1 = str1.length;
        const len2 = str2.length;

        // If either string is empty, similarity is 0
        if (len1 === 0 || len2 === 0) return 0;

        // Calculate Levenshtein distance
        const distance = levenshteinDistance(str1, str2);

        // Convert distance to similarity score (0-1)
        return 1 - distance / Math.max(len1, len2);
      };

      // Levenshtein distance implementation
      const levenshteinDistance = (str1, str2) => {
        const m = str1.length;
        const n = str2.length;

        // Create a matrix of size (m+1) x (n+1)
        const dp = Array(m + 1)
          .fill()
          .map(() => Array(n + 1).fill(0));

        // Fill the first row and column
        for (let i = 0; i <= m; i++) dp[i][0] = i;
        for (let j = 0; j <= n; j++) dp[0][j] = j;

        // Fill the rest of the matrix
        for (let i = 1; i <= m; i++) {
          for (let j = 1; j <= n; j++) {
            const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
            dp[i][j] = Math.min(
              dp[i - 1][j] + 1, // deletion
              dp[i][j - 1] + 1, // insertion
              dp[i - 1][j - 1] + cost // substitution
            );
          }
        }

        return dp[m][n];
      };

      // Simple function to compare cell content
      const compareCellContent = (origCell, editedCell) => {
        if (!origCell || !editedCell) return;

        const origContent = origCell.innerHTML;
        const editedContent = editedCell.innerHTML;

        // If content is identical, no need to process
        if (origContent === editedContent) return;

        // Compare using diff
        const diffResult = Diff.diffWords(origContent, editedContent);

        let newHtml = "";
        diffResult.forEach((part) => {
          if (part.added) {
            newHtml += `<span style="background-color: #CCFFCC; color: green; font-weight: bold;">${part.value}</span>`;
          } else if (part.removed) {
            newHtml += `<span style="background-color: #FFCCCC; color: red; text-decoration: line-through;">${part.value}</span>`;
          } else {
            newHtml += part.value;
          }
        });

        editedCell.innerHTML = newHtml;
      };

      // Function to mark a row as added (green)
      const markRowAsAdded = (row) => {
        if (!row) return;

        const cells = row.querySelectorAll("td, th");
        cells.forEach((cell) => {
          if (cell) {
            // Check if the cell already has a highlighting span
            const hasHighlightSpan = cell.querySelector(
              'span[style*="background-color: #CCFFCC"]'
            );

            if (!hasHighlightSpan) {
              const content = cell.innerHTML;
              cell.innerHTML = `<span style="background-color: #CCFFCC; color: green; font-weight: bold; display:inline-block; width:100%">${content}</span>`;
            }
          }
        });
      };

      // Function to mark a cell as added (green)
      const markCellAsAdded = (cell) => {
        if (!cell) return;

        // Check if the cell already has a highlighting span
        const hasHighlightSpan = cell.querySelector(
          'span[style*="background-color: #CCFFCC"]'
        );

        if (!hasHighlightSpan) {
          const content = cell.innerHTML;

          // Special case for content with parentheses like "(Đồng)"
          if (content.includes("(") && content.includes(")")) {
            // Try to highlight just the parenthetical part if possible
            const parentheticalRegex = /(\([^)]+\))/g;
            if (parentheticalRegex.test(content)) {
              // Replace all occurrences of parenthetical text with highlighted version
              cell.innerHTML = content.replace(
                parentheticalRegex,
                '<span style="background-color: #CCFFCC; color: green; font-weight: bold;">$1</span>'
              );
              return;
            }
          }

          // Default case - highlight entire cell content
          cell.innerHTML = `<span style="background-color: #CCFFCC; color: green; font-weight: bold; display:inline-block; width:100%">${content}</span>`;
        }
      };

      // Function to mark a table as added (green)
      const markTableAsAdded = (table) => {
        if (!table) return;

        const cells = table.querySelectorAll("td, th");
        cells.forEach((cell) => {
          if (cell) {
            // Check if the cell already has a highlighting span
            const hasHighlightSpan = cell.querySelector(
              'span[style*="background-color: #CCFFCC"]'
            );

            if (!hasHighlightSpan) {
              const content = cell.innerHTML;
              cell.innerHTML = `<span style="background-color: #CCFFCC; color: green; font-weight: bold; display:inline-block; width:100%">${content}</span>`;
            }
          }
        });
      };

      // First process tables
      processTables();

      // Then process remaining non-table text
      processNonTableText();

      // Convert the highlighted document back to HTML
      setDiffHtmlContent(diffDoc.body.innerHTML);
    } catch (err) {
      console.error(err);
      setError("Lỗi khi tạo bản so sánh: " + err.message);
    }
  }, [originalHtmlContent, editedHtmlContent, processingMode]);

  useEffect(() => {
    if (originalHtmlContent && editedHtmlContent) {
      generateDiff();
    }
  }, [originalHtmlContent, editedHtmlContent, generateDiff]);

  const handleDownloadOriginalDocx = () => {
    if (!originalFile) return;
    const url = URL.createObjectURL(originalFile);
    const a = document.createElement("a");
    a.href = url;
    a.download = originalFile.name || "tai-lieu-goc.docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleDownloadEditedDocx = () => {
    if (!editedFile) return;
    const url = URL.createObjectURL(editedFile);
    const a = document.createElement("a");
    a.href = url;
    a.download = editedFile.name || "tai-lieu-da-sua.docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <div>
        <label>
          <strong>Tải lên tài liệu gốc:</strong>
          <input
            type="file"
            accept=".docx"
            onChange={handleOriginalFileChange}
            style={{ marginLeft: 8 }}
          />
        </label>
      </div>
      <div style={{ marginTop: 16 }}>
        <label>
          <strong>Tải lên tài liệu đã sửa:</strong>
          <input
            type="file"
            accept=".docx"
            onChange={handleEditedFileChange}
            style={{ marginLeft: 8 }}
          />
        </label>
      </div>

      <div style={{ marginTop: 16 }}>
        <label>
          <strong>Chế độ xử lý:</strong>
          <select
            value={processingMode}
            onChange={(e) => setProcessingMode(e.target.value)}
            style={{ marginLeft: 8 }}
          >
            <option value="auto">Tự động phát hiện</option>
            <option value="simple">Bảng đơn giản</option>
            <option value="complex">Bảng phức tạp (có merged cells)</option>
          </select>
        </label>
      </div>

      {error && <div style={{ color: "red" }}>{error}</div>}
      <div style={{ display: "flex", marginTop: 16 }}>
        {originalFile && (
          <button
            onClick={handleDownloadOriginalDocx}
            style={{ marginRight: 8 }}
          >
            Tải file DOCX gốc
          </button>
        )}
        {editedFile && (
          <button onClick={handleDownloadEditedDocx} style={{ marginRight: 8 }}>
            Tải file DOCX đã sửa
          </button>
        )}
        {originalFile && editedFile && (
          <button
            onClick={() => setShowDiff(!showDiff)}
            style={{
              marginRight: 8,
              backgroundColor: showDiff ? "#4CAF50" : "",
            }}
          >
            {showDiff ? "Ẩn thay đổi" : "Hiển thị thay đổi"}
          </button>
        )}
      </div>
      <div style={{ display: "flex", marginTop: 16 }}>
        {/* Hiển thị file gốc */}
        <div style={{ flex: 1, margin: "0 12px" }}>
          <h3>File gốc</h3>
          <div
            style={{
              border: "1px solid #ccc",
              padding: 12,
              minHeight: 200,
              background: "#f9f9f9",
              overflow: "auto",
            }}
            dangerouslySetInnerHTML={{ __html: originalHtmlContent }}
          />
        </div>
        {/* Hiển thị file sửa */}
        <div style={{ flex: 1, margin: "0 12px" }}>
          <h3>File sửa</h3>
          <div
            style={{
              border: "1px solid #ccc",
              padding: 12,
              minHeight: 200,
              background: "#f9f9f9",
              overflow: "auto",
            }}
            dangerouslySetInnerHTML={{ __html: editedHtmlContent }}
          />
        </div>
        {/* Hiển thị diff giữa 2 file */}
        <div style={{ flex: 1, margin: "0 12px" }}>
          <h3>So sánh thay đổi</h3>
          <div
            style={{
              border: "1px solid #ccc",
              padding: 12,
              minHeight: 200,
              background: "#fffbe7",
              overflow: "auto",
            }}
            dangerouslySetInnerHTML={{ __html: diffHtmlContent }}
          />
        </div>
      </div>
    </div>
  );
};

export default DocxToHtml;
