import React, { useState, useRef, useEffect, useCallback } from "react";
import { renderAsync } from "docx-preview";
import mammoth from "mammoth";
import * as Diff from "diff";

const TrackingDiffDocx = () => {
  const [originalFile, setOriginalFile] = useState(null);
  const [editedFile, setEditedFile] = useState(null);
  const [originalHtml, setOriginalHtml] = useState("");
  const [editedHtml, setEditedHtml] = useState("");
  const [diffHtml, setDiffHtml] = useState("");
  const [isComparing, setIsComparing] = useState(false);

  const originalPreviewRef = useRef(null);
  const editedPreviewRef = useRef(null);
  const inlineDiffRef = useRef(null);

  // Style map for mammoth conversion
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
    "p[style-name='List Number'] => ol > li:fresh",
    "p[style-name='List Number 2'] => ol > li:fresh",
    "p[style-name='List Number 3'] => ol > li:fresh",
    "p[style-name='List Paragraph'] => ol > li:fresh",
  ];

  // Helper to clean HTML before comparison
  const cleanHtml = useCallback((html) => {
    // Remove extra spaces and normalize
    return html
      .replace(/\s+/g, " ")
      .replace(/<\/p>\s*<p/g, "</p><p")
      .trim();
  }, []);

  // Detect merged cells in a table
  const hasTableMergedCells = useCallback((tableHtml) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(tableHtml, "text/html");
    return Array.from(doc.querySelectorAll("td, th")).some(
      (cell) => cell.hasAttribute("rowspan") || cell.hasAttribute("colspan")
    );
  }, []);

  // Mark a row as added
  const markRowAsAdded = useCallback((row) => {
    Array.from(row.querySelectorAll("td, th")).forEach((cell) => {
      cell.innerHTML = `<span class='added'>${cell.innerHTML}</span>`;
    });
  }, []);

  // Mark a row as removed
  const markRowAsRemoved = useCallback((row) => {
    Array.from(row.querySelectorAll("td, th")).forEach((cell) => {
      cell.innerHTML = `<span class='removed'>${cell.innerHTML}</span>`;
    });
  }, []);

  // Compare and highlight cell content
  const compareCellContent = useCallback((oCell, eCell) => {
    const o = oCell?.textContent || "";
    const n = eCell?.textContent || "";
    const diffs = Diff.diffWords(o, n);
    let out = "";
    diffs.forEach((part) => {
      if (part.added) out += `<span class='added'>${part.value}</span>`;
      else if (part.removed)
        out += `<span class='removed'>${part.value}</span>`;
      else out += part.value;
    });
    if (eCell) eCell.innerHTML = out;
  }, []);

  // Compare non-indexed rows by string similarity (Levenshtein)
  const levenshteinDistance = useCallback((a, b) => {
    const m = a.length,
      n = b.length;
    const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++)
      for (let j = 1; j <= n; j++)
        dp[i][j] =
          a[i - 1] === b[j - 1]
            ? dp[i - 1][j - 1]
            : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    return dp[m][n];
  }, []);

  const stringSimilarity = useCallback(
    (a, b) => {
      if (!a && !b) return 1;
      if (!a || !b) return 0;
      const maxLen = Math.max(a.length, b.length);
      if (maxLen === 0) return 1;
      return 1 - levenshteinDistance(a, b) / maxLen;
    },
    [levenshteinDistance]
  );

  const compareNonIndexedRows = useCallback(
    (origRows, editRows) => {
      // Greedy match by similarity > 0.5
      const used = new Set();
      origRows.forEach((oRow) => {
        let bestIdx = -1,
          bestSim = 0.5;
        const oText = oRow.textContent.trim();
        editRows.forEach((eRow, idx) => {
          if (used.has(idx)) return;
          const sim = stringSimilarity(oText, eRow.textContent.trim());
          if (sim > bestSim) {
            bestSim = sim;
            bestIdx = idx;
          }
        });
        if (bestIdx !== -1) {
          // Matched
          const eRow = editRows[bestIdx];
          used.add(bestIdx);
          const oCells = Array.from(oRow.querySelectorAll("td, th"));
          const eCells = Array.from(eRow.querySelectorAll("td, th"));
          const min = Math.min(oCells.length, eCells.length);
          for (let i = 0; i < min; i++)
            compareCellContent(oCells[i], eCells[i]);
          for (let i = min; i < oCells.length; i++)
            oCells[
              i
            ].innerHTML = `<span class='removed'>${oCells[i].innerHTML}</span>`;
          for (let i = min; i < eCells.length; i++)
            eCells[
              i
            ].innerHTML = `<span class='added'>${eCells[i].innerHTML}</span>`;
        } else {
          // Not matched, mark removed
          markRowAsRemoved(oRow);
        }
      });
      // Mark any remaining editRows as added
      editRows.forEach((eRow, idx) => {
        if (!used.has(idx)) markRowAsAdded(eRow);
      });
    },
    [compareCellContent, markRowAsAdded, markRowAsRemoved, stringSimilarity]
  );

  // Process complex table (merged cells, or non-numeric first cell)
  const processComplexTable = useCallback(
    (origTableHtml, editTableHtml) => {
      const parser = new DOMParser();
      const origDoc = parser.parseFromString(origTableHtml, "text/html");
      const editDoc = parser.parseFromString(editTableHtml, "text/html");
      const origRows = Array.from(origDoc.querySelectorAll("tr"));
      const editRows = Array.from(editDoc.querySelectorAll("tr"));
      // Map rows by first cell if numeric
      const getNumericIndex = (cell) => {
        const val = cell?.textContent.trim();
        return /^\d+$/.test(val) ? val : null;
      };
      const origMap = new Map(),
        editMap = new Map();
      const origNonIdx = [],
        editNonIdx = [];
      origRows.forEach((row) => {
        const idx = getNumericIndex(row.cells[0]);
        if (idx) origMap.set(idx, row);
        else origNonIdx.push(row);
      });
      editRows.forEach((row) => {
        const idx = getNumericIndex(row.cells[0]);
        if (idx) editMap.set(idx, row);
        else editNonIdx.push(row);
      });
      // Compare indexed rows
      origMap.forEach((oRow, idx) => {
        const eRow = editMap.get(idx);
        if (eRow) {
          const oCells = Array.from(oRow.querySelectorAll("td, th"));
          const eCells = Array.from(eRow.querySelectorAll("td, th"));
          const min = Math.min(oCells.length, eCells.length);
          for (let i = 0; i < min; i++)
            compareCellContent(oCells[i], eCells[i]);
          for (let i = min; i < oCells.length; i++)
            oCells[
              i
            ].innerHTML = `<span class='removed'>${oCells[i].innerHTML}</span>`;
          for (let i = min; i < eCells.length; i++)
            eCells[
              i
            ].innerHTML = `<span class='added'>${eCells[i].innerHTML}</span>`;
          editMap.delete(idx);
        } else {
          markRowAsRemoved(oRow);
        }
      });
      // Mark remaining editMap rows as added
      editMap.forEach((eRow) => markRowAsAdded(eRow));
      // Compare non-indexed rows
      compareNonIndexedRows(origNonIdx, editNonIdx);
      // Return HTML
      const t =
        editDoc.querySelector("table") || origDoc.querySelector("table");
      return t ? t.outerHTML : "";
    },
    [
      compareCellContent,
      markRowAsAdded,
      markRowAsRemoved,
      compareNonIndexedRows,
    ]
  );

  // Process simple table: DOM-based cell-by-cell diff preserving formatting
  const processSimpleTable = useCallback(
    (origHtml, editHtml) => {
      const parser = new DOMParser();
      const oDoc = parser.parseFromString(origHtml, "text/html");
      const eDoc = parser.parseFromString(editHtml, "text/html");
      const oTable = oDoc.querySelector("table");
      const eTable = eDoc.querySelector("table") || oTable;
      if (!oTable && !eTable) return "";
      for (
        let i = 0;
        i < Math.max(oTable?.rows.length || 0, eTable?.rows.length || 0);
        i++
      ) {
        const oRow = oTable?.rows[i];
        const eRow = eTable?.rows[i];
        if (oRow && eRow) {
          const cols = Math.max(oRow.cells.length, eRow.cells.length);
          for (let j = 0; j < cols; j++) {
            const oCell = oRow.cells[j];
            const eCell = eRow.cells[j];
            if (oCell && eCell) compareCellContent(oCell, eCell);
            else if (oCell) {
              const newCell = eDoc.createElement(oCell.tagName);
              newCell.innerHTML = `<span class='removed'>${oCell.innerHTML}</span>`;
              eRow.insertBefore(newCell, eRow.cells[j] || null);
            } else if (eCell)
              eCell.innerHTML = `<span class='added'>${eCell.innerHTML}</span>`;
          }
        } else if (oRow && !eRow) {
          const newRow = oRow.cloneNode(true);
          markRowAsRemoved(newRow);
          eTable.appendChild(newRow);
        } else if (!oRow && eRow) {
          markRowAsAdded(eRow);
        }
      }
      return eTable.outerHTML;
    },
    [compareCellContent, markRowAsAdded, markRowAsRemoved]
  );

  // Main diffTable: auto-select mode
  const diffTable = useCallback(
    (origTableHtml, editTableHtml) => {
      if (
        hasTableMergedCells(origTableHtml) ||
        hasTableMergedCells(editTableHtml)
      ) {
        return processComplexTable(origTableHtml, editTableHtml);
      }
      return processSimpleTable(origTableHtml, editTableHtml);
    },
    [hasTableMergedCells, processComplexTable, processSimpleTable]
  );

  // Generate structured diff that handles tables inline
  const generateStructuredDiff = useCallback(
    (originalHtml, editedHtml) => {
      const cleanOriginal = cleanHtml(originalHtml);
      const cleanEdited = cleanHtml(editedHtml);
      const segmentRegex = /(<table[\s\S]*?<\/table>)/g;
      const originalSegments = cleanOriginal.split(segmentRegex);
      const editedSegments = cleanEdited.split(segmentRegex);
      const diffSegments = Diff.diffArrays(originalSegments, editedSegments);
      let diffOutput = "";
      for (let i = 0; i < diffSegments.length; i++) {
        const part = diffSegments[i];
        if (part.removed) {
          const next = diffSegments[i + 1];
          if (next && next.added) {
            const removed = part.value;
            const added = next.value;
            const len = Math.max(removed.length, added.length);
            for (let j = 0; j < len; j++) {
              const origSeg = removed[j] || "";
              const newSeg = added[j] || "";
              if (
                origSeg.trim().startsWith("<table") &&
                newSeg.trim().startsWith("<table")
              ) {
                diffOutput += diffTable(origSeg, newSeg);
              } else {
                const wordDiff = Diff.diffWords(origSeg, newSeg);
                let segOut = "";
                wordDiff.forEach((wp) => {
                  if (wp.added)
                    segOut += `<span class="added">${wp.value}</span>`;
                  else if (wp.removed)
                    segOut += `<span class="removed">${wp.value}</span>`;
                  else segOut += wp.value;
                });
                diffOutput += `<p>${segOut}</p>`;
              }
            }
            i++;
          } else {
            part.value.forEach((seg) => {
              if (seg.trim().startsWith("<table"))
                diffOutput += `<div class="removed-table">${seg}</div>`;
              else diffOutput += `<p class="removed">${seg}</p>`;
            });
          }
        } else if (part.added) {
          part.value.forEach((seg) => {
            if (seg.trim().startsWith("<table"))
              diffOutput += `<div class="added-table">${seg}</div>`;
            else diffOutput += `<p class="added">${seg}</p>`;
          });
        } else {
          part.value.forEach((seg) => {
            diffOutput += seg;
          });
        }
      }
      setDiffHtml(diffOutput);
    },
    [cleanHtml, diffTable]
  );

  // Handle original file upload
  const handleOriginalFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setOriginalFile(file);
      convertDocxToHtml(file, setOriginalHtml);
      setDiffHtml(""); // Reset diff when a new file is uploaded
    }
  };

  // Handle edited file upload
  const handleEditedFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      setEditedFile(file);
      convertDocxToHtml(file, setEditedHtml);
      setDiffHtml(""); // Reset diff when a new file is uploaded
    }
  };

  // Convert DOCX to HTML using mammoth
  const convertDocxToHtml = (file, setHtmlFunc) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target.result;
      mammoth
        .convertToHtml({ arrayBuffer }, { styleMap })
        .then((result) => {
          setHtmlFunc(result.value);
        })
        .catch((error) => {
          console.error("Error converting DOCX to HTML:", error);
        });
    };
    reader.readAsArrayBuffer(file);
  };

  // Generate diff when both HTML contents are available
  useEffect(() => {
    if (originalHtml && editedHtml && isComparing) {
      // Create a more detailed diff that preserves structure
      generateStructuredDiff(originalHtml, editedHtml);
    }
  }, [originalHtml, editedHtml, isComparing, generateStructuredDiff]);

  // Handle compare button click
  const handleCompare = () => {
    if (originalHtml && editedHtml) {
      setIsComparing(true);
    }
  };

  // Preview original file when it changes
  useEffect(() => {
    if (originalFile && originalPreviewRef.current) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;

          // Clear previous content
          originalPreviewRef.current.innerHTML = "";

          // Render DOCX in the container
          await renderAsync(arrayBuffer, originalPreviewRef.current, null, {
            className: "docx-viewer",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
          });
        } catch (error) {
          console.error("Error previewing original file:", error);
          originalPreviewRef.current.innerHTML =
            '<div class="error">Error previewing document</div>';
        }
      };
      reader.readAsArrayBuffer(originalFile);
    }
  }, [originalFile]);

  // Preview edited file when it changes
  useEffect(() => {
    if (editedFile && editedPreviewRef.current) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;

          // Clear previous content
          editedPreviewRef.current.innerHTML = "";

          // Render DOCX in the container
          await renderAsync(arrayBuffer, editedPreviewRef.current, null, {
            className: "docx-viewer",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
          });
        } catch (error) {
          console.error("Error previewing edited file:", error);
          editedPreviewRef.current.innerHTML =
            '<div class="error">Error previewing document</div>';
        }
      };
      reader.readAsArrayBuffer(editedFile);
    }
  }, [editedFile]);

  // Display differences when diffHtml changes
  useEffect(() => {
    if (diffHtml) {
      if (inlineDiffRef.current) inlineDiffRef.current.innerHTML = diffHtml;
    }
  }, [diffHtml]);

  return (
    <div className="tracking-diff-container">
      <div className="upload-section">
        <div className="upload-file">
          <h3>Tài liệu gốc</h3>
          <input
            type="file"
            id="originalFile"
            accept=".docx"
            onChange={handleOriginalFileUpload}
          />
          <label htmlFor="originalFile" className="upload-button">
            Tải lên tài liệu gốc
          </label>
          {originalFile && <p className="file-name">{originalFile.name}</p>}
        </div>

        <div className="upload-file">
          <h3>Tài liệu chỉnh sửa</h3>
          <input
            type="file"
            id="editedFile"
            accept=".docx"
            onChange={handleEditedFileUpload}
          />
          <label htmlFor="editedFile" className="upload-button">
            Tải lên tài liệu chỉnh sửa
          </label>
          {editedFile && <p className="file-name">{editedFile.name}</p>}
        </div>

        {originalFile && editedFile && (
          <div className="compare-section">
            <button
              className="compare-button"
              onClick={handleCompare}
              disabled={isComparing}
            >
              {isComparing ? "Đang so sánh..." : "So sánh tài liệu"}
            </button>
          </div>
        )}
      </div>

      <div className="preview-section">
        <div className="document-preview">
          <h3>Tài liệu gốc</h3>
          <div ref={originalPreviewRef} className="document-container">
            {!originalFile && (
              <p className="placeholder">Vui lòng tải lên tài liệu gốc</p>
            )}
          </div>
          {diffHtml && (
            <div className="inline-diff-section">
              <h4>Hiển thị thay đổi ngay bên dưới</h4>
              <div ref={inlineDiffRef} className="inline-diff-container"></div>
            </div>
          )}
        </div>

        <div className="document-preview">
          <h3>Tài liệu chỉnh sửa</h3>
          <div ref={editedPreviewRef} className="document-container">
            {!editedFile && (
              <p className="placeholder">Vui lòng tải lên tài liệu chỉnh sửa</p>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default TrackingDiffDocx;
