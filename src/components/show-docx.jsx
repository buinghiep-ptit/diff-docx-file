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

  // Helper to clean HTML before comparison
  const cleanHtml = (html) => {
    // Remove extra spaces and normalize
    return html
      .replace(/\s+/g, " ")
      .replace(/<\/p>\s*<p/g, "</p><p")
      .trim();
  };

  // Helper function to find the best matching paragraph
  const findBestMatch = (paragraph, paragraphArray) => {
    // Simple implementation: return the first paragraph that shares at least 50% of the words
    const words = paragraph.split(/\s+/);

    for (const p of paragraphArray) {
      const targetWords = p.split(/\s+/);
      const sharedWords = words.filter((word) => targetWords.includes(word));

      if (sharedWords.length / words.length > 0.5) {
        return p;
      }
    }

    return null;
  };

  // Generate structured diff that preserves HTML structure
  const generateStructuredDiff = useCallback((originalHtml, editedHtml) => {
    // Use a wrapper to help preserve structure
    const cleanOriginal = cleanHtml(originalHtml);
    const cleanEdited = cleanHtml(editedHtml);

    // Create diff at paragraph level first
    const originalParagraphs = cleanOriginal.split(/<\/p>\s*<p[^>]*>/g);
    const editedParagraphs = cleanEdited.split(/<\/p>\s*<p[^>]*>/g);

    // Prepare HTML to display differences
    let diffOutput = "";

    // Find differences using diff-match-patch approach
    const diffParagraphs = Diff.diffArrays(
      originalParagraphs,
      editedParagraphs
    );

    diffParagraphs.forEach((part) => {
      if (part.added) {
        // Added paragraphs - highlight in green
        part.value.forEach((p) => {
          diffOutput += `<p class="added">${p}</p>`;
        });
      } else if (part.removed) {
        // Removed paragraphs - highlight in red with strikethrough
        part.value.forEach((p) => {
          diffOutput += `<p class="removed">${p}</p>`;
        });
      } else {
        // Same paragraphs - check for word-level differences
        part.value.forEach((p) => {
          const correspondingEdited = editedParagraphs.find((ep) => ep === p);

          if (correspondingEdited) {
            // Exact match, no need for word diff
            diffOutput += `<p>${p}</p>`;
          } else {
            // Find the best match in edited paragraphs
            const bestMatch = findBestMatch(p, editedParagraphs);
            if (bestMatch) {
              // Do word-level diff
              const wordDiff = Diff.diffWords(p, bestMatch);

              let wordDiffOutput = "";
              wordDiff.forEach((wordPart) => {
                if (wordPart.added) {
                  wordDiffOutput += `<span class="added">${wordPart.value}</span>`;
                } else if (wordPart.removed) {
                  wordDiffOutput += `<span class="removed">${wordPart.value}</span>`;
                } else {
                  wordDiffOutput += wordPart.value;
                }
              });

              diffOutput += `<p>${wordDiffOutput}</p>`;
            } else {
              // No good match, keep original
              diffOutput += `<p>${p}</p>`;
            }
          }
        });
      }
    });

    // Set the diff HTML content
    setDiffHtml(diffOutput);
  }, []);

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
        .convertToHtml({ arrayBuffer })
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
