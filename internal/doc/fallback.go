package doc

import "docstudio/internal/docx"

// docxFallback handles a .docx file that was saved with a .doc extension.
func docxFallback(b []byte) (string, string, error) { return docx.ToHTML(b) }
