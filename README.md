# Vocabulario Mágico V3

V3 adds OCR for scanned PDFs and photos without an AI model or API cost.

## Uploads
PDF, DOCX, TXT, JPG, JPEG, PNG, WebP, HEIC, HEIF, TIFF. Multiple photographed pages are supported.

## Deploy
1. Upload everything in this folder to a new GitHub repository.
2. Keep `packages.txt`; it installs Tesseract and the Spanish language pack on Streamlit Community Cloud.
3. Deploy `streamlit_app.py` with Python 3.12.

## OCR modes
- Auto: embedded PDF text first; OCR fallback for image-only PDFs.
- Force OCR: ignores a poor PDF text layer and OCRs every page.
- Off: never runs OCR.

## Good scans
Photograph pages square-on, with even light, minimal shadow and readable type. Review the editable preview, especially accents and names.

## Limits
Handwriting and complex multi-column pages may be inaccurate. Large documents consume more memory. Only process material you are authorised to use.
