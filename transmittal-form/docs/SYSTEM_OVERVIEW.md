# System Overview — Smart Transmittal

## What this system is

Smart Transmittal is a client-side web application that helps you generate a formal **Transmittal Form** by importing documents (or lists of documents), extracting metadata, letting you review/edit the results, and exporting the final transmittal as **PDF** and **DOCX**.

It is built as a **single-page app** (no routing) and runs entirely in the browser.

## Primary user workflows

### 1) Import documents / file lists

Supported sources:

- Local uploads:
  - PDF files (`application/pdf`)
  - Images (`image/*`)
- Google Drive:
  - Drive picker selection (download file contents and parse)
  - Drive folder link (list file names and create placeholder items)
- Smart paste:
  - Paste file names / text lines (AI parsing, with local regex fallback)

### 2) Review & edit

- Edit transmittal header fields:
  - Recipient
  - Project
  - Sender branding (including logo)
  - Signatories / time released
  - Transmission method checkboxes
  - Notes / instructions
- Edit the item table:
  - Change values (qty, doc number, description, remarks)
  - Reorder items via drag-and-drop
  - Remove items
  - Bulk import rows from spreadsheet-style text

### 3) Export / deliver

- Export **PDF** (HTML → canvas → PDF)
- Export **DOCX** (generated programmatically with `docx`)
- Preview DOCX (generated DOCX → HTML via `mammoth`)
- Export **CSV** of item rows
- Open a pre-filled **mailto:** link (does not attach files)

## Core modules (by responsibility)

- **UI orchestration**: `App.tsx`
  - Owns the full `AppData` state
  - Handles imports, parsing, merging, export, history snapshots
- **Transmittal layout/template**: `components/NewReportTemplate.tsx`
  - Receives `AppData` and callbacks
  - Renders the printable transmittal layout
- **AI parsing**: `services/geminiService.ts`
  - Calls Gemini (`@google/genai`) with a structured JSON schema
  - Provides fallback parsing for text lists
- **Google Drive integration**: `services/googleDriveService.ts`
  - Initializes Google API + picker
  - OAuth token flow
  - List folder contents
  - Download file contents
- **DOCX generation**: `services/docxGenerator.ts`
  - Generates Word document using `docx` → `Blob`
- **Batch file processing**: `hooks/useFileProcessing.ts`
  - Reads files as base64
  - Resizes images
  - Concurrency-limited processing + progress reporting

## Data model

The single source of truth is `AppData` (`types.ts`).

Key objects:

- `AppData`
  - `recipient`, `project`, `sender`, `signatories`, `receivedBy`, `footerNotes`, `notes`
  - `items: TransmittalItem[]`

`TransmittalItem` represents one line in the item table.

## Persistence

- History snapshots are stored in `localStorage` under `transmittal_history` (limited to the latest 20).
- Google OAuth Client ID is stored in `localStorage` under `google_client_id`.

## External services

- **Gemini** (Google Generative AI)
  - Key is supplied as `GEMINI_API_KEY` via Vite env injection.
- **Google Drive**
  - Uses Google Identity Services and Drive API (readonly scopes).

## Known constraints / tradeoffs

- This app runs in the browser; **Gemini API key is effectively client-exposed** in production builds. This is acceptable for internal tools/prototyping but not recommended for public deployment.
- Google Drive integration requires a valid OAuth Client ID and user consent.
