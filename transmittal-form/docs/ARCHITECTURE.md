# Architecture & Data Flow

## High-level architecture

This project is a **single-page React app**.

- Entry: `index.html` → `index.tsx` → `App.tsx`
- `App.tsx` is the orchestrator (state + handlers)
- `TransmittalTemplate` is primarily a presentational component
- Services and hooks provide:
  - AI parsing
  - Drive access
  - File processing
  - DOCX generation

## Entry points

- `index.html`
  - Includes CDN scripts:
    - Tailwind (`cdn.tailwindcss.com`)
    - `html2canvas` and `html2pdf.js`
    - Google scripts (GIS + `gapi`)
  - Mounts React at `<div id="root"></div>`
  - Loads `/index.tsx`

- `index.tsx`
  - Creates React root and renders `<App />`

- `App.tsx`
  - Defines the core state (`AppData`)
  - Owns import / parse / export actions
  - Persists history snapshots

## Main state

`AppContent` owns:

- `data: AppData`
- Import state:
  - `smartInput`, `isAnalyzingText`, parsing progress via `useFileProcessing`
- Export state:
  - `isGeneratingPdf`, `isGeneratingDocx`, docx preview modal state
- Drive state:
  - `googleClientId`, `isDriveReady`, `isDriveAuthorized`, `driveAuthError`
- UI state:
  - `activeTab`, `columnWidths`
- Persistence:
  - `history` (loaded/saved to `localStorage`)

## Data flow (imports)

### A) Local upload (PDF/images)

1. User selects files
2. `handleBatchUpload` calls `processDocs(files, onItemComplete)`
3. `useFileProcessing`:
   - Validates MIME type
   - Reads base64 (`readAsDataURL`) and strips prefix
   - If image: resizes with canvas for speed/limits
   - Runs concurrency-limited processing
4. For each file:
   - `parseTransmittalDocument(base64, mimeType, false, fileName)`
   - Maps result → `TransmittalItem[]`
5. `addItems` appends items and reindexes `noOfItems`
6. `mergeHeaderData` merges extracted header fields into `recipient`/`project`

### B) Google Drive picker (file contents)

1. User connects Drive (OAuth token)
2. User opens picker and selects files
3. For each selected file:
   - `getFileContentAsBase64(file.id)` downloads the file bytes
   - `parseTransmittalDocument(...)`
   - Items/header are merged into state

### C) Smart input (text or Drive folder link)

1. User pastes text in textarea
2. `handleSmartAnalysis`:
   - If Drive folder link: `extractFolderIdFromLink` → `listFilesInFolder`
     - Adds placeholder items based on file names
   - Else: `parseTransmittalDocument(text, 'text/plain', true)`
     - On failures in text mode: local regex fallback in `geminiService.ts`

## Data flow (exports)

### A) PDF

- `handlePrint`:
  - Uses `window.html2pdf()` on the element `#print-container`
  - Adds page numbering after render

PDF is a capture of the DOM layout produced by `TransmittalTemplate`.

### B) DOCX

- `handleDownloadDocx`:
  - Calls `generateTransmittalDocx(data)` → `Blob`
  - Downloads via `URL.createObjectURL`

DOCX is generated programmatically in `services/docxGenerator.ts`.

### C) DOCX Preview

- `handlePreviewDocx`:
  - Generates the same DOCX blob
  - Converts to HTML via `mammoth.convertToHtml`
  - Displays inside a modal

## Key files and responsibilities

- `App.tsx`
  - Orchestration, state, handlers, persistence
- `components/NewReportTemplate.tsx`
  - Printable transmittal layout + item table editing/reordering
- `services/geminiService.ts`
  - Gemini structured extraction + fallback parsing
- `services/googleDriveService.ts`
  - Google API init + OAuth + picker + download/list
- `services/docxGenerator.ts`
  - DOCX generation
- `hooks/useFileProcessing.ts`
  - Batch processing + resizing + progress tracking
- `types.ts`
  - Type definitions

## Environment and configuration

- `vite.config.ts` injects:
  - `process.env.API_KEY` from `GEMINI_API_KEY`

This is why the browser code can call `new GoogleGenAI({ apiKey: process.env.API_KEY })`.
