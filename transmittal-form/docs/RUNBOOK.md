# Runbook

## Prerequisites

- Node.js installed
- A Gemini API key
- (Optional) A Google OAuth Client ID if you want Google Drive integration

## Install

```bash
npm install
```

## Configure environment

Create `.env.local` at the project root and set:

- `GEMINI_API_KEY=<your_gemini_api_key>`
- `BETTER_AUTH_SECRET=<32+ char secret>`
- `BETTER_AUTH_URL=http://localhost:8000`
- `VITE_BETTER_AUTH_URL=http://localhost:8000`
- `GOOGLE_CLIENT_ID=<google client id>`
- `GOOGLE_CLIENT_SECRET=<google client secret>`

Notes:

- `vite.config.ts` maps `GEMINI_API_KEY` into the browser bundle as `process.env.API_KEY`.
- Treat this key as **exposed to the client** when deployed.

## Run locally

```bash
npm run dev
```

Vite runs on:

- `http://localhost:3000`

## Build

```bash
npm run build
```

## Preview production build

```bash
npm run preview
```

## Enabling Google Drive integration

### 1) Provide OAuth Client ID

- In the app UI, open settings (gear icon) and paste the **Google OAuth Client ID**.
- Click **Update Cloud Bridge**.
- Refresh the page to re-initialize.

The Client ID is stored in `localStorage` under `google_client_id`.

### 2) Scopes used

The app requests readonly Drive scopes:

- `https://www.googleapis.com/auth/drive.readonly`
- `https://www.googleapis.com/auth/drive.metadata.readonly`

### 3) What Drive features do

- **Browse Drive**: picker selection → download file contents → Gemini parse
- **Paste a Drive folder link**: lists filenames only (does not download content)

## Export behavior

- **PDF export** uses `html2pdf.js` + `html2canvas` loaded from CDN in `index.html`.
- **DOCX export** uses `docx` and downloads a `Blob`.
- **DOCX preview** converts the generated DOCX to HTML using `mammoth`.

## History snapshots

- Snapshot is added automatically before generating PDF or DOCX.
- Stored in `localStorage` under `transmittal_history`.
- UI keeps the latest 20 snapshots.

## Troubleshooting

### Gemini errors

- **401 / Invalid API Key**: verify `GEMINI_API_KEY` in `.env.local`.
- **429 / Quota**: retry later or use a different key.
- If parsing fails for text input, the app falls back to a local regex/line-based parser.

### Google Drive not available

- Ensure you provided a valid Client ID.
- Ensure the Google scripts load (network access).
- If scripts do not load, `initializeGoogleApi` will warn and Drive integration will remain unavailable.

### Google OAuth redirect URI (Better Auth)

For local development, add this to your Google OAuth app:

```
http://localhost:8000/api/auth/callback/google
```

### Upload parsing seems slow

- Image resizing is performed for larger images to speed up AI parsing.
- Batch parsing runs concurrently (limit 5 files at a time).
