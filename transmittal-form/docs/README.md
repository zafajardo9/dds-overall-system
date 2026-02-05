# Smart Transmittal (Transmittal Form) — Documentation

This `docs/` folder describes how the Smart Transmittal system is structured, how data flows through it, and how to run/export.

## Documents

- `SYSTEM_OVERVIEW.md`
- `ARCHITECTURE.md`
- `RUNBOOK.md`
- `planned-things-to-add.md`

## Quick context

- **Framework**: Vite + React + TypeScript
- **Core capabilities**:
  - Import documents / file lists (upload, Drive picker, pasted text, Drive folder link)
  - Extract structured metadata using Gemini (`@google/genai`) with a structured JSON schema
  - Edit/reorder items and fill transmittal fields
  - Export PDF (HTML → `html2pdf`) and DOCX (`docx`)
  - Keep local history snapshots in `localStorage`
