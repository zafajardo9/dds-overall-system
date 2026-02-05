# Roadmap — Evolving Smart Transmittal into a heavier system

This roadmap is written to **retain the current Vite/React SPA** (your existing UI + workflows) and progressively add a backend, auth, database, and integrations.

## Goal

Keep the current app behavior:

- Import (upload, Drive picker, pasted text, Drive folder link) (DONE)
- Extract items via Gemini ()
- Edit/reorder items (DONE)
- Export PDF/DOCX (DONE)

While adding:

- Authentication + organizations/roles
- Persistent storage (Prisma)
- Google Drive + Google Sheets integration
- Organizational standards / filtering

## Target architecture (recommended)

- **Frontend**: keep this repo as the SPA UI (Vite + React).
- **Backend API**: add a small server app (Node) that handles:
  - BetterAuth sessions
  - Prisma DB access
  - Google OAuth token storage (if needed)
  - Gemini calls (recommended: move Gemini off the client)
- **Data**: Postgres (or MySQL) as primary DB.

This avoids a risky rewrite while giving you a real “system” foundation.

## Packages to add (backend + integrations)

> These are **new** packages for the backend/server layer. The current SPA components remain unchanged.

### Core backend + auth

- `better-auth` — authentication framework
- `express` (or `fastify`) — API server
- `cors` — CORS control for SPA → API
- `cookie-parser` — session cookie handling
- `zod` — request/response validation
- `dotenv` — env config

### Database (Prisma)

- `prisma` (dev)
- `@prisma/client`

### Google integrations

- `googleapis` — Drive + Sheets APIs

### Utilities (optional but useful)

- `pino` (or `winston`) — logging
- `multer` — multipart uploads if you decide to proxy file uploads to the backend
- `nanoid` — IDs (if needed outside DB)

## Phase 1 — Backend foundation (no UX changes)

- **Add a backend** (API server) and keep the SPA calling it via HTTP.
- **Move Gemini calls server-side**
  - Replace `parseTransmittalDocument()` usage in the SPA with `POST /api/parse`.
  - Keep the same response shape so the UI doesn’t change.
- **Add BetterAuth**
  - Basic sign-in/sign-out
  - Session-based auth for API routes

Deliverable: You can sign in and parse documents without exposing API keys in the browser.

## Phase 2 — Database with Prisma (persistence)

- **Introduce Prisma** and a DB schema.
- Persist:
  - Users
  - Organizations
  - Transmittals (header fields + status)
  - Transmittal items (rows)
  - Audit / history (replace local-only snapshots, or sync them)

Deliverable: “History” becomes a real dataset (queryable, shareable, filterable).

## Phase 3 — Organizational standards / multi-tenant filtering

- Define organization rules:
  - Organization membership and roles (admin/member/viewer)
  - Standard templates per organization (sender defaults, footer text, naming conventions)
  - Filtering/searching transmittals by organization/project/recipient/date
- Add server-side authorization checks:
  - Users can only read/write transmittals in their org

Deliverable: The app behaves like an internal system (multi-org / multi-user) rather than a single-user tool.

## Phase 4 — Google integrations (Drive + Sheets)

Split this into two capabilities:

### A) Google Drive (documents)

- Keep the current picker UX, but decide where OAuth lives:
  - Option 1: OAuth tokens stay client-side (simpler, but weaker control)
  - Option 2: OAuth handled by backend (recommended for enterprise; enables token storage + governance)
- Store Drive file metadata on transmittal items:
  - file ID, webViewLink, mime type, original name

### B) Google Sheets (structured item sources)

- Add “Import from Sheet”:
  - Select a spreadsheet/range
  - Map columns → `qty`, `documentNumber`, `description`, `remarks`
- Optional: “Export items to Sheet” for audit/tracking

Deliverable: A transmittal can be generated from standard operational trackers.

## Phase 5 — Hardening (what makes it “heavy rated”)

- Centralized logging + error reporting
- Background jobs (if parsing large batches)
- Rate limiting and request validation
- Role-based access control
- Activity feed / audit trail
- File storage strategy (if you decide to store attachments)

## Decisions to confirm (so I can implement cleanly)

- **Do you want to keep Vite SPA + add an API server** (recommended), or **migrate to Next.js**?
- **Database** preference:
  - Postgres (recommended) vs something else
- **Google OAuth**:
  - Client-only vs backend-managed OAuth
- **Multi-tenancy**:
  - One organization per user vs users in many organizations

## Original items (tracked)

- Connecting G-Drive and Sheet to the app
- Connecting Google OAuth with BetterAuth to the system
- Have Organizational standards in the system to filter out transmittals that happened for specific company/group
- With Database with prisma


