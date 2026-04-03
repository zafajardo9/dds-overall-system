# DDS Overall System

A comprehensive collection of automation systems and tools for DDS (Document & Data Solutions) operations, legal services, and marketing. This repository contains Google Apps Script automations, web applications, and utilities designed to streamline document management, client communications, compliance tracking, and SEO auditing.

---

## Folder Overview

### CAR/ — Certificate Authorizing Registration Monitoring
**Goal:** Automates compliance tracking and deadline management for CAR (Certificate Authorizing Registration) projects, primarily focused on Transfer of Shares and related corporate filings.

**Key Features:**
- Deadline calculation engine (DST Due Date: 5th of next month, CGT/DOD Due Date: 30 days after notarization)
- Remaining days tracker with color-coded status indicators
- Automated reminder emails to staff and clients
- Activity logging with timestamp and user tracking
- Status auto-update based on completion keywords
- Daily automation triggers at 8 AM

**Main Files:**
- `code.gs` — Core automation engine with 4-step setup wizard
- `CAR_Special_Projects_Monitoring_Summary.md` — Detailed documentation
- `ReminderEmailTester.html` — Email template testing UI

---

### amlc-automation/ — AMLC Expiry Automation
**Goal:** Manages expiry notifications and compliance tracking for AMLC (Anti-Money Laundering Council) requirements.

**Main Files:**
- `ExpiryAutomation.gs` — Automation script for AMLC expiry management

---

### automated-expiry-notification/ — Universal Expiry Notification System
**Goal:** A robust, auditable reminder system that sends expiry reminder emails from Google Sheets, tracks client reply acknowledgements, and provides guided setup for non-technical users.

**Key Features:**
- Sends reminder emails based on `Expiry Date` + `Notice Date` per row
- Collects reply acknowledgements from clients using configurable keywords
- Scheduled and manual runs across multiple configured sheet tabs
- Complete logging to `LOGS` tab and back to the sheet
- Guided setup wizard requiring no scripting knowledge
- Twice-daily reply scans to keep acknowledgement status current
- Configurable daily send time per team preference

**Main Files:**
- `code.gs` — Core notification automation
- `full-code.gs` — Complete implementation
- `unified-library.gs` — Shared utility library
- `unified-wrapper-template.gs` — Template for new implementations

---

### automatic-listing-document/ — Checklist Sync Automation
**Goal:** Links Google Sheets with Google Drive folders to automatically sync and highlight checklist items, links, remarks, and copies based on file presence in linked drive folders.

**Key Features:**
- Drive folder linking per sheet tab
- Automatic checklist column detection and setup
- Color-coded highlighting (Checklist, Links, Remarks, Copies)
- 1-minute interval trigger support

**Main Files:**
- `code.gs` — Checklist synchronization engine

---

### file-center/ — DDS & FilePino Document Processing
**Goal:** AI-powered document operations automation for DDS and FilePino brands. Scans department folders, extracts document titles using Gemini AI, and renames/reorganizes files into date-based folder structures.

**Key Features:**
- Multi-department Google Drive folder scanning
- AI title extraction (Gemini → OCR → Filename fallback)
- Automatic renaming with date-based organization (YYYY/MMMM/DD MMMM)
- Duplicate detection and prevention
- Daily/weekly summary emails with batching
- Alert emails for no-files and error conditions
- Comprehensive audit trails

**Main Files:**
- `code-dds.gs` — DDS brand automation engine
- `code-filepino.gs` — FilePino brand automation engine
- `DDS_FILEPINO_SYSTEM_MEMORY.md` — System architecture documentation
- `APPS_SCRIPT_WEB_APP_IMPLEMENTATION_GUIDE.md` — Web app deployment guide
- `GEMINI_MODEL_MENU_GUIDE.md` — AI model configuration guide

---

### gsheet-jira/ — Google Sheets Jira-Style Board
**Goal:** Transforms a Google Spreadsheet into a draggable 3-lane Kanban task board (Todo → In Progress → Done) with real-time status updates.

**Key Features:**
- Draggable cards across three lanes
- Automatic status updates based on lane placement
- LockService protection against duplicate writes
- Web app deployment for team access
- Authorization checks for spreadsheet owners and editors

**Spreadsheet Configuration:**
- Required tabs: `List of Task`, `Completed Tasks/For Monitoring`, `Cancelled`
- Required headers: ID, COMPANY, DESCRIPTION, STATUS, REMARKS, NOTES, Links, Progress

**Main Files:**
- `Code.gs` — Server-side logic and sheet operations
- `Index.html` — Main HTML shell
- `Styles.html` — Board and modal styling
- `Scripts.html` — Drag/drop and modal client logic

---

### litigation/ — Client Document Monitoring System
**Goal:** AI-powered legal document management for Duran Schulze law firm. Monitors two Google Drive folders for client uploads and scanned pleadings, with automatic AI renaming and user tracking.

**Monitored Folders:**
1. **"01 Client Documents"** — Files uploaded by clients
2. **"5.02 Scanned Pleadings"** — Scanned legal pleadings filed by or against the firm

**Key Features:**
- Detects new, modified, moved, or deleted files in both drives
- Tracks file uploader/modifier by email address
- AI-renames pleading files using Gemini into `YYYY-MM-DD Descriptive Title` format
- Daily scheduled scans at 6 PM Manila time (Mon–Fri)
- Comprehensive logging with 1000-row auto-trim
- System Config sheet for hidden key-value storage

**Main Files:**
- `Code-Client-Document-Management.gs` — Main automation script
- `DOCUMENTATION.md` — Complete technical documentation

---

### marketing-seo-website/ — SEO Automation & Analysis
**Goal:** Comprehensive SEO automation system for analyzing websites, generating recommendations, and tracking search performance metrics. Includes both Apps Script and Python-based scanning tools.

**Key Features:**
- Website SEO scanning and analysis
- Search Console and GA4 integration
- AI-powered SEO prioritization scoring
- Title tag and meta description recommendations
- Quick Wins Dashboard for actionable insights
- Indexability problem detection and plain-English explanations
- CTR improvement suggestions based on real search data

**Main Files:**
- `code.gs` — Google Apps Script SEO automation
- `KeywordViewer.html` — Keyword analysis UI
- `FEATURE_RECOMMENDATIONS.md` — Future feature roadmap
- `python-seo-scanner/` — Python-based SEO scanning tools

---

### notary/ — Notary Document Management
**Goal:** Document processing and management system for notary operations, supporting multiple notary staff members.

**Main Files:**
- `Code.gs` — Base automation script
- `Code-March.gs`, `Code-Meryshel.gs`, `Code-Steph.gs` — Staff-specific implementations

---

### site-audit/ — Website Audit Framework
**Goal:** Presentation-ready website audit system that translates technical audit data into clear, business-facing reports using a monochrome, corporate design language.

**Key Features:**
- Monochrome presentation UI for stakeholder reports
- Hero section with summary metric cards
- Executive Summary in plain language
- Priority findings categorized by business impact
- Risk assessment with strengths and weaknesses
- Action plan by timeline (Immediate, Near-Term, Ongoing)
- Priority assets list with expandable details
- Responsive design (desktop → tablet → mobile)
- Print-friendly PDF output

**Audit Sites:**
- `duran/` — Duran Schulze audit data
- `filedocsphil/` — FileDocsPhil audit data
- `filepino/` — FilePino audit data

**Main Files:**
- `telehr.html` — Main audit presentation template
- `presentation.UI.md` — UI design system documentation
- `squirrel.toml` — Configuration for squirrel audit tool

---

### spc/ — Automated Liquidation Processing System
**Goal:** AI-powered liquidation voucher processing system that extracts data from submitted documents, organizes them by month, and maintains comprehensive processing logs.

**Key Features:**
- Gemini AI-powered voucher data extraction
- Multi-voucher support (single file containing multiple vouchers)
- Monthly subfolder organization (MM Month format)
- Rate limiting for API quota management
- Batch processing (3 files per batch)
- Comprehensive error handling and logging
- Dynamic configuration via Script Properties

**Architecture:**
- `Code.gs` — Main processing engine
- `DataValidator.gs` — Input validation
- `DriveManager.gs` — Google Drive operations
- `GeminiParser.gs` — AI extraction logic
- `Logger.gs` — System logging
- `RateLimiterManager.gs` — API rate limiting
- `SheetManager_v2.gs` — Spreadsheet operations

---

### transmittal-form/ — Modern Web Application
**Goal:** A full-stack web application (React/Next.js with Prisma) for document transmittal and management. Built with AI Studio integration and modern UI components.

**Technology Stack:**
- **Frontend:** React, Next.js, TypeScript, Tailwind CSS
- **Backend:** Next.js API routes, Prisma ORM
- **UI Components:** shadcn/ui
- **Database:** PostgreSQL (via Prisma)

**Key Features:**
- Document transmittal form interface
- Server-side data processing
- Type-safe API with TypeScript
- Responsive design with Tailwind CSS

**Main Directories:**
- `app/` — Next.js app router pages
- `components/` — UI components (41 components)
- `server/` — Server-side utilities
- `services/` — API service layers
- `prisma/` — Database schema and migrations
- `hooks/` — Custom React hooks
- `lib/` — Utility libraries

---

## Common Technologies Used

| Technology | Purpose |
|------------|---------|
| Google Apps Script | Serverless automation for Google Workspace |
| Gemini AI | Document title extraction and content analysis |
| Google Drive API | File storage and organization |
| Google Sheets API | Data storage and user interface |
| Gmail API | Email notifications and reply tracking |
| React/Next.js | Modern web application frontend |
| Prisma | Database ORM and schema management |
| Tailwind CSS | Utility-first styling |

---

## Getting Started

Most folders contain their own README.md with specific setup instructions. Generally:

1. **Google Apps Script Projects:**
   - Open the bound Google Spreadsheet
   - Go to **Extensions → Apps Script**
   - Copy the script files into the editor
   - Run the setup/initialization function from the custom menu

2. **Web Applications (transmittal-form):**
   - Install dependencies: `npm install`
   - Set environment variables in `.env.local`
   - Run locally: `npm run dev`

---

## Documentation Standards

Each major system includes:
- `README.md` — Setup and usage instructions
- `DOCUMENTATION.md` — Technical documentation (where applicable)
- `plans.md` — Roadmap and planning documents
- Inline code comments explaining complex logic

---

## Contact & Support

For questions about specific systems, refer to the individual folder README files or contact the system administrator.
