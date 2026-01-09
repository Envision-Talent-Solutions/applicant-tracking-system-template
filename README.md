 # Applicant Tracking System Template

  Source code for the Google Sheets Applicant Tracking System Template by Envision Talent Solutions.

  ## About

  This repository contains the Apps Script code that powers the Applicant Tracking System template. The code is provided for transparency so users can review exactly what it does before installing.

  ## Get the Template

  Install from the [Google Workspace Marketplace](#) (link coming soon)

  ## Files Overview

  | File | Purpose |
  |------|---------|
  | `appsscript.json` | Manifest with OAuth scopes and configuration |
  | `Main.js` | Core initialization |
  | `Triggers.js` | Menu creation and event handlers |
  | `CandidatesSync.js` | Sync between Candidate Database and Active Candidates |
  | `Requisitions.js` | Job ID generation and requisition management |
  | `FormProcessor.js` | Google Form submission handling |
  | `FormInit.js` | Form setup and configuration |
  | `Validation.js` | Dropdown menu management |
  | `Settings.js` | Settings tab processing |
  | `Import.js` | Bulk candidate import |
  | `ResumeLinker.js` | Resume import and contact extraction |
  | `Sidebar.js` | Sidebar UI management |
  | `ImportSidebar.html` | Sidebar HTML interface |
  | `LinkHygiene.js` | URL formatting and cleanup |
  | `StateManager.js` | Document properties wrapper |
  | `DebounceQueue.js` | Operation queuing system |
  | `DashboardData.js` | Dashboard metrics |
  | `Constants.js` | Shared constants |
  | `Authorization.js` | OAuth scope management |
  | `Diagnostics.js` | System diagnostics |
  | `Util.js` | Logging and utility functions |

  ## How It Works

  The template uses three main tabs that stay synchronized:

  - **Requisitions** – Job openings with status tracking and timeline metrics
  - **Candidate Database** – Complete candidate records and talent pool
  - **Active Candidates** – Filtered view of candidates for open positions

  ## Privacy & Security

  - All data stays in your Google account
  - Resume processing happens in your browser (never uploaded to external servers)
  - No third-party services or external APIs

  ## Documentation

  Full documentation is included with the template:
  - Quick Start Guide
  - User Operations Guide
  - Admin Reference Guide

  ## License

  This source code is provided for transparency and review purposes only. See [LICENSE](LICENSE) for details.

  All rights reserved. © 2026 Envision Talent Solutions
