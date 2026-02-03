# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains a Google Apps Script (GAS) project for Pousada Porto Fino / Areia do Forte: **Auditoria Noturna** (Nightly Audit). It runs on the V8 runtime and is deployed to Google Sheets via CLASP.

The nightly audit system imports data from Omnibees (hotel booking platform), Niara, and Bee2Pay (payment processor), reconciles reservations, and generates daily audit sheets.

## Deployment Commands

```bash
# Push to Google Apps Script
cd /c/Users/Usuario/auditoria-noturna
clasp push

# Pull latest from remote
clasp pull

# Open project in Apps Script editor
clasp open
```

## Architecture

### Language and Runtime
- Pure JavaScript (Google Apps Script), V8 runtime
- No package.json, no npm dependencies, no build step, no bundler
- HTML files serve as UI templates (sidebars) using `HtmlService`
- No test framework or linter configured

### Naming Conventions
- **Private/helper functions**: suffixed with underscore (`ensureSheet_()`, `toast_()`, `extractOmniReport_()`)
- **Public functions**: no underscore, callable from UI menus or HTML templates (`setupAll()`, `validateEnvironment()`)
- **Constants**: ALL_CAPS for config objects (`SHEETS`, `LISTS`, `UX`, `THEME`, `LAYOUT`, `UPLOAD`)
- **Sheet names**: defined as constants, never hardcoded in logic

### Data Flow
```
Excel/CSV upload (chunked for large files)
  -> Temporary Google Sheet via Drive API conversion
  -> Extract & normalize (extractOmniReport_, extractNiaraReport_, extractBee2PayReport_)
  -> Create AUDIT_<date> sheet with block-based layout
  -> Merge data from multiple sources (Niara, Bee2Pay)
  -> Log to AUDIT_LOG
```

Each audit sheet uses a fixed block layout: `LAYOUT.blockHeight = 5` rows per reservation (Header + Omni + PMS + Checks + Spacer), starting at `LAYOUT.startRow = 7`, across `LAYOUT.cols = 13` columns (A..M).

This project requires **Drive API Advanced Service** enabled in Apps Script (used for file conversion and management). The `validateEnvironment()` function checks this.

### Resumable Upload Protocol
Large files use a chunked upload via `CacheService`: `resumableStart_()` initializes a session, `resumableChunk_()` appends base64 chunks, then the assembled blob is processed. Session state is stored in script cache with `UPLOAD.cacheTtlSeconds = 900`.

## Key Configuration

- Timezone: `America/Sao_Paulo`
- Currency format: `R$ #,##0.00`
- Soft limit: 250 audit blocks per sheet (`UX.maxBlocksSoftLimit`)
- Raw data limit: 8000 rows (`UX.rawMaxRows`)
- Google Apps Script execution timeout: 6 minutes (platform limit)

## Important Considerations

- All code runs server-side in Google Apps Script. There is no Node.js environment.
- HTML files use `google.script.run` to call server-side functions from the client.
- The `onOpen()` trigger creates custom menus in the Google Sheets UI.
- The `onEdit(e)` trigger strips emoji prefixes from cell values in column L.
- Sheet operations should be batched (use `getValues()`/`setValues()` over loops) to stay within the 6-minute execution limit.
- The `.clasp.json` file contains the `scriptId` that links to the Google Apps Script project. Do not change it.
