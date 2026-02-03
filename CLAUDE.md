# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains two Google Apps Script (GAS) projects for a pousada (small hotel) called Porto Fino / Areia do Forte. Both run on the V8 runtime and are deployed to Google Sheets via CLASP.

### Project 1: Pousada Estoque & Gastos (root directory)
Inventory and expense management system. Tracks products, purchase invoices, stock movements, weekly inventory counts, and expense reporting.

### Project 2: Auditoria Noturna (`meu-projeto-gas/`)
Nightly audit system that imports data from Omnibees (hotel booking platform), Niara, and Bee2Pay (payment processor), reconciles reservations, and generates daily audit sheets.

## Deployment Commands

```bash
# Push root project (Estoque & Gastos) to Google Apps Script
cd /c/Users/Usuario/projetos
clasp push

# Push audit project
cd /c/Users/Usuario/projetos/meu-projeto-gas
clasp push

# Pull latest from remote
clasp pull

# Open project in Apps Script editor
clasp open
```

The root `.claspignore` excludes `meu-projeto-gas/**`, so the two projects deploy independently with separate `scriptId` values.

## Architecture

### Language and Runtime
- Pure JavaScript (Google Apps Script), V8 runtime
- No package.json, no npm dependencies, no build step, no bundler
- HTML files serve as UI templates (modal dialogs, sidebars) using `HtmlService`
- No test framework or linter configured

### Naming Conventions
- **Private/helper functions**: suffixed with underscore (`ensureSheet_()`, `toast_()`, `extractOmniReport_()`)
- **Public functions**: no underscore, callable from UI menus or HTML templates (`saveCompra()`, `setupAll()`, `validateEnvironment()`)
- **Constants**: ALL_CAPS for config objects (`SHEETS`, `LISTS`, `UX`, `THEME`, `LAYOUT`, `UPLOAD`)
- **Sheet names**: defined as constants, never hardcoded in logic

### Estoque & Gastos Data Model
Sheets form a normalized data model: `Produtos` (master), `Notas` (invoices), `Itens_Nota` (line items linked via ARRAYFORMULA/VLOOKUP), `Movimentacoes` (immutable transaction log), `Inventarios`/`Inventario_Itens` (counts), `Estoque_Atual` and `Consumo_Semanal` (computed reports), `Painel_Gastos` (expense dashboard).

Key pattern: `setup*_()` functions create/recreate sheet structure (headers, validations, formatting, formulas). Reports are rebuilt by clearing and rewriting data ranges.

### Auditoria Noturna Data Flow
```
Excel/CSV upload (chunked for large files)
  -> Temporary Google Sheet via Drive API conversion
  -> Extract & normalize (extractOmniReport_, extractNiaraReport_, extractBee2PayReport_)
  -> Create AUDIT_<date> sheet with block-based layout
  -> Merge data from multiple sources (Niara, Bee2Pay)
  -> Log to AUDIT_LOG
```

Each audit sheet uses a fixed block layout: `LAYOUT.blockHeight = 5` rows per reservation (Header + Omni + PMS + Checks + Spacer), starting at `LAYOUT.startRow = 7`, across `LAYOUT.cols = 13` columns (A..M).

The audit project requires **Drive API Advanced Service** enabled in Apps Script (used for file conversion and management). The `validateEnvironment()` function checks this.

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
- The `onOpen()` trigger in each project creates custom menus in the Google Sheets UI.
- The `onEdit(e)` trigger in the audit project strips emoji prefixes from cell values in column L.
- Sheet operations should be batched (use `getValues()`/`setValues()` over loops) to stay within the 6-minute execution limit.
- The `.clasp.json` files contain `scriptId` values that link to specific Google Apps Script projects. Do not change these.
