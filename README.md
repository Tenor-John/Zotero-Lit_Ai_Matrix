# Zotero Lit AI Matrix

<div align="center">

[English](README.md) | [简体中文](doc/README-zhCN.md)

</div>

An in-Zotero intelligent literature matrix plugin for research workflows.

It combines Zotero metadata (title, author, journal, year, tags) with structured AI notes into one searchable, sortable, and exportable matrix.

## Highlights

- Smart Matrix View
  - Built directly inside Zotero.
  - Shows metadata columns plus AI-structured fields.
- AI Note Parsing
  - Parses structured fields from notes and updates matrix cache automatically.
  - Works well with Zotero GPT style note generation.
- GitHub-style Usage Heatmap
  - Visualizes daily usage frequency.
  - Click one day to filter table by that day; click again to reset.
- One-click Reading Jump
  - Click title to open PDF reader directly.
- Productivity Utilities
  - Full-library cache rebuild, batch read-status update, CSV/Markdown export.

## Installation

### Option 1: Release build (Recommended)

1. Go to GitHub Releases and download the latest `.xpi`.
2. In Zotero: `Tools` -> `Plugins` -> gear icon -> `Install Plugin From File...`.
3. Select the `.xpi` and restart Zotero.

Repository: <https://github.com/Tenor-John/Zotero-Lit_Ai_Matrix>

### Option 2: Development mode

```bash
npm install
npm start
```

Configure `.env` correctly:

- `ZOTERO_PLUGIN_ZOTERO_BIN_PATH`
- `ZOTERO_PLUGIN_PROFILE_PATH`

## Quick Start

### 1. Open the Matrix page

Use any entry:

- Toolbar AI Matrix icon
- File menu matrix entry

### 2. Prepare structured AI notes

Recommended note keys:

- `领域基础知识::`
- `研究背景::`
- `作者的问题意识::`
- `研究意义::`
- `研究结论::`
- `未来研究方向提及::`
- `未来研究方向思考::`

Use one key per line with `::` separator.

### 3. Rebuild matrix cache

- Selected items: item right-click menu -> rebuild matrix cache
- Whole library: file menu -> rebuild full-library matrix cache

### 4. Filter and analyze

- Filter by keyword, status, year, journal, tags
- Sort by headers
- Click title to open PDF (or locate item if no PDF)

### 5. Heatmap day filtering

- Click a day cell with data: table filters to that day
- Click the same cell again: reset to full results

## Recommended Workflow with Zotero GPT

1. Generate structured notes for selected papers.
2. Keep field names consistent with matrix keys.
3. Rebuild cache.
4. Compare, filter, and export in matrix view.

## FAQ

### Why are column headers missing?

- Make sure you installed the latest release.
- Check matrix top-right `UI:xxxx` version.
- If version is old, reinstall latest `.xpi` and restart Zotero.

### Why does `npm start` report "Zotero binary not found"?

Check `.env` paths:

- `ZOTERO_PLUGIN_ZOTERO_BIN_PATH` must point to Zotero executable
- `ZOTERO_PLUGIN_PROFILE_PATH` must point to Zotero profile directory

## Privacy

- The plugin mainly uses local Zotero item/note data.
- It does not proactively upload your library data.
- Any external AI traffic depends on your GPT plugin setup.

## Build and Release

```bash
npm run build
npm run release
```

Recommended release process:

1. Develop and test in feature branch
2. Merge via PR into `main`
3. Release with semantic version tags (e.g. `v0.1.5`)

## License

AGPL-3.0-or-later

## Author

Tenor-John
