# XLSM Unlocker

A free, browser-based tool to remove password protection from `.xlsm` and `.xlsx` files — including **e-IPCRF** forms used by DepEd teachers in the Philippines.

🌐 **Live tool:** https://trashpenguin.github.io/xlsm-unlocker

---

## Features

- ✅ Removes workbook-level and sheet-level protection
- ✅ Works with SHA-512 encrypted locks (e-IPCRF, DepEd forms)
- ✅ **Batch processing** — drop multiple files at once, download individually or as a ZIP
- ✅ **Sheet selection** — choose exactly which sheets to unlock per file
- ✅ **Sheet names in log** — shows actual sheet names (e.g. "Ratings" unlocked), not just counts
- ✅ **Progress bar** — real-time repacking progress for large files
- ✅ **Paste support** — Ctrl+V a file directly from clipboard
- ✅ 100% client-side — your file is never uploaded to any server
- ✅ **True offline support** via Service Worker (works after first load)
- ✅ No installation, no sign-up, no ads

## How It Works

Excel `.xlsm`/`.xlsx` files are ZIP archives containing XML files. Protection is stored as `<sheetProtection>` and `<workbookProtection>` XML tags. This tool uses [JSZip](https://stuk.github.io/jszip/) (bundled locally — no CDN) to unzip the file in your browser, strips those tags, and repacks the file for download.

## Usage

1. Go to https://trashpenguin.github.io/xlsm-unlocker
2. Drop your `.xlsm` or `.xlsx` file(s) onto the page — or paste with Ctrl+V
3. Choose which sheets to unlock (single-file mode)
4. Download the unlocked file(s)

## Local Development

Just open `index.html` in a browser — no build step needed.

```bash
git clone https://github.com/trashpenguin/xlsm-unlocker
cd xlsm-unlocker
# Open index.html in your browser
```

## Important Note for e-IPCRF Users

After unlocking, **avoid clicking the "Finalize" or "Lock" buttons** inside the form. The VBA macros will re-apply protection. Edit the cells directly instead.

## Tech Stack

- Vanilla HTML, CSS, JavaScript — no frameworks, no build step
- [JSZip 3.10.1](https://stuk.github.io/jszip/) — bundled locally
- [Playfair Display](https://fonts.google.com/specimen/Playfair+Display) — Google Fonts (serif italic accent)
- Service Worker for offline caching

## License

MIT — free to use, share, and modify.

