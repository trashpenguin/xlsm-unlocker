# XLSM Unlocker 🔓

A free, browser-based tool to remove password protection from `.xlsm` and `.xlsx` files — including **e-IPCRF** forms used by DepEd teachers in the Philippines.

🌐 **Live tool:** https://trashpenguin.github.io/xlsm-unlocker

---

## Features

- ✅ Removes workbook-level and sheet-level protection
- ✅ Works with SHA-512 encrypted locks (e-IPCRF, DepEd forms)
- ✅ 100% client-side — your file never leaves your device
- ✅ No installation, no sign-up, no ads
- ✅ Works offline after first load

## How It Works

Excel `.xlsm`/`.xlsx` files are ZIP archives containing XML files. Protection is stored as `<sheetProtection>` and `<workbookProtection>` XML tags. This tool uses [JSZip](https://stuk.github.io/jszip/) to unzip the file in your browser, strips those tags, and repacks the file for download.

## Usage

1. Go to https://trashpenguin.github.io/xlsm-unlocker
2. Drop your `.xlsm` or `.xlsx` file onto the page
3. Download the unlocked file

## Local Development

Just open `index.html` in a browser — no build step needed.

```bash
git clone https://github.com/trashpenguin/xlsm-unlocker
cd xlsm-unlocker
# Open index.html in your browser
```

## Important Note for e-IPCRF Users

After unlocking, **avoid clicking the "Finalize" or "Lock" buttons** inside the form. The VBA macros will re-apply protection. Edit the cells directly instead.

## License

MIT — free to use, share, and modify.
