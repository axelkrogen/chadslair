# ChadLair — GitHub Pages + XLSX

This build removes the Google Apps Script runtime dependency and runs as a static GitHub Pages site.

## Files

- `index.html` — the ChadLair interface adapted for static hosting.
- `xlsx-backend.js` — browser-side XLSX persistence and the replacement for `google.script.run`.
- `data/chadlair-main.xlsx` — Activity, Tasks, TaskArchive, and StatsRollup.
- `data/chadlair-pump.xlsx` — Affirmations, Taunts, and Names.
- `media/manifest.json` — optional static PUMP media list.

## Publish

1. Put every file and folder in the root of a GitHub repository.
2. In GitHub, open **Settings → Pages → Deploy from a branch**.
3. Select the branch and `/ (root)`.
4. Open the Pages URL after deployment.

## XLSX behavior

The default workbooks under `data/` are fetched when the site first opens.

- **Connect XLSX** selects existing ChadLair workbooks.
- **Save XLSX** saves to browser storage and writes to selected files when the browser grants a writable file handle.
- **Export** downloads the current main and PUMP workbooks.

GitHub Pages cannot modify repository files by itself. Direct write-back uses the browser File System Access API when available. Other browsers retain changes in IndexedDB and use Export for durable copies.

## Use existing workbooks

Replace the template files in `data/` with your existing `.xlsx` files, retaining the exact sheet names and headers, or connect them from the page.

## Static PUMP media

Put media files under `media/`, then list them in `media/manifest.json`:

```json
{
  "media": [
    {
      "name": "r1f.gif",
      "path": "media/r1f.gif",
      "mimeType": "image/gif",
      "role": "round1"
    }
  ]
}
```

Recognized system names include `r1f.gif`, `r2f.gif`, `r3f.gif`, `final.gif`, `ko.png`, and `1.gif` through `9.gif`.
