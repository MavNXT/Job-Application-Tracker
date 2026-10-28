# Job Application Tracker

A single-user job application tracker. The browser app talks to a small local Node server, which saves every change to `job-applications.xlsx` in this folder. There is no database.

## Quick start

1. Install Node.js (check with `node -v`). No `npm install` is needed, the project has no npm dependencies.
2. Double-click `start.bat`. It opens `http://localhost:3000` and runs `server.js`.
3. Keep the black window open while you use the app. Close it (or press `Ctrl+C`) to stop.

Do not open `index.html` by double-clicking it. Without the server it cannot load or save and shows "Failed to fetch".

## Features

- Dashboard: counts, a bar chart by status, and a follow-up list (overdue or due in the next 3 days).
- Table: search (company, role, job description, notes), status filter, sortable columns, Edit and Delete per row, and "Remove all rejected".
- Board: Kanban columns by status. Drag a card to another column to change its status.
- Add / edit form: company, role, status, applied date, follow-up date, job link, referral toggle (with "Referred by" name), job description, notes. The applied date is empty for Wishlist and defaults to today for other statuses.
- Referral jobs: the table has a "Referred by" column and an "All / Referral only / Non-referral" filter, and search covers the referrer name.
- Portals tab: one entry per company (case-insensitive, so a company can only appear once) with email ID, career portal link, login type, username, password and notes. Turn on "OTP login (no password)" for portals that use one-time codes: the password is cleared, the table shows an OTP badge, and the Excel LOGIN TYPE column says OTP (otherwise Password). Click a password to show or hide it. Stored in the second sheet, "Portals".
- Import: adds rows from an `.xlsx` or `.csv` file to the existing data. Rows whose id already exists are skipped.
- Download .xlsx: downloads a copy of the current spreadsheet.
- Saving is automatic and silent. A warning bar appears only when a save fails.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | The whole front end: HTML, CSS and JavaScript in one file |
| `server.js` | Local HTTP server and the Excel reader/writer |
| `start.bat` | Starts the server and opens the browser || `xlsx.full.min.js` | SheetJS 0.18.5 (community). Reads Excel on the server and in the browser. Do not use it to write styled files |
| `job-applications.xlsx` | The data file (created automatically if missing) |

Keep these files together. The server looks for the spreadsheet next to `server.js`.

## How it works

```
Browser (index.html)  --GET /api/applications-->  server.js  --reads-->  job-applications.xlsx
                      --POST /api/applications->             --writes->
                      --GET /api/download------>             (file download)
```

- On page load the app calls `load()`, which fetches all applications and keeps them in the `apps` array.
- Any change (add, edit, delete, drag, import) updates `apps` and calls `save()`, which POSTs the whole array. The server rewrites the entire spreadsheet.
- `save()` refuses to run until `load()` has succeeded. This stops a failed load from overwriting the file with a near-empty list.
- The server writes to `job-applications.xlsx.tmp` and then renames it over the real file, so a crash mid-write does not corrupt the data.
- On startup the server reads and rewrites the file once, which refreshes the headings and styling.

### API (`server.js`)

| Method and path | What it does |
| --- | --- |
| `GET /` and `GET /index.html` | Serves the app |
| `GET /xlsx.full.min.js` | Serves the SheetJS library |
| `GET /api/applications` | Returns all applications as a JSON array |
| `POST /api/applications` | Body is the full JSON array. Replaces the spreadsheet contents |
| `GET /api/portals` | Returns all career portal entries as a JSON array |
| `POST /api/portals` | Body is the full JSON array. Replaces the "Portals" sheet and keeps the Applications sheet |
| `GET /api/download` | Downloads `job-applications.xlsx` |
New static files must be added to the `STATIC` map in `server.js` or they will return 404.

### Data model

One application is an object with string values. The same keys are used in the browser, in the JSON API and (via `HEADERS`) in the spreadsheet:

`id`, `company`, `role`, `status`, `applied` (YYYY-MM-DD, empty for Wishlist), `followUp` (YYYY-MM-DD), `link`, `referral` (`Yes` or empty), `referredBy` (name, only when `referral` is `Yes`), `description`, `notes`

`status` is one of `Wishlist`, `Applied`, `Interview`, `Offer`, `Rejected`.

A portal entry has the keys `id`, `company`, `emailId`, `portalLink`, `loginType` (`OTP` or `Password`, empty counts as Password), `username`, `password`, `notes`. Its fields, headings and widths are in the `PORTAL` object in `server.js`, and the applications ones are in `APP`.

Security: passwords are stored as plain text in `job-applications.xlsx`. Anyone who can open the file can read them. Keep the folder private and do not share or upload the file.

### Excel format

- Sheet 1 is "Applications" and sheet 2 is "Portals".
- `server.js` writes the `.xlsx` itself (a small zip and XML writer using Node's `zlib`) because the free SheetJS cannot write bold text. Headings are bold capitals from `HEADERS`.
- Reading ignores case and punctuation in headings (`FOLLOW-UP`, `followUp` and `follow up` all map to `followUp`). Old files with lowercase headings still load.
- Cells are stored as text (inline strings), including dates.

## Where to change things (`index.html`)

The script is split into commented sections:

| Section | Contents |
| --- | --- |
| Top of script | `STATUSES`, `FIELDS`, global state (`apps`, `loaded`, `saveError`, sort state) |
| `// ---------- Excel ----------` | `fromXLSX()` for import |
| `// ---------- persistence ----------` | `save()`, `load()`, download and import handlers, `renderSaveBar()` |
| `// ---------- helpers ----------` | Date helpers, follow-up status and label |
| `// ---------- render ----------` | `renderStats()`, `renderTable()`, `renderBoard()`; `render()` calls them all |
| `// ---------- events ----------` | Tabs, search, sorting, delete buttons, drag and drop |
| `// ---------- form ----------` | Add/edit dialog, `openForm()`, submit handler |

The pattern for any change is: update `apps`, then call `save()`. `save()` persists and re-renders.

Status colours are CSS variables: the classes `.s-Wishlist`, `.s-Applied` and so on set `--c`, which the badges, board columns, cards and chart bars use.

## How to add things

### Add a new field (for example "salary")

Touch every place the field list appears:

1. `index.html`: add the key to `FIELDS`.
2. `index.html`: add an input to the form (`name="salary"`) and add `'salary'` to the two field lists in `openForm()` and the submit handler.
3. `index.html`: add a column to `renderTable()` (the `cols` array and the row template) and update the `colspan` in the empty-state row.
4. `server.js`: add the key to `FIELDS`, a heading to `HEADERS` (capitals) and a width to `WIDTHS`. All three arrays must stay in the same order.
5. Restart `start.bat`. Existing rows get an empty value for the new field.

### Add a new status (for example "Screening")

1. `index.html`: add it to `STATUSES` (the order is the board column order).
2. `index.html`: add a colour rule next to the other `.s-...` lines, for example `.s-Screening { --c:#0891b2; }`.
3. The filter, form dropdown, board columns and chart are generated from `STATUSES`. The dashboard tiles in `renderStats()` are hand-listed, so add one there if you want a count.

### Add a new API endpoint

Add a branch in the request handler in `server.js` before the static-file lookup. Use `send(res, code, body, type)` for responses. Restart the server to pick up changes.

### Add a new tab or view

1. Add a `<button data-view="name">` in the `<nav>` and a `<section id="name" class="hidden">`.
2. Add `'name'` to the view list in the tab click handler (`['dashboard', 'table', 'board']`).
3. Add a `renderName()` function and call it from `render()`.

## Development notes

- Edit `index.html` and press `Ctrl+F5` in the browser. No restart is needed.
- Edit `server.js` and restart `start.bat`.
- Syntax check the server without running it: `node --check server.js`.
- There are no automated tests. After a change, test by hand: add, edit, delete and drag an application, import a file, refresh the page, and open the spreadsheet in Excel (with the app stopped, or close the file again before using the app).
- The server listens on `127.0.0.1` only, so it is not reachable from other machines. There is no login, by design.
- Change the port with `PORT` in `server.js` and update the URL in `start.bat`.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| "Failed to fetch" bar | `start.bat` is not running, or the page was opened as a file. Start it and use `http://localhost:3000` |
| Save fails with `EBUSY` or `EPERM` | The spreadsheet is open in Excel. Close it, then change something again |
| "address already in use" | The app is already running. Use the existing window, or close it and start again |
| "node is not recognized" | Install Node.js and run `start.bat` again |
| Page looks old after an edit | Hard refresh with `Ctrl+F5` |
| Headings still lowercase | Restart `start.bat` so the server rewrites the file |

## Ideas for later

- Bulk status change or multi-select delete.
- Undo for deletes.
- Interview dates and contact person fields.
- Export to CSV, or a backup copy of the spreadsheet on every start.
- Tests for `buildXlsx()` and `readApps()` round-trips.
