# Cell Culture Tracker

A Google Apps Script web app for planning and tracking cell-culture runs on an interactive Gantt chart. Data is stored in a Google Sheet created automatically on first use.

## Features

- Schedule fixed-duration, open-ended, and milestone steps
- Reuse saved run templates
- See today’s and upcoming work at a glance
- Collaborate safely with a single-editor lock and read-only viewers
- Export tracker data as JSON

## Setup

1. Create a standalone project at [Google Apps Script](https://script.google.com/).
2. Add `Code.gs` and an HTML file named `Index`, then copy in the contents of this repository’s files.
3. Optionally configure `ALLOWED_EMAILS` near the top of `Code.gs`.
4. Deploy the project as a web app and choose access settings appropriate for your Google Workspace or Google account setup.
5. Open the deployment URL. The app creates its backing spreadsheet on first use, in the Google Drive of the account under which the script executes.

The deployment owner must retain access to the generated spreadsheet. When deploying as the user accessing the app, share that spreadsheet with each editor.

## License

Licensed under the [GNU General Public License v3.0](LICENSE).
