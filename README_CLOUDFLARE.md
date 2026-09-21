# Deploying St. Ann's College Report Generator to Cloudflare Pages

This web application is **100% client-side** (built with pure HTML5, Tailwind CSS, and ExcelJS).
- Zero server maintenance, zero compute fees (100% free forever on Cloudflare Pages).
- Instant report generation (<0.5 seconds).
- 100% private: Student batches and faculty contacts never leave the user's browser.

---

## Method 1: 1-Click Drag-and-Drop (Easiest - No Git Required)

1. Log into your [Cloudflare Dashboard](https://dash.cloudflare.com/).
2. On the left sidebar, click **Workers & Pages**.
3. Click **Create application** -> select the **Pages** tab.
4. Click **Upload assets**.
5. Give your project a name (e.g. `st-anns-reports`).
6. Select or drag-and-drop the files from this folder:
   - `index.html`
   - `course_report_engine.js`
   - `app.js`
   - `faculty_data.js`
   - `exceljs.min.js`
   - `lucide.min.js`
   - `wrangler.toml`
   - `_headers`
7. Click **Deploy site**.
8. In 5 seconds, your app will be live at:
   `https://st-anns-reports.pages.dev`

---

## Method 2: Connect via GitHub (Automatic Deployments on Commit)

1. Push this folder to a GitHub repository.
2. In Cloudflare Dashboard -> **Workers & Pages** -> **Create application** -> **Pages**.
3. Click **Connect to Git** and select your repository.
4. Set the build settings:
   - **Framework preset**: None
   - **Build command**: (Leave blank)
   - **Build output directory**: `.`
5. Click **Save and Deploy**.

---

## Running Locally on Windows Anytime

Just double-click **`Run_Web_App.bat`** in this folder. It opens your browser immediately at `http://localhost:8000`.
