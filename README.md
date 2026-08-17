# Raj Meesho — Orders, P&L, Inventory & GST Dashboard

Internal seller-analytics tool for Unigen Lifestyle. Runs entirely in the browser (no backend server) — data is processed on your device and stored only in your browser (localStorage), never uploaded anywhere.

## Files in this repo

| File | What it is | Where it goes |
|---|---|---|
| `raj_meesho.html` | The dashboard itself | GitHub Pages (or just open locally) |
| `Code.gs` | Login/session + Google Sheet sync backend | Paste into a Google Sheet's Apps Script editor |
| `Upload.html` | In-Sheet file-import sidebar | Paste into the same Apps Script project (as an HTML file named `Upload`) |

## One-time setup

### 1. Google Sheet backend
1. Create (or open) a Google Sheet.
2. **Extensions → Apps Script**. Delete the default code, paste in `Code.gs`.
3. **Files → + → HTML**, name it exactly `Upload`, paste in `Upload.html`.
4. In `Code.gs`, change `ADMIN_PASSWORD` at the top to your own secret.
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the deployed Web App URL.
7. Open your Google Sheet, reload it — a new menu **"📊 Raj Meesho"** appears.
   Run **🛠️ One-Time Setup (Create All Tabs)** to create all required tabs.
8. Run `addFirstUserFromScript()` once from the Apps Script editor (edit the userId/password inside it first) to create your first login — or use the in-app Admin panel once deployed.

### 2. The dashboard (`raj_meesho.html`)
- Open the file directly (double-click) **or** host it via GitHub Pages (see below).
- On first load, log in → click **"🔐 Admin Login"** → enter your `ADMIN_PASSWORD` → add your real user.
- On the login screen's setup box, paste the Apps Script Web App URL from step 1.6 above and save — this is stored only in your browser.
- Go to **Upload Data** and either upload files directly, or connect the Google Sheet sync (tab names are pre-filled; leave the "Sheet Link" field blank to use the Sheet the script is bound to, or paste a link to a *different* seller's Sheet if you've been given Viewer access to it).

### 3. Hosting on GitHub Pages (optional)
1. Push this repo to GitHub.
2. **Settings → Pages → Deploy from a branch → main → / (root)**.
3. Your dashboard will be live at `https://<username>.github.io/<repo>/raj_meesho.html`.

> ⚠️ **Important:** `raj_meesho.html` contains no business data and no secrets by default — data lives only in each visitor's own browser (localStorage), and the Apps Script URL is entered per-browser via the in-app setup box, not hardcoded in the file. If your repo is **public**, anyone can view the page's source code (this is unavoidable for any client-side web page) — but they will not see your Orders/Payments/GST data unless they also have your login and your Apps Script URL. If you want extra privacy, make the GitHub repo **Private** (GitHub Pages from a private repo needs a paid GitHub plan) or just keep using the file locally instead of hosting it.

## Notes
- The dashboard's JavaScript is minified/obfuscated to discourage casual reading of the business logic. This does **not** provide real security — a technically determined person can still de-obfuscate it. Treat this as a light deterrent only.
- Login uses single-session enforcement (same User ID can't be active on two devices at once) and hashed passwords (SHA-256) — see `Code.gs`.
