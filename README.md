# Drone Fleet Manager — PWA

## How to install on Android (Add to Home Screen)

1. Host these files anywhere (Netlify, GitHub Pages, etc.)
2. Open the URL in Chrome on Android
3. Tap the **⋮ menu → "Add to Home Screen"**
4. It installs like a native app — no Play Store needed!

---

## Google Sheets Setup

### Step 1 — Prepare your sheet

Your Google Sheet should have a tab (e.g. `Drones`) with these columns:

| Drone ID | Name | Group | Status | Reason | Notes |
|---|---|---|---|---|---|
| DA01 | Drone Alpha 01 | Alpha | Good | | |
| DA02 | Drone Alpha 02 | Alpha | Fail | Battery Issue | |

Column names are flexible — the app auto-detects columns containing "name", "status", "reason", "note", etc.

### Step 2 — Get a Google Sheets API Key

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project (or use an existing one)
3. Enable the **Google Sheets API**
4. Go to **Credentials → Create Credentials → API Key**
5. (Optional but recommended) Restrict the key to Sheets API only

### Step 3 — Share your sheet

The sheet must be accessible for reading:
- **Share → Anyone with the link → Viewer** (for reading)
- The app writes back using the API key (which requires the sheet to not be fully private)

> **Note:** For production use with private sheets, you'd use OAuth2 or a backend proxy. The API key method works for sheets shared as "Anyone with link can view."

### Step 4 — Configure the app

Open the app → tap the ⚙️ gear icon → enter:
- **Sheet ID** — found in the URL: `docs.google.com/spreadsheets/d/SHEET_ID/edit`
- **Tab name** — the name of your drones tab (e.g. `Drones`)
- **API Key** — from Step 2

---

## Free Hosting Options

### Netlify (easiest)
1. Go to [netlify.com](https://netlify.com) → Sign up free
2. Drag and drop this folder onto the dashboard
3. Get a URL like `https://drone-fleet-abc.netlify.app`

### GitHub Pages
1. Push this folder to a GitHub repo
2. Settings → Pages → Deploy from branch → main
3. Get a URL like `https://yourusername.github.io/drone-fleet`

---

## Features
- ✅ Fleet overview with Good / Fail / Un-slotted counts
- ✅ Tap any drone to update its status
- ✅ Fail reasons: Battery Issue, Motor Failure, Camera Not Working, Signal Loss, Physical Damage, Software Error, Other
- ✅ Optional notes per update
- ✅ Update history log (stored locally)
- ✅ Search drones by name or group
- ✅ Offline support (cached via service worker)
- ✅ Demo mode (no sheet required)
- ✅ Works on Android & iOS home screen
