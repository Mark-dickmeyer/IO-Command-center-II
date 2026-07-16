# Sales Command Center

This is your sales pipeline, forecast, accounts, and task tracker, packaged as a
standalone project so it can be hosted outside of Claude.

## ⚠️ Read this before you do anything else

This app was built inside Claude, which provides a special built-in storage
system (`window.storage`) that saves data on Anthropic's servers and shares
it live between everyone who opens the same link. That system **does not
exist outside Claude.**

To keep this project runnable right away, a temporary stand-in
(`src/storageShim.js`) is included. It saves data to your own browser's
local storage instead. That means:

- ✅ The app will load and work when you run it.
- ✅ You can click around and confirm everything looks right.
- ❌ Your data will **not** be shared between you and your team.
- ❌ Data only lives in one browser, on one device. Clearing your browser
  data, or opening the app on a different computer, starts you from empty.

**Before this is a real tool your team uses together, `storageShim.js` needs
to be replaced with real calls to a database** (Postgres, Supabase, etc.),
through a small API. That's a follow-up step, not something included here.

## What's in this folder

```
sales-command-center-app/
├── index.html          The page that loads the app
├── package.json         List of libraries the app needs
├── vite.config.js       Build tool configuration
├── src/
│   ├── main.jsx          Starts the app
│   ├── App.jsx            <- The actual command center (all your work lives here)
│   └── storageShim.js    Temporary local-only data storage (see warning above)
└── .gitignore            Tells Git which files to ignore
```

## Getting this onto GitHub (no command line required)

1. Go to [github.com](https://github.com) and create a free account if you don't have one.
2. Click the **+** icon (top right) → **New repository**.
3. Give it a name (e.g. `sales-command-center`), leave it **Private** if you don't want it public, and click **Create repository**.
4. On the next page, look for a link that says **"uploading an existing file"**.
5. Drag this entire folder's contents into that upload box (or click to browse and select all the files/folders — including the `src` folder). GitHub will preserve the folder structure automatically.
6. Scroll down, add a short message like "Initial upload," and click **Commit changes**.

That's it — your code is now on GitHub. No installing anything, no terminal.

## Testing it on your own computer (optional, requires more setup)

If you want to actually run the app and click through it before deploying:

1. Install [Node.js](https://nodejs.org) (choose the "LTS" version) — this is a one-time install, like installing any other program.
2. Open a terminal in this folder (on Mac: right-click the folder → "New Terminal at Folder"; on Windows: open the folder, type `cmd` in the address bar and press Enter).
3. Type `npm install` and press Enter. This downloads the libraries the app needs — takes a minute or two.
4. Type `npm run dev` and press Enter. It will print a local web address (usually `http://localhost:5173`).
5. Open that address in your browser. The app should load.

## Next steps to make this a real, shared tool

1. **Deploy it somewhere public** — Vercel is the natural choice, and it connects directly to a GitHub repo (import the repo you just created, it builds automatically).
2. **Replace the storage shim with a real database** — this is the part that brings back shared, persistent, multi-user data. Ask Claude (in a fresh conversation, pointing at this code) or a developer to wire `storageShim.js` up to a real backend such as Vercel Postgres or Supabase.
3. Once both of those are done, you'll have a permanent, team-shared version of this tool living outside Claude entirely.
