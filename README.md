# Seller Signal — Farm Prioritization Tool

A browser-based tool that ranks homes in a farm export (from Title Toolbox or
a title rep) by how likely the owner is to sell soon, based on tax
delinquency, vacancy, and ownership-stage patterns (upsize/downsize signals).

All file processing happens client-side in the browser — uploaded farm data
is never sent to a server.

## Running it locally

1. Install [Node.js](https://nodejs.org/) (v18 or later) if you don't have it.
2. In this folder, install dependencies:
   ```bash
   npm install
   ```
3. Start the local dev server:
   ```bash
   npm run dev
   ```
4. Open the URL it prints (usually `http://localhost:5173`).

## Deploying on Vercel

1. Push this folder to a new GitHub repository.
2. Go to [vercel.com](https://vercel.com), sign in, and click **Add New → Project**.
3. Select this repository. Vercel auto-detects the Vite/React setup — no
   configuration changes are needed.
4. Click **Deploy**. You'll get a live URL (e.g. `your-tool.vercel.app`)
   within a minute or two.
5. (Optional) Add a custom domain under the project's **Settings → Domains**.

Every time you push a change to the repo's main branch, Vercel automatically
rebuilds and redeploys the live site.

## Project structure

```
seller-signal-tool/
├── index.html          # HTML entry point
├── package.json        # Dependencies and scripts
├── vite.config.js       # Build configuration
├── src/
│   ├── main.jsx         # Mounts the App component
│   └── App.jsx          # The tool itself
└── README.md
```

## Editing the tool

The scoring logic, column-matching rules, and UI all live in `src/App.jsx`.
Scoring weights are defined near the top of the file in `computeRow()`.
