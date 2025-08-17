# Spreadsheet Duplicator

A small front-end prototype to duplicate Excel workbooks from a master template using values from a replacement sheet.

Features
- Select a master template (.xlsx) and a replacement sheet (table of values)
- Map replacement-sheet columns to one or more target cells (sheet + A1 address) in the template
- Two replace modes: conditional (skip empty values) and force (always write)
- Filename control: prefix + column or token pattern (e.g. `Invoice_{Name}`)
- Browser: downloads generated workbooks as files
- Desktop (Electron): can write outputs directly into the same folder as the template

Note: The browser version cannot write files back to the original folder due to browser sandboxing — use the Electron desktop build for native saving into the template folder.

Quickstart — run in a local static server (recommended)

1. Clone or copy this repository locally.

2. Serve the folder with a simple static server (recommended so assets and scripts load consistently):

- Using Python 3 (built-in):

  python3 -m http.server 8000

- Using Node (http-server):

  npm install -g http-server
  http-server -p 8000

3. Open your browser and navigate to:

  http://localhost:8000

4. Use the UI:
- Click "Choose file" to select your master template (`.xlsx`).
- Select the replacement sheet from the dropdown and click "Load Mapping".
- Map headers to target sheet/cell addresses (the mapping sheet will be excluded from the target list).
- Configure naming (pattern OR prefix + header) and click "Preview" to verify output filenames.
- Click "Generate" to create and download the generated workbooks.

Running the Electron desktop app (optional)

The project includes an Electron scaffold so you can run a desktop version that can save outputs into the same folder as the template.

1. Install dependencies:

  npm install

2. Start Electron (if there's a script in package.json, use it; otherwise run):

  npx electron .

3. In the desktop app, choose the template and mapping as above. When generating, files will be written into the same directory as the template file.

Note: Electron may require development dependencies (electron) and a proper package.json script. Adjust as needed.

Push to GitHub and enable GitHub Pages

1. Create a new repository on GitHub (via web UI or `gh` CLI).

2. From your project folder, run (replace `<url>` with your repository URL):

  git init
  git add .
  git commit -m "Initial commit"
  git branch -M main
  git remote add origin <url>
  git push -u origin main

3. Enable GitHub Pages for the repository:
- Go to the repository Settings → Pages
- Under "Source" choose the branch `main` and folder `/ (root)` and save.
- Wait a minute, then visit the published URL shown by GitHub.

Important details about GitHub Pages
- The app will run in the browser on GitHub Pages, but the browser build cannot perform native filesystem writes. Generated files will be offered as downloads to the user (just like using the app locally in a browser).
- If you need the app to save outputs directly back into the template folder on disk, use the Electron build on a desktop machine instead.

Security & Privacy
- Files are selected in-browser via a file input and are not uploaded anywhere by the UI. The app runs entirely client-side when used via GitHub Pages or a static server.

Suggestions / Next steps
- Add overwrite protection / uniquification when saving to the template folder (Electron) so existing files are not overwritten silently.
- Add a progress indicator for batch generation.
- Optionally add a small server-side component if you want server-side processing or centralized saving.

If you want, I can also add a small `deploy` script or GitHub Actions workflow to automatically publish to GitHub Pages.
