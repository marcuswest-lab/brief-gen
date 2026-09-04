# Production Brief Generator

Standalone, browser-only app for filling out BAD Marketing creative briefs. Pick a client, pick a brief type (Static / Video / Body Copy), fill the form, click **Generate Brief** — get a properly formatted `.docx` with dropdowns set, no Cayce comments, and only the creative tables you actually used.

## Run locally

The app is pure static — no build step, no backend.

```bash
cd production-brief-generator
python3 -m http.server 8000
# open http://localhost:8000
```

(Or use any other static server. `npx serve`, VS Code Live Server, etc.)

> `file://` will not work because the app uses ES module imports and `fetch()`. You need an HTTP server.

## Deploy

### Cloudflare Pages (recommended)
1. Push this folder to a Git repo
2. Cloudflare Dashboard → Pages → Create project → Connect to Git
3. Build command: *(leave empty)*
4. Build output directory: `/` (or whatever subfolder this project lives in)
5. Deploy

Free plan, instant cold-start, custom domain support.

### GitHub Pages
1. Push to a GitHub repo
2. Repo Settings → Pages → Source: `main` branch, `/` folder
3. URL appears in a minute

## Project structure

```
production-brief-generator/
├── index.html                Page shell
├── app.js                    Form state + UI + generate orchestration
├── styles.css                Form styling
├── clients.json              Client list + per-client defaults
├── lib/
│   ├── templates-config.js   Row→field mappings + dropdown options
│   ├── docx-filler.js        Core .docx fill logic (text + SDT dropdowns)
│   └── docx-cleaner.js       Runtime safety: strips review comments
├── templates/
│   ├── static.docx           Cleaned static template
│   ├── video.docx            Cleaned video template
│   └── body-copy.docx        Body copy template
└── tools/
    └── strip_comments.py     One-time CLI: clean review comments out of a .docx
```

## Adding a client

Edit `clients.json`:

```json
{
  "clients": [
    {
      "id": "vam",
      "name": "Value Added Moving",
      "defaults": {
        "Copywriter": "Nate",
        "Conversion Objective": "Lead",
        "Ad Platform": "Meta"
      }
    },
    {
      "id": "tcc",
      "name": "TCC",
      "defaults": {
        "Copywriter": "Nate",
        "Landing Page URL": "https://...",
        "Conversion Objective": "Lead"
      }
    }
  ]
}
```

`defaults` keys must match the field names in the brief (e.g. `Landing Page URL`, `Copywriter`, `Avatar`, etc.). Defaults pre-fill empty form fields when the client is selected; user-entered values are never overwritten.

For the PM tool, `id`, `name`, and `defaults` are enough to make a client available in the client picker and Ad Categorizer. Add these optional fields when they are available:

- `tracker_url` — enables the one-click **Open Tracker** button.
- `creative_folder_url` — pre-fills the tracker Folder Link.
- `tracker_overrides` — handles a client's nonstandard tracker columns.

If the client's tracker filenames use a shortened or alternate name, add that alias to `detectClientName()` in `pm/lib/tracker-pipeline.js` so Meeting Notes and Weekly Creative Updates label it consistently. A Creative Dashboard URL and Strategist KPI thresholds are separate, optional integrations.

In-browser additions (the **+ Add client** button) are saved to localStorage on the user's machine only — for permanent additions across users, edit `clients.json`.

## Updating templates

If BAD Marketing publishes a new version of a brief template:

1. Drop the new `.docx` into a temp location.
2. Run `python3 tools/strip_comments.py path/to/new.docx templates/<static|video|body-copy>.docx` to strip any review comments and write directly into the templates folder.
3. If row positions or field names changed, update [`lib/templates-config.js`](lib/templates-config.js) accordingly. Use this snippet to inspect a template:

   ```python
   from docx import Document
   doc = Document('templates/static.docx')
   for ti, t in enumerate(doc.tables):
       print(f'Table {ti}: {len(t.rows)} rows')
       for ri, row in enumerate(t.rows):
           print(f'  R{ri}: {[c.text.strip()[:40] for c in row.cells]}')
   ```

4. Reload the app — no rebuild needed.

## What gets generated

- Header: `[CLIENT]` → selected client name, `[CAMPAIGN/LAUNCH NAME]` → Idea Name, `[DATE OF CREATION]` → today's date
- Overview table fully populated
- Only the creative tables you submitted (1–5) are kept; unused ones are removed
- All review comments stripped
- Dropdowns set to selected values (with alias normalization for common variants)

Filename pattern: `{Client}_{BriefType}_{Idea Name}_{YYYY-MM-DD}.docx`
