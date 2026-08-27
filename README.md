# William McLaughlan CV Site

Static GitHub Pages CV site for William McLaughlan.

## Files

- `index.html` — CV page
- `styles.css` — layout and visual styling
- `assets/William_McLaughlan_CV.pdf` — downloadable styled PDF CV
- `assets/William_McLaughlan_CV.docx` — editable DOCX CV
- `assets/headshot.png` — CV headshot
- `assets/cv-content.json` — source content used for the CV draft
- `scripts/build_downloads.py` — regenerates the styled PDF/DOCX/Markdown downloads from `assets/cv-content.json`

## Rebuild downloads

```bash
uv run --with reportlab --with python-docx scripts/build_downloads.py
```

## Local preview

```bash
python3 -m http.server 8080
```

Then open `http://127.0.0.1:8080/`.

## GitHub Pages

This repo is ready to publish from the `main` branch root using GitHub Pages.

