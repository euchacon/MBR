# La-Z-Boy LATAM · MBR Dashboard

Web app that fills the official LZB Dealer MBR template and exports a ready-to-use PPTX directly from the browser.

## Deploy to Render

1. Push this folder to a GitHub repo (public or private)
2. Go to render.com → New → Web Service
3. Connect your GitHub repo
4. Render auto-detects the settings from `render.yaml`
5. Deploy — takes ~2 minutes
6. Your app URL will be: `https://lzb-mbr-dashboard.onrender.com`

## Local run (for testing)

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5050
```

## File structure

```
mbr-app/
├── app.py              # Flask backend — fills template and returns PPTX
├── template.pptx       # Official LZB Dealer MBR Template (do not rename)
├── templates/
│   └── index.html      # Dashboard frontend
├── requirements.txt
├── Procfile
└── render.yaml
```

## How it works

1. Fill in each section of the dashboard
2. Data auto-saves to browser localStorage (persists between sessions on same browser)
3. Click "Download PPTX" — the server fills the official template and returns the file
4. Open the downloaded PPTX — design is 100% preserved, only text is updated
