# MobileViT2 Anomaly Detector – Slide Generator

This repository generates a single PowerPoint slide matching the specified layout and branding (light theme, pastel blue/green, no title).

## Quick start (local)

```bash
pip install -r requirements.txt
python generate_pptx.py
# Output: MobileViT2-Anomaly-Detector-Key-Insights.pptx
```

## CI build and download

On every push to `main` (or via “Run workflow”):
1. Go to the **Actions** tab.
2. Open the latest **Build PPTX** run.
3. Download the artifact named `MobileViT2-Anomaly-Detector-Key-Insights` — it contains `MobileViT2-Anomaly-Detector-Key-Insights.pptx`. 
