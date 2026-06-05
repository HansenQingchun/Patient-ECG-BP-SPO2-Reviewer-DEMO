# Central Station Alarm Data Analysis (DEMO)

> 中央站报警数据分析系统 · 患者监护报警审计 DEMO

A patient-monitoring **alarm data analysis** demo. Upload an alarm audit export (Excel) from a central monitoring station and the tool instantly turns it into interactive dashboards — alarm category distribution, per-bed alarm counts, time-series trends, and dedicated drill-downs for arrhythmia and pressure alarms. Results can be exported to a PDF report.

> ⚕️ Demonstration prototype. Do not commit real patient/clinical data — load your own Excel export at runtime. (The original sample dataset is intentionally excluded from this repository.)

## ✨ Features

- **One-click data import** — parse `.xlsx` alarm audit exports in the browser (SheetJS / xlsx).
- **Interactive dashboards** (ECharts) organized into tabs:
  - **Overview** — alarm category distribution, per-bed counts, Top 15 alarm types, severity distribution, daily / hourly / weekly trends
  - **Arrhythmia** — arrhythmia type distribution, per-bed counts, severity, daily trend
  - **Pressure** — pressure-alarm type distribution, per-bed counts, severity, daily trend
- **PDF report export** — render the analysis to a shareable PDF (`generate-pdf.js`, Puppeteer).
- **Desktop edition** — an optional Electron build (`desktop/`) packaged as a Windows installer.

## 🧩 Tech Stack

Vanilla HTML / CSS / JavaScript · [SheetJS (xlsx)](https://sheetjs.com) · [ECharts](https://echarts.apache.org) · Puppeteer (PDF) · Electron (desktop edition)

## 🚀 Quick Start (Web)

Just open `index.html` in a browser, then upload an alarm audit `.xlsx` file.

To regenerate the PDF report:

```bash
npm install
node generate-pdf.js
```

## 🖥️ Desktop Edition (Electron)

```bash
cd desktop
npm install
npm start          # run the app
npm run build      # build a Windows installer (output: desktop/release/)
```

## 🗂️ Project Structure

```
.
├── index.html          # Web app UI
├── app.js              # Data parsing, analysis & charts
├── styles.css          # Styles
├── generate-pdf.js     # PDF report generator (Puppeteer)
├── PRD.html            # Product Requirements Document (HTML)
├── PRD_*.pdf           # Product Requirements Document (PDF)
└── desktop/            # Electron desktop edition
    ├── main.js
    ├── src/            # app.js / index.html / styles.css
    └── libs/           # bundled echarts & xlsx
```

## 📄 Documentation

See `PRD.html` / `PRD_中央站报警数据分析系统_产品需求文档.pdf` for the full Product Requirements Document.
