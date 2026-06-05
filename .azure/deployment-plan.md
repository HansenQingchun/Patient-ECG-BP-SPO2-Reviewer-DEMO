# Deployment Plan — 飞利浦监护仪报警管理器

## Status: Deployed ✅

## Overview
| Field | Value |
|-------|-------|
| **Application** | 飞利浦监护仪报警管理器 (Philips Patient Monitor Alarm Manager) |
| **Type** | Static Web Application (HTML + CSS + JavaScript) |
| **Mode** | MODIFY (existing project, adding Azure deployment) |
| **Azure Service** | Azure Static Web Apps (Free tier) |
| **Target Account** | hansen_gzb@hotmail.com |
| **Subscription** | Visual Studio Enterprise 订阅 (2bda3292-d13f-498d-a6df-521bec74bf2a) |
| **Region** | East Asia (eastasia) |
| **Resource Group** | rg-monitor-alarm |
| **Public URL** | https://brave-sea-09a76db00.7.azurestaticapps.net |

## Architecture

```
[Browser] ──HTTPS──▶ [Azure Static Web Apps]
                          │
                          ├── index.html
                          ├── styles.css
                          ├── app.js
                          └── (user uploads .xlsx client-side)
```

- **No backend required** — all data processing happens in the browser
- **No database** — users upload .xlsx files locally, no data leaves the browser
- **CDN included** — Azure Static Web Apps provides global CDN automatically
- **HTTPS included** — Free SSL/TLS certificate

## Why Azure Static Web Apps?

1. **Free tier** — suitable for static sites, no cost
2. **Global CDN** — fast access worldwide
3. **Auto HTTPS** — secure by default
4. **Simple deployment** — `az staticwebapp` CLI or `swa` CLI
5. **No build step needed** — plain HTML/CSS/JS, no bundler

## Deployment Method

Use **Azure CLI (`az`)** directly:
1. Create resource group
2. Create Static Web App resource
3. Deploy files via SWA CLI or deployment token

## Steps

- [x] Step 1: Create resource group `rg-monitor-alarm` in `eastasia`
- [x] Step 2: Create Azure Static Web App resource
- [x] Step 3: Deploy static files via StaticSitesClient
- [x] Step 4: Verify public URL is accessible

## Files to Deploy

| File | Purpose |
|------|---------|
| `index.html` | Main page |
| `styles.css` | Styling |
| `app.js` | Application logic |

> Note: The `.xlsx` data file, `node_modules/`, and `package.json` are NOT deployed — they are development/local artifacts only.

## Security Considerations

- No backend, no API keys, no secrets
- All data processing is client-side (XLSX parsing in browser)
- HTTPS enforced by Azure Static Web Apps
- No user authentication required (public dashboard tool)
