<!-- CONSOLIDATED & EXPANDED GUIDE -->
# OnlyOffice Plugin – Complete Step‑by‑Step Guide (Beginner Friendly)

This guide assumes zero prior knowledge. Follow in order; stop when everything works.

## 0. What You Get
Edit Office docs (DOCX / XLSX / PPTX / PDF) directly inside Obsidian – locally – using an OnlyOffice Document Server Docker container.

## 1. Prerequisites
| Need | Why |
| ---- | --- |
| Obsidian | Host application for your vault |
| Docker Desktop (Win/macOS) or Docker Engine (Linux) | Runs the OnlyOffice server |
| Plugin runtime files | Provide integration |

### Install Docker Desktop (Windows/macOS)
1. Download: https://www.docker.com/products/docker-desktop/
2. Run installer (keep “Use WSL2”). Reboot if prompted.
3. Launch Docker Desktop → wait for Running.
4. Verify:
```bash
docker --version
```

### If Docker Desktop Won’t Start (Windows quick fix)
```powershell
wsl --update
wsl --shutdown
```
Restart Docker Desktop. For deep reset: unregister `docker-desktop*` distributions (see older guide / SERVER_SETUP.md).

### Linux (brief)
Install Docker Engine per distro docs; add user to `docker` group; re‑login; run `docker run --rm hello-world`.

## 2. Place Plugin in Vault
1. Obsidian: Settings → Community Plugins → disable Safe Mode.
2. Install the OnlyOffice plugin.
3. Or install manually: “Open plugins folder” → `<vault>/.obsidian/plugins/`.
4. Create folder `OnlyOffice`.
5. Copy runtime files:
    - `manifest.json`, `main.js`, `embedded-editor.html`
    - `assets/Start.docx` (+ optional `Start.xlsx`, `Start.pptx`)
    - Scripts: `start-onlyoffice.*`, `stop-*`, `update-*`, `remove-*`, `check-setup.*`
    (Source `.ts`, build configs not needed on end‑user machines.)

## 3. Choose JWT Secret
All auth uses a shared secret. Default: `your-secret-key-please-change` (OK for testing). Recommended: long random (≥32 chars). Generate (PowerShell):
```powershell
[guid]::NewGuid().ToString()+[guid]::NewGuid().ToString()
```
Edit every script that sets `JWT_SECRET=`.

## 4. Start the OnlyOffice Server
Windows: double‑click `start-onlyoffice.bat`.
What it does (first run): pulls image, creates container `obsidian-onlyoffice`, maps port 8080, sets headers + JWT.
macOS/Linux: run `start-onlyoffice.sh` (chmod +x) or replicate its `docker run` line.

Verify container:
```bash
docker ps --filter name=obsidian-onlyoffice
```
Optional: open http://127.0.0.1:8080 in a browser.

## 5. Activate Plugin
1. Settings → Community Plugins → enable “OnlyOffice”.
2. Open plugin settings:
    - Server URL: `http://127.0.0.1:8080`
    - JWT Secret: match scripts
3. Open a `.docx` file to test.

## 6. Everyday Usage
### Open
Click any `.docx`, `.xlsx`, `.pptx`, `.pdf`. 
If you have other plugins that reads for example .pdf you might need to right click on the file and choose open with OnlyOffice.

### Create
Ribbon icon or Command Palette → “New OnlyOffice Document” → choose type & name. Templates: place `Start.<ext>` in `assets/`.

### Save
Autosave handled by OnlyOffice. Use “Save As” (if exposed) for duplicates (extension auto‑applied).

### Multiple Docs
Open many simultaneously; they share the one container.

## 7. Manage Server
| Task | Script (Win) | Script (Unix) | Manual |
| ---- | ------------ | ------------- | ------ |
| Start/create | start-onlyoffice | start-onlyoffice | docker run/start |
| Stop | stop-onlyoffice | stop-onlyoffice | docker stop obsidian-onlyoffice |
| Update image | update-onlyoffice | update-onlyoffice | docker pull onlyoffice/documentserver |
| Remove | remove-onlyoffice | remove-onlyoffice | docker rm -f obsidian-onlyoffice |
| Check | check-setup | check-setup | docker ps -a |

Change port: edit `HOST_PORT` + update plugin setting.  
Change secret: edit scripts, remove container, start again, update plugin setting.

## 8. Troubleshooting
| Symptom | Cause | Fix |
| ------- | ----- | --- |
| Token / security error | Secret mismatch | Align JWT secret & recreate container |
| Read‑only toolbar | CORS / iframe headers missing | Ensure env vars in start script |
| Blank pane | Container stopped / wrong URL | Check `docker ps`, port, URL |
| XLSX/PPTX creation fails | Template missing | Add `assets/Start.xlsx` / `Start.pptx` |
| Scrollbar in modal | Old build | Update `main.js` |
| Edits not appearing | Delay / callback issue | Wait; check console & `docker logs` |

Diagnostics:
```powershell
docker logs obsidian-onlyoffice --tail 60
```
Obsidian dev console: Ctrl+Shift+I (Win/Linux) / Cmd+Opt+I (macOS).

Reset sequence: close panes → remove-onlyoffice → start-onlyoffice → reopen doc.

## 9. Security Notes
| Item | Recommendation |
| ---- | -------------- |
| JWT secret | Long random; don’t share |
| Exposure | Keep container bound to localhost |
| Backups | Still back up your vault regularly |

## 10. FAQ
**Need internet?** Only for first Docker desktop installation image pull.  
**Remote server?** Yes; must allow `app://obsidian.md` in frame/CORS + matching secret.  
**Disable JWT?** Possible but discouraged.   
**Where are files?** In your vault; container caches temporary editing state.

## 11. Cheat Sheet
Start: start-onlyoffice  |  Stop: stop-onlyoffice  |  Update: update-onlyoffice  |  Remove: remove-onlyoffice  |  Logs: docker logs obsidian-onlyoffice  |  New Doc: Ribbon / Command Palette | Templates: assets/Start.<ext>

You’re done. See `SERVER_SETUP.md` for deeper Nginx / reverse proxy options.
