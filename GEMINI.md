# Trendcarpet_Interiors: project rules for Antigravity (security, 2026-10-04)

This folder is behind a public tunnel (QA hub `hatshop_server.py`, port 8092, cloudflared, trendcarpet-qa.vercel.app).

## Where the secrets live (values are NEVER in this folder)
- **Gemini API key and Reforma API key:** `C:\Users\AndronikLindgren\reforma_automation\` (`GEMINI_API_KEY.env`,
  `Reforma-API_KEY-26-06-08.env`). All scripts here already load them from that path (e.g. `clean_interiors_gemini.py`,
  `process_iterative.py`, `orchestrator/config.py`, `reforma-turboflow/reforma_pipeline/config.py`). They load the keys from that path
  themselves; never copy, read or print the key files.
- **GitHub (repo trendcarpet-qa):** no token in the remote URL. Git authenticates through the GitHub CLI
  (`credential.helper = gh auth git-credential`, account obab33369-dotcom). `git push` works as before.
- Archive of the old duplicate key files: `C:\Users\AndronikLindgren\.secrets\` (user-only access). Not for normal use.

## Rules
1. Never copy, create or write a key, token, password or `.env` file inside this folder or any folder the hub serves.
2. Never put a token in a git remote URL, a commit, a log or a chat message.
3. `hatshop_server.py` serves a WHITELIST only (the HTML routes, `/api/...`, `/image/...`, `/manifest.json`,
   `/trendcarpet_qa.ico`). Everything else is 404. Do not restore the `super().do_GET()` fallback.
   New static files need an explicit whitelist entry.
4. `/api/open-folder` is local-only (not through the tunnel). Keep it that way.
5. Global §11.6 applies: Antigravity never calls the paid Gemini API itself or from new code, but may run the existing pipeline
   scripts (e.g. `clean_interiors_gemini.py`, `process_iterative.py`) when the task needs them, saying so in the answer. The keys
   exist for those scripts, never for the agent to read, copy or print.

## QA-hubben: regelkort (CEO 2026-10-15; källa: agent-home/ag-rules/projects/Trendcarpet_Interiors/GEMINI.md)
1. Kollegorna använder hubben hela dagen. Redigera ALDRIG live-filerna (`hatshop_ui.html`, `hatshop_mobile.html`, `hatshop_server.py`); all redigering sker i `_staging/` (`_staging/hatshop_ui.html`, `_staging/hatshop_mobile.html`).
2. Testa staging via proxyn på port 8093 (`node C:\Users\AndronikLindgren\agent-home\scripts\hub-staging-proxy.mjs`): den serverar staging-HTML och proxar `/api` och `/image` till live-hubben på 8092.
3. Inbyte i live sker bara nattetid (00–05) med `node C:\Users\AndronikLindgren\agent-home\scripts\hub-deploy-staging.mjs` (atomisk kopia + backup), och först efter ett grönt smoke-skript som loggar till fil (HTTP 200 på `/`, `/mobile`, `/api/...` och en bild via proxyn). Omstart av server eller tunnel: bara nattetid.
4. Serverfrågor (VPS, domän, named tunnel) är pausade tills CEO:n tar upp dem; lista dem inte som beslut.
5. Scope: bara den här mappen (`PROJECT_STATE.json` scope_jail); aldrig rekursiva sökningar i `OneDrive - CaMa Gruppen AB` eller andra föräldrakataloger. Reforma-möbelbilder hör till Reforma-sektionen, aldrig till Trendcarpet-/Hatshop-sektioner.
6. Hubbens mp4/fullskärm/bildordning får inte brytas av en sidändring: smoke-skriptet kontrollerar att videor spelas, fullskärm öppnar och ordningen är oförändrad innan inbyte.
7. En uppgift per tråd; efter komprimering: läs `PROJECT_STATE.json` + det här kortet först.
