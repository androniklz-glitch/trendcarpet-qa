# Trendcarpet_Interiors – FACTS (QA-hubben; CEO-beslut 2026-10-04 … 2026-10-15)

_Källa: `agent-home/ag-rules/projects/Trendcarpet_Interiors/FACTS.md`, installeras av Claude Code. Gäller före `scratch/FACTS.md`._

Den här mappen ligger bakom en publik tunnel: QA-hubben `hatshop_server.py` på port 8092, via cloudflared och trendcarpet-qa.vercel.app. Kollegorna använder hubben hela dagen.

## Hemligheter (värdena finns aldrig i den här mappen)
- **Gemini- och Reforma-nycklarna** ligger i `C:\Users\AndronikLindgren\reforma_automation\` (`GEMINI_API_KEY.env`, `Reforma-API_KEY-26-06-08.env`). Skripten här (`clean_interiors_gemini.py`, `process_iterative.py`, `orchestrator/config.py`, `reforma-turboflow/reforma_pipeline/config.py`) laddar dem själva. Du läser, kopierar och skriver aldrig ut nyckelfilerna.
- **GitHub (repot trendcarpet-qa)** autentiseras via `gh auth git-credential`. Lägg aldrig en token i en remote-URL, en commit, en logg eller ett chattmeddelande.
- **Skapa aldrig** en nyckel, en token eller en `.env` i en mapp som hubben serverar.

## Hubben
1. **Vitlista:** `hatshop_server.py` serverar bara HTML-routes, `/api/...`, `/image/...`, `/manifest.json` och `/trendcarpet_qa.ico`. Allt annat ger 404. Återinför aldrig fallbacken `super().do_GET()`. En ny statisk fil kräver en ny rad i vitlistan.
2. **`/api/open-folder` är bara lokal** och går aldrig via tunneln.
3. **Redigera aldrig live-filerna** (`hatshop_ui.html`, `hatshop_mobile.html`, `hatshop_server.py`). Allt redigeras under `_staging/`.
4. **Testa staging via proxyn på 8093:** `node C:\Users\AndronikLindgren\agent-home\scripts\hub-staging-proxy.mjs`.
5. **Inbyte bara 00–05:** `node C:\Users\AndronikLindgren\agent-home\scripts\hub-deploy-staging.mjs`. Gör det först efter ett grönt smoke-skript med loggfil: HTTP 200 på `/`, `/mobile`, `/api/...` och en bild, videor spelas, fullskärm öppnar och bildordningen är oförändrad. Server och tunnel startas också om bara nattetid.
6. **Serverfrågor** (VPS, domän, named tunnel) är pausade tills CEO tar upp dem.
7. **Reforma-möbelbilder** hör till hubbens Reforma-sektion, aldrig till Trendcarpet- eller Hatshop-sektionerna.
8. **Upplägg och publicering (CEO 2026-10-23):** bara via `node C:\Users\AndronikLindgren\agent-home\scripts\qa-publish.mjs --path <mapp> --name <namn> --category <kategori> --org <org> [--publish]`, kört av en egen Publicerare-subagent efter domarens PASS i `QA_VERDICT.json`. Skriptet lägger först upp ett opublicerat utkast, stämmer av varje bildfil mot `/api/products` och publicerar bara vid full match. Egna anrop mot `/api/batches/add`, `toggle_publish` m.fl. nekas av hooken; `batches_config.json` skrivs aldrig direkt. Länken till CEO bara med skriptets tabell.
   - **Från 2026-11-11:** hubben tar bara batchmappar i Pictures (annars 400) och AG skriver aldrig där, så AG-batcher levereras i stället till hubbens `_inkorg` enligt `C:\Users\AndronikLindgren\.gemini\antigravity\scratch\QA_HUBB_INKORG.md`; en Claude-tråd tar in dem.
9. **En publicerad eller levererad mapp struktureras aldrig om** (inga flyttar, inga `rmtree`): bygg en ny versionsmapp (`…_v3`), låt domaren döma den och publicera den.
10. **Mattbilder i Flow (hook 13, CEO 2026-10-23):** pilot först, PASS från en domare (TypeName "judge") mot `flow-content\style\JUDGE_CHECKLIST.md` innan resten körs. Mattprompter: möblerna står PÅ mattan (minst framben) och mattan ligger parallell med golvbrädorna/väggarna; aldrig "built around the cut-out" eller "all four edges visible/free". run-batch kör recipe-lint själv (exit 14).
