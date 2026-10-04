# Trendcarpet_Interiors: project rules for Antigravity (security, 2026-10-04)

This folder is behind a public tunnel (QA hub `hatshop_server.py`, port 8092, cloudflared, trendcarpet-qa.vercel.app).

## Where the secrets live (values are NEVER in this folder)
- **Gemini API key and Reforma API key:** `C:\Users\AndronikLindgren\reforma_automation\` (`GEMINI_API_KEY.env`,
  `Reforma-API_KEY-26-06-08.env`). All scripts here already load them from that path (e.g. `clean_interiors_gemini.py`,
  `process_iterative.py`, `orchestrator/config.py`, `reforma-turboflow/reforma_pipeline/config.py`). Use that path; do not copy it.
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
5. The CEO's rule stands: Antigravity itself does not call the paid Gemini API. The keys are for the existing pipeline scripts.
