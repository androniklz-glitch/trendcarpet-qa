"""Scheduled run of pictures_godkanda.py (Koordinator 4.9, 2026-11-11: "QA-hubb Godkända" must follow the hub without a manual
sync): `sync` (mass-move guard kept, never --allow-mass-move), then `verify` only when sync exited 0. Registered as Windows task
AgentHome-QaHubGodkanda (agent-home scripts/startup/register-qa-hub-godkanda.ps1), every 30 min, cwd = hub dir.
Log dir ~/qa-hub-backup/pictures-godkanda/task/: runs.log (one line per run), last-sync.json / last-verify.json (stdout+stderr
of the last run), alarm.log + ALARM.json while failing. Never pushes: the coordinator reads ALARM.json and relays.
Exit 4 from sync (lock held = another run) is a skip, not a failure. Exit: 0 ok/skip, 1 sync or verify failed."""
import datetime
import json
import os
import subprocess
import sys

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(HUB_DIR, "scripts", "pictures_godkanda.py")
LOG_DIR = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "task")
TIMEOUT_S = {"sync": 15 * 60, "verify": 5 * 60}


def now():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def log(name, line):
    with open(os.path.join(LOG_DIR, name), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(cmd):
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    try:
        p = subprocess.run([sys.executable, "-u", SCRIPT, cmd], cwd=HUB_DIR, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=TIMEOUT_S[cmd], env=env)
        code, out, err = p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as exc:
        code, out, err = 124, exc.stdout or "", f"timeout after {TIMEOUT_S[cmd]} s"
    with open(os.path.join(LOG_DIR, f"last-{cmd}.json"), "w", encoding="utf-8") as f:
        json.dump({"at": now(), "cmd": cmd, "exit": code, "stdout": out[-20000:], "stderr": err[-5000:]}, f, ensure_ascii=False, indent=1)
    summary = next((l for l in reversed(out.splitlines()) if l.startswith('{"summary"') or l.startswith('{"ALARM"')), "")
    return code, summary[:600], err.strip().splitlines()[-1][:300] if err.strip() else ""


def main():
    os.makedirs(LOG_DIR, exist_ok=True)
    started = now()
    code, summary, err = run("sync")
    if code == 4:
        log("runs.log", f"{started} SKIP sync lock held")
        return 0
    vcode, vsummary, verr = run("verify") if code == 0 else (None, "", "")
    ok = code == 0 and vcode == 0
    log("runs.log", f"{started} {'ok' if ok else 'FAIL'} sync={code} verify={vcode} {summary} {vsummary} {err} {verr}".rstrip())
    alarm = os.path.join(LOG_DIR, "ALARM.json")
    if ok:
        if os.path.exists(alarm):
            os.remove(alarm)
            log("alarm.log", f"{now()} RECOVERED pictures-godkanda")
        return 0
    if not os.path.exists(alarm):
        with open(alarm, "w", encoding="utf-8") as f:
            json.dump({"since": started, "sync_exit": code, "verify_exit": vcode, "summary": summary or vsummary,
                       "stderr": err or verr}, f, ensure_ascii=False, indent=1)
        log("alarm.log", f"{now()} ALARM pictures-godkanda sync={code} verify={vcode} {err or verr}".rstrip())
    return 1


if __name__ == "__main__":
    sys.exit(main())
