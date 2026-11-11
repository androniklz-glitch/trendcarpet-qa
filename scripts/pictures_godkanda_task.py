"""Scheduled run of pictures_godkanda.py (Koordinator 4.9, 2026-11-11: "QA-hubb Godkända" must follow the hub without a manual
sync): `sync` (mass-move guard kept, never --allow-mass-move), then `verify` after every sync that ran (also exit 1/3/124); only
exit 4 (lock held = another run) skips verify and is not a failure. Registered as Windows task AgentHome-QaHubGodkanda (agent-home
scripts/startup/register-qa-hub-godkanda.ps1), every 30 min, limit 25 min, cwd = hub dir.
Log dir ~/qa-hub-backup/pictures-godkanda/task/: runs.log (one line per run), last-sync.json / last-verify.json (stdout+stderr
of the last run that ran; a skip never overwrites them), alarm.log + ALARM.json while failing, skip-since.txt while skipping.
ALARM also when skips last > SKIP_ALARM_S in a row, or on any exception in this wrapper (pythonw shows nothing otherwise).
Never pushes: the coordinator reads ALARM.json and relays. Exit: 0 ok/skip, 1 sync/verify failed, long skip or wrapper error."""
import datetime
import json
import os
import subprocess
import sys
import traceback

HUB_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(HUB_DIR, "scripts", "pictures_godkanda.py")
LOG_DIR = os.path.join(os.path.expanduser("~"), "qa-hub-backup", "pictures-godkanda", "task")
TIMEOUT_S = {"sync": 20 * 60, "verify": 4 * 60}  # sum < the task limit of 25 min (4.10 gate point 4)
SKIP_ALARM_S = 60 * 60


def now():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def log(name, line):
    with open(os.path.join(LOG_DIR, name), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def call(cmd):
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    try:
        p = subprocess.run([sys.executable, "-u", SCRIPT, cmd], cwd=HUB_DIR, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=TIMEOUT_S[cmd], env=env)
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as exc:
        out = exc.stdout.decode("utf-8", "replace") if isinstance(exc.stdout, bytes) else exc.stdout or ""
        return 124, out, f"timeout after {TIMEOUT_S[cmd]} s"


def run(cmd):
    code, out, err = call(cmd)
    if code != 4:  # a skip keeps the last real run's output (4.10 gate point 3)
        with open(os.path.join(LOG_DIR, f"last-{cmd}.json"), "w", encoding="utf-8") as f:
            json.dump({"at": now(), "cmd": cmd, "exit": code, "stdout": out[-20000:], "stderr": err[-5000:]}, f,
                      ensure_ascii=False, indent=1)
    summary = next((l for l in reversed(out.splitlines()) if l.startswith('{"summary"') or l.startswith('{"ALARM"')), "")
    return code, summary[:600], err.strip().splitlines()[-1][:300] if err.strip() else ""


def raise_alarm(started, info, line):
    alarm = os.path.join(LOG_DIR, "ALARM.json")
    if not os.path.exists(alarm):
        with open(alarm, "w", encoding="utf-8") as f:
            json.dump({"since": started, **info}, f, ensure_ascii=False, indent=1)
    log("alarm.log", f"{now()} ALARM pictures-godkanda {line}".rstrip())


def main():
    os.makedirs(LOG_DIR, exist_ok=True)
    started = now()
    skip_file = os.path.join(LOG_DIR, "skip-since.txt")
    code, summary, err = run("sync")
    if code == 4:
        if not os.path.exists(skip_file):
            with open(skip_file, "w", encoding="utf-8") as f:
                f.write(datetime.datetime.now().astimezone().isoformat(timespec="seconds"))
        with open(skip_file, encoding="utf-8") as f:
            since = f.read().strip()
        skipped_s = (datetime.datetime.now().astimezone() - datetime.datetime.fromisoformat(since)).total_seconds()
        log("runs.log", f"{started} SKIP sync lock held (skipping since {since}, {int(skipped_s / 60)} min)")
        if skipped_s > SKIP_ALARM_S:
            raise_alarm(started, {"sync_exit": 4, "skipping_since": since, "stderr": err},
                        f"sync skipped (lock held) since {since}: {err}")
            return 1
        return 0
    if os.path.exists(skip_file):
        os.remove(skip_file)
    vcode, vsummary, verr = run("verify")
    ok = code == 0 and vcode == 0
    log("runs.log", f"{started} {'ok' if ok else 'FAIL'} sync={code} verify={vcode} {summary} {vsummary} {err} {verr}".rstrip())
    alarm = os.path.join(LOG_DIR, "ALARM.json")
    if ok:
        if os.path.exists(alarm):
            os.remove(alarm)
            log("alarm.log", f"{now()} RECOVERED pictures-godkanda")
        return 0
    raise_alarm(started, {"sync_exit": code, "verify_exit": vcode, "summary": summary or vsummary, "stderr": err or verr},
                f"sync={code} verify={vcode} {err or verr}")
    return 1


def guarded_main():
    """Any exception -> ALARM.json + alarm.log + exit 1 (4.10 gate point 2)."""
    started = now()
    try:
        return main()
    except Exception as exc:  # noqa: BLE001 - the task must never fail silently under pythonw
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            tb = traceback.format_exc()[-2000:]
            log("runs.log", f"{started} ERROR wrapper {type(exc).__name__}: {str(exc)[:300]}")
            raise_alarm(started, {"wrapper_error": f"{type(exc).__name__}: {exc}", "traceback": tb},
                        f"wrapper error {type(exc).__name__}: {str(exc)[:300]}")
        except Exception:  # noqa: BLE001
            pass
        return 1

if __name__ == "__main__":
    sys.exit(guarded_main())
