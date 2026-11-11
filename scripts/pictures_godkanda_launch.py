"""Entry point of the Windows task AgentHome-QaHubGodkanda (Koordinator 4.11 N4, 2026-11-11). Stdlib only, kept tiny so it
cannot break: imports pictures_godkanda_task and runs guarded_main(). Anything that escapes it (syntax/import error in the
task module, unwritable LOG_DIR) is invisible under pythonw, so it writes FALLBACK_ALARM (%LOCALAPPDATA%) and exits 1. A later
clean run (exit 0) removes FALLBACK_ALARM and notes RECOVERED in the task's alarm.log. The coordinator checks both alarm files."""
import datetime
import json
import os
import sys
import traceback

FALLBACK_ALARM = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "qa-hub-godkanda-ALARM.json")


def fallback_alarm(stage, exc):
    try:
        with open(FALLBACK_ALARM, "w", encoding="utf-8") as f:
            json.dump({"since": datetime.datetime.now().astimezone().isoformat(timespec="seconds"), "stage": stage,
                       "error": f"{type(exc).__name__}: {exc}", "traceback": traceback.format_exc()[-2000:]}, f,
                      ensure_ascii=False, indent=1)
    except Exception:  # noqa: BLE001 - nothing left to report to
        pass


def launch():
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import pictures_godkanda_task as task
    except BaseException as exc:  # noqa: BLE001 - SyntaxError/ImportError in the task module
        fallback_alarm("import", exc)
        return 1
    try:
        code = task.guarded_main()
    except BaseException as exc:  # noqa: BLE001 - guarded_main could not write its own alarm (LOG_DIR)
        fallback_alarm("run", exc)
        return 1
    if code == 0 and os.path.exists(FALLBACK_ALARM):
        try:
            os.remove(FALLBACK_ALARM)
            task.log("alarm.log", f"{task.now()} RECOVERED pictures-godkanda fallback alarm ({FALLBACK_ALARM})")
        except Exception as exc:  # noqa: BLE001
            fallback_alarm("recover", exc)
            return 1
    return code


if __name__ == "__main__":
    sys.exit(launch())
