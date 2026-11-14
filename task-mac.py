"""Manage the current macOS user's WASDE launchd agent."""
import argparse
import os
from pathlib import Path
import plistlib
import re
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parent
RUNTIME = ROOT / "runtime"
LABEL = "com.wasde.excel-update"
AGENTS = Path.home() / "Library" / "LaunchAgents"


def launch(*args, check=True, capture=False):
    return subprocess.run(
        ["/bin/launchctl", *args], check=check, text=True, capture_output=capture,
    )


def write_agent(path, label, arguments, log, *, repeat=False):
    data = {
        "Label": label,
        "WorkingDirectory": str(ROOT),
        "ProgramArguments": [
            str(ROOT / ".venv/bin/python"), str(ROOT / "wasde_automation.py"), *arguments,
        ],
        "StandardOutPath": str(log),
        "StandardErrorPath": str(log),
        "LimitLoadToSessionType": "Aqua",
    }
    if repeat:
        data["StartInterval"] = 120
    path.write_bytes(plistlib.dumps(data))
    subprocess.run(["/usr/bin/plutil", "-lint", str(path)], check=True)


def test_agent(domain):
    label = LABEL + ".test"
    target = f"{domain}/{label}"
    if launch("print", target, check=False, capture=True).returncode == 0:
        raise RuntimeError("A WASDE scheduler test is already loaded.")
    definition = RUNTIME / (label + ".plist")
    log = RUNTIME / "launchd-test.log"
    log.unlink(missing_ok=True)
    write_agent(definition, label, ["--check-install"], log)
    launch("bootstrap", domain, str(definition))
    try:
        launch("kickstart", target)
        until = time.monotonic() + 120
        while time.monotonic() < until:
            status = launch("print", target, capture=True).stdout
            match = re.search(r"last exit code = (\d+)", status)
            if match and not re.search(r"^\s*state = running\s*$", status, re.M):
                if match[1] != "0":
                    raise RuntimeError(f"Scheduled Excel test failed. See {log}")
                output = log.read_text(encoding="utf-8") if log.exists() else ""
                if "PASS: Excel for Mac" in output:
                    print(output, end="")
                    return
            time.sleep(1)
        raise RuntimeError(f"Scheduled Excel test timed out. See {log}")
    finally:
        launch("bootout", target, check=False)
        definition.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("register", "test", "status", "disable", "enable", "remove"))
    action = parser.parse_args().action
    if sys.platform != "darwin":
        parser.error("Run this command on macOS.")
    domain = f"gui/{os.getuid()}"
    target = f"{domain}/{LABEL}"
    definition = AGENTS / (LABEL + ".plist")
    if action == "status":
        launch("print", target)
        return
    if action == "test":
        RUNTIME.mkdir(exist_ok=True)
        test_agent(domain)
        return
    if action in ("register", "enable"):
        if action == "register":
            for required in (ROOT / ".venv/bin/python", ROOT / "config.json"):
                if not required.is_file():
                    raise RuntimeError("Run bash install-mac.sh first.")
            if definition.exists():
                raise RuntimeError("WASDE is already registered. Use enable, or remove and register again.")
            AGENTS.mkdir(parents=True, exist_ok=True)
            RUNTIME.mkdir(exist_ok=True)
            write_agent(definition, LABEL, ["--scheduled"], RUNTIME / "launchd.log", repeat=True)
        elif not definition.is_file():
            raise RuntimeError("Register the WASDE job first.")
        launch("enable", target)
        if launch("print", target, check=False, capture=True).returncode:
            launch("bootstrap", domain, str(definition))
        print(f"WASDE scheduler enabled: {definition}")
        return
    loaded = launch("print", target, check=False, capture=True)
    if re.search(r"^\s*state = running\s*$", loaded.stdout, re.M):
        raise RuntimeError("Wait for the current run to finish before changing the schedule.")
    launch("disable", target)
    if loaded.returncode == 0:
        launch("bootout", target)
    if action == "remove":
        definition.unlink(missing_ok=True)
    print(f"WASDE scheduler {'removed' if action == 'remove' else 'disabled'}.")


if __name__ == "__main__":
    try:
        main()
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        sys.exit(str(exc))
