"""
Deploy all 5 Google Apps Script projects using clasp CLI and verify web app entry points.
Ensures appsscript.json has correct webapp configuration (USER_DEPLOYING, ANYONE_ANONYMOUS).
Extracts live deployment URLs from the Google Apps Script API.
"""
import os
import sys
import json
import subprocess
import requests

BASE_DIR = r"d:\Project Aplikasi\Digital-Kasbon"

PROJECTS = [
    {
        "folder": "gas-dashboard",
        "title": "Kasbon - Dashboard Monitoring",
        "key": "dashboard"
    },
    {
        "folder": "gas-form-pengajuan",
        "title": "Kasbon - Form Pengajuan (Step 1)",
        "key": "formPengajuan"
    },
    {
        "folder": "gas-portal-manager",
        "title": "Kasbon - Portal Manager (Step 2 & 7)",
        "key": "portalManager"
    },
    {
        "folder": "gas-portal-cashier",
        "title": "Kasbon - Portal Kasir (Step 3 & 6)",
        "key": "portalCashier"
    },
    {
        "folder": "gas-portal-requester",
        "title": "Kasbon - Portal Pemohon (Step 4 & 5)",
        "key": "portalRequester"
    }
]

CORRECT_APPSSCRIPT = {
    "timeZone": "Asia/Jakarta",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
        "executeAs": "USER_DEPLOYING",
        "access": "ANYONE_ANONYMOUS"
    }
}

def get_token():
    rc_path = r"C:\Users\Admin\.clasprc.json"
    with open(rc_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["tokens"]["default"]["access_token"]

def run_cmd(args, cwd):
    print(f"[{os.path.basename(cwd)}] Running: {' '.join(args)}")
    result = subprocess.run(args, cwd=cwd, capture_output=True, text=True, shell=True)
    if result.returncode != 0:
        print(f"  STDERR: {result.stderr.strip()}")
    if result.stdout:
        print(f"  STDOUT: {result.stdout.strip()}")
    return result

def ensure_webapp_config(folder_path):
    manifest_path = os.path.join(folder_path, "appsscript.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(CORRECT_APPSSCRIPT, f, indent=2)

def main():
    token = get_token()
    results = {}

    for proj in PROJECTS:
        folder_path = os.path.join(BASE_DIR, proj["folder"])
        clasp_json_path = os.path.join(folder_path, ".clasp.json")
        
        with open(clasp_json_path, "r", encoding="utf-8") as f:
            script_id = json.load(f)["scriptId"]

        print(f"\n==========================================")
        print(f"Processing: {proj['title']} ({script_id})")
        print(f"==========================================")

        # 1. Update appsscript.json
        ensure_webapp_config(folder_path)

        # 2. Push files
        run_cmd(["npx.cmd", "@google/clasp", "push", "-f"], folder_path)

        # 3. Deploy
        run_cmd(["npx.cmd", "@google/clasp", "deploy", "--description", '"Live Web App v1.0"'], folder_path)

        # 4. Fetch latest deployment from API
        api_url = f"https://script.googleapis.com/v1/projects/{script_id}/deployments"
        resp = requests.get(api_url, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            deps = resp.json().get("deployments", [])
            # Find the latest deployment with WEB_APP entry point
            web_app_url = None
            dep_id = None
            for d in reversed(deps):
                entry_points = d.get("entryPoints", [])
                for ep in entry_points:
                    if ep.get("entryPointType") == "WEB_APP":
                        web_app_url = ep.get("webApp", {}).get("url")
                        dep_id = d.get("deploymentId")
                        break
                if web_app_url:
                    break

            if not web_app_url and deps:
                # Use the latest deployment ID
                dep_id = deps[-1].get("deploymentId")
                web_app_url = f"https://script.google.com/macros/s/{dep_id}/exec"

            results[proj["key"]] = {
                "title": proj["title"],
                "folder": proj["folder"],
                "script_id": script_id,
                "script_editor_url": f"https://script.google.com/d/{script_id}/edit",
                "deployment_id": dep_id,
                "url": web_app_url
            }
            print(f"  >>> WEB APP URL: {web_app_url}")
            print(f"  >>> EDITOR URL:  https://script.google.com/d/{script_id}/edit")
        else:
            print(f"  >>> ERROR fetching API: {resp.text}")

    # Save to json file
    out_file = os.path.join(BASE_DIR, "gas_deployments.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)
    print(f"\nAll deployment results saved to {out_file}")

if __name__ == "__main__":
    main()
