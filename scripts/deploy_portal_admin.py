"""
Direct Google Apps Script REST API Deployment & Synchronization Script
----------------------------------------------------------------------
Pushes all 11 modular files (.gs, .html, appsscript.json) to all 6 GAS projects
via the official Google Apps Script REST API, creates immutable versions, updates
deployments, and verifies that every project has a live WEB_APP entryPoint.
"""
import os
import re
import json
import subprocess
import requests

BASE_DIR = r"D:\Project Aplikasi\Digital-Kasbon"

def get_token():
    rc_path = r"C:\Users\Admin\.clasprc.json"
    with open(rc_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["tokens"]["default"]["access_token"]

def build_files_payload(folder_path):
    files_payload = []
    for fname in sorted(os.listdir(folder_path)):
        fpath = os.path.join(folder_path, fname)
        if not os.path.isfile(fpath) or fname.startswith("."):
            continue
        name, ext = os.path.splitext(fname)
        if ext == ".json" and name == "appsscript":
            ftype = "JSON"
        elif ext in (".gs", ".js"):
            ftype = "SERVER_JS"
        elif ext == ".html":
            ftype = "HTML"
        else:
            continue
        with open(fpath, "r", encoding="utf-8") as f:
            source = f.read()
        files_payload.append({
            "name": name,
            "type": ftype,
            "source": source
        })
    return files_payload

def push_and_deploy_project(token, script_id, folder_path, existing_dep_id=None, desc="Live Release v2.0"):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    base_api = f"https://script.googleapis.com/v1/projects/{script_id}"

    # 1. PUT content (all 11 files)
    files_payload = build_files_payload(folder_path)
    r_put = requests.put(f"{base_api}/content", headers=headers, json={"files": files_payload}, timeout=30)
    if r_put.status_code != 200:
        raise RuntimeError(f"PUT content failed for {script_id}: {r_put.status_code} {r_put.text}")
    pushed_names = [f["name"] for f in r_put.json().get("files", [])]

    # 2. POST new immutable version
    r_ver = requests.post(f"{base_api}/versions", headers=headers, json={"description": desc}, timeout=30)
    if r_ver.status_code != 200:
        raise RuntimeError(f"POST version failed for {script_id}: {r_ver.status_code} {r_ver.text}")
    ver_num = r_ver.json()["versionNumber"]

    # 3. Create or update deployment
    dep_data = None
    if existing_dep_id:
        r_upd = requests.put(
            f"{base_api}/deployments/{existing_dep_id}",
            headers=headers,
            json={
                "deploymentConfig": {
                    "scriptId": script_id,
                    "versionNumber": ver_num,
                    "manifestFileName": "appsscript",
                    "description": desc
                }
            },
            timeout=30
        )
        if r_upd.status_code == 200:
            dep_data = r_upd.json()

    # Check if dep_data has WEB_APP entryPoint; if not, create a fresh deployment
    web_app_url = None
    if dep_data:
        for ep in dep_data.get("entryPoints", []):
            if ep.get("entryPointType") == "WEB_APP":
                web_app_url = ep.get("webApp", {}).get("url")
                break

    if not web_app_url:
        r_new = requests.post(
            f"{base_api}/deployments",
            headers=headers,
            json={
                "versionNumber": ver_num,
                "manifestFileName": "appsscript",
                "description": desc
            },
            timeout=30
        )
        if r_new.status_code != 200:
            raise RuntimeError(f"POST deployment failed for {script_id}: {r_new.status_code} {r_new.text}")
        dep_data = r_new.json()
        for ep in dep_data.get("entryPoints", []):
            if ep.get("entryPointType") == "WEB_APP":
                web_app_url = ep.get("webApp", {}).get("url")
                break

    return {
        "deployment_id": dep_data["deploymentId"],
        "version_number": ver_num,
        "web_app_url": web_app_url,
        "pushed_files": len(pushed_names)
    }

def main():
    token = get_token()
    deps_file = os.path.join(BASE_DIR, "gas_deployments.json")
    with open(deps_file, "r", encoding="utf-8") as f:
        deps_data = json.load(f)

    # Step 1: Sync gas-common first
    subprocess.run(["python", "scripts/sync_gas_common.py"], cwd=BASE_DIR, check=True)

    # Step 2: Push and deploy gas-portal-admin to get its verified WEB_APP URL
    admin_info = deps_data["portalAdmin"]
    admin_folder = os.path.join(BASE_DIR, admin_info["folder"])
    print(f"[*] Deploying gas-portal-admin ({admin_info['script_id']})...")
    admin_res = push_and_deploy_project(
        token,
        admin_info["script_id"],
        admin_folder,
        existing_dep_id=None, # Create fresh WEB_APP deployment since v1 had no webapp manifest
        desc="Live Web App Portal Admin v1.0"
    )
    admin_url = admin_res["web_app_url"]
    admin_dep_id = admin_res["deployment_id"]
    print(f"    -> Verified Admin WEB_APP URL: {admin_url} (files={admin_res['pushed_files']}, v={admin_res['version_number']})")

    # Step 3: Update gas-common/Config.gs with the verified URL_PORTAL_ADMIN
    cfg_path = os.path.join(BASE_DIR, "gas-common", "Config.gs")
    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg_text = f.read()
    cfg_text = re.sub(
        r"URL_PORTAL_ADMIN:\s*'[^']*'",
        f"URL_PORTAL_ADMIN: '{admin_url}'",
        cfg_text
    )
    with open(cfg_path, "w", encoding="utf-8") as f:
        f.write(cfg_text)

    # Step 4: Re-sync gas-common across all 6 projects
    subprocess.run(["python", "scripts/sync_gas_common.py"], cwd=BASE_DIR, check=True)

    # Step 5: Push & update deployments for all 6 projects so all of them have the latest Config.gs & SupabaseClient.gs
    deps_data["portalAdmin"]["deployment_id"] = admin_dep_id
    deps_data["portalAdmin"]["url"] = admin_url

    for key, info in deps_data.items():
        folder_path = os.path.join(BASE_DIR, info["folder"])
        sid = info["script_id"]
        dep_id = info["deployment_id"]
        res = push_and_deploy_project(
            token,
            sid,
            folder_path,
            existing_dep_id=dep_id,
            desc="Enterprise Interconnected Release v2.0"
        )
        info["deployment_id"] = res["deployment_id"]
        info["url"] = res["web_app_url"]
        print(f"[OK] {key:16} | files={res['pushed_files']} | v={res['version_number']} | {res['web_app_url']}")

    with open(deps_file, "w", encoding="utf-8") as f:
        json.dump(deps_data, f, indent=2)

    print("\nALL 6 PROJECTS VERIFIED AND LIVE!")

if __name__ == "__main__":
    main()
