"""
Sync Shared GAS Modules (`gas-common/`) into all 5 Modular Apps Script Projects
-------------------------------------------------------------------------------
Ensures each Apps Script project directory (`gas-dashboard`, `gas-form-pengajuan`,
`gas-portal-manager`, `gas-portal-cashier`, `gas-portal-requester`) has the latest
shared core modules and a standardized `appsscript.json` V8 manifest.
"""
import os
import json
import shutil

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMMON_DIR = os.path.join(BASE_DIR, "gas-common")

TARGET_PROJECTS = [
    "gas-dashboard",
    "gas-form-pengajuan",
    "gas-portal-manager",
    "gas-portal-cashier",
    "gas-portal-requester",
    "gas-portal-admin",
]

SHARED_FILES = [
    "Config.gs",
    "SupabaseClient.gs",
    "StorageService.gs",
    "AuthService.gs",
    "KasbonHelper.gs",
    "SharedStyles.html",
    "SharedUtils.html",
]

# Standard manifest for Portals (executes as deploying developer)
MANIFEST_DEPLOYING = {
    "timeZone": "Asia/Jakarta",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
        "executeAs": "USER_DEPLOYING",
        "access": "ANYONE"
    }
}

# Dedicated manifest for Form Pengajuan (executes as accessing user to capture verified Google email)
MANIFEST_ACCESSING = {
    "timeZone": "Asia/Jakarta",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
        "executeAs": "USER_ACCESSING",
        "access": "ANYONE"
    },
    "oauthScopes": [
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/script.external_request",
        "https://www.googleapis.com/auth/drive"
    ]
}

def main():
    for proj in TARGET_PROJECTS:
        proj_dir = os.path.join(BASE_DIR, proj)
        os.makedirs(proj_dir, exist_ok=True)

        # Write appsscript.json manifest (project-specific)
        manifest_path = os.path.join(proj_dir, "appsscript.json")
        selected_manifest = MANIFEST_ACCESSING if proj == "gas-form-pengajuan" else MANIFEST_DEPLOYING
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(selected_manifest, f, indent=2)

        # Sync shared files from gas-common/
        for fname in SHARED_FILES:
            src = os.path.join(COMMON_DIR, fname)
            dst = os.path.join(proj_dir, fname)
            shutil.copy2(src, dst)

        print(f"[OK] Synced {len(SHARED_FILES)} shared modules + appsscript.json -> {proj}/")


if __name__ == "__main__":
    main()
