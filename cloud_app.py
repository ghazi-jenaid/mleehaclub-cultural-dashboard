"""نقطة تشغيل النسخة الإلكترونية للوحة اللجنة الثقافية والمجتمعية."""

from __future__ import annotations

import json
import os
import runpy
from pathlib import Path

import streamlit as st

from drive_sync import start_drive_sync


APP_DIR = Path(__file__).resolve().parent
MIRROR_ROOT = Path(os.environ.get("COMMITTEE_PROJECT_ROOT", "/tmp/committee-project"))
ROOT_FOLDER_ID = os.environ.get(
    "GOOGLE_DRIVE_ROOT_FOLDER_ID",
    "1oOgdmfprLWUku9yOIdBljxlNPk5RMGFB",
).strip()
SYNC_INTERVAL = int(os.environ.get("GOOGLE_DRIVE_SYNC_SECONDS", "60"))

os.environ["COMMITTEE_PROJECT_ROOT"] = str(MIRROR_ROOT)
if not os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") and "gcp_service_account" in st.secrets:
    os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"] = json.dumps(
        dict(st.secrets["gcp_service_account"]),
        ensure_ascii=False,
    )

try:
    start_drive_sync(ROOT_FOLDER_ID, MIRROR_ROOT, SYNC_INTERVAL)
except Exception as exc:
    st.set_page_config(page_title="اللجنة الثقافية والمجتمعية", layout="wide")
    st.error("تعذر الاتصال بمستندات اللجنة على Google Drive.")
    st.caption(str(exc))
    st.stop()

runpy.run_path(str(APP_DIR / "app.py"), run_name="__main__")
