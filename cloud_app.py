"""نقطة تشغيل النسخة الإلكترونية للوحة اللجنة الثقافية والمجتمعية."""

from __future__ import annotations

import json
import hmac
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


def require_dashboard_login() -> None:
    """Block the public Streamlit endpoint until the shared access key is verified."""
    try:
        expected_password = str(st.secrets["dashboard_access"]["password"])
    except (KeyError, TypeError):
        st.error(
            "\u0644\u0645 \u064a\u062a\u0645 \u0625\u0639\u062f\u0627\u062f "
            "\u0645\u0641\u062a\u0627\u062d \u0627\u0644\u062f\u062e\u0648\u0644 "
            "\u0627\u0644\u0622\u0645\u0646 \u0644\u0644\u0648\u062d\u0629."
        )
        st.stop()

    if st.session_state.get("dashboard_authenticated"):
        return

    st.title("\u0627\u0644\u0644\u062c\u0646\u0629 \u0627\u0644\u062b\u0642\u0627\u0641\u064a\u0629 \u0648\u0627\u0644\u0645\u062c\u062a\u0645\u0639\u064a\u0629")
    st.caption("\u0644\u0648\u062d\u0629 \u062f\u0627\u062e\u0644\u064a\u0629 \u0645\u062d\u0645\u064a\u0629 \u2014 \u0646\u0627\u062f\u064a \u0645\u0644\u064a\u062d\u0629 \u0627\u0644\u062b\u0642\u0627\u0641\u064a \u0627\u0644\u0631\u064a\u0627\u0636\u064a")
    entered_password = st.text_input(
        "\u0645\u0641\u062a\u0627\u062d \u0627\u0644\u062f\u062e\u0648\u0644",
        type="password",
    )
    if st.button("\u062f\u062e\u0648\u0644", type="primary", use_container_width=True):
        if hmac.compare_digest(entered_password, expected_password):
            st.session_state["dashboard_authenticated"] = True
            st.rerun()
        else:
            st.error("\u0645\u0641\u062a\u0627\u062d \u0627\u0644\u062f\u062e\u0648\u0644 \u063a\u064a\u0631 \u0635\u062d\u064a\u062d.")
    st.stop()


require_dashboard_login()

os.environ["COMMITTEE_PROJECT_ROOT"] = str(MIRROR_ROOT)
if not os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") and "gcp_service_account" in st.secrets:
    os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"] = json.dumps(
        dict(st.secrets["gcp_service_account"]),
        ensure_ascii=False,
    )

try:
    start_drive_sync(
        ROOT_FOLDER_ID,
        MIRROR_ROOT,
        SYNC_INTERVAL,
        initial_wait_seconds=0,
    )
except Exception as exc:
    st.set_page_config(page_title="اللجنة الثقافية والمجتمعية", layout="wide")
    st.error("تعذر الاتصال بمستندات اللجنة على Google Drive.")
    st.caption(str(exc))
    st.stop()

runpy.run_path(str(APP_DIR / "app.py"), run_name="__main__")
