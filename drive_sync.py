"""مزامنة مجلد اللجنة من Google Drive إلى مرآة محلية للعرض والتحليل.

تعمل المزامنة للقراءة فقط: لا يرسل التطبيق أي تعديل إلى Google Drive.
"""

from __future__ import annotations

import io
import json
import logging
import os
import threading
import time
from pathlib import Path
from typing import Any

import google.auth
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


LOGGER = logging.getLogger("committee-drive-sync")
DRIVE_READONLY_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
FOLDER_MIME = "application/vnd.google-apps.folder"
GOOGLE_EXPORTS = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".docx",
    ),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pptx",
    ),
}

_LOCK = threading.Lock()
_THREAD: threading.Thread | None = None
_READY = threading.Event()
_LAST_ERROR: str | None = None


def _safe_name(name: str) -> str:
    """يعيد اسماً صالحاً على نظام الملفات مع الحفاظ على العربية."""

    replacements = {'<': '‹', '>': '›', ':': '꞉', '"': '″', '/': '／', '\\': '＼', '|': '¦', '?': '؟', '*': '٭'}
    cleaned = "".join(replacements.get(char, char) for char in name).strip().rstrip(".")
    return cleaned or "بدون اسم"


def _credentials() -> Any:
    """ينشئ بيانات اعتماد Google بأقل صلاحية لازمة للقراءة."""

    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        info = json.loads(raw_json)
        return service_account.Credentials.from_service_account_info(
            info, scopes=[DRIVE_READONLY_SCOPE]
        )
    credentials, _ = google.auth.default(scopes=[DRIVE_READONLY_SCOPE])
    if not credentials.valid:
        credentials.refresh(Request())
    return credentials


def _service() -> Any:
    """يبني عميل Drive API دون تخزين اكتشاف الواجهة على القرص."""

    return build("drive", "v3", credentials=_credentials(), cache_discovery=False)


def _list_children(service: Any, folder_id: str) -> list[dict[str, Any]]:
    """يجلب جميع العناصر المباشرة داخل مجلد Drive."""

    items: list[dict[str, Any]] = []
    token: str | None = None
    while True:
        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken,files(id,name,mimeType,modifiedTime,size,md5Checksum)",
            pageSize=1000,
            pageToken=token,
            orderBy="name_natural",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        items.extend(response.get("files", []))
        token = response.get("nextPageToken")
        if not token:
            return items


def _download(service: Any, item: dict[str, Any], destination: Path) -> None:
    """ينزل ملفاً عادياً أو يصدّر ملف Google أصلياً بصورة ذرية."""

    mime_type = item["mimeType"]
    export = GOOGLE_EXPORTS.get(mime_type)
    if export:
        export_mime, extension = export
        if destination.suffix.lower() != extension:
            destination = destination.with_suffix(extension)
        request = service.files().export_media(fileId=item["id"], mimeType=export_mime)
    else:
        request = service.files().get_media(
            fileId=item["id"], supportsAllDrives=True
        )

    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f".{destination.name}.drive-part")
    with temporary.open("wb") as handle:
        downloader = MediaIoBaseDownload(handle, request, chunksize=8 * 1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk(num_retries=3)
    temporary.replace(destination)


def sync_once(root_folder_id: str, mirror_root: Path) -> dict[str, int]:
    """ينفذ مزامنة كاملة آمنة ويعيد إحصاء الملفات والمجلدات."""

    service = _service()
    mirror_root.mkdir(parents=True, exist_ok=True)
    discovered_paths: set[Path] = set()
    counts = {"files": 0, "folders": 0, "updated": 0, "removed": 0}

    def visit(folder_id: str, local_folder: Path) -> None:
        local_folder.mkdir(parents=True, exist_ok=True)
        discovered_paths.add(local_folder.resolve())
        for item in _list_children(service, folder_id):
            name = _safe_name(item["name"])
            destination = local_folder / name
            if item["mimeType"] == FOLDER_MIME:
                counts["folders"] += 1
                visit(item["id"], destination)
                continue

            export = GOOGLE_EXPORTS.get(item["mimeType"])
            if export and destination.suffix.lower() != export[1]:
                destination = destination.with_suffix(export[1])
            discovered_paths.add(destination.resolve())
            counts["files"] += 1

            remote_modified = item.get("modifiedTime", "")
            stamp_file = destination.with_name(f".{destination.name}.drive-meta")
            expected_stamp = f"{item['id']}|{remote_modified}|{item.get('size', '')}"
            current_stamp = stamp_file.read_text("utf-8") if stamp_file.exists() else ""
            if destination.exists() and current_stamp == expected_stamp:
                discovered_paths.add(stamp_file.resolve())
                continue

            _download(service, item, destination)
            stamp_file.write_text(expected_stamp, encoding="utf-8")
            discovered_paths.add(stamp_file.resolve())
            counts["updated"] += 1

    visit(root_folder_id, mirror_root)

    # المرآة حصرية للتطبيق السحابي؛ تزال منها فقط العناصر التي اختفت من Drive.
    for path in sorted(mirror_root.rglob("*"), key=lambda value: len(value.parts), reverse=True):
        if path.resolve() in discovered_paths:
            continue
        if path.is_file():
            path.unlink(missing_ok=True)
            counts["removed"] += 1
        elif path.is_dir():
            try:
                path.rmdir()
            except OSError:
                pass
    return counts


def _worker(root_folder_id: str, mirror_root: Path, interval_seconds: int) -> None:
    """يشغّل المزامنة المتكررة دون حجب واجهة Streamlit."""

    global _LAST_ERROR
    while True:
        try:
            with _LOCK:
                counts = sync_once(root_folder_id, mirror_root)
            _LAST_ERROR = None
            LOGGER.info("Drive sync completed: %s", counts)
        except Exception as exc:  # يحافظ على الخدمة حتى عند خطأ مؤقت من Drive.
            _LAST_ERROR = f"{type(exc).__name__}: {exc}"
            LOGGER.exception("Drive sync failed")
        finally:
            _READY.set()
        time.sleep(max(30, interval_seconds))


def start_drive_sync(
    root_folder_id: str,
    mirror_root: Path,
    interval_seconds: int = 60,
    initial_wait_seconds: int = 300,
) -> None:
    """يبدأ خيط المزامنة مرة واحدة وينتظر اكتمال الفهرسة الأولى."""

    global _THREAD
    if _THREAD is None or not _THREAD.is_alive():
        _READY.clear()
        _THREAD = threading.Thread(
            target=_worker,
            args=(root_folder_id, mirror_root, interval_seconds),
            name="google-drive-readonly-sync",
            daemon=True,
        )
        _THREAD.start()
    if not _READY.wait(timeout=initial_wait_seconds):
        raise TimeoutError("لم تكتمل المزامنة الأولى مع Google Drive في الوقت المحدد.")
    if _LAST_ERROR and not any(mirror_root.iterdir()):
        raise RuntimeError(f"تعذرت المزامنة الأولى: {_LAST_ERROR}")

