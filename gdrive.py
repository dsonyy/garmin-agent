import logging
import os
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

log = logging.getLogger(__name__)

CLIENT_SECRET_FILE = os.getenv(
    "GDRIVE_CLIENT_SECRET_FILE",
    str(Path(__file__).parent / "client_secret.json"),
)
TOKEN_FILE = str(Path(__file__).parent / "secrets" / "gdrive_token.json")
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

_service = None


def _get_service():
    global _service
    if _service is not None:
        return _service

    creds = None
    if Path(TOKEN_FILE).exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not Path(CLIENT_SECRET_FILE).exists():
                raise FileNotFoundError(
                    f"Client secret not found: {CLIENT_SECRET_FILE}")
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        Path(TOKEN_FILE).write_text(creds.to_json())

    _service = build("drive", "v3", credentials=creds)
    return _service


def _find_existing_file(service, filename: str, folder_id: str) -> str | None:
    query = f"name = '{filename}' and '{folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def download_from_drive(filename: str, folder_id: str, local_path: Path | str) -> bool:
    service = _get_service()
    file_id = _find_existing_file(service, filename, folder_id)
    if not file_id:
        return False
    content = service.files().get_media(fileId=file_id).execute()
    Path(local_path).write_bytes(content)
    log.info(f"Downloaded {filename} from Drive")
    return True


def download_google_doc(doc_name: str, folder_id: str, local_path: Path | str) -> bool:
    """Download a native Google Doc as plain text."""
    service = _get_service()
    file_id = _find_existing_file(service, doc_name, folder_id)
    if not file_id:
        return False
    content = service.files().export(fileId=file_id, mimeType="text/plain").execute()
    Path(local_path).write_bytes(content)
    log.info(f"Exported Google Doc '{doc_name}' from Drive")
    return True


def upload_google_doc(file_path: Path | str, doc_name: str, folder_id: str) -> str:
    """Upload/update a text file as a native Google Doc."""
    file_path = Path(file_path)
    service = _get_service()
    media = MediaFileUpload(str(file_path), mimetype="text/plain", resumable=True)

    existing_id = _find_existing_file(service, doc_name, folder_id)

    if existing_id:
        service.files().update(fileId=existing_id, media_body=media).execute()
        log.info(f"Updated Google Doc '{doc_name}' (ID: {existing_id})")
        return existing_id
    else:
        file_metadata = {
            "name": doc_name,
            "parents": [folder_id],
            "mimeType": "application/vnd.google-apps.document",
        }
        file = service.files().create(
            body=file_metadata, media_body=media, fields="id"
        ).execute()
        file_id = file.get("id")
        log.info(f"Created Google Doc '{doc_name}' (ID: {file_id})")
        return file_id


def upload_to_drive(file_path: Path | str, folder_id: str) -> str:
    file_path = Path(file_path)
    service = _get_service()

    mime_map = {
        ".json": "application/json",
        ".md": "text/markdown",
        ".csv": "text/csv",
        ".txt": "text/plain",
    }
    mime_type = mime_map.get(file_path.suffix, "application/octet-stream")
    media = MediaFileUpload(str(file_path), mimetype=mime_type, resumable=True)

    existing_id = _find_existing_file(service, file_path.name, folder_id)

    if existing_id:
        service.files().update(fileId=existing_id, media_body=media).execute()
        log.info(f"Updated {file_path.name} (ID: {existing_id})")
        return existing_id
    else:
        file_metadata = {"name": file_path.name, "parents": [folder_id]}
        file = service.files().create(
            body=file_metadata, media_body=media, fields="id"
        ).execute()
        file_id = file.get("id")
        log.info(f"Uploaded {file_path.name} (ID: {file_id})")
        return file_id
