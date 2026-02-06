import json
from pathlib import Path
from typing import List

from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_names_from_sheets(
    spreadsheet_id: str,
    range_name: str,
    credentials_path: Path,
    token_path: Path,
) -> List[str]:
    raw = json.loads(credentials_path.read_text(encoding="utf-8"))

    creds = None
    if raw.get("type") == "service_account":
        creds = service_account.Credentials.from_service_account_info(raw, scopes=SCOPES)
    else:
        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(credentials_path), SCOPES
                )
                creds = flow.run_local_server(port=0)

            token_path.parent.mkdir(parents=True, exist_ok=True)
            token_path.write_text(creds.to_json(), encoding="utf-8")

    service = build("sheets", "v4", credentials=creds)
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute()
    )

    values = result.get("values", [])
    names: List[str] = []
    for row in values:
        for cell in row:
            if isinstance(cell, str):
                names.append(cell)

    return names
