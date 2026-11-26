"""Mirror extracted workbook data into matching Google Sheets tabs."""

import os
import re
import ssl
import tempfile
import time as time_module
from datetime import date, datetime, time as datetime_time
from typing import Any, Dict, List, Optional, Sequence

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, to_excel


SCOPES = (
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
)
SPREADSHEET_MIME_TYPE = "application/vnd.google-apps.spreadsheet"
FORMAT_BATCH_SIZE = 1000
VALUE_BATCH_ROWS = 1000
MAX_RATE_LIMIT_RETRIES = 3
RATE_LIMIT_RETRY_SECONDS = 60
TRANSIENT_HTTP_STATUSES = {429, 500, 502, 503, 504}
TRANSIENT_RETRY_BASE_SECONDS = 3.0


def _google_ca_bundle(settings: Dict[str, Any], project_root: str) -> str:
    configured_bundle = settings.get("ca_bundle")
    if configured_bundle:
        bundle_path = configured_bundle
        if not os.path.isabs(bundle_path):
            bundle_path = os.path.join(project_root, bundle_path)
        if not os.path.isfile(bundle_path):
            raise FileNotFoundError(
                f"Google API CA bundle not found: {os.path.abspath(bundle_path)}"
            )
        return bundle_path

    import certifi

    certifi_path = certifi.where()
    if not hasattr(ssl, "enum_certificates"):
        return certifi_path

    server_auth_oid = "1.3.6.1.5.5.7.3.1"
    windows_roots = []
    for certificate, encoding, trust in ssl.enum_certificates("ROOT"):
        if encoding != "x509_asn":
            continue
        if trust is not True and trust and server_auth_oid not in trust:
            continue
        windows_roots.append(ssl.DER_cert_to_PEM_cert(certificate).encode("ascii"))
    if not windows_roots:
        raise RuntimeError(
            "The Windows ROOT certificate store returned no trusted certificates."
        )

    with open(certifi_path, "rb") as certifi_file:
        bundle_contents = certifi_file.read() + b"\n" + b"\n".join(windows_roots)

    cache_root = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    cache_dir = os.path.join(cache_root, "RevenueAnalysisUpdater")
    os.makedirs(cache_dir, exist_ok=True)
    bundle_path = os.path.join(cache_dir, "google-api-ca-bundle.pem")
    bundle_is_current = False
    if os.path.isfile(bundle_path):
        with open(bundle_path, "rb") as current_bundle:
            bundle_is_current = current_bundle.read() == bundle_contents
    if not bundle_is_current:
        file_descriptor, temporary_path = tempfile.mkstemp(
            prefix="google-api-ca-", suffix=".pem", dir=cache_dir
        )
        try:
            with os.fdopen(file_descriptor, "wb") as bundle_file:
                bundle_file.write(bundle_contents)
            os.replace(temporary_path, bundle_path)
        finally:
            if os.path.exists(temporary_path):
                os.remove(temporary_path)
    return bundle_path


def _configured_path(path: str, project_root: str, setting_name: str) -> str:
    resolved_path = os.path.expandvars(os.path.expanduser(path))
    if "%" in resolved_path:
        raise ValueError(
            f"Could not resolve {setting_name}; check environment variables in its path."
        )
    if not os.path.isabs(resolved_path):
        resolved_path = os.path.join(project_root, resolved_path)
    return os.path.abspath(resolved_path)


OAUTH_REFRESH_RETRIES = 3
OAUTH_REFRESH_BASE_DELAY_SECONDS = 2.0
OAUTH_INTERACTIVE_TIMEOUT_SECONDS = 300


def _refresh_oauth_credentials(credentials: Any, ca_bundle: str) -> None:
    """Refresh an expired OAuth token with retries on transient network errors."""
    import requests
    from google.auth.exceptions import RefreshError, TransportError
    from google.auth.transport.requests import Request

    last_exc: Optional[Exception] = None
    for attempt in range(OAUTH_REFRESH_RETRIES):
        session = requests.Session()
        session.verify = ca_bundle
        try:
            credentials.refresh(Request(session=session))
            return
        except RefreshError:
            raise
        except (TransportError, requests.exceptions.RequestException, TimeoutError, OSError) as exc:
            if isinstance(exc, (ssl.SSLError, requests.exceptions.SSLError)):
                raise
            last_exc = exc
            if attempt == OAUTH_REFRESH_RETRIES - 1:
                raise
            time_module.sleep(OAUTH_REFRESH_BASE_DELAY_SECONDS * (attempt + 1))
        finally:
            session.close()
    if last_exc is not None:
        raise last_exc


def _oauth_user_credentials(
    settings: Dict[str, Any],
    project_root: str,
    *,
    allow_interactive: bool = False,
    force_reauth: bool = False,
):
    client_file = settings.get("oauth_client_file")
    token_file = settings.get("oauth_token_file")
    if not client_file or not token_file:
        raise ValueError(
            "Configure google_sheets.oauth_client_file and "
            "google_sheets.oauth_token_file for OAuth user authentication."
        )

    client_path = _configured_path(
        client_file, project_root, "google_sheets.oauth_client_file"
    )
    token_path = _configured_path(
        token_file, project_root, "google_sheets.oauth_token_file"
    )
    if not os.path.isfile(client_path):
        raise FileNotFoundError(
            f"Google OAuth desktop client file not found: {client_path}"
        )

    try:
        from google.auth.exceptions import RefreshError
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError as exc:
        raise RuntimeError(
            "Google user OAuth requires google-auth-oauthlib. "
            "Install the project requirements and retry."
        ) from exc

    ca_bundle = _google_ca_bundle(settings, project_root)
    credentials = None
    needs_save = False
    if not force_reauth and os.path.isfile(token_path):
        credentials = Credentials.from_authorized_user_file(
            token_path, scopes=SCOPES
        )
        if credentials.expired and credentials.refresh_token:
            try:
                _refresh_oauth_credentials(credentials, ca_bundle)
                needs_save = True
            except RefreshError as exc:
                if not allow_interactive:
                    raise RuntimeError(
                        "The saved Google OAuth token could not be refreshed (expired or revoked; "
                        "if your Google Cloud OAuth app is in 'Testing' status, tokens expire every 7 days). "
                        "Run '.venv\\Scripts\\python.exe scripts\\authorize_google_sheets.py' to re-authorize, "
                        "then retry with '--resume'."
                    ) from exc
                credentials = None

    if credentials is None or not credentials.valid:
        if not allow_interactive:
            raise RuntimeError(
                f"Google OAuth token is missing or invalid ({token_path}) and interactive login "
                "is disabled during automated ETL runs. Run "
                "'.venv\\Scripts\\python.exe scripts\\authorize_google_sheets.py' interactively, "
                "then retry with '--resume'."
            )
        os.makedirs(os.path.dirname(token_path), exist_ok=True)
        flow = InstalledAppFlow.from_client_secrets_file(
            client_path, scopes=SCOPES
        )
        flow.oauth2session.verify = ca_bundle
        credentials = flow.run_local_server(
            port=0,
            access_type="offline",
            prompt="select_account consent",
            timeout_seconds=OAUTH_INTERACTIVE_TIMEOUT_SECONDS,
        )
        if not getattr(credentials, "refresh_token", None):
            raise RuntimeError(
                "Google OAuth authorization succeeded without returning a refresh_token. "
                "Revoke existing app access in your Google Account security settings and "
                "re-run 'scripts\\authorize_google_sheets.py --force'."
            )
        needs_save = True

    if needs_save:
        os.makedirs(os.path.dirname(token_path), exist_ok=True)
        file_descriptor, temporary_path = tempfile.mkstemp(
            prefix="google-oauth-", suffix=".json", dir=os.path.dirname(token_path)
        )
        try:
            with os.fdopen(file_descriptor, "w", encoding="utf-8") as token_stream:
                token_stream.write(credentials.to_json())
                token_stream.flush()
                os.fsync(token_stream.fileno())
            for attempt in range(5):
                try:
                    os.replace(temporary_path, token_path)
                    break
                except PermissionError:
                    if attempt == 4:
                        raise
                    time_module.sleep(0.25 * (attempt + 1))
        finally:
            if os.path.exists(temporary_path):
                os.remove(temporary_path)

    return credentials


def _a1_sheet_name(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def _google_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (datetime, date, datetime_time)):
        return to_excel(value, CALENDAR_WINDOWS_1900)
    if isinstance(value, (str, bool, int, float)):
        return value
    return str(value)


def _rgb_color(color: Any) -> Optional[Dict[str, float]]:
    if color is None or color.type != "rgb" or not color.rgb:
        return None
    value = str(color.rgb)[-6:]
    try:
        return {
            "red": int(value[0:2], 16) / 255,
            "green": int(value[2:4], 16) / 255,
            "blue": int(value[4:6], 16) / 255,
        }
    except ValueError:
        return None


def _number_format(cell: Any) -> Optional[Dict[str, str]]:
    pattern = cell.number_format
    if not pattern or pattern.casefold() == "general":
        return None
    pattern = re.sub(r"_(.)", "", pattern)
    pattern = re.sub(r"\*(.)?", "", pattern).replace("\\", "")
    lower = pattern.casefold()
    if "@" in pattern:
        kind = "TEXT"
    elif any(token in lower for token in ("yy", "dd")):
        kind = "DATE_TIME" if any(token in lower for token in ("hh", "ss")) else "DATE"
    elif any(token in lower for token in ("hh", "ss")):
        kind = "TIME"
    elif "%" in pattern:
        kind = "PERCENT"
    elif any(symbol in pattern for symbol in ("$", "€", "£", "¥")):
        kind = "CURRENCY"
    else:
        kind = "NUMBER"
    return {"type": kind, "pattern": pattern}


def _border_side(side: Any) -> Optional[Dict[str, Any]]:
    if side is None or not side.style:
        return None
    styles = {
        "thin": "SOLID",
        "medium": "SOLID_MEDIUM",
        "thick": "SOLID_THICK",
        "hair": "DOTTED",
        "dotted": "DOTTED",
        "dashed": "DASHED",
        "dashDot": "DASH_DOT",
        "dashDotDot": "DASH_DOT_DOT",
        "double": "DOUBLE",
    }
    result = {"style": styles.get(side.style, "SOLID")}
    color = _rgb_color(side.color)
    if color:
        result["color"] = color
    return result


def _cell_format(cell: Any) -> Dict[str, Any]:
    result: Dict[str, Any] = {}
    font: Dict[str, Any] = {}
    if cell.font.bold is not None:
        font["bold"] = cell.font.bold
    if cell.font.italic is not None:
        font["italic"] = cell.font.italic
    if cell.font.underline:
        font["underline"] = True
    if cell.font.strike is not None:
        font["strikethrough"] = cell.font.strike
    if cell.font.name:
        font["fontFamily"] = cell.font.name
    if cell.font.sz:
        font["fontSize"] = max(1, int(round(float(cell.font.sz))))
    font_color = _rgb_color(cell.font.color)
    if font_color:
        font["foregroundColor"] = font_color
    if font:
        result["textFormat"] = font

    fill_color = _rgb_color(cell.fill.fgColor)
    if cell.fill.fill_type == "solid" and fill_color:
        result["backgroundColor"] = fill_color

    horizontal = {
        "left": "LEFT",
        "center": "CENTER",
        "centerContinuous": "CENTER",
        "right": "RIGHT",
        "justify": "JUSTIFY",
        "distributed": "JUSTIFY",
    }.get(cell.alignment.horizontal)
    if horizontal:
        result["horizontalAlignment"] = horizontal
    vertical = {
        "top": "TOP",
        "center": "MIDDLE",
        "bottom": "BOTTOM",
        "justify": "BOTTOM",
        "distributed": "BOTTOM",
    }.get(cell.alignment.vertical)
    if vertical:
        result["verticalAlignment"] = vertical
    if cell.alignment.wrap_text is not None:
        result["wrapStrategy"] = "WRAP" if cell.alignment.wrap_text else "CLIP"
    if cell.alignment.text_rotation:
        result["textRotation"] = {"angle": int(round(float(cell.alignment.text_rotation)))}

    number_format = _number_format(cell)
    if number_format:
        result["numberFormat"] = number_format

    borders = {}
    for source_name, target_name in (
        ("top", "top"),
        ("bottom", "bottom"),
        ("left", "left"),
        ("right", "right"),
    ):
        border = _border_side(getattr(cell.border, source_name))
        if border:
            borders[target_name] = border
    if borders:
        result["borders"] = borders
    return result


def _format_requests(
    worksheet: Any,
    sheet_id: int,
    source_start_row: int,
    source_start_column: int,
    row_count: int,
    column_count: int,
) -> List[Dict[str, Any]]:
    """Compress identical adjacent Excel row styles into Sheets format ranges."""
    row_styles = []
    for row in range(source_start_row, source_start_row + row_count):
        row_styles.append(
            tuple(
                worksheet.cell(row, column).style_id
                for column in range(
                    source_start_column, source_start_column + column_count
                )
            )
        )

    row_groups = []
    group_start = 0
    for index in range(1, len(row_styles) + 1):
        if index == len(row_styles) or row_styles[index] != row_styles[group_start]:
            row_groups.append((group_start, index, row_styles[group_start]))
            group_start = index

    requests = []
    for relative_start, relative_end, styles in row_groups:
        column_offset = 0
        while column_offset < len(styles):
            style_id = styles[column_offset]
            run_end = column_offset + 1
            while run_end < len(styles) and styles[run_end] == style_id:
                run_end += 1
            if style_id:
                cell = worksheet.cell(
                    source_start_row + relative_start,
                    source_start_column + column_offset,
                )
                cell_format = _cell_format(cell)
                if cell_format:
                    requests.append(
                        {
                            "repeatCell": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": source_start_row - 1 + relative_start,
                                    "endRowIndex": source_start_row - 1 + relative_end,
                                    "startColumnIndex": source_start_column - 1 + column_offset,
                                    "endColumnIndex": source_start_column - 1 + run_end,
                                },
                                "cell": {"userEnteredFormat": cell_format},
                                "fields": "userEnteredFormat",
                            }
                        }
                    )
            column_offset = run_end
    return requests


def _dimension_requests(
    worksheet: Any,
    sheet_id: int,
    source_start_row: int,
    source_start_column: int,
    row_count: int,
    column_count: int,
) -> List[Dict[str, Any]]:
    requests = []
    dimensions = []
    for column in range(source_start_column, source_start_column + column_count):
        width = worksheet.column_dimensions[get_column_letter(column)].width
        if width is not None:
            dimensions.append(("COLUMNS", column - 1, round(width * 7 + 5)))
    for row in range(source_start_row, source_start_row + row_count):
        height = worksheet.row_dimensions[row].height
        if height is not None:
            dimensions.append(("ROWS", row - 1, round(height * 4 / 3)))

    dimensions.sort(key=lambda item: (item[0], item[1]))
    index = 0
    while index < len(dimensions):
        dimension, start, pixel_size = dimensions[index]
        end = start + 1
        index += 1
        while (
            index < len(dimensions)
            and dimensions[index][0] == dimension
            and dimensions[index][1] == end
            and dimensions[index][2] == pixel_size
        ):
            end += 1
            index += 1
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": dimension,
                        "startIndex": start,
                        "endIndex": end,
                    },
                    "properties": {"pixelSize": pixel_size},
                    "fields": "pixelSize",
                }
            }
        )
    return requests


class GoogleSheetsSync:
    """Sync selected local workbook tabs into existing Google Sheets files."""

    def __init__(
        self,
        config: Dict[str, Any],
        *,
        drive_service: Any = None,
        sheets_service: Any = None,
        allow_interactive: bool = False,
        force_reauth: bool = False,
    ):
        settings = config.get("google_sheets", {})
        self.folder_id = settings.get("folder_id")
        if not self.folder_id:
            raise ValueError("Configure google_sheets.folder_id.")

        self.drive_service = drive_service
        self.sheets_service = sheets_service
        if self.drive_service is None or self.sheets_service is None:
            try:
                import httplib2
                from google_auth_httplib2 import AuthorizedHttp
                from googleapiclient.discovery import build
            except ImportError as exc:
                raise RuntimeError(
                    "Google Sheets sync requires certifi, google-api-python-client, "
                    "google-auth, and google-auth-httplib2."
                ) from exc
            auth_type = settings.get("auth_type", "service_account")
            if auth_type == "oauth_user":
                credentials = _oauth_user_credentials(
                    settings,
                    config.get("_project_root", os.getcwd()),
                    allow_interactive=allow_interactive,
                    force_reauth=force_reauth,
                )
            elif auth_type == "service_account":
                credentials_path = settings.get("service_account_file")
                if not credentials_path:
                    raise ValueError(
                        "Configure google_sheets.service_account_file for "
                        "service-account authentication."
                    )
                credentials_path = _configured_path(
                    credentials_path,
                    config.get("_project_root", os.getcwd()),
                    "google_sheets.service_account_file",
                )
                if not os.path.isfile(credentials_path):
                    raise FileNotFoundError(
                        f"Google service-account key not found: {credentials_path}"
                    )
                from google.oauth2.service_account import Credentials

                credentials = Credentials.from_service_account_file(
                    credentials_path, scopes=SCOPES
                )
            else:
                raise ValueError(
                    f"Unsupported google_sheets.auth_type {auth_type!r}; "
                    "use 'oauth_user' or 'service_account'."
                )
            ca_bundle = _google_ca_bundle(
                settings, config.get("_project_root", os.getcwd())
            )
            self.drive_service = build(
                "drive",
                "v3",
                http=AuthorizedHttp(
                    credentials, http=httplib2.Http(ca_certs=ca_bundle)
                ),
                cache_discovery=False,
            )
            self.sheets_service = build(
                "sheets",
                "v4",
                http=AuthorizedHttp(
                    credentials, http=httplib2.Http(ca_certs=ca_bundle)
                ),
                cache_discovery=False,
            )

    @staticmethod
    def _execute_write(request: Any, logger: Any = None) -> Any:
        from googleapiclient.errors import HttpError

        for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
            try:
                return request.execute()
            except HttpError as exc:
                status = getattr(getattr(exc, "resp", None), "status", None)
                if status not in TRANSIENT_HTTP_STATUSES:
                    raise
                if attempt == MAX_RATE_LIMIT_RETRIES:
                    raise
                if status == 429:
                    retry_after = getattr(exc, "resp", {}).get("retry-after")
                    try:
                        delay = (
                            max(float(retry_after), 1.0)
                            if retry_after
                            else RATE_LIMIT_RETRY_SECONDS
                        )
                    except (TypeError, ValueError):
                        delay = RATE_LIMIT_RETRY_SECONDS
                    if logger:
                        logger.warning(
                            "Google Sheets write quota reached; retrying after %.0f seconds.",
                            delay,
                        )
                else:
                    delay = TRANSIENT_RETRY_BASE_SECONDS * (attempt + 1)
                    if logger:
                        logger.warning(
                            "Google API transient HTTP %s; retrying in %.1f seconds (%d/%d).",
                            status,
                            delay,
                            attempt + 1,
                            MAX_RATE_LIMIT_RETRIES,
                        )
                time_module.sleep(delay)
            except (TimeoutError, ConnectionError, OSError) as exc:
                if isinstance(exc, ssl.SSLError):
                    raise
                if attempt == MAX_RATE_LIMIT_RETRIES:
                    raise
                delay = TRANSIENT_RETRY_BASE_SECONDS * (attempt + 1)
                if logger:
                    logger.warning(
                        "Google API transient network error (%s); retrying in %.1f seconds (%d/%d).",
                        exc,
                        delay,
                        attempt + 1,
                        MAX_RATE_LIMIT_RETRIES,
                    )
                time_module.sleep(delay)

    def _find_spreadsheet(self, title: str, logger: Any = None) -> str:
        escaped_title = title.replace("\\", "\\\\").replace("'", "\\'")
        query = (
            f"'{self.folder_id}' in parents and name = '{escaped_title}' "
            f"and mimeType = '{SPREADSHEET_MIME_TYPE}' and trashed = false"
        )
        files = []
        page_token = None
        while True:
            response = self._execute_write(
                self.drive_service.files().list(
                    q=query,
                    pageSize=100,
                    fields="nextPageToken,files(id,name)",
                    pageToken=page_token,
                ),
                logger,
            )
            files.extend(response.get("files", []))
            page_token = response.get("nextPageToken")
            if not page_token:
                break

        if len(files) != 1:
            raise RuntimeError(
                f"Expected exactly one Google spreadsheet named {title!r} in the "
                f"configured Drive folder; found {len(files)}."
            )
        return files[0]["id"]

    def _spreadsheet_sheets(self, spreadsheet_id: str, logger: Any = None) -> Dict[str, Dict[str, Any]]:
        spreadsheet = self._execute_write(
            self.sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                includeGridData=False,
                fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
            ),
            logger,
        )
        return {
            sheet["properties"]["title"]: sheet["properties"]
            for sheet in spreadsheet.get("sheets", [])
        }

    def _ensure_grid(
        self,
        spreadsheet_id: str,
        properties: Dict[str, Any],
        rows: int,
        columns: int,
        logger: Any = None,
    ) -> Dict[str, Any]:
        grid = properties.get("gridProperties", {})
        row_count = max(grid.get("rowCount", 1000), rows)
        column_count = max(grid.get("columnCount", 26), columns)
        if (
            row_count != grid.get("rowCount")
            or column_count != grid.get("columnCount")
        ):
            updated = dict(properties)
            updated["gridProperties"] = {
                **grid,
                "rowCount": row_count,
                "columnCount": column_count,
            }
            self._execute_write(
                self.sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={
                        "requests": [
                            {
                                "updateSheetProperties": {
                                    "properties": updated,
                                    "fields": "gridProperties.rowCount,gridProperties.columnCount",
                                }
                            }
                        ]
                    },
                ),
                logger,
            )
            properties = updated
        return properties

    def _replace_sheet_range(
        self,
        spreadsheet_id: str,
        properties: Dict[str, Any],
        worksheet: Any,
        values: Sequence[Sequence[Any]],
        *,
        start_row: int,
        start_column: int,
        source_start_row: int,
        source_start_column: int,
        clear_column_count: int,
        logger: Any = None,
    ) -> None:
        title = properties["title"]
        row_count = len(values)
        if not row_count or not clear_column_count:
            raise ValueError(f"Refusing to sync an empty range for tab {title!r}.")
        column_count = max(len(row) for row in values)
        if column_count > clear_column_count:
            raise ValueError(f"Sync range exceeds configured columns for tab {title!r}.")

        end_row = start_row + row_count - 1
        end_column = start_column + clear_column_count - 1
        properties = self._ensure_grid(
            spreadsheet_id, properties, end_row, end_column, logger
        )
        grid = properties["gridProperties"]
        sheet_name = _a1_sheet_name(title)
        clear_range = (
            f"{sheet_name}!{get_column_letter(start_column)}{start_row}:"
            f"{get_column_letter(end_column)}{grid['rowCount']}"
        )
        self._execute_write(
            self.sheets_service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id, range=clear_range, body={}
            ),
            logger,
        )

        style_clear = {
            "repeatCell": {
                "range": {
                    "sheetId": properties["sheetId"],
                    "startRowIndex": start_row - 1,
                    "endRowIndex": grid["rowCount"],
                    "startColumnIndex": start_column - 1,
                    "endColumnIndex": end_column,
                },
                "cell": {},
                "fields": "userEnteredFormat",
            }
        }
        self._execute_write(
            self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [style_clear]}
            ),
            logger,
        )

        for offset in range(0, row_count, VALUE_BATCH_ROWS):
            batch = values[offset : offset + VALUE_BATCH_ROWS]
            self._execute_write(
                self.sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!{get_column_letter(start_column)}{start_row + offset}",
                    valueInputOption="RAW",
                    body={"majorDimension": "ROWS", "values": batch},
                ),
                logger,
            )

        format_requests = _format_requests(
            worksheet,
            properties["sheetId"],
            source_start_row,
            source_start_column,
            row_count,
            clear_column_count,
        )
        format_requests.extend(
            _dimension_requests(
                worksheet,
                properties["sheetId"],
                source_start_row,
                source_start_column,
                row_count,
                clear_column_count,
            )
        )
        for offset in range(0, len(format_requests), FORMAT_BATCH_SIZE):
            self._execute_write(
                self.sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={
                        "requests": format_requests[
                            offset : offset + FORMAT_BATCH_SIZE
                        ]
                    },
                ),
                logger,
            )
        if logger:
            logger.info(
                "Synced %s rows x %s columns to Google Sheets tab %s.",
                row_count,
                clear_column_count,
                title,
            )

    def sync_workbook_tabs(
        self,
        workbook_path: str,
        spreadsheet_title: str,
        sheet_names: Any,
        *,
        logger: Any = None,
    ) -> None:
        spreadsheet_id = self._find_spreadsheet(spreadsheet_title, logger=logger)
        target_sheets = self._spreadsheet_sheets(spreadsheet_id, logger=logger)
        workbook = openpyxl.load_workbook(
            workbook_path, data_only=True, read_only=False
        )
        try:
            tab_mapping = (
                dict(sheet_names)
                if isinstance(sheet_names, dict)
                else {name: name for name in sheet_names}
            )
            missing_local = [
                name for name in tab_mapping if name not in workbook.sheetnames
            ]
            missing_remote = [
                name for name in tab_mapping.values() if name not in target_sheets
            ]
            if missing_local:
                raise ValueError(
                    f"Local workbook {workbook_path!r} is missing tab(s): "
                    f"{', '.join(missing_local)}."
                )
            if missing_remote:
                raise ValueError(
                    f"Google spreadsheet {spreadsheet_title!r} is missing tab(s): "
                    f"{', '.join(missing_remote)}."
                )
            for source_name, target_name in tab_mapping.items():
                worksheet = workbook[source_name]
                rows = worksheet.max_row
                columns = worksheet.max_column
                values = [
                    [
                        _google_value(worksheet.cell(row, column).value)
                        for column in range(1, columns + 1)
                    ]
                    for row in range(1, rows + 1)
                ]
                self._replace_sheet_range(
                    spreadsheet_id,
                    target_sheets[target_name],
                    worksheet,
                    values,
                    start_row=1,
                    start_column=1,
                    source_start_row=1,
                    source_start_column=1,
                    clear_column_count=columns,
                    logger=logger,
                )
        finally:
            workbook.close()

    def sync_pos_payments(
        self,
        workbook_path: str,
        spreadsheet_title: str,
        headers: Sequence[str],
        rows: Sequence[Sequence[Any]],
        *,
        logger: Any = None,
    ) -> None:
        sheet_name = "BO_extract"
        spreadsheet_id = self._find_spreadsheet(spreadsheet_title, logger=logger)
        target_sheets = self._spreadsheet_sheets(spreadsheet_id, logger=logger)
        if sheet_name not in target_sheets:
            raise ValueError(
                f"Google spreadsheet {spreadsheet_title!r} has no tab {sheet_name!r}."
            )
        workbook = openpyxl.load_workbook(
            workbook_path, data_only=False, read_only=False
        )
        try:
            if sheet_name not in workbook.sheetnames:
                raise ValueError(
                    f"Local workbook {workbook_path!r} has no tab {sheet_name!r}."
                )
            worksheet = workbook[sheet_name]
            values = [
                [_google_value(value) for value in headers],
                *[
                    [_google_value(value) for value in row]
                    for row in rows
                ],
            ]
            if any(len(row) != len(headers) for row in rows):
                raise ValueError("POS payment data has an unexpected number of columns.")
            self._replace_sheet_range(
                spreadsheet_id,
                target_sheets[sheet_name],
                worksheet,
                values,
                start_row=3,
                start_column=1,
                source_start_row=3,
                source_start_column=1,
                clear_column_count=len(headers),
                logger=logger,
            )
        finally:
            workbook.close()
