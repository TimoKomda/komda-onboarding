import azure.functions as func
import base64
import json
import os
import re
import urllib.request
import urllib.parse
import urllib.error

app = func.FunctionApp()

TENANT_ID     = os.environ.get("TENANT_ID", "")
CLIENT_ID     = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
SITE_ID       = os.environ.get("SITE_ID", "")
LIST_ID       = os.environ.get("LIST_ID", "")
UPDATE_SECRET = os.environ.get("UPDATE_SECRET", "")

DOC_FIELD = {
    "sepa":           "DocSepa",
    "email_rechnung": "DocEmailRechnung",
    "fernwartung":    "DocFernwartung",
    "avv":            "DocAvv",
    "vorlagen":       "DocVorlagen",
    "debitoren":      "DocDebitoren",
    "mitarbeiter":    "DocMitarbeiter",
    "lohnarten":      "DocLohnarten",
    "verguetung":     "DocVerguetung",
    "datenubernahme": "DocDatenubernahme",
    "preisliste":     "DocPreisliste",
}

GET_SELECT_FIELDS = (
    "Kundennummer,Sachbearbeiter,SPUrl,SPUrlCloud,SPUrlMobile,SPUrlAuftrag,Optionen,Erstschulung,"
    "DocSepa,DocEmailRechnung,DocFernwartung,DocAvv,"
    "DocVorlagen,DocDebitoren,DocMitarbeiter,DocLohnarten,"
    "DocVerguetung,DocDatenubernahme,DocPreisliste,LogoUrl"
)

CORS_HEADERS = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
}


def get_app_token() -> str:
    data = urllib.parse.urlencode({
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }).encode()
    req = urllib.request.Request(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data=data, method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())["access_token"]


def sp_get_item(item_id: str) -> dict:
    token = get_app_token()
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
        f"/lists/{LIST_ID}/items/{item_id}/fields"
        f"?$select={GET_SELECT_FIELDS}"
    )
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def sp_patch(item_id: str, field: str, value: bool):
    token = get_app_token()
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
        f"/lists/{LIST_ID}/items/{item_id}/fields"
    )
    payload = json.dumps({field: value}).encode()
    req = urllib.request.Request(url, data=payload, method="PATCH", headers={
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    })
    with urllib.request.urlopen(req) as resp:
        resp.read()


def _is_sharing_link(sp_url: str) -> bool:
    """Detect SharePoint sharing links (/:f:/, /:b:/, /:fl:/ etc.)"""
    return bool(re.search(r'/:[a-z]+:/', sp_url))


def sp_list_subfolder(sp_url: str, subfolder_name: str) -> list:
    """
    List files in a named subfolder of the SharePoint folder identified by sp_url.
    Returns [{name, size, downloadUrl}, ...].
    """
    app_token = get_app_token()
    headers = {"Authorization": f"Bearer {app_token}"}

    if _is_sharing_link(sp_url):
        share_token = _encode_sharing_token(sp_url)
        resolve_url = (
            f"https://graph.microsoft.com/v1.0/shares/{share_token}"
            f"/driveItem?$select=id,parentReference"
        )
        req = urllib.request.Request(resolve_url, headers=headers)
        with urllib.request.urlopen(req) as resp:
            item = json.loads(resp.read())
        drive_id = item["parentReference"]["driveId"]
        parent_id = item["id"]
        children_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
            f"/items/{parent_id}:/{urllib.parse.quote(subfolder_name, safe='')}:/children"
            f"?$select=name,size,file,@microsoft.graph.downloadUrl"
        )
    else:
        parsed = urllib.parse.urlparse(sp_url)
        path = urllib.parse.unquote(parsed.path)
        m = re.match(r"^/sites/[^/]+/[^/]+/(.+)$", path)
        if m:
            folder_path = m.group(1)
        else:
            m2 = re.match(r"^/[^/]+/[^/]+/(.+)$", path)
            folder_path = m2.group(1) if m2 else path.lstrip("/")
        children_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
            f"/drive/root:/{urllib.parse.quote(folder_path, safe='/')}"
            f"/{urllib.parse.quote(subfolder_name, safe='')}:/children"
            f"?$select=name,size,file,@microsoft.graph.downloadUrl"
        )

    req = urllib.request.Request(children_url, headers=headers)
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())

    files = []
    for entry in data.get("value", []):
        if "file" not in entry:
            continue  # skip folders
        files.append({
            "name":        entry.get("name", ""),
            "size":        entry.get("size", 0),
            "downloadUrl": entry.get("@microsoft.graph.downloadUrl", ""),
        })
    return files


def _encode_sharing_token(url: str) -> str:
    """Encode a sharing URL as a Graph API shares token (u!<base64>)."""
    b64 = base64.b64encode(url.encode('utf-8')).decode()
    return 'u!' + b64.rstrip('=').replace('/', '_').replace('+', '-')


def sp_upload_via_sharing_link(sp_url: str, filename: str, file_bytes: bytes, subfolder: str = "") -> tuple:
    """
    Upload to a SharePoint folder identified by a sharing link.
    Uses /v1.0/shares/{token}/driveItem to resolve the folder, then uploads.
    """
    safe_name  = re.sub(r'[<>:"/\\|?*]', '_', filename)
    app_token  = get_app_token()
    share_token = _encode_sharing_token(sp_url)

    # Resolve sharing link → drive item
    resolve_url = f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem?$select=id,parentReference"
    req = urllib.request.Request(resolve_url, headers={"Authorization": f"Bearer {app_token}"})
    with urllib.request.urlopen(req) as resp:
        item = json.loads(resp.read())

    drive_id = item['parentReference']['driveId']
    item_id  = item['id']

    if subfolder:
        safe_sub = urllib.parse.quote(subfolder, safe='')
        upload_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
            f"/items/{item_id}:/{safe_sub}/{urllib.parse.quote(safe_name)}:/content"
        )
    else:
        upload_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
            f"/items/{item_id}:/{urllib.parse.quote(safe_name)}:/content"
        )
    req = urllib.request.Request(
        upload_url, data=file_bytes, method="PUT",
        headers={"Authorization": f"Bearer {app_token}", "Content-Type": "application/octet-stream"},
    )
    with urllib.request.urlopen(req) as resp:
        resp.read()
    return True, f"driveId={drive_id} item={item_id} sub={subfolder} file={safe_name}"


def sp_upload_via_path(sp_url: str, filename: str, file_bytes: bytes, subfolder: str = "") -> tuple:
    """
    Upload to a SharePoint folder identified by a direct URL.
    e.g. https://tenant.sharepoint.com/sites/ZSH/Shared Documents/Onboarding/Kunde
    """
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', filename)
    parsed    = urllib.parse.urlparse(sp_url)
    path      = urllib.parse.unquote(parsed.path)

    # Strip /sites/{site}/{library}/ prefix → relative folder path
    m = re.match(r'^/sites/[^/]+/[^/]+/(.+)$', path)
    if m:
        folder_path = m.group(1)
    else:
        m2 = re.match(r'^/[^/]+/[^/]+/(.+)$', path)
        folder_path = m2.group(1) if m2 else path.lstrip('/')

    if subfolder:
        folder_path = folder_path.rstrip('/') + '/' + subfolder

    app_token  = get_app_token()
    upload_url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
        f"/drive/root:/{urllib.parse.quote(folder_path, safe='/')}/{urllib.parse.quote(safe_name)}:/content"
    )
    req = urllib.request.Request(
        upload_url, data=file_bytes, method="PUT",
        headers={"Authorization": f"Bearer {app_token}", "Content-Type": "application/octet-stream"},
    )
    with urllib.request.urlopen(req) as resp:
        resp.read()
    return True, f"path={folder_path}/{safe_name}"


def sp_upload_file(sp_url: str, filename: str, file_bytes: bytes, subfolder: str = "") -> tuple:
    """
    Upload a file to a SharePoint folder (optionally into a named subfolder).
    Automatically detects sharing links (/:f:/ etc.) vs. direct folder URLs.
    Returns (success: bool, debug_message: str).
    """
    try:
        if _is_sharing_link(sp_url):
            return sp_upload_via_sharing_link(sp_url, filename, file_bytes, subfolder)
        else:
            return sp_upload_via_path(sp_url, filename, file_bytes, subfolder)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return False, f"HTTP {e.code}: {body[:300]}"
    except Exception as ex:
        return False, str(ex)


def decode_token(token: str) -> str:
    padded = token + "=" * (4 - len(token) % 4 if len(token) % 4 else 0)
    return base64.b64decode(padded).decode("utf-8")


@app.route(route="status", methods=["GET", "POST", "OPTIONS"],
           auth_level=func.AuthLevel.ANONYMOUS)
def update_status(req: func.HttpRequest) -> func.HttpResponse:

    if req.method == "OPTIONS":
        return func.HttpResponse("", status_code=200, headers=CORS_HEADERS)

    if req.method == "GET":
        token_param = req.params.get("token", "").strip()
        action      = req.params.get("action", "").strip()

        if not token_param:
            return func.HttpResponse(
                json.dumps({"ok": True, "service": "komda-onboarding"}),
                status_code=200, headers=CORS_HEADERS
            )

        # ── Action: list files in a subfolder ───────────────────────────────
        if action == "list-folder":
            folder_name = req.params.get("folder", "Datenübernahme").strip()
            try:
                item_id = decode_token(token_param)
                fields  = sp_get_item(item_id)
                sp_url  = fields.get("SPUrl", "")
                if not sp_url:
                    return func.HttpResponse(
                        json.dumps({"ok": False, "error": "Kein SharePoint-Ordner konfiguriert"}),
                        status_code=200, headers=CORS_HEADERS
                    )
                files = sp_list_subfolder(sp_url, folder_name)
                return func.HttpResponse(
                    json.dumps({"ok": True, "files": files}),
                    status_code=200, headers=CORS_HEADERS
                )
            except Exception as exc:
                return func.HttpResponse(
                    json.dumps({"ok": False, "error": str(exc)}),
                    status_code=200, headers=CORS_HEADERS
                )

        try:
            item_id = decode_token(token_param)
            fields  = sp_get_item(item_id)
            # Build doc-status map so the client can sync across devices
            docs = {doc_id: bool(fields.get(sp_field, False))
                    for doc_id, sp_field in DOC_FIELD.items()}
            return func.HttpResponse(
                json.dumps({
                    "ok":             True,
                    "kundennummer":   fields.get("Kundennummer",   ""),
                    "sachbearbeiter": fields.get("Sachbearbeiter", ""),
                    "spUrl":          fields.get("SPUrl",          ""),
                    "spUrlCloud":     fields.get("SPUrlCloud",     ""),
                    "spUrlMobile":    fields.get("SPUrlMobile",    ""),
                    "spUrlAuftrag":   fields.get("SPUrlAuftrag",   ""),
                    "optionen":       fields.get("Optionen",       ""),
                    "erstschulung":   fields.get("Erstschulung",   ""),
                    "docs":           docs,
                    "logoUrl":        fields.get("LogoUrl", ""),
                }),
                status_code=200, headers=CORS_HEADERS
            )
        except Exception as exc:
            return func.HttpResponse(
                json.dumps({"error": str(exc)}),
                status_code=500, headers=CORS_HEADERS
            )

    # ── POST ──────────────────────────────────────────────────────────────
    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse(
            json.dumps({"error": "Ungültiges JSON"}),
            status_code=400, headers=CORS_HEADERS
        )

    if body.get("secret") != UPDATE_SECRET:
        return func.HttpResponse(
            json.dumps({"error": "Nicht autorisiert"}),
            status_code=401, headers=CORS_HEADERS
        )

    cust_id   = str(body.get("custId",    "")).strip()
    doc_id    = str(body.get("docId",     "")).strip()
    value     = bool(body.get("value",    False))
    file_b64  = str(body.get("file",      "")).strip()
    filename  = str(body.get("filename",  "")).strip()
    sp_url    = str(body.get("spUrl",     "")).strip()
    subfolder = str(body.get("subfolder", "")).strip()
    field     = DOC_FIELD.get(doc_id)

    if not field or not cust_id:
        return func.HttpResponse(
            json.dumps({"error": "Ungültige Parameter"}),
            status_code=400, headers=CORS_HEADERS
        )

    # Optional: upload signed PDF to SharePoint folder
    uploaded      = False
    upload_error  = ""
    if file_b64 and filename and sp_url:
        try:
            file_bytes = base64.b64decode(file_b64)
            uploaded, upload_error = sp_upload_file(sp_url, filename, file_bytes, subfolder)
        except Exception as ex:
            upload_error = str(ex)

    # Always update the boolean status field
    try:
        sp_patch(cust_id, field, value)
        return func.HttpResponse(
            json.dumps({"ok": True, "uploaded": uploaded, "uploadError": upload_error}),
            status_code=200, headers=CORS_HEADERS
        )
    except Exception as exc:
        return func.HttpResponse(
            json.dumps({"error": str(exc), "uploaded": uploaded}),
            status_code=500, headers=CORS_HEADERS
        )
