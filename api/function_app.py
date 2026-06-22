import azure.functions as func
import base64
import datetime
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

# Notification settings
# NOTIFY_FROM_EMAIL : UPN/E-Mail of the mailbox used to send notifications (must have Mail.Send permission)
# NOTIFY_EMAILS     : comma-separated list of recipients for all notifications
NOTIFY_FROM   = os.environ.get("NOTIFY_FROM_EMAIL", "")
NOTIFY_EMAILS = os.environ.get("NOTIFY_EMAILS", "")

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
    "Kundennummer,Firma,Email,Sachbearbeiter,SPUrl,SPUrlCloud,SPUrlMobile,SPUrlAuftrag,Optionen,Erstschulung,"
    "DocSepa,DocEmailRechnung,DocFernwartung,DocAvv,"
    "DocVorlagen,DocDebitoren,DocMitarbeiter,DocLohnarten,"
    "DocVerguetung,DocDatenubernahme,DocPreisliste,LogoUrl,SchulungDurchgefuehrt"
)

# Block A = Pflichtunterlagen (Vertragsunterlagen)
BLOCK_A_FIELDS = ["DocSepa", "DocEmailRechnung", "DocFernwartung", "DocAvv"]
BLOCK_A_IDS    = {"sepa", "email_rechnung", "fernwartung", "avv"}
BLOCK_A_LABELS = {
    "DocSepa":           "SEPA-Mandat",
    "DocEmailRechnung":  "E-Mail Rechnung",
    "DocFernwartung":    "Fernwartungsvereinbarung",
    "DocAvv":            "AVV",
}

# Block B Pflicht = Pflicht-Vorbereitungsunterlagen
# Always required: datenubernahme + vorlagen
# Conditionally required based on Optionen field: verguetung, preisliste
BLOCK_B_PFLICHT_ALWAYS = ["DocDatenubernahme", "DocVorlagen"]
BLOCK_B_PFLICHT_OPTIONAL_MAP = {
    "verguetung": "DocVerguetung",
    "preisliste":  "DocPreisliste",
}
BLOCK_B_PFLICHT_IDS = {"datenubernahme", "vorlagen", "verguetung", "preisliste"}
BLOCK_B_PFLICHT_LABELS = {
    "DocDatenubernahme": "Datenübernahme",
    "DocVorlagen":       "Vorlagen, Logos & Briefbogen",
    "DocVerguetung":     "Vergütungsvereinbarung",
    "DocPreisliste":     "Preisliste",
}

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


# ── Notification helpers ────────────────────────────────────────────────────


def _get_notify_recipients() -> list:
    """Return list of email addresses from NOTIFY_EMAILS env var."""
    return [e.strip() for e in NOTIFY_EMAILS.split(",") if e.strip()]


def _get_all_recipients(fields: dict) -> list:
    """Merge global NOTIFY_EMAILS with per-customer ZusatzEmails (deduped)."""
    base = _get_notify_recipients()
    extra = [e.strip() for e in fields.get("ZusatzEmails", "").split(",") if e.strip()]
    seen = set()
    result = []
    for addr in base + extra:
        if addr.lower() not in seen:
            seen.add(addr.lower())
            result.append(addr)
    return result


def send_email(subject: str, body: str, recipients: list) -> None:
    """Send a plain-text email via Microsoft Graph using the NOTIFY_FROM mailbox."""
    if not recipients or not NOTIFY_FROM:
        return
    token = get_app_token()
    payload = json.dumps({
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": r}} for r in recipients],
        },
        "saveToSentItems": False,
    }).encode()
    url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(NOTIFY_FROM)}/sendMail"
    req = urllib.request.Request(url, data=payload, method="POST", headers={
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            resp.read()
    except Exception:
        pass  # Don't let notification failure break the main request


def _block_b_pflicht_fields(optionen: str) -> list:
    """Return the Block-B-Pflicht SP field names relevant for this customer."""
    opts = [o.strip() for o in optionen.split(",") if o.strip() and not o.startswith("!")]
    fields = list(BLOCK_B_PFLICHT_ALWAYS)
    for opt_id, sp_field in BLOCK_B_PFLICHT_OPTIONAL_MAP.items():
        if opt_id in opts:
            fields.append(sp_field)
    return fields


def _all_pflicht_complete(fields: dict) -> bool:
    """Return True when Block A AND Block B Pflicht are all done."""
    if not all(fields.get(f, False) for f in BLOCK_A_FIELDS):
        return False
    b_fields = _block_b_pflicht_fields(fields.get("Optionen", ""))
    return all(fields.get(f, False) for f in b_fields)


def send_completion_email(item_id: str) -> None:
    """Send notification when all Pflichtunterlagen (A + B Pflicht) are complete."""
    if not NOTIFY_FROM:
        return
    try:
        fields = sp_get_item(item_id)
        if not _all_pflicht_complete(fields):
            return  # not yet fully complete
        recipients = _get_all_recipients(fields)
        if not recipients:
            return
        subject = f"✅ Onboarding: Alle Pflichtunterlagen vollständig – {kundennummer} {firma}".strip()
        body = (
            f"Alle Pflichtunterlagen (Vertragsunterlagen und Vorbereitungsunterlagen) "
            f"für den Kunden {firma} (Kundennummer: {kundennummer}) "
            f"wurden vollständig hochgeladen und stehen im Portal bereit.\n\n"
            f"Zuständiger Betreuer: {sachbearbeiter}\n\n"
            f"Bitte prüfen Sie das Onboarding-Portal für weitere Details."
        )
        send_email(subject, body, recipients)
    except Exception:
        pass


def sp_list_all_customers() -> list:
    """Fetch all customer items from the SharePoint list."""
    token = get_app_token()
    fields = (
        "id,Kundennummer,Firma,Email,Sachbearbeiter,Erstschulung,Optionen,SchulungDurchgefuehrt,"
        "DocSepa,DocEmailRechnung,DocFernwartung,DocAvv,"
        "DocDatenubernahme,DocVorlagen,DocVerguetung,DocPreisliste,ZusatzEmails"
    )
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}"
        f"/lists/{LIST_ID}/items"
        f"?$expand=fields($select={fields})&$top=500"
    )
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    return data.get("value", [])


def business_days_until(target_date: datetime.date) -> int:
    """Count business days (Mon–Fri) from today up to (not including) target_date."""
    today = datetime.date.today()
    if target_date <= today:
        return 0
    days = 0
    current = today
    while current < target_date:
        if current.weekday() < 5:  # Mon=0 … Fri=4
            days += 1
        current += datetime.timedelta(days=1)
    return days


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
                    "docs":                docs,
                    "logoUrl":             fields.get("LogoUrl", ""),
                    "schulungDurchgefuehrt": bool(fields.get("SchulungDurchgefuehrt", False)),
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
    action    = str(body.get("action",    "")).strip()

    # ── Special action: mark training as completed ────────────────────────
    if action == "schulung-abgeschlossen" and cust_id:
        try:
            sp_patch(cust_id, "SchulungDurchgefuehrt", True)
            return func.HttpResponse(
                json.dumps({"ok": True}),
                status_code=200, headers=CORS_HEADERS
            )
        except Exception as exc:
            return func.HttpResponse(
                json.dumps({"error": str(exc)}),
                status_code=500, headers=CORS_HEADERS
            )

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
        # If a Pflicht doc was just marked complete, check if everything is now done → notify
        if value and doc_id in (BLOCK_A_IDS | BLOCK_B_PFLICHT_IDS):
            send_completion_email(cust_id)
        return func.HttpResponse(
            json.dumps({"ok": True, "uploaded": uploaded, "uploadError": upload_error}),
            status_code=200, headers=CORS_HEADERS
        )
    except Exception as exc:
        return func.HttpResponse(
            json.dumps({"error": str(exc), "uploaded": uploaded}),
            status_code=500, headers=CORS_HEADERS
        )


# ── Daily deadline check ─────────────────────────────────────────────────────
# Runs every day at 07:00 UTC.
# Sends a warning when Block A is incomplete and the training date is
# exactly 7, 3 or 1 business day(s) away.

@app.timer_trigger(schedule="0 0 7 * * *", arg_name="timer", run_on_startup=False)
def check_deadline_notifications(timer: func.TimerRequest) -> None:
    if not NOTIFY_FROM:
        return
    global_recipients = _get_notify_recipients()

    try:
        customers = sp_list_all_customers()
    except Exception:
        return

    for item in customers:
        fields = item.get("fields", {})
        erstschulung_str = fields.get("Erstschulung", "")
        if not erstschulung_str:
            continue

        # Skip if all Pflichtunterlagen (A + B Pflicht) are already complete
        if _all_pflicht_complete(fields):
            continue

        # Parse the training date (SharePoint returns ISO date or datetime)
        try:
            schulung_date = datetime.date.fromisoformat(erstschulung_str[:10])
        except Exception:
            continue

        bdays = business_days_until(schulung_date)
        if bdays not in (7, 3, 1):
            continue  # only notify at these checkpoints

        kundennummer   = fields.get("Kundennummer",   "")
        firma          = fields.get("Firma",           "")
        sachbearbeiter = fields.get("Sachbearbeiter", "")
        recipients = _get_all_recipients(fields) or global_recipients
        if not recipients:
            continue

        # Collect all missing Pflicht docs (Block A + Block B Pflicht)
        missing = []
        for f in BLOCK_A_FIELDS:
            if not fields.get(f, False):
                missing.append(BLOCK_A_LABELS[f])
        for f in _block_b_pflicht_fields(fields.get("Optionen", "")):
            if not fields.get(f, False):
                missing.append(BLOCK_B_PFLICHT_LABELS.get(f, f))

        days_label    = f"{bdays} Werktag{'e' if bdays != 1 else ''}"
        missing_lines = "".join(f"  • {m}\n" for m in missing)

        # ── Internal notification to Komda staff ────────────────────────────────────────────
        internal_subject = (
            f"⚠️ Onboarding: Pflichtunterlagen fehlen – "
            f"{schulung_date.strftime('%d.%m.%Y')} – {kundennummer} {firma}".strip()
        )
        internal_body = (
            f"Der Schulungstermin für {firma} (Kundennummer: {kundennummer}) "
            f"ist am {schulung_date.strftime('%d.%m.%Y')} – noch {days_label}.\n\n"
            "Folgende Pflichtunterlagen wurden noch nicht hochgeladen:\n"
            + missing_lines
            + f"\nZuständig: {sachbearbeiter}\n\n"
            "Bitte nehmen Sie Kontakt mit dem Kunden auf."
        )
        send_email(internal_subject, internal_body, recipients)

        # ── Customer reminder email ─────────────────────────────────────────────────────
        customer_email = fields.get("Email", "").strip()
        if customer_email and NOTIFY_FROM:
            customer_subject = (
                f"Erinnerung: Ihr Komda® Onboarding – "
                f"Schulung am {schulung_date.strftime('%d.%m.%Y')}"
            )
            customer_body = (
                f"Sehr geehrte Damen und Herren,\n\n"
                f"Ihr Schulungstermin bei Komda® Software ist am "
                f"{schulung_date.strftime('%d.%m.%Y')} – noch {days_label}.\n\n"
                f"Damit wir den Termin optimal vorbereiten können, benötigen wir "
                f"noch folgende Unterlagen von Ihnen:\n"
                + missing_lines
                + "\nBitte laden Sie diese Dokumente über Ihr persönliches "
                "Onboarding-Portal hoch.\n\n"
                "Bei Fragen steht Ihnen Ihr Betreuer gerne zur Verfügung.\n\n"
                "Mit freundlichen Grüßen\n"
                "Ihr Komda® Software Team"
            )
            send_email(customer_subject, customer_body, [customer_email])
