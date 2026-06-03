"""
Microsoft Graph API email fetcher for Roxhill alerts.

Fetches emails from a configured mailbox, filtering by sender
(roxhillmedia.com) and date window.

Requires Azure AD app registration with Mail.Read permission.
"""

import os
import requests
from datetime import datetime, timedelta
from rich.console import Console

console = Console()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL  = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"


def _get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = TOKEN_URL.format(tenant_id=tenant_id)
    data = {
        "grant_type":    "client_credentials",
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def fetch_roxhill_emails(
    hours: int = 24,
    user_email: str | None = None,
    tenant_id: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
) -> list[dict]:
    """
    Fetch Roxhill alert emails from the last `hours` hours.

    Returns list of dicts: {subject, body_html, received_at, sender}

    Credentials fall back to environment variables:
      AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, OUTLOOK_USER_EMAIL
    """
    tenant_id     = tenant_id     or os.environ.get("AZURE_TENANT_ID", "")
    client_id     = client_id     or os.environ.get("AZURE_CLIENT_ID", "")
    client_secret = client_secret or os.environ.get("AZURE_CLIENT_SECRET", "")
    user_email    = user_email    or os.environ.get("OUTLOOK_USER_EMAIL", "")

    if not all([tenant_id, client_id, client_secret, user_email]):
        console.print("[red]Missing Azure/Outlook credentials. "
                      "Set AZURE_TENANT_ID, AZURE_CLIENT_ID, "
                      "AZURE_CLIENT_SECRET, OUTLOOK_USER_EMAIL.[/red]")
        return []

    token = _get_access_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token}"}

    since = (datetime.utcnow() - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")

    filter_str = (
        f"receivedDateTime ge {since} and "
        f"contains(from/emailAddress/address, 'roxhillmedia.com')"
    )

    url = (
        f"{GRAPH_BASE}/users/{user_email}/messages"
        f"?$filter={filter_str}"
        f"&$select=subject,body,receivedDateTime,from"
        f"&$top=50"
        f"&$orderby=receivedDateTime desc"
    )

    console.print(f"[cyan]Fetching Roxhill emails from last {hours}h...[/cyan]")

    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    emails = []
    for msg in data.get("value", []):
        emails.append({
            "subject":     msg.get("subject", ""),
            "body_html":   msg.get("body", {}).get("content", ""),
            "received_at": msg.get("receivedDateTime", ""),
            "sender":      msg.get("from", {}).get("emailAddress", {}).get("address", ""),
        })

    console.print(f"[green]Found {len(emails)} Roxhill email(s)[/green]")
    return emails
