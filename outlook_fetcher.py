"""
outlook_fetcher.py
──────────────────
Fetches Roxhill alert emails from Outlook via the Microsoft Graph API.

Authentication:
  Uses OAuth2 client credentials flow (app-only) or delegated access.
  Set credentials in environment variables — never hardcode them.

Required environment variables:
  MS_TENANT_ID      Your Azure AD tenant ID
  MS_CLIENT_ID      App registration client ID
  MS_CLIENT_SECRET  App registration client secret
  MS_USER_EMAIL     The mailbox to read (e.g. lea@activedmc.com)

Roxhill alert emails always come from:
  monitoring-emails@roxhillmedia.com

Usage:
  from outlook_fetcher import OutlookFetcher

  fetcher = OutlookFetcher()
  emails  = fetcher.fetch_roxhill_alerts(lookback_hours=24)

  for email in emails:
      print(email["subject"], email["to_address"])
      # email["body_html"] contains the full HTML body to pass to the parser
"""

from __future__ import annotations

import os
import time
import logging
from datetime import datetime, timedelta, timezone
from typing import Optional

import requests

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

ROXHILL_SENDER   = "monitoring-emails@roxhillmedia.com"
GRAPH_BASE       = "https://graph.microsoft.com/v1.0"
TOKEN_URL_TMPL   = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
GRAPH_SCOPE      = "https://graph.microsoft.com/.default"

# Outlook search API requires OData filter syntax
# We filter by sender address and a received date window
SEARCH_FILTER_TMPL = (
    "from/emailAddress/address eq '{sender}' "
    "and receivedDateTime ge {after}"
)


# ─────────────────────────────────────────────────────────────────────────────
# Token cache (simple in-memory, resets on process restart)
# ─────────────────────────────────────────────────────────────────────────────

class _TokenCache:
    def __init__(self):
        self._token: str = ""
        self._expires_at: float = 0.0

    def get(self) -> str:
        if self._token and time.time() < self._expires_at - 60:
            return self._token
        return ""

    def set(self, token: str, expires_in: int) -> None:
        self._token = token
        self._expires_at = time.time() + expires_in


_cache = _TokenCache()


# ─────────────────────────────────────────────────────────────────────────────
# Fetcher class
# ─────────────────────────────────────────────────────────────────────────────

class OutlookFetcher:
    """
    Fetches Roxhill alert emails from a specific Outlook mailbox.

    Initialisation reads credentials from environment variables.
    Raises EnvironmentError if any required variable is missing.
    """

    def __init__(self):
        self.tenant_id     = self._require_env("MS_TENANT_ID")
        self.client_id     = self._require_env("MS_CLIENT_ID")
        self.client_secret = self._require_env("MS_CLIENT_SECRET")
        self.user_email    = self._require_env("MS_USER_EMAIL")
        self.session       = requests.Session()
        self.session.headers.update({"Content-Type": "application/json"})

    @staticmethod
    def _require_env(key: str) -> str:
        value = os.getenv(key)
        if not value:
            raise EnvironmentError(
                f"Missing required environment variable: {key}\n"
                f"Set it in your .env file or system environment."
            )
        return value

    # ── Authentication ────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        cached = _cache.get()
        if cached:
            return cached

        token_url = TOKEN_URL_TMPL.format(tenant=self.tenant_id)
        resp = self.session.post(token_url, data={
            "grant_type":    "client_credentials",
            "client_id":     self.client_id,
            "client_secret": self.client_secret,
            "scope":         GRAPH_SCOPE,
        })
        resp.raise_for_status()
        data = resp.json()
        _cache.set(data["access_token"], data.get("expires_in", 3600))
        return data["access_token"]

    def _auth_header(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}"}

    # ── Core fetch ────────────────────────────────────────────────────────────

    def fetch_roxhill_alerts(
        self,
        lookback_hours: int = 24,
        client_email_alias: Optional[str] = None,
        max_results: int = 50,
    ) -> list[dict]:
        """
        Fetch all Roxhill alert emails received in the last `lookback_hours`.

        Args:
            lookback_hours:       How far back to look (default 24h).
            client_email_alias:   Optional — filter by a specific To: address
                                  e.g. "uipath@activedmc.com" to only get UiPath alerts.
            max_results:          Maximum emails to return per call.

        Returns:
            List of dicts, each with:
                id              str   Outlook message ID
                subject         str   Email subject line
                to_address      str   First To: recipient address
                received_dt     str   ISO datetime string
                body_html       str   Full HTML body of the email
                body_preview    str   Plain text preview (first ~255 chars)
        """
        after_dt = (
            datetime.now(timezone.utc) - timedelta(hours=lookback_hours)
        ).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Build OData filter
        odata_filter = SEARCH_FILTER_TMPL.format(
            sender=ROXHILL_SENDER,
            after=after_dt,
        )

        url = (
            f"{GRAPH_BASE}/users/{self.user_email}/messages"
            f"?$filter={odata_filter}"
            f"&$select=id,subject,toRecipients,receivedDateTime,body,bodyPreview"
            f"&$top={max_results}"
            f"&$orderby=receivedDateTime desc"
        )

        logger.info(f"Fetching Roxhill alerts after {after_dt}")
        resp = self.session.get(url, headers=self._auth_header())

        if resp.status_code != 200:
            logger.error(f"Graph API error {resp.status_code}: {resp.text[:500]}")
            resp.raise_for_status()

        messages = resp.json().get("value", [])
        logger.info(f"Found {len(messages)} Roxhill alert email(s)")

        results = []
        for msg in messages:
            to_address = ""
            to_recipients = msg.get("toRecipients", [])
            if to_recipients:
                to_address = to_recipients[0].get("emailAddress", {}).get("address", "")

            # If filtering by alias, skip non-matching emails
            if client_email_alias:
                if client_email_alias.lower() not in to_address.lower():
                    continue

            body = msg.get("body", {})
            body_html = body.get("content", "") if body.get("contentType") == "html" else ""

            results.append({
                "id":           msg.get("id", ""),
                "subject":      msg.get("subject", ""),
                "to_address":   to_address,
                "received_dt":  msg.get("receivedDateTime", ""),
                "body_html":    body_html,
                "body_preview": msg.get("bodyPreview", ""),
            })

        return results

    # ── Pagination helper (for large inboxes) ────────────────────────────────

    def fetch_all_pages(self, lookback_hours: int = 24) -> list[dict]:
        """
        Same as fetch_roxhill_alerts but follows @odata.nextLink pagination
        to retrieve all matching emails even if they exceed the Graph API page size.
        Use this when you expect 50+ alerts in a single lookback window.
        """
        after_dt = (
            datetime.now(timezone.utc) - timedelta(hours=lookback_hours)
        ).strftime("%Y-%m-%dT%H:%M:%SZ")

        odata_filter = SEARCH_FILTER_TMPL.format(
            sender=ROXHILL_SENDER,
            after=after_dt,
        )

        url = (
            f"{GRAPH_BASE}/users/{self.user_email}/messages"
            f"?$filter={odata_filter}"
            f"&$select=id,subject,toRecipients,receivedDateTime,body,bodyPreview"
            f"&$top=25"
            f"&$orderby=receivedDateTime desc"
        )

        all_results = []
        page = 0

        while url:
            page += 1
            logger.info(f"Fetching page {page}...")
            resp = self.session.get(url, headers=self._auth_header())
            resp.raise_for_status()
            data = resp.json()

            for msg in data.get("value", []):
                to_address = ""
                to_recipients = msg.get("toRecipients", [])
                if to_recipients:
                    to_address = to_recipients[0].get("emailAddress", {}).get("address", "")

                body = msg.get("body", {})
                body_html = body.get("content", "") if body.get("contentType") == "html" else ""

                all_results.append({
                    "id":           msg.get("id", ""),
                    "subject":      msg.get("subject", ""),
                    "to_address":   to_address,
                    "received_dt":  msg.get("receivedDateTime", ""),
                    "body_html":    body_html,
                    "body_preview": msg.get("bodyPreview", ""),
                })

            url = data.get("@odata.nextLink")

        logger.info(f"Total emails fetched across {page} page(s): {len(all_results)}")
        return all_results
