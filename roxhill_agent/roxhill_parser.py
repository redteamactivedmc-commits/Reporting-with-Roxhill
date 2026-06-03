"""
Roxhill HTML parser + URL decoder.

Parses Roxhill alert email HTML to extract:
  - Article URLs (decoded from tracking wrappers)
  - Headlines
  - Publication names
  - Dates
  - Tier classification
  - Language detection (English/Arabic)
"""

import re
from urllib.parse import unquote
from bs4 import BeautifulSoup
from datetime import datetime


CLIENT_ALIAS_MAP = {
    "uipath":        "UiPath",
    "scc":           "SCC Middle East",
    "illumio":       "Illumio",
    "denodo":        "Denodo",
    "phosphorus":    "Phosphorus Cybersecurity",
    "knowbe4":       "KnowBe4",
    "netscout":      "NETSCOUT",
    "dynatrace":     "Dynatrace",
    "qlik":          "Qlik",
    "commscope":     "CommScope",
    "heidrick":      "Heidrick & Struggles",
    "levelinfinite": "Level Infinite",
}

TIER_A_PUBLICATIONS = {
    "arabian business", "gulf news", "khaleej times", "arab news",
    "the national", "zawya", "al arabiya", "cnbc arabia",
    "forbes middle east", "itp.net", "tahawultech", "cxo insight me",
    "security advisor middle east", "computerweekly",
}

TIER_B_PUBLICATIONS = {
    "securitymea", "techpulsemea", "meatechwatch", "techafricanews",
    "gulftech-news", "datacentremagazine", "intelligentcio",
    "enterpriseit", "channelpostmea",
}

ARABIC_DOMAINS = {
    "alriyadh.com", "aleqt.com", "al-jazirah.com", "okaz.com.sa",
    "aawsat.com", "albayan.ae", "alittihad.ae", "emaratalyoum.com",
    "annahar.com", "alrai.com", "alwatan.com.sa",
}


def resolve_client_name(alias: str) -> str:
    return CLIENT_ALIAS_MAP.get(alias.lower().strip(), alias)


def decode_roxhill_url(tracking_url: str) -> str:
    """
    Roxhill wraps article URLs in tracking redirects:
      https://email-tracking.pimlico.roxhillmedia.com/CL0/https:%2F%2Fwww.example.com%2Farticle/1/<hash>
    Extract and decode the real URL.
    """
    match = re.search(r"/CL0/(https?[^/]+)/", tracking_url)
    if match:
        return unquote(match.group(1))

    match = re.search(r"/CL0/(https?.*?)(?:/\d+/|$)", tracking_url)
    if match:
        return unquote(match.group(1))

    if tracking_url.startswith("http") and "roxhillmedia.com" not in tracking_url:
        return tracking_url

    return tracking_url


def detect_language(url: str, headline: str = "") -> str:
    from urllib.parse import urlparse
    domain = urlparse(url).netloc.lower()
    for arabic_domain in ARABIC_DOMAINS:
        if arabic_domain in domain:
            return "Arabic"

    if headline and any("؀" <= c <= "ۿ" for c in headline):
        return "Arabic"

    return "English"


def classify_tier(publication: str, url: str = "") -> str:
    pub_lower = publication.lower().strip()
    from urllib.parse import urlparse
    domain = urlparse(url).netloc.lower().replace("www.", "") if url else ""

    if pub_lower in TIER_A_PUBLICATIONS or domain in TIER_A_PUBLICATIONS:
        return "Tier A"

    for name in TIER_A_PUBLICATIONS:
        if name in pub_lower or name in domain:
            return "Tier A"

    for name in TIER_B_PUBLICATIONS:
        if name in pub_lower or name in domain:
            return "Tier B"

    return "Tier C"


def _extract_publication_from_url(url: str) -> str:
    from urllib.parse import urlparse
    domain = urlparse(url).netloc.lower()
    domain = domain.replace("www.", "")
    parts = domain.split(".")
    if parts:
        name = parts[0]
        name = re.sub(r"[-_]", " ", name)
        return name.title()
    return domain


def parse_roxhill_email_html(html_body: str) -> list[dict]:
    """
    Parse a Roxhill alert email HTML body and extract all articles.

    Returns list of dicts with keys:
      headline, publication, url, date, language, tier, print_online
    """
    soup = BeautifulSoup(html_body, "html.parser")
    articles = []

    rows = soup.find_all("tr")
    for row in rows:
        links = row.find_all("a", href=True)
        for link in links:
            href = link.get("href", "")
            if not href or "unsubscribe" in href.lower() or "mailto:" in href.lower():
                continue

            real_url = decode_roxhill_url(href)
            if "roxhillmedia.com" in real_url:
                continue

            headline_text = link.get_text(strip=True)
            if not headline_text or len(headline_text) < 10:
                parent_td = link.find_parent("td")
                if parent_td:
                    headline_text = parent_td.get_text(strip=True)

            if not headline_text or len(headline_text) < 10:
                continue

            pub_cells = row.find_all("td")
            publication = ""
            for cell in pub_cells:
                cell_text = cell.get_text(strip=True)
                if cell_text and cell_text != headline_text and len(cell_text) < 60:
                    publication = cell_text
                    break

            if not publication:
                publication = _extract_publication_from_url(real_url)

            language = detect_language(real_url, headline_text)
            tier = classify_tier(publication, real_url)

            articles.append({
                "headline":     headline_text,
                "publication":  publication,
                "url":          real_url,
                "date":         datetime.today().strftime("%d/%m/%Y"),
                "language":     language,
                "tier":         tier,
                "print_online": "Online",
            })

    seen_urls = set()
    unique = []
    for art in articles:
        if art["url"] not in seen_urls:
            seen_urls.add(art["url"])
            unique.append(art)

    return unique


def parse_direct_urls(urls: list[str], client: str = "") -> list[dict]:
    """
    Parse a list of direct article URLs (not from email HTML).
    Fetches each page to extract headline and publication info.
    """
    import requests

    articles = []
    for url in urls:
        url = url.strip()
        if not url:
            continue

        headline = ""
        publication = _extract_publication_from_url(url)

        try:
            resp = requests.get(url, timeout=15, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/121.0.0.0 Safari/537.36"
            })
            if resp.ok:
                soup = BeautifulSoup(resp.text, "html.parser")
                title_tag = soup.find("title")
                if title_tag:
                    headline = title_tag.get_text(strip=True)

                og_title = soup.find("meta", property="og:title")
                if og_title and og_title.get("content"):
                    headline = og_title["content"]

                og_site = soup.find("meta", property="og:site_name")
                if og_site and og_site.get("content"):
                    publication = og_site["content"]

        except Exception:
            if not headline:
                path = url.rstrip("/").split("/")[-1]
                headline = re.sub(r"[-_]", " ", path).title()

        date_match = re.search(r"/(\d{4})/(\d{2})/(\d{2})/", url)
        if date_match:
            y, m, d = date_match.groups()
            date_str = f"{d}/{m}/{y}"
        else:
            date_str = datetime.today().strftime("%d/%m/%Y")

        language = detect_language(url, headline)
        tier = classify_tier(publication, url)

        articles.append({
            "headline":     headline,
            "publication":  publication,
            "url":          url,
            "date":         date_str,
            "language":     language,
            "tier":         tier,
            "print_online": "Online",
        })

    return articles
