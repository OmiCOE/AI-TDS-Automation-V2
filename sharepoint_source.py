"""
sharepoint_source.py
=====================
Lets the TDS tool use a SharePoint / OneDrive folder as the TDS database,
via Microsoft Graph. Only the PDFs that are actually needed are downloaded.

How the folder is located
-------------------------
You paste the folder's URL (the address-bar URL while browsing the folder in
SharePoint, or a "Copy link" sharing link). Graph's /shares endpoint resolves
either form to a drive + folder id. The folder is then listed recursively.

Authentication (no Azure subscription, no cost)
-----------------------------------------------
Uses a free Entra ID app registration as a *public client* with the
"device code" sign-in: the person using the web app signs in with their own
Microsoft 365 account and only sees what that account can already see.
Delegated scopes: Files.Read.All, Sites.Read.All.
"""

import base64
import re
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import requests

try:
    import msal
except ImportError:  # only needed for sign-in, not for the provider class itself
    msal = None

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Files.Read.All", "Sites.Read.All"]


# --------------------------------------------------------------------------
# Sign-in (device code flow)
# --------------------------------------------------------------------------

def start_device_login(client_id, tenant_id="organizations"):
    """Returns (msal_app, flow). Show flow['user_code'] and flow['verification_uri']
    to the user, then call finish_device_login(app, flow)."""
    if msal is None:
        raise RuntimeError("pip install msal")
    app = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start sign-in: {flow.get('error_description', flow)}")
    return app, flow


def finish_device_login(app, flow):
    """Blocks until the user completes sign-in (or the code expires). Returns access token."""
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Sign-in failed: {result.get('error_description', result)}")
    return result["access_token"]


# --------------------------------------------------------------------------
# Provider
# --------------------------------------------------------------------------

def encode_share_url(url):
    """Graph 'shares' id for any SharePoint/OneDrive URL: 'u!' + base64url(url) without padding."""
    b64 = base64.urlsafe_b64encode(url.strip().encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + b64


class SharePointTDSProvider:
    """Drop-in replacement for LocalTDSProvider: .find(model) -> local Path | None, .count."""

    def __init__(self, access_token, folder_url, cache_dir=None, session=None, progress_cb=None):
        self.session = session or requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {access_token}"})
        self.cache_dir = Path(cache_dir or tempfile.mkdtemp(prefix="tds_cache_"))
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.files = []        # every PDF: dict(drive, id, name, folder, size)
        self.index = {}        # exact filename stem -> (drive_id, item_id, name)
        self.duplicates = []
        self._downloaded = {}  # stem -> local Path

        root = self._get(f"{GRAPH}/shares/{encode_share_url(folder_url)}/driveItem")
        if "folder" not in root:
            raise ValueError("That URL points to a file, not a folder. Paste the TDS folder's URL.")
        self.drive_id = root["parentReference"]["driveId"]
        self._walk(root["id"], progress_cb)
        self.total_mb = sum(f["size"] for f in self.files) / 1e6

    # -- HTTP helpers ------------------------------------------------------
    def _get(self, url, **kw):
        r = self.session.get(url, timeout=60, **kw)
        if r.status_code == 401:
            raise PermissionError("Microsoft sign-in expired. Please sign in again.")
        if r.status_code == 403:
            raise PermissionError("Your account doesn't have access to that folder (HTTP 403).")
        if r.status_code == 404:
            raise FileNotFoundError("Folder not found. Check the URL and your access.")
        r.raise_for_status()
        return r.json()

    # -- recursive listing -------------------------------------------------
    def _walk(self, folder_id, progress_cb=None):
        stack = [(folder_id, "")]
        while stack:
            fid, rel = stack.pop()
            url = (f"{GRAPH}/drives/{self.drive_id}/items/{fid}/children"
                   "?$top=200&$select=id,name,folder,file,size")
            while url:
                data = self._get(url)
                for it in data.get("value", []):
                    if "folder" in it:
                        stack.append((it["id"], (rel + " " + it["name"]).strip()))
                    elif it["name"].lower().endswith(".pdf"):
                        stem = it["name"].rsplit(".", 1)[0].strip()
                        self.files.append({"drive": self.drive_id, "id": it["id"], "name": it["name"],
                                           "folder": rel, "size": it.get("size", 0) or 0})
                        if stem in self.index:
                            self.duplicates.append(stem)
                        else:
                            self.index[stem] = (self.drive_id, it["id"], it["name"])
                url = data.get("@odata.nextLink")
            if progress_cb:
                progress_cb(f"Indexed {len(self.files)} PDFs...")

    def _download(self, drive_id, item_id, name):
        safe = "".join(c if c.isalnum() or c in " ._-()" else "_" for c in name)
        dest = self.cache_dir / (re.sub(r"[^A-Za-z0-9]", "", item_id)[-12:] + "_" + safe)
        if not dest.exists():
            r = self.session.get(f"{GRAPH}/drives/{drive_id}/items/{item_id}/content",
                                 timeout=180, allow_redirects=True)
            r.raise_for_status()
            dest.write_bytes(r.content)
        return dest

    def build_library(self, progress_cb=None, workers=8):
        """Download every TDS (parallel) and read its text -> tds_matcher.TDSLibrary."""
        from tds_matcher import TDSLibrary
        lib = TDSLibrary()
        done = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = [(f, ex.submit(self._download, f["drive"], f["id"], f["name"])) for f in self.files]
            for f, fut in futs:
                try:
                    path = fut.result()
                except Exception:
                    continue          # unreadable file: skip, don't kill the whole run
                lib.add(path, f["name"], f["folder"])
                done += 1
                if progress_cb and done % 5 == 0:
                    progress_cb(f"Reading TDS files... {done}/{len(self.files)}")
        lib.finalize()
        return lib

    # -- provider interface -----------------------------------------------
    @property
    def count(self):
        return len(self.index)

    def find(self, model_value):
        model_value = (model_value or "").strip()
        if not model_value or model_value not in self.index:
            return None
        if model_value in self._downloaded:
            return self._downloaded[model_value]
        drive_id, item_id, name = self.index[model_value]
        dest = self._download(drive_id, item_id, name)
        self._downloaded[model_value] = dest
        return dest
