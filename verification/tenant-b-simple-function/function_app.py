import json
import logging
import os
import time
from typing import Any, Dict, List, Optional

import azure.functions as func
import msal
import requests
from azure.identity import ManagedIdentityCredential

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"


def _get_env(name: str, required: bool = True, default: Optional[str] = None) -> str:
    value = os.getenv(name, default)
    if required and (value is None or value.strip() == ""):
        raise ValueError(f"Missing required environment variable: {name}")
    return (value or "").strip()


def _get_graph_token() -> str:
    auth_mode = os.getenv("GRAPH_AUTH_MODE", "client_secret").strip().lower()

    if auth_mode == "managed_identity":
        # ⚠️ クロステナントシナリオ（テナントB → テナントA）では動作しません。
        # テナントB の MI が取得するトークンはテナントB の Entra ID が発行したものであり、
        # テナントA の Graph リソースへのアクセス権がありません。
        # 同一テナント内の検証目的にのみ使用してください。
        # クロステナントアクセスには client_secret モードを使用してください。
        mi_client_id = os.getenv("MANAGED_IDENTITY_CLIENT_ID")
        credential = ManagedIdentityCredential(client_id=mi_client_id) if mi_client_id else ManagedIdentityCredential()
        token = credential.get_token(GRAPH_SCOPE)
        return token.token

    if auth_mode == "client_secret":
        tenant_id = _get_env("GRAPH_TENANT_ID")
        client_id = _get_env("GRAPH_CLIENT_ID")
        client_secret = _get_env("GRAPH_CLIENT_SECRET")
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        msal_app = msal.ConfidentialClientApplication(
            client_id=client_id,
            authority=authority,
            client_credential=client_secret,
        )
        result = msal_app.acquire_token_for_client(scopes=[GRAPH_SCOPE])
        token = result.get("access_token")
        if not token:
            raise RuntimeError(f"Failed to acquire token by client_secret: {result}")
        return token

    raise ValueError("GRAPH_AUTH_MODE must be managed_identity or client_secret")


def _graph_get(path: str, token: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE_URL}{path}"

    for attempt in range(4):
        response = requests.get(url, headers=headers, params=params, timeout=30)
        if response.status_code in (429, 503, 504) and attempt < 3:
            retry_after = int(response.headers.get("Retry-After", "2"))
            time.sleep(min(10, retry_after + attempt))
            continue
        if response.status_code >= 400:
            raise RuntimeError(f"Graph request failed: {response.status_code} {response.text}")
        return response.json()

    raise RuntimeError("Graph request failed after retries")


def _resolve_site(hostname: str, site_path: str, token: str) -> Dict[str, Any]:
    return _graph_get(
        f"/sites/{hostname}:{site_path}",
        token,
        params={"$select": "id,displayName,webUrl"},
    )


def _list_root_children(site_id: str, token: str, top: int) -> List[Dict[str, Any]]:
    data = _graph_get(
        f"/sites/{site_id}/drive/root/children",
        token,
        params={
            "$top": str(top),
            "$select": "id,name,webUrl,size,createdDateTime,lastModifiedDateTime,eTag,cTag,file,folder,parentReference",
        },
    )
    return data.get("value", [])


def _scan_sites() -> Dict[str, Any]:
    hostname = _get_env("SP_HOSTNAME")
    site_paths_raw = _get_env("SP_SITE_PATHS")
    site_paths = [p.strip() for p in site_paths_raw.split(",") if p.strip()]
    if not site_paths:
        raise ValueError("SP_SITE_PATHS must include at least one site path")

    max_items = int(os.getenv("MAX_ITEMS_PER_SITE", "20"))
    token = _get_graph_token()

    sites_result: List[Dict[str, Any]] = []
    for site_path in site_paths:
        site = _resolve_site(hostname, site_path, token)
        children = _list_root_children(site["id"], token, max_items)
        file_items = []

        for item in children:
            file_items.append(
                {
                    "id": item.get("id"),
                    "title": item.get("name"),
                    "url": item.get("webUrl"),
                    "size": item.get("size"),
                    "createdDateTime": item.get("createdDateTime"),
                    "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                    "isFolder": "folder" in item,
                    "parentPath": item.get("parentReference", {}).get("path"),
                    "eTag": item.get("eTag"),
                }
            )

        sites_result.append(
            {
                "sitePath": site_path,
                "siteId": site.get("id"),
                "siteTitle": site.get("displayName"),
                "siteUrl": site.get("webUrl"),
                "items": file_items,
            }
        )

    return {
        "authMode": os.getenv("GRAPH_AUTH_MODE", "client_secret"),
        "hostname": hostname,
        "siteCount": len(sites_result),
        "sites": sites_result,
    }


@app.function_name(name="GraphTenantScanHttp")
@app.route(route="graph/tenant-scan", methods=["GET"])
def graph_tenant_scan_http(req: func.HttpRequest) -> func.HttpResponse:
    try:
        result = _scan_sites()
        return func.HttpResponse(
            body=json.dumps(result, ensure_ascii=False, indent=2),
            status_code=200,
            mimetype="application/json",
        )
    except Exception as exc:
        logging.exception("GraphTenantScanHttp failed")
        return func.HttpResponse(
            body=json.dumps({"error": str(exc)}, ensure_ascii=False),
            status_code=500,
            mimetype="application/json",
        )


@app.function_name(name="GraphTenantScanTimer")
@app.schedule(
    schedule="%GRAPH_SCAN_SCHEDULE%",
    arg_name="timer",
    run_on_startup=False,
    use_monitor=True,
)
def graph_tenant_scan_timer(timer: func.TimerRequest) -> None:
    try:
        result = _scan_sites()
        logging.info(
            "GraphTenantScanTimer completed. siteCount=%s authMode=%s",
            result.get("siteCount"),
            result.get("authMode"),
        )
        for site in result.get("sites", []):
            logging.info(
                "site=%s items=%s",
                site.get("sitePath"),
                len(site.get("items", [])),
            )
    except Exception:
        logging.exception("GraphTenantScanTimer failed")
        raise
