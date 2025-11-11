import requests
from urllib.parse import urljoin


class AuthError(RuntimeError):
    """Raised when authentication fails."""


def _extract_from_path(data, path):
    if not path:
        return data
    current = data
    for part in str(path).split("."):
        if isinstance(current, dict):
            current = current.get(part)
        else:
            current = None
        if current is None:
            break
    return current


def fetch_auth_token(base_url, auth_config):
    if not base_url:
        raise AuthError("Base URL is required to fetch auth token.")
    if not auth_config:
        raise AuthError("Auth configuration missing.")

    endpoint = auth_config.get("endpoint")
    if not endpoint:
        raise AuthError("Auth endpoint not provided in configuration.")

    method = (auth_config.get("method") or "POST").upper()
    body = auth_config.get("body") or {}
    headers = auth_config.get("headers") or {"Content-Type": "application/json"}
    verify = auth_config.get("verify", True)
    timeout = auth_config.get("timeout", 30)

    url = urljoin(base_url.rstrip("/") + "/", endpoint.lstrip("/"))

    try:
        response = requests.request(method, url, json=body, headers=headers, timeout=timeout, verify=verify)
        response.raise_for_status()
    except Exception as exc:
        raise AuthError(f"Failed to call auth endpoint: {exc}") from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise AuthError("Auth response is not valid JSON.") from exc

    token_path = auth_config.get("token_path", "token")
    token = _extract_from_path(payload, token_path)
    print(token)
    if not token:
        raise AuthError(f"Token not found at path '{token_path}'.")

    return str(token)





