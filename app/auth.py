"""
API key authentication, backed by the api_clients table (see app/db.py).

Every protected request must carry a valid key in the X-API-Key header.
require_api_key is a FastAPI dependency - use it as
`client: dict = Depends(require_api_key)` on any route that needs auth, and
FastAPI's generated OpenAPI docs will show the X-API-Key header requirement
automatically.
"""
import logging

from fastapi import Header, HTTPException

from app.db import get_client_by_key
from app.observability import capture_service_error

logger = logging.getLogger(__name__)


async def require_api_key(x_api_key: str = Header(default="", alias="X-API-Key")) -> dict:
    if not x_api_key:
        raise HTTPException(status_code=401, detail="Missing API key. Include it in the X-API-Key header.")

    try:
        client = await get_client_by_key(x_api_key)
    except Exception as e:
        capture_service_error(e, where="require_api_key")
        raise HTTPException(status_code=503, detail="Service temporarily unavailable. Please try again shortly.")

    if not client:
        raise HTTPException(status_code=401, detail="Invalid API key.")

    return client
