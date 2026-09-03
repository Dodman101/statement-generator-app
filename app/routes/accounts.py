"""
Accounts: signup, login, logout, and a dashboard where a logged-in user can
issue their own API key instead of needing the founder to run
scripts/create_api_key.py for every new customer. This is the "self-serve"
piece the earlier billing decision depends on.

Session handling matches Vett's exactly (see app/accounts.py) - same
cookie name, same token format, same secret env var - so turning on true
cross-product SSO later is a config change, not a rewrite.
"""
import logging
import os

from fastapi import APIRouter, Request, Response, Cookie, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from app.accounts import hash_password, verify_password, make_session_token, decode_session_token
from app.db import create_user, get_user_by_email, get_user_by_id, create_api_client, list_api_clients_for_user

logger = logging.getLogger(__name__)
router = APIRouter(tags=["accounts"])

templates = Jinja2Templates(directory=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates"))


async def get_optional_web_user(dodman_session: str | None) -> dict | None:
    if not dodman_session:
        return None
    user_id = decode_session_token(dodman_session)
    if not user_id:
        return None
    return await get_user_by_id(user_id)


def _set_session_cookie(resp: Response, user_id: str) -> None:
    token = make_session_token(user_id)
    resp.set_cookie(
        key="dodman_session",
        value=token,
        httponly=True,
        samesite="lax",
        secure=(os.getenv("APP_ENV", "production") == "production"),
        max_age=60 * 60 * 24 * 7,  # 7 days - matches Vett
    )


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@router.get("/signup", response_class=HTMLResponse, include_in_schema=False)
async def signup_page(request: Request):
    return templates.TemplateResponse(request, "accounts_signup.html")


@router.get("/login", response_class=HTMLResponse, include_in_schema=False)
async def login_page(request: Request):
    return templates.TemplateResponse(request, "accounts_login.html")


@router.get("/dashboard", response_class=HTMLResponse, include_in_schema=False)
async def dashboard_page(request: Request, dodman_session: str | None = Cookie(default=None)):
    user = await get_optional_web_user(dodman_session)
    if not user:
        return RedirectResponse(url="/accounts/login", status_code=303)
    clients = await list_api_clients_for_user(user["id"])
    return templates.TemplateResponse(request, "accounts_dashboard.html", {"user": user, "clients": clients})


# ---------------------------------------------------------------------------
# Actions (called via fetch() from the pages above)
# ---------------------------------------------------------------------------

@router.post("/signup", include_in_schema=False)
async def signup_action(email: str = Form(...), password: str = Form(...)):
    email = email.lower().strip()
    if len(password) < 8:
        raise HTTPException(status_code=400, detail="Password must be at least 8 characters.")

    existing = await get_user_by_email(email)
    if existing:
        raise HTTPException(status_code=409, detail="An account with that email already exists. Try logging in instead.")

    try:
        user_id = await create_user(email, hash_password(password))
    except Exception as e:
        logger.error(f"Signup failed for {email}: {e}")
        raise HTTPException(status_code=503, detail="Could not create account. Please try again shortly.")

    resp = RedirectResponse(url="/accounts/dashboard", status_code=303)
    _set_session_cookie(resp, user_id)
    logger.info(f"New account created: {email}")
    return resp


@router.post("/login", include_in_schema=False)
async def login_action(email: str = Form(...), password: str = Form(...)):
    row = await get_user_by_email(email)
    invalid = HTTPException(status_code=401, detail="Incorrect email or password.")
    if not row:
        raise invalid
    if not verify_password(password, row["password_hash"]):
        raise invalid

    resp = RedirectResponse(url="/accounts/dashboard", status_code=303)
    _set_session_cookie(resp, str(row["id"]))
    return resp


@router.post("/logout", include_in_schema=False)
async def logout_action():
    resp = RedirectResponse(url="/accounts/login", status_code=303)
    resp.delete_cookie(key="dodman_session", httponly=True, samesite="lax")
    return resp


@router.post("/dashboard/create-key", include_in_schema=False)
async def create_key_action(request: Request, label: str = Form(...), dodman_session: str | None = Cookie(default=None)):
    user = await get_optional_web_user(dodman_session)
    if not user:
        return RedirectResponse(url="/accounts/login", status_code=303)

    label = label.strip() or f"{user['email']}'s key"
    try:
        raw_key = await create_api_client(label, plan='free', user_id=user["id"])
    except Exception as e:
        logger.error(f"Self-serve key creation failed for user {user['id']}: {e}")
        raise HTTPException(status_code=503, detail="Could not create a key right now. Please try again shortly.")

    clients = await list_api_clients_for_user(user["id"])
    return templates.TemplateResponse(request, "accounts_dashboard.html", {
        "user": user, "clients": clients, "new_key": raw_key, "new_key_label": label,
    })
