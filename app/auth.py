"""
API key authentication, backed by the api_clients table (see app/db.py).

Every protected request must carry a valid key in the X-API-Key header. On
success, g.api_client is set to a dict: {id, label, plan, monthly_job_limit}.
"""
import logging
from functools import wraps
from flask import request, jsonify, g

from app.db import get_client_by_key

logger = logging.getLogger(__name__)


def require_api_key(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        supplied_key = request.headers.get('X-API-Key', '')
        if not supplied_key:
            return jsonify({"message": "Missing API key. Include it in the X-API-Key header."}), 401

        try:
            client = get_client_by_key(supplied_key)
        except Exception as e:
            logger.error(f"Database error while validating API key: {e}")
            return jsonify({"message": "Service temporarily unavailable. Please try again shortly."}), 503

        if not client:
            return jsonify({"message": "Invalid API key."}), 401

        g.api_client = client
        return view_func(*args, **kwargs)

    return wrapped
