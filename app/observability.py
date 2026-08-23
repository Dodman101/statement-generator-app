"""
Error tracking, via Sentry - entirely opt-in.

If SENTRY_DSN isn't set, init_error_tracking() does nothing and every
capture_* call below becomes a silent no-op. That means you can deploy this
app today with zero Sentry account and nothing breaks; set SENTRY_DSN
whenever you actually create one and error tracking turns on with no other
code changes.
"""
import os
import logging

logger = logging.getLogger(__name__)

_enabled = False


def init_error_tracking(app_env='production'):
    global _enabled
    dsn = os.getenv('SENTRY_DSN')
    if not dsn:
        logger.info("SENTRY_DSN not set - error tracking is disabled. Server logs are still the fallback.")
        return

    try:
        import sentry_sdk
        from sentry_sdk.integrations.flask import FlaskIntegration
        sentry_sdk.init(
            dsn=dsn,
            integrations=[FlaskIntegration()],
            environment=app_env,
            # Full error capture; keep performance-trace sampling low since this
            # app doesn't need request-latency tracing to be useful here.
            traces_sample_rate=0.0,
        )
        _enabled = True
        logger.info("Error tracking initialized (Sentry).")
    except Exception as e:
        # Never let observability setup take the app down.
        logger.error(f"Could not initialize Sentry: {e}")


def capture_job_failure(exc, *, temp_id, client_label, client_id):
    """Report a failed statement-generation job, tagged so you can tell which
    client and which job from the Sentry alert alone - not just a stack trace."""
    logger.error(f"Job {temp_id} failed for client '{client_label}' (id={client_id}): {exc}")
    if not _enabled:
        return
    import sentry_sdk
    with sentry_sdk.push_scope() as scope:
        scope.set_tag("temp_id", temp_id)
        scope.set_tag("client_label", client_label)
        scope.set_tag("client_id", client_id)
        sentry_sdk.capture_exception(exc)


def capture_service_error(exc, *, where, **extra):
    """Report an infrastructure-level failure (DB unreachable, etc.) - the
    503s a client would otherwise have to report to you manually."""
    logger.error(f"Service error in {where}: {exc}")
    if not _enabled:
        return
    import sentry_sdk
    with sentry_sdk.push_scope() as scope:
        scope.set_tag("where", where)
        for key, value in extra.items():
            scope.set_extra(key, value)
        sentry_sdk.capture_exception(exc)
