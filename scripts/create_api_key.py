"""
Issue a new Statement Generator API key.

The raw key is only ever printed here, once - only its hash is stored in the
database, so if it's lost the only fix is issuing a new one for that client.

Usage:
    python scripts/create_api_key.py "Acme Corp" --plan paid --limit 500
    python scripts/create_api_key.py "Pilot Client" --plan free --limit 20
    python scripts/create_api_key.py "Internal - Dodman" --plan internal   # no --limit = unlimited

Requires DATABASE_URL to be set (same Neon connection string as the app uses).
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from app.db import init_schema, create_api_client  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Issue a new Statement Generator API key.")
    parser.add_argument("label", help="Human-readable name for this client, e.g. 'Acme Corp'")
    parser.add_argument("--plan", default="free", help="Plan name, e.g. free / paid / internal (default: free)")
    parser.add_argument("--limit", type=int, default=None, help="Monthly job limit. Omit for unlimited.")
    args = parser.parse_args()

    init_schema()
    raw_key = create_api_client(args.label, plan=args.plan, monthly_job_limit=args.limit)

    print(f"\nCreated API client '{args.label}' (plan: {args.plan}, "
          f"monthly limit: {args.limit if args.limit is not None else 'unlimited'})")
    print(f"\nAPI key (save this now - it will not be shown again):\n\n  {raw_key}\n")


if __name__ == "__main__":
    main()
