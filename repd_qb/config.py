"""Runtime settings. Secrets come from environment variables (GitHub secrets), never from code."""
import os

QB_REALM = os.getenv("QB_REALM", "enerdatics.quickbase.com")
QB_USER_TOKEN = os.getenv("QB_USER_TOKEN", "")
QB_TABLE_ID = os.getenv("QB_TABLE_ID", "bs8th2g3g")
QB_REPORT_ID = os.getenv("QB_REPORT_ID", "1000638")     # ONE report: Country = United Kingdom, all rows

OPENCAGE_API_KEY = os.getenv("OPENCAGE_API_KEY", "").strip()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
GPT_MAX_CALLS = int(os.getenv("GPT_MAX_CALLS", "800"))
