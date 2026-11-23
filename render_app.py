from app import app, warm_report_caches


# Preload Excel data once in the Gunicorn master; workers inherit the memory cache.
warm_report_caches()
