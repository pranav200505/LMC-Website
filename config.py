import os

class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-change-in-production")
    MAX_CONTENT_LENGTH = 10 * 1024 * 1024  # 10 MB upload limit
    WTF_CSRF_ENABLED = True
