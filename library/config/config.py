import os

class Config:
    DEPLOY_URL = ""
    DEPLOY_KEY = os.getenv("DEPLOY_KEY", "")
    DATABASE = "artifactory.db"
    STORAGE_FOLDER = "storage"
    DEBUG = False
    TESTING = False


class DevelopmentConfig(Config):
    DEPLOY_URL = "https://github.com/Ali-Abasqulizada/TESTCASE/commit/"
    DEPLOY_KEY = os.getenv("DEPLOY_KEY", "very very secret code")
    DEBUG = True