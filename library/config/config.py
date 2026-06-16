import os

class Config:
    DEPLOY_URL = ""
    DEPLOY_KEY = os.getenv("DEPLOY_KEY", "")
    GITHUB_COMMIT_URL = ""
    DATABASE = "artifactory.db"
    STORAGE_FOLDER = "storage"
    DEBUG = False
    TESTING = False


class DevelopmentConfig(Config):
    DEPLOY_URL = ""
    DEPLOY_KEY = os.getenv("DEPLOY_KEY", "testCode")
    GITHUB_COMMIT_URL = "https://github.com/Ali-Abasqulizada/Project/commit/"
    DEBUG = True