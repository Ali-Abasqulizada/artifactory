class Config:
    DEPLOY_URL = ""
    DEPLOY_KEY = ""
    DATABASE = "artifactory.db"
    STORAGE_FOLDER = "storage"
    DEBUG = False
    TESTING = False

class DevelopmentConfig(Config):
    DEPLOY_URL = "http://localhost:5000"
    DEPLOY_KEY = "very very secred code"
    DEBUG = True