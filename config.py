"""
config.py - Configuración centralizada del Sistema de Inventario (versión segura)
"""
import os
from datetime import timedelta
from dotenv import load_dotenv

# Carga variables desde el archivo .env (solo en entornos locales)
load_dotenv()

class Config:
    """Configuración base de la aplicación"""
    
    # ==================== FLASK ====================
    SECRET_KEY = os.getenv("SECRET_KEY", "default-key")  # usar valor por defecto solo localmente
    DEBUG = os.getenv("FLASK_DEBUG", "False").lower() == "true"

    # ==================== SERVER ====================
    HOST = os.getenv("FLASK_HOST", "0.0.0.0")
    PORT = int(os.getenv("FLASK_PORT", 5000))

    # ==================== BASES DE DATOS COMPARTIDAS ====================
    SHARED_DB_SERVER = os.getenv("SHARED_DB_SERVER", "localhost")
    SHARED_DB_PATH = os.getenv("SHARED_DB_PATH", r"\\servidor\db")

    INVENTARIO_DB = os.getenv("INVENTARIO_DB", rf"{SHARED_DB_PATH}\inventarioWeb\inventario_computadoras.db")
    OFICINAS_CNE_DB = os.getenv("OFICINAS_CNE_DB", rf"{SHARED_DB_PATH}\oficinasCne.db")
    IMPRESORAS_DB = os.getenv("IMPRESORAS_DB", rf"{SHARED_DB_PATH}\inventarioWeb\impresoras.db")
    NOTEBOOKS_DB = os.getenv("NOTEBOOKS_DB", rf"{SHARED_DB_PATH}\inventarioWeb\inventario_notebooks.db")

    # ==================== BASES DE DATOS LOCALES ====================
    LOCAL_DB_DIR = os.getenv("LOCAL_DB_DIR", "./db")
    LOCAL_FALLBACK_DB = os.path.join(LOCAL_DB_DIR, "inventario_local_fallback.db")
    LOCAL_IMPRESORAS_DB = os.path.join(LOCAL_DB_DIR, "impresoras_local_fallback.db")
    LOCAL_NOTEBOOKS_DB = os.path.join(LOCAL_DB_DIR, "notebooks_local_fallback.db")

    # ==================== TIMEOUTS Y CONEXIONES ====================
    DB_TIMEOUT = int(os.getenv("DB_TIMEOUT", 5))
    DB_CHECK_SAME_THREAD = False

    # ==================== LOGGING ====================
    LOG_LEVEL = os.getenv("LOG_LEVEL", "DEBUG")
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

    # ==================== SESIONES ====================
    SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "False").lower() == "true"
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    PERMANENT_SESSION_LIFETIME = timedelta(hours=int(os.getenv("SESSION_LIFETIME_HOURS", 24)))

    # ==================== ARCHIVOS Y UPLOADS ====================
    MAX_CONTENT_LENGTH = int(os.getenv("MAX_CONTENT_LENGTH", 16 * 1024 * 1024))
    ALLOWED_EXTENSIONS = {"xlsx", "csv", "db"}

    # ==================== EXPORTACIÓN ====================
    EXPORT_DIR = os.getenv("EXPORT_DIR", "./exports")
    EXPORT_FILENAME_FORMAT = "%Y%m%d_%H%M%S"
    EXCEL_HEADER_BG_COLOR = "FF4472C4"
    EXCEL_HEADER_FONT_COLOR = "FFFFFFFF"

    # ==================== ESTADOS ====================
    ESTADOS_COMPUTADORAS = ["Ok", "En reparación", "Dado de baja", "Reserva"]
    ESTADOS_IMPRESORAS = ["Operativa", "En reparación", "Fuera de servicio", "En mantenimiento"]
    ESTADOS_NOTEBOOKS = ["Disponible", "Asignada", "En reparación", "Dada de baja"]

    # ==================== VALIDACIONES ====================
    MAX_LENGTH_USUARIO = 100
    MAX_LENGTH_PC_NOMBRE = 100
    MAX_LENGTH_IP = 15
    MAX_LENGTH_MAC = 17
    MAX_LENGTH_MARCA = 50
    MAX_LENGTH_MODELO = 100
    MAX_LENGTH_SERIE = 100
    MAX_LENGTH_OBSERVACIONES = 500

    # ==================== PAGINACIÓN ====================
    ITEMS_PER_PAGE = int(os.getenv("ITEMS_PER_PAGE", 50))

    # ==================== FEATURE FLAGS ====================
    ENABLE_API = os.getenv("ENABLE_API", "True").lower() == "true"
    ENABLE_DEBUG_ROUTES = os.getenv("ENABLE_DEBUG_ROUTES", "True").lower() == "true"
    ENABLE_EXPORT = os.getenv("ENABLE_EXPORT", "True").lower() == "true"

    @staticmethod
    def init_app(app):
        os.makedirs(Config.LOCAL_DB_DIR, exist_ok=True)
        if Config.ENABLE_EXPORT:
            os.makedirs(Config.EXPORT_DIR, exist_ok=True)


class DevelopmentConfig(Config):
    DEBUG = True
    LOG_LEVEL = "DEBUG"
    ENABLE_DEBUG_ROUTES = True


class ProductionConfig(Config):
    DEBUG = False
    LOG_LEVEL = "INFO"
    SESSION_COOKIE_SECURE = True
    ENABLE_DEBUG_ROUTES = False

    @classmethod
    def init_app(cls, app):
        Config.init_app(app)
        import logging
        from logging.handlers import SysLogHandler
        syslog_handler = SysLogHandler()
        syslog_handler.setLevel(logging.WARNING)
        app.logger.addHandler(syslog_handler)


class TestingConfig(Config):
    TESTING = True
    DEBUG = True
    WTF_CSRF_ENABLED = False
    INVENTARIO_DB = "./test_inventario.db"
    OFICINAS_CNE_DB = "./test_oficinas.db"
    IMPRESORAS_DB = "./test_impresoras.db"
    NOTEBOOKS_DB = "./test_notebooks.db"


config = {
    "development": DevelopmentConfig,
    "production": ProductionConfig,
    "testing": TestingConfig,
    "default": DevelopmentConfig
}


def get_config():
    env = os.getenv("FLASK_ENV", "development")
    return config.get(env, config["default"])
