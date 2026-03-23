import os
import sys
from pathlib import Path

# 1. Ajustar o caminho para o Python encontrar os ficheiros na pasta config
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# 2. Importar a classe (Agora de forma simples)
try:
    from config.settings_class import Settings
except ImportError:
    from settings_class import Settings

# 3. Caminho do .env
ENV_PATH = BASE_DIR / "deploy" / ".env"

try:
    # 4. Carregar as configurações
    settings = Settings.from_env(env_file=str(ENV_PATH))
    
    # 5. Exportar variáveis para o robô
    EMAIL_SOMA = settings.site_user
    SENHA_SOMA = settings.site_password
    SOMA_URL_INICIAL = settings.site_login_url
    HEADLESS = settings.headless 
    SELENIUM_TIMEOUT = settings.timeout_seconds
    CREDENCIAIS_PATH = settings.google_credentials_path
    SPREADSHEET_URL = settings.spreadsheet_url
    SHEET_NAME = settings.sheet_contaordem
    
    print(f"✅ Configurações carregadas (Modo Visível: {not HEADLESS})")

except Exception as e:
    print(f"❌ Erro fatal ao carregar configurações: {e}")
    raise