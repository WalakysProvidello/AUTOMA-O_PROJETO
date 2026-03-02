# ==========================================
# SCRIPT 2 - DOWNLOAD WHATSAPP SESSÕES
# ==========================================

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from pathlib import Path
import time
import zipfile

# ========================
# CONFIGURAÇÃO CHROME
# ========================

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

wait = WebDriverWait(driver, 60)

# ========================
# FUNÇÕES AUXILIARES
# ========================

def aguardar_pagina(driver, timeout=60):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def aguardar_download(pasta, extensao, timeout=180):
    inicio = time.time()
    while time.time() - inicio < timeout:
        arquivos = list(pasta.glob(f"*{extensao}"))
        parciais = list(pasta.glob("*.crdownload"))
        if arquivos and not parciais:
            return max(arquivos, key=lambda f: f.stat().st_ctime)
        time.sleep(1)
    raise TimeoutException("Download não finalizou.")

# ========================
# DATA
# ========================

agora = datetime.now()
ano = agora.year
mes = agora.month
dia = agora.day

data_referencia = f"01/{mes:02d}/{ano} 00:00"

# ========================
# LOGIN
# ========================

url_login = "https://accounts.robbu.global/"
ambiente = "Concilig 3"
usuario = "lucascavaguti"
senha = "Trocar@11990"

try:
    print("🟢 Iniciando processo WhatsApp Sessões")

    # Abre login
    driver.get(url_login)
    aguardar_pagina(driver)

    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome da empresa')]/following-sibling::input"))
    ).send_keys(ambiente)

    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome de usuário') or contains(text(),'Email')]/following-sibling::input"))
    ).send_keys(usuario)

    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Senha')]/following-sibling::input"))
    ).send_keys(senha)

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[normalize-space()='Entrar']"))
    ).click()

    # Aguarda redirecionamento
    wait.until(EC.url_contains("robbu.global"))
    aguardar_pagina(driver)

    # ========================
    # FATURAMENTO
    # ========================

    driver.get("https://inveniocenter.robbu.global/painel/faturamento")
    aguardar_pagina(driver)

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    print("✅ Tabela carregada.")

    linha = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, f"//tr[.//td[normalize-space(text())='{data_referencia}']]")
        )
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", linha)

    # Aguarda botão menu
    menu = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//tr[.//td[normalize-space(text())='{data_referencia}']]//div[contains(@class,'dots')]")
        )
    )
    menu.click()

    # Aguarda botão download
    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@title,'download')]"))
    ).click()

    # Aguarda modal abrir
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container')]"))
    )

    # Aguarda botão WhatsApp ficar clicável
    botao = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[normalize-space()='Analítico WhatsApp Sessões']/ancestor::a")
        )
    )

    driver.execute_script("arguments[0].click();", botao)

    # ========================
    # DOWNLOAD
    # ========================

    pasta = Path(r"C:\Users\lucascavaguti\Downloads")
    arquivo_zip = aguardar_download(pasta, ".zip")

    print("📦 ZIP baixado:", arquivo_zip.name)

    # Extração segura
    with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
        zip_ref.extractall(pasta)

    arquivo_zip.unlink()

    print("✅ WhatsApp Sessões concluído com sucesso.")

except Exception as e:
    print("❌ Erro:", e)

finally:
    driver.quit()
    print("🔴 Navegador fechado.")
