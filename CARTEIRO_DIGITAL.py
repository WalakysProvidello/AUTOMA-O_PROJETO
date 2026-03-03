# ==========================================
# SCRIPT 2 - DOWNLOAD WHATSAPP SESSÕES (CORRIGIDO)
# ==========================================

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from datetime import datetime
from pathlib import Path
import time
import zipfile
import os

# ========================
# CONFIGURAÇÃO CHROME
# ========================

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

# Configurar pasta de download
options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\lucascavaguti\Downloads",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

wait = WebDriverWait(driver, 30)  # Reduzi para 30s para testes mais rápidos

# ========================
# FUNÇÕES AUXILIARES
# ========================

def aguardar_pagina(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(2)  # Pequena pausa extra para renderização
    except:
        pass

def aguardar_download(pasta, extensao, timeout=180):
    inicio = time.time()
    arquivos_iniciais = list(pasta.glob(f"*{extensao}"))
    
    while time.time() - inicio < timeout:
        arquivos = list(pasta.glob(f"*{extensao}"))
        parciais = list(pasta.glob("*.crdownload"))
        
        # Verifica se há novo arquivo e nenhum parcial
        if len(arquivos) > len(arquivos_iniciais) and not parciais:
            return max(arquivos, key=lambda f: f.stat().st_ctime)
        
        time.sleep(2)
    raise TimeoutException("Download não finalizou.")

def clicar_com_erro(elemento, timeout=10):
    """Tenta clicar com fallback para JavaScript"""
    try:
        elemento.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", elemento)

# ========================
# DADOS (USE VARIÁVEIS DE AMBIENTE NA PRODUÇÃO)
# ========================

url_login = "https://accounts.robbu.global/"
ambiente = "Concilig 3"
usuario = "SEU_USUARIO"  # ⚠️ Não use senhas no código!
senha = "SUA_SENHA"

# ========================
# INICIALIZAÇÃO
# ========================

try:
    print("🟢 Iniciando processo WhatsApp Sessões")
    driver.get(url_login)
    aguardar_pagina(driver)

    # ========================
    # LOGIN
    # ========================
    
    # Campo empresa
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome da empresa')]/following-sibling::input"))
    ).send_keys(ambiente)

    # Campo usuário
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome de usuário') or contains(text(),'Email')]/following-sibling::input"))
    ).send_keys(usuario)

    # Campo senha
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Senha')]/following-sibling::input"))
    ).send_keys(senha)

    # Botão entrar
    btn_entrar = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[normalize-space()='Entrar']"))
    )
    clicar_com_erro(btn_entrar)

    # Aguarda redirecionamento
    wait.until(EC.url_contains("robbu.global"))
    aguardar_pagina(driver)

    # ========================
    # FATURAMENTO
    # ========================

    driver.get("https://inveniocenter.robbu.global/painel/faturamento")
    aguardar_pagina(driver)

    # Aguarda tabela
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    print("✅ Tabela carregada.")

    # Data de referência
    agora = datetime.now()
    data_referencia = f"01/{agora.month:02d}/{agora.year} 00:00"

    # Localiza linha
    linha = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, f"//tr[.//td[normalize-space(text())='{data_referencia}']]")
        )
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", linha)
    time.sleep(2)

    # Aguarda botão menu (dots)
    menu = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//tr[.//td[normalize-space(text())='{data_referencia}']]//div[contains(@class,'dots')]")
        )
    )
    clicar_com_erro(menu)
    time.sleep(1)

    # Aguarda botão download
    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@title,'download')]"))
    ).click()

    # Aguarda modal abrir
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container')]"))
    )
    time.sleep(1)

    # Aguarda botão WhatsApp ficar clicável
    botao = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[normalize-space()='Analítico WhatsApp Sessões']/ancestor::a")
        )
    )
    clicar_com_erro(botao)

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
    print("❌ Erro:", str(e))
    # Opcional: tirar print da tela para debug
    # driver.save_screenshot("erro.png")

finally:
    driver.quit()
    print("🔴 Navegador fechado.")
