# ==============================
# SCRIPT 1 - CARTEIRO DIGITAL
# ==============================

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
import shutil

# ========================
# CONFIG
# ========================

options = Options()
options.add_argument("--start-maximized")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# ========================
# DATA
# ========================

agora = datetime.now()
ano = agora.year
mes = agora.month
dia = agora.day

data_referencia = f"01/{mes:02d}/{ano} 00:00"
data_nome = f"{dia:02d}{mes:02d}"

# ========================
# LOGIN
# ========================

url_login = "https://accounts.robbu.global/"
ambiente = "Concilig 3"
usuario = "lucascavaguti"
senha = "Trocar@11990"

def aguardar_download(pasta, extensao):
    inicio = time.time()
    while time.time() - inicio < 180:
        arquivos = list(pasta.glob(f"*{extensao}"))
        if arquivos:
            return max(arquivos, key=lambda f: f.stat().st_ctime)
        time.sleep(1)
    raise Exception("Download não finalizou")

try:
    driver.get(url_login)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome da empresa')]/following-sibling::input"))
    ).send_keys(ambiente)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Nome de usuário')]/following-sibling::input"))
    ).send_keys(usuario)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Senha')]/following-sibling::input"))
    ).send_keys(senha)

    WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//div[normalize-space()='Entrar']"))
    ).click()

    time.sleep(3)

    # Ir para faturamento
    driver.get("https://inveniocenter.robbu.global/painel/faturamento")

    linha = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, f"//tr[.//td[normalize-space(text())='{data_referencia}']]"))
    )

    linha.find_element(By.XPATH, ".//div[contains(@class,'dots')]").click()

    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@title,'download')]"))
    ).click()

    modal = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container')]"))
    )

    botao = modal.find_element(By.XPATH, ".//div[normalize-space()='Analítico Carteiro Digital']/ancestor::a")
    driver.execute_script("arguments[0].click();", botao)

    pasta = Path(r"C:\Users\lucascavaguti\Downloads")
    arquivo = aguardar_download(pasta, ".csv")

    destino = Path(fr"C:\Users\lucascavaguti\Desktop\CARTEIRO_{data_nome}.csv")
    shutil.move(str(arquivo), destino)

    print("Carteiro Digital concluído.")

except Exception as e:
    print("Erro:", e)

finally:
    driver.quit()