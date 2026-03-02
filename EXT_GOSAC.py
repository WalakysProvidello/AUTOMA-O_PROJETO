from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from pathlib import Path
import time
import os
import shutil

# =======================================================
# CONFIGURAÇÃO DE DATAS (PADRÃO BR)
# =======================================================
hoje = datetime.now()
ano_atual = hoje.strftime("%Y")
mes_atual_num = hoje.strftime("%m")
dia_atual_num = hoje.strftime("%d")

meses_br = {
    "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL",
    "05": "MAIO", "06": "JUNHO", "07": "JULHO", "08": "AGOSTO",
    "09": "SETEMBRO", "10": "OUTUBRO", "11": "NOVEMBRO", "12": "DEZEMBRO"
}
nome_mes_pt = meses_br[mes_atual_num]

# Data para o filtro do sistema
data_inicial = datetime(hoje.year, hoje.month, 1).strftime("%d-%m-%Y T00:00")
data_final_dt = hoje - timedelta(days=1)
data_final = data_final_dt.strftime("%d-%m-%Y T23:59")

# Data para o nome do arquivo (ex: 2502)
data_nome_arquivo = hoje.strftime("%d%m")

# =======================================================
# CONFIG CHROME
# =======================================================
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
# Configuração para garantir o diretório de download local antes da rede
downloads_local = Path.home() / "Downloads"

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)
wait = WebDriverWait(driver, 30)

try:
    # =======================================================
    # ACESSO E LOGIN
    # =======================================================
    driver.get("https://gestor.gosac.com.br/dashboard")

    wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys("concilig")
    wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys("Concilig102030")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entrar')]"))).click()

    time.sleep(4)
    driver.get("https://gestor.gosac.com.br/dashboard")

    # EXIBIR FILTROS
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'EXIBIR FILTROS')]"))).click()

    # PREENCHER DATAS
    inputs_datetime = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='datetime-local']")))
    
    inputs_datetime[0].clear()
    inputs_datetime[0].send_keys(data_inicial)
    
    inputs_datetime[1].clear()
    inputs_datetime[1].send_keys(data_final)

    # APLICAR
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Aplicar')]"))).click()
    
    print("Filtros aplicados, aguardando carregamento...")
    time.sleep(10)

    # =======================================================
    # EXPORTAR
    # =======================================================
    antes = set(os.listdir(downloads_local))
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Exportar')]"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Analítico de Campanhas')]"))).click()

    print("Download solicitado...")

    # =======================================================
    # AGUARDAR DOWNLOAD
    # =======================================================
    timeout = time.time() + 600
    novo_arquivo = None

    while time.time() < timeout:
        depois = set(os.listdir(downloads_local))
        novos = depois - antes
        if novos:
            tmp = list(novos)[0]
            if not tmp.endswith(".crdownload") and not tmp.endswith(".tmp"):
                novo_arquivo = tmp
                break
        time.sleep(2)

    if not novo_arquivo:
        raise Exception("Nenhum arquivo novo detectado na pasta de Downloads.")

    arquivo_origem = downloads_local / novo_arquivo

    # =======================================================
    # ORGANIZAÇÃO FINAL (REDE G:)
    # =======================================================
    pasta_mes_formatada = f"{mes_atual_num} - {nome_mes_pt}"
    
    # Caminho seguindo o padrão dos outros scripts
    pasta_destino = Path(r"G:/DIGITAL/COTA DISPARO MASSIVO/ALIMENTAÇÃO") / ano_atual / pasta_mes_formatada / dia_atual_num / "GOSAC"
    
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # Nome final: GOSAC_2502.xlsx
    extensao = arquivo_origem.suffix
    caminho_final = pasta_destino / f"GOSAC_{data_nome_arquivo}{extensao}"

    # Se já existir, remove antes de mover
    if caminho_final.exists():
        caminho_final.unlink()

    shutil.move(str(arquivo_origem), str(caminho_final))

    print(f"SUCESSO: Arquivo movido para {caminho_final}")

except Exception as e:
    print(f"ERRO NO PROCESSO GOSAC: {e}")

finally:
    driver.quit()