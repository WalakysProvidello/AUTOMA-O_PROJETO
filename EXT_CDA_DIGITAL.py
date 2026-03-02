from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from pathlib import Path
import os
import shutil
import time


# ===============================
# 1. CONFIGURAÇÃO DE DATAS
# ===============================
hoje = datetime.now()

dia = hoje.strftime("%d")
mes_num = hoje.strftime("%m")
ano = hoje.strftime("%Y")

# Data para o nome do arquivo: APENAS DIA E MÊS (ex: 2502)
data_nome_arquivo = hoje.strftime("%d%m")

data_inicio = f"01-{mes_num}-{ano}"
data_fim = (hoje - timedelta(days=1)).strftime("%d-%m-%Y")

meses_br = {
    "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL",
    "05": "MAIO", "06": "JUNHO", "07": "JULHO", "08": "AGOSTO",
    "09": "SETEMBRO", "10": "OUTUBRO", "11": "NOVEMBRO", "12": "DEZEMBRO"
}
nome_mes = meses_br[mes_num]

# ===============================
# 2. CAMINHOS PADRONIZADOS
# ===============================
downloads = Path.home() / "Downloads"
caminho_base = Path(r"G:/DIGITAL/COTA DISPARO MASSIVO/ALIMENTAÇÃO")

# Pasta do mês como "02 - FEVEREIRO"
pasta_mes_nome = f"{mes_num} - {nome_mes}"
destino_final = caminho_base / ano / pasta_mes_nome / dia / "CDA_DIGITAL"

# Cria estrutura automaticamente
destino_final.mkdir(parents=True, exist_ok=True)

print(f"Pasta destino: {destino_final}")

# ===============================
# 3. CONFIGURAÇÃO DO CHROME
# ===============================
options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": str(downloads),
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True
})

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# ===============================
# 4. FLUXO PRINCIPAL
# ===============================
try:
    print("Acessando sistema...")
    driver.get("https://cromosapp.com.br/login/conciligdigital")

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "user"))
    ).send_keys("rafaelmdias")

    driver.find_element(By.ID, "pass").send_keys("@Trocar123")
    driver.find_element(By.ID, "pass").send_keys(Keys.ENTER)
    time.sleep(3)

    print("Acessando relatório...")
    driver.get("https://cromosapp.com.br/relatorio/rel_metricas")

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "dt_ini")))

    driver.find_element(By.ID, "dt_ini").clear()
    driver.find_element(By.ID, "dt_ini").send_keys(data_inicio)

    driver.find_element(By.ID, "dt_fim").clear()
    driver.find_element(By.ID, "dt_fim").send_keys(data_fim)

    driver.find_element(By.XPATH, "//button[contains(text(),'Enviar')]").click()
    time.sleep(5)

    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
    time.sleep(2)

    driver.find_element(By.XPATH, "//button[contains(text(),'Gerar CSV')]").click()
    print("Download solicitado...")

    # ===============================
    # 5. AGUARDAR E RENOMEAR ARQUIVO 699*.csv
    # ===============================
    timeout = time.time() + 60
    arquivo_encontrado = None

    print("Aguardando arquivo 699*.csv...")
    while time.time() < timeout:
        agora = datetime.now()
        # Busca arquivos que começam com 699
        arquivos = list(downloads.glob("699*.csv"))
        
        # Filtra apenas os criados nos últimos 5 minutos
        recentes = [f for f in arquivos if (agora - datetime.fromtimestamp(f.stat().st_mtime)) <= timedelta(minutes=5)]

        if recentes:
            recentes.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            arquivo_encontrado = recentes[0]
            break
        time.sleep(2)

    # ===============================
    # 6. MOVER E RENOMEAR
    # ===============================
    if arquivo_encontrado and arquivo_encontrado.exists():
        # Aguarda finalizar download se estiver pendente
        while arquivo_encontrado.with_suffix(".csv.crdownload").exists():
            time.sleep(1)

        # DEFINE O NOME: CDA_ddmm.csv
        nome_final = f"CDA_{data_nome_arquivo}.csv"
        caminho_final = destino_final / nome_final

        # Remove se já existir no destino (para evitar erro de permissão)
        if caminho_final.exists():
            caminho_final.unlink()

        shutil.move(str(arquivo_encontrado), str(caminho_final))
        print(f"Sucesso! Arquivo renomeado para: {nome_final}")
        print(f"Localizado em: {destino_final}")

    else:
        print("Erro: Nenhum arquivo recente (699*.csv) foi encontrado.")

except Exception as e:
    print("Erro durante o processo:", e)

finally:
    driver.quit()
    print("Processo finalizado.")