from selenium import webdriver 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from pathlib import Path
import os
import shutil
import time
import pandas as pd

# ==========================================================
# CONFIGURAÇÃO DE DATA
# ==========================================================
hoje = datetime.now()
data_inicio = hoje.replace(day=1).strftime("%d/%m/%Y")
data_fim = (hoje - timedelta(days=1)).strftime("%d/%m/%Y")

ano = hoje.strftime("%Y")
mes_num = hoje.strftime("%m")
dia = hoje.strftime("%d")

meses_br = {
    "01": "JANEIRO",
    "02": "FEVEREIRO",
    "03": "MARÇO",
    "04": "ABRIL",
    "05": "MAIO",
    "06": "JUNHO",
    "07": "JULHO",
    "08": "AGOSTO",
    "09": "SETEMBRO",
    "10": "OUTUBRO",
    "11": "NOVEMBRO",
    "12": "DEZEMBRO"
}

mes_nome = meses_br[mes_num]

# ==========================================================
# CAMINHO BASE REDE G:
# ==========================================================
caminho_base = Path(r"G:/DIGITAL/COTA DISPARO MASSIVO/ALIMENTAÇÃO")

if not caminho_base.exists():
    raise Exception("Unidade G: não está acessível.")

fornecedor = "BRAGI"

pasta_destino = (
    caminho_base
    / ano
    / f"{mes_num} - {mes_nome}"
    / dia
    / fornecedor
)

pasta_destino.mkdir(parents=True, exist_ok=True)

# ==========================================================
# CONFIGURAÇÃO DO NAVEGADOR
# ==========================================================
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("prefs", {
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False
})

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.set_page_load_timeout(30)

# ==========================================================
# LISTA DE SITES
# ==========================================================
sites = [
    'https://bvrodastradicional.noahomni.com.br/',
    'https://bvrodasrenegociados.noahomni.com.br/',
    'https://bvrodasbompagador.noahomni.com.br/',
    'https://solagora.noahomni.com.br/',
    'https://bvrodaswo.noahomni.com.br/'
]

# ==========================================================
# CREDENCIAIS
# ==========================================================
def get_credentials(url):

    if any(chave in url for chave in [
        "nubank.noahomni.com.br",
        "banqi.noahomni.com.br"
    ]):
        return "mis_reports", "alterar@123"

    elif any(chave in url for chave in [
        "bvrodastradicional.noahomni.com.br",
        "bvrodasrenegociados.noahomni.com.br",
        "bvrodasbompagador.noahomni.com.br",
        "bvrodaswo.noahomni.com.br",
        "https://solagora.noahomni.com.br",
        "bvrodascontencioso.noahomni.com.br"
    ]):
        return "admin", "IoH@na@2026"

    else:
        return "mis_reports", "Alterar@123"

falhas_login = []

# ==========================================================
# LOOP PRINCIPAL
# ==========================================================
for base_url in sites:

    nome_sistema = base_url.split("//")[1].split(".")[0].lower()
    print(f"Acessando: {base_url}")

    usuario, senha = get_credentials(base_url)

    try:
        driver.get(base_url)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='text' or @name='username']"))
        )

        driver.find_element(By.XPATH, "//input[@type='text' or @name='username']").send_keys(usuario)
        senha_input = driver.find_element(By.XPATH, "//input[@type='password']")
        senha_input.send_keys(senha)
        senha_input.send_keys(Keys.ENTER)

        print("Login realizado.")

    except Exception as e:
        print(f"Erro ao logar em {base_url}: {e}")
        falhas_login.append(base_url)
        continue

    try:
        time.sleep(4)

        estat_url = base_url.rstrip("/") + "/#/historico-atendimento"
        driver.get(estat_url)
        time.sleep(3)

        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//i[text()='filter_list']"))
        ).click()

        campo_inicio = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Data inicial']"))
        )
        campo_inicio.send_keys(Keys.CONTROL, "a")
        campo_inicio.send_keys(Keys.BACKSPACE)
        campo_inicio.send_keys(data_inicio)

        campo_fim = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Data final']"))
        )
        campo_fim.send_keys(Keys.CONTROL, "a")
        campo_fim.send_keys(Keys.BACKSPACE)
        campo_fim.send_keys(data_fim)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='block' and text()='Pesquisar']"))
        ).click()

        time.sleep(5)

        driver.find_element(By.XPATH, "//i[contains(@class,'mdi-close')]").click()

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//i[text()='more_vert']"))
        ).click()

        exportar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//i[contains(@class,'mdi-microsoft-excel')]"))
        )

        downloads = Path.home() / "Downloads"

        for antigo in downloads.glob("report_tickets*.xlsx"):
            try:
                antigo.unlink()
            except:
                pass

        exportar.click()
        print("Exportando Excel...")

        # Esperar arquivo
        timeout = 300
        inicio_tempo = time.time()
        novo_arquivo = None

        while time.time() - inicio_tempo < timeout:
            arquivos = list(downloads.glob("report_tickets*.xlsx"))
            if arquivos:
                candidato = max(arquivos, key=os.path.getctime)
                if not Path(str(candidato) + ".crdownload").exists():
                    novo_arquivo = candidato
                    break
            time.sleep(2)

        if not novo_arquivo:
            print(f"Arquivo não apareceu para {nome_sistema}")
            falhas_login.append(base_url)
            continue

        destino_final = pasta_destino / f"{nome_sistema}.xlsx"

        contador = 2
        while destino_final.exists():
            destino_final = pasta_destino / f"{nome_sistema}_{contador}.xlsx"
            contador += 1

        shutil.move(str(novo_arquivo), str(destino_final))
        print(f"Arquivo salvo: {destino_final}")

    except Exception as e:
        print(f"Erro no fluxo da página {base_url}: {e}")
        falhas_login.append(base_url)
        continue

# ==========================================================
# CONSOLIDAÇÃO
# ==========================================================
dfs = []

for arq in pasta_destino.glob("*.xlsx"):
    if "CONSOLIDADO" in arq.name:
        continue
    try:
        df = pd.read_excel(arq)
        df["ARQUIVO_ORIGEM"] = arq.name
        dfs.append(df)
    except Exception as e:
        print(f"Erro ao ler {arq.name}: {e}")

if dfs:
    consolidado = pd.concat(dfs, ignore_index=True)
    consolidado.to_excel(pasta_destino / "CONSOLIDADO.xlsx", index=False)
    print("CONSOLIDADO.xlsx gerado com sucesso.")
else:
    print("Nenhum arquivo para consolidar.")

# ==========================================================
# LOG
# ==========================================================
if falhas_login:
    log_path = pasta_destino / "log.txt"
    with open(log_path, "w") as f:
        for url in falhas_login:
            f.write(f"Falha ao acessar: {url}\n")
    print(f"log.txt gerado com {len(falhas_login)} falha(s).")

driver.quit()

print("Automação finalizada.")
