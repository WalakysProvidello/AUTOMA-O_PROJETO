import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import threading
from datetime import datetime, timedelta
import win32com.client as win32

# ===============================
# CONFIGURAÇÕES
# ===============================

SCRIPTS = [
    "EXT_BRAGI.py",
    "EXT_CDA.py",
    "EXT_CDA_DIGITAL.py",
    "EXT_CDA_MMS.py",
    "EXT_GOSAC.py",
    "EXT_SALES.py"
]

PASTA_DIGITAL = r"C:\CAMINHO\PARA\DIGITAL"  # ALTERAR AQUI

DESTINATARIOS = "julianecaraca@concilig.com.br; joicerodrigues@concilig.com.br; guilhermedasilva@concilig.com.br"
CC = "lucascavaguti@concilig.com.br; walakysprovidello@concilig.com.br"


# ===============================
# FUNÇÃO DATA D-1
# ===============================

def calcular_data_d1():
    hoje = datetime.today()

    if hoje.weekday() == 0:  # Segunda
        d1 = hoje - timedelta(days=3)
    else:
        d1 = hoje - timedelta(days=1)

    return d1.strftime("%d/%m/%Y")


# ===============================
# EXECUTAR SCRIPTS
# ===============================

def executar_scripts():
    btn_iniciar.config(state="disabled")
    log("Iniciando execução...\n")

    for script in SCRIPTS:
        try:
            log(f"Executando {script}...\n")
            subprocess.run(["python", script], check=True)
            log(f"{script} finalizado com sucesso.\n\n")
        except subprocess.CalledProcessError:
            log(f"Erro ao executar {script}\n")
            messagebox.showerror("Erro", f"Erro ao executar {script}")
            btn_iniciar.config(state="normal")
            return

    log("Todos scripts finalizados.\n")
    enviar_email()
    btn_iniciar.config(state="normal")


# ===============================
# ENVIAR EMAIL OUTLOOK
# ===============================

def enviar_email():
    try:
        log("Abrindo Outlook...\n")

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        data_d1 = calcular_data_d1()

        mail.To = DESTINATARIOS
        mail.CC = CC
        mail.Subject = f"BASES PARA ANALISE DISPARO MASSIVO HOJE - {data_d1}"

        mail.Body = f"""
Prezados,

Segue em anexo as bases para análise do disparo massivo referente à data {data_d1}.

Atenciosamente,
        """

        # Anexar arquivos da pasta DIGITAL
        for arquivo in os.listdir(PASTA_DIGITAL):
            caminho_arquivo = os.path.join(PASTA_DIGITAL, arquivo)
            if os.path.isfile(caminho_arquivo):
                mail.Attachments.Add(caminho_arquivo)

        mail.Display()  # Abre para revisão (usar mail.Send() para enviar direto)

        log("Email criado com sucesso.\n")

    except Exception as e:
        messagebox.showerror("Erro Email", str(e))


# ===============================
# LOG
# ===============================

def log(mensagem):
    txt_log.insert(tk.END, mensagem)
    txt_log.see(tk.END)


def iniciar_thread():
    thread = threading.Thread(target=executar_scripts)
    thread.start()


# ===============================
# INTERFACE
# ===============================

app = tk.Tk()
app.title("DISPARO MASSIVO - AUTOMAÇÃO")
app.geometry("700x500")
app.configure(bg="#1e1e2f")

titulo = tk.Label(app, text="AUTOMAÇÃO DISPARO MASSIVO", 
                  font=("Segoe UI", 16, "bold"), 
                  bg="#1e1e2f", fg="white")
titulo.pack(pady=15)

btn_iniciar = ttk.Button(app, text="Executar Processos", command=iniciar_thread)
btn_iniciar.pack(pady=10)

frame_log = tk.Frame(app)
frame_log.pack(padx=20, pady=10, fill="both", expand=True)

txt_log = tk.Text(frame_log, bg="black", fg="lime", font=("Consolas", 10))
txt_log.pack(fill="both", expand=True)

app.mainloop()
