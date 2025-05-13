import os
import re
import logging
import datetime
import win32com.client

# === CONFIGURAÇÕES ===
# Caminho da pasta a ser escaneada
PASTA_ARQUIVOS = r"C:\\Caminho\\Para\\Arquivos"  # <-- Ajuste para o seu caminho

# Dicionário de e-mails por cliente
EMAILS_CLIENTES = {
    "Cliente A": "clientea@exemplo.com",
    "Cliente B": "clienteb@exemplo.com",
    # Adicione mais clientes conforme necessário
}

# Corpo do e-mail
CORPO_EMAIL = """
Prezado(a),\n\nSegue em anexo os documentos referentes ao mês.\n\nAtenciosamente,\nSua Empresa
"""

# === LOGGING ===
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Regex para extrair informações do arquivo
PADRAO_ARQUIVO = re.compile(r"^(\\d+) - (.+?) - (.+)\\.[^.]+$")

def obter_mes_ano():
    agora = datetime.datetime.now()
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return f"{meses[agora.month-1]}/{agora.year}"

def escanear_pasta(pasta):
    arquivos_por_cliente = {}
    for raiz, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            match = PADRAO_ARQUIVO.match(arquivo)
            if match:
                numero, cliente, tipo = match.groups()
                caminho_arquivo = os.path.join(raiz, arquivo)
                arquivos_por_cliente.setdefault(cliente, []).append(caminho_arquivo)
    return arquivos_por_cliente

def enviar_email(cliente, anexos):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    email_destino = EMAILS_CLIENTES.get(cliente)
    if not email_destino:
        raise ValueError(f"E-mail não cadastrado para o cliente: {cliente}")
    mail.To = email_destino
    mail.Subject = f"Documentos Referentes - {cliente} - {obter_mes_ano()}"
    mail.Body = CORPO_EMAIL
    for anexo in anexos:
        mail.Attachments.Add(anexo)
    mail.Send()

def main():
    logging.info("Iniciando escaneamento da pasta...")
    arquivos_clientes = escanear_pasta(PASTA_ARQUIVOS)
    if not arquivos_clientes:
        logging.info("Nenhum arquivo encontrado para envio.")
        return
    for cliente, anexos in arquivos_clientes.items():
        try:
            enviar_email(cliente, anexos)
            logging.info(f"E-mail enviado com sucesso para {cliente} ({EMAILS_CLIENTES.get(cliente, 'e-mail não cadastrado')}) com {len(anexos)} anexos.")
        except Exception as e:
            logging.error(f"Erro ao enviar e-mail para {cliente}: {e}")

if __name__ == "__main__":
    main()
