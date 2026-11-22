#!/usr/bin/env python3
import imaplib
import email
import sys
from datetime import datetime
from email.header import decode_header

PROVEDORES = {
    "gmail": {"host": "imap.gmail.com", "port": 993, "nome": "Gmail"},
    "outlook": {"host": "imap-mail.outlook.com", "port": 993, "nome": "Outlook / Hotmail"},
    "yahoo": {"host": "imap.mail.yahoo.com", "port": 993, "nome": "Yahoo"},
    "icloud": {"host": "imap.mail.me.com", "port": 993, "nome": "iCloud"}
}

def decodificar_cabecalho(valor):
    if not valor:
        return "(Sem assunto)"
    decodificado = decode_header(valor)
    partes = []
    for texto, encoding in decodificado:
        if isinstance(texto, bytes):
            partes.append(texto.decode(encoding or 'utf-8', errors='replace'))
        else:
            partes.append(str(texto))
    return "".join(partes)

def parse_data_email(data_email):
    try:
        return email.utils.parsedate_to_datetime(data_email)
    except:
        return None

def detectar_provedor(email_usuario):
    dominio = email_usuario.lower().split('@')[-1]
    if "gmail" in dominio: return "gmail"
    elif "outlook" in dominio or "hotmail" in dominio or "live" in dominio: return "outlook"
    elif "yahoo" in dominio: return "yahoo"
    elif "icloud" in dominio or "me.com" in dominio: return "icloud"
    return None

def conectar_email(email_usuario, senha_app, provedor_escolhido=None):
    if not provedor_escolhido or provedor_escolhido == "auto":
        provedor_escolhido = detectar_provedor(email_usuario)
    
    if not provedor_escolhido:
        print("❌ Não consegui identificar seu provedor. Selecione manualmente.")
        return None, None
    
    if provedor_escolhido not in PROVEDORES:
        print(f"❌ Provedor '{provedor_escolhido}' não suportado.")
        return None, None
    
    config = PROVEDORES[provedor_escolhido]
    try:
        mail = imaplib.IMAP4_SSL(config["host"], config["port"])
        mail.login(email_usuario, senha_app)
        return mail, config["nome"]
    except imaplib.IMAP4.error:
        print(f"❌ Falha no login! Verifique seu e-mail e a Senha de App.")
        print(f"   💡 Lembre-se: não use sua senha normal, crie uma Senha de App.")
        return None, None
    except Exception as e:
        print(f"❌ Erro de conexão: {str(e)}")
        return None, None

def listar_emails(email_usuario, senha_app, limite=20, provedor_escolhido=None):
    mail, nome_provedor = conectar_email(email_usuario, senha_app, provedor_escolhido)
    if not mail:
        return

    print(f"📧 Conectando em: {email_usuario}")
    print("-" * 60)

    try:
        mail.select("INBOX", readonly=True)
        status, dados = mail.search(None, "ALL")
        
        if status != "OK":
            print("❌ Erro ao listar e-mails.")
            mail.logout()
            return

        ids_emails = dados[0].split()
        total = len(ids_emails)
        print(f"✅ Conectado! Total na caixa de entrada: {total} e-mail(s)")

        # Pegar os últimos 'limite' e-mails
        ids_ultimos = ids_emails[-limite:] if total >= limite else ids_emails
        ids_ultimos.reverse()  # Do mais recente para o mais antigo

        print("=" * 60)
        print(f"📧 Últimos {len(ids_ultimos)} e-mails — do mais recente → mais antigo")
        print("=" * 60)

        for pos, email_id in enumerate(ids_ultimos, 1):
            status, dados_msg = mail.fetch(email_id, "(RFC822.HEADER)")
            if status != "OK":
                continue
            msg = email.message_from_bytes(dados_msg[0][1])
            assunto = decodificar_cabecalho(msg["Subject"])
            remetente = decodificar_cabecalho(msg["From"])
            data_obj = parse_data_email(msg["Date"])
            data_exib = data_obj.strftime("%d/%m/%Y %H:%M") if data_obj else "Data desconhecida"

            print(f"\n{pos}. 📅 {data_exib}")
            print(f"   De:    {remetente}")
            print(f"   Assunto: {assunto}")

        mail.logout()
        print("\n" + "=" * 60)
        print(f"✅ Concluído! Exibidos os {len(ids_ultimos)} mais recentes de {total} total.")

    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        try: mail.logout()
        except: pass

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        email_usuario = sys.argv[1]
        senha_app = sys.argv[2]
        listar_emails(email_usuario, senha_app)
    else:
        print("❌ Parâmetros insuficientes.")
