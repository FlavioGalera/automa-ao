#!/usr/bin/env python3
import imaplib
import email
import sys
from datetime import datetime
from email.header import decode_header

# ==============================================
PROVEDORES = {
    "gmail": {"host": "imap.gmail.com", "port": 993},
    "outlook": {"host": "imap-mail.outlook.com", "port": 993},
    "hotmail": {"host": "imap-mail.outlook.com", "port": 993},
    "icloud": {"host": "imap.mail.me.com", "port": 993},
    "yahoo": {"host": "imap.mail.yahoo.com", "port": 993}
}
LIMITE_EXIBIR = 20
# ==============================================

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
    elif "icloud" in dominio or "me.com" in dominio: return "icloud"
    elif "yahoo" in dominio: return "yahoo"
    return None

def conectar_email(email_usuario, senha_app):
    provedor = detectar_provedor(email_usuario)
    if not provedor:
        print("❌ Provedor não reconhecido.")
        return None
    config = PROVEDORES[provedor]
    try:
        mail = imaplib.IMAP4_SSL(config["host"], config["port"])
        mail.login(email_usuario, senha_app)
        return mail
    except imaplib.IMAP4.error:
        print("❌ E-mail ou senha de app incorretos. Tente novamente.")
        return None
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        return None

def listar_emails(email_usuario, senha_app):
    print(f"📧 Conectando em: {email_usuario}")
    print("-" * 60)

    mail = conectar_email(email_usuario, senha_app)
    if not mail:
        return

    try:
        status, dados = mail.select("INBOX", readonly=True)
        total = int(dados[0])
        print(f"✅ Conectado! Total na caixa de entrada: {total} e-mail(s)\n")

        status, ids = mail.search(None, 'ALL')
        lista_ids = ids[0].split()

        if not lista_ids:
            print("ℹ️ Sua caixa de entrada está vazia.")
            mail.logout()
            return

        qtd_pegar = min(LIMITE_EXIBIR, len(lista_ids))
        ids_recentes = lista_ids[-qtd_pegar:]

        emails_lista = []
        for email_id in ids_recentes:
            status, dados_msg = mail.fetch(email_id, "(RFC822.HEADER)")
            if status != "OK":
                continue
            msg = email.message_from_bytes(dados_msg[0][1])
            assunto = decodificar_cabecalho(msg["Subject"])
            remetente = decodificar_cabecalho(msg["From"])
            data_obj = parse_data_email(msg["Date"])

            emails_lista.append({
                "assunto": assunto,
                "remetente": remetente,
                "data_obj": data_obj
            })

        mail.logout()

        # ✅ Mais novo primeiro
        emails_lista.sort(
            key=lambda x: x["data_obj"] or datetime.min,
            reverse=True
        )

        print("=" * 60)
        print(f"📧 Últimos {len(emails_lista)} e-mails — do mais novo ao mais antigo")
        print("=" * 60)

        for pos, item in enumerate(emails_lista, 1):
            data_exib = item["data_obj"].strftime("%d/%m/%Y %H:%M") if item["data_obj"] else "Data desconhecida"
            print(f"\n{pos}. 📅 {data_exib}")
            print(f"   De:    {item['remetente']}")
            print(f"   Assunto: {item['assunto']}")

        print("\n" + "=" * 60)
        print(f"✅ Concluído! Exibidos os {len(emails_lista)} mais recentes de {total} total.")

    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        try:
            mail.logout()
        except:
            pass

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        email_usuario = sys.argv[1]
        senha_app = sys.argv[2]
        listar_emails(email_usuario, senha_app)
    else:
        print("❌ Parâmetros insuficientes.")

