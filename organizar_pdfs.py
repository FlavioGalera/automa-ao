#!/usr/bin/env python3
import os
import sys
import shutil
from datetime import datetime

# ================= CONFIGURAÇÕES =================
if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
    PASTA_ORIGEM = sys.argv[1]
    print(f"📂 Pasta escolhida: {PASTA_ORIGEM}")
else:
    PASTA_ORIGEM = "/home/flavio/Área de trabalho"
    print(f"📂 Usando pasta padrão: {PASTA_ORIGEM}")

SOMENTE_PDFS = True
FORMATO_PASTA_DATA = "%Y-%m"
# =================================================

def organizar_pdfs():
    if not os.path.exists(PASTA_ORIGEM):
        print(f"❌ Pasta não encontrada: {PASTA_ORIGEM}")
        return

    print("-" * 50)
    todos_itens = os.listdir(PASTA_ORIGEM)
    print(f"📋 Total de itens na pasta: {len(todos_itens)}")

    arquivos_pdf = []
    for item in todos_itens:
        caminho_item = os.path.join(PASTA_ORIGEM, item)
        if os.path.isfile(caminho_item) and item.lower().endswith(".pdf"):
            arquivos_pdf.append(item)

    print(f"📄 PDFs encontrados: {len(arquivos_pdf)}")

    if len(arquivos_pdf) == 0:
        print("⚠️ Nenhum PDF encontrado! Arquivos na pasta:")
        for a in todos_itens[:15]:
            print(f"   - {a}")
        return

    contador = 0
    for nome_arquivo in arquivos_pdf:
        caminho_completo = os.path.join(PASTA_ORIGEM, nome_arquivo)
        
        data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_completo))
        nome_pasta = data_mod.strftime(FORMATO_PASTA_DATA)
        pasta_destino = os.path.join(PASTA_ORIGEM, nome_pasta)

        os.makedirs(pasta_destino, exist_ok=True)

        caminho_destino = os.path.join(pasta_destino, nome_arquivo)

        cont = 1
        while os.path.exists(caminho_destino):
            base, ext = os.path.splitext(nome_arquivo)
            caminho_destino = os.path.join(pasta_destino, f"{base} ({cont}){ext}")
            cont += 1

        shutil.move(caminho_completo, caminho_destino)
        print(f"✅ {nome_arquivo} → {nome_pasta}/")
        contador += 1

    print("-" * 50)
    print(f"🎉 Concluído! {contador} arquivo(s) organizado(s).")

if __name__ == "__main__":
    organizar_pdfs()
