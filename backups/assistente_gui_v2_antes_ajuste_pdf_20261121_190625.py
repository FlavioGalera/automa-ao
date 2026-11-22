import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext, simpledialog
import os
import re
import sys
import shutil
import json
import imaplib
import webbrowser
import email
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
import threading
import subprocess
from html.parser import HTMLParser


APP_VERSION = "1.1.4"
APP_NAME = "Sistema de Automação"
APP_AUTHOR = "Flávio Galera"
APP_AI_DISCLOSURE = "Desenvolvido com auxílio de Inteligência Artificial (ChatGPT)."


# ==========================================
# 🛠️ FUNÇÕES AUXILIARES
# ==========================================
class ExtrairTextoHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.texto = []
        self.dentro_script = False
        self.dentro_estilo = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            if tag == "script":
                self.dentro_script = True
            else:
                self.dentro_estilo = True
        elif tag in ("p", "br", "div", "h1", "h2", "h3", "tr"):
            self.texto.append("\n")

    def handle_endtag(self, tag):
        if tag == "script":
            self.dentro_script = False
        elif tag == "style":
            self.dentro_estilo = False

    def handle_data(self, dados):
        if not self.dentro_script and not self.dentro_estilo:
            self.texto.append(dados.strip() + " ")

    def pegar_texto(self):
        texto = "".join(self.texto)
        linhas = [linha.strip() for linha in texto.split("\n") if linha.strip()]
        return "\n".join(linhas)

def extrair_texto_do_html(html):
    if not html:
        return "(Sem conteúdo)"
    if "<" not in html and ">" not in html:
        return html.strip()
    parser = ExtrairTextoHTML()
    try:
        parser.feed(html)
        resultado = parser.pegar_texto()
        return resultado if resultado else "(Sem conteúdo de texto)"
    except Exception as e:
        return html.strip() if html else "(Sem conteúdo)"


def diretorio_dados_aplicativo():
    """Diretório gravável do usuário, também quando o app estiver empacotado."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.join(
            os.path.expanduser("~"), ".local", "share"
        )
    pasta = os.path.join(base, "SistemaAutomacao")
    os.makedirs(pasta, exist_ok=True)
    if os.name != "nt":
        try:
            os.chmod(pasta, 0o700)
        except OSError:
            pass
    return pasta


def caminho_arquivo_historico():
    """Mantém o histórico ao lado do .py em desenvolvimento e no perfil do usuário no app instalado."""
    if getattr(sys, "frozen", False):
        return os.path.join(diretorio_dados_aplicativo(), "historico.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico.json")


def salvar_no_historico(tipo, detalhe, registros):
    """Salva no máximo 100 ações, com gravação atômica e acesso privado no Linux."""
    import tempfile

    arquivo = caminho_arquivo_historico()
    historico = []
    if os.path.exists(arquivo):
        try:
            with open(arquivo, "r", encoding="utf-8") as f:
                dados_existentes = json.load(f)
            if isinstance(dados_existentes, list):
                historico = dados_existentes
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            # Um histórico inválido não deve impedir o uso do restante do sistema.
            historico = []

    entrada = {
        "data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "tipo": str(tipo)[:120],
        "detalhe": str(detalhe)[:1000],
        "quantidade": registros if isinstance(registros, int) else 0,
    }
    historico.insert(0, entrada)
    historico = historico[:100]

    pasta = os.path.dirname(arquivo)
    os.makedirs(pasta, exist_ok=True)
    fd, temporario = tempfile.mkstemp(prefix="historico_", suffix=".tmp", dir=pasta)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporario, arquivo)
        if os.name != "nt":
            os.chmod(arquivo, 0o600)
    finally:
        if os.path.exists(temporario):
            os.remove(temporario)


# ==========================================
# 📊 DETECÇÃO DO FORMATO DAS PLANILHAS EXCEL
# ==========================================
MESES_EXCEL = {
    1: ("JAN", "JANEIRO", "JANUARY"),
    2: ("FEV", "FEB", "FEVEREIRO", "FEBRUARY"),
    3: ("MAR", "MARCO", "MARCH"),
    4: ("ABR", "APR", "ABRIL", "APRIL"),
    5: ("MAI", "MAY", "MAIO"),
    6: ("JUN", "JUNE", "JUNHO"),
    7: ("JUL", "JULY", "JULHO"),
    8: ("AGO", "AUG", "AGOSTO", "AUGUST"),
    9: ("SET", "SEP", "SEPT", "SETEMBRO", "SEPTEMBER"),
    10: ("OUT", "OCT", "OUTUBRO", "OCTOBER"),
    11: ("NOV", "NOVEMBRO", "NOVEMBER"),
    12: ("DEZ", "DEC", "DEZEMBRO", "DECEMBER"),
}


def normalizar_celula_excel(valor):
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", texto.lower()).strip()


def detectar_mes_excel(valor):
    """Retorna o número do mês quando o cabeçalho contém JAN, JANEIRO etc."""
    texto = normalizar_celula_excel(valor)
    if not texto or texto in {"nan", "none"}:
        return None
    partes = texto.split()
    for numero, aliases in MESES_EXCEL.items():
        if any(parte.upper() in aliases for parte in partes):
            return numero
    return None


def detectar_layout_excel(df_bruto):
    """Detecta cabeçalho de transações ou cabeçalho de categorias x meses.

    Lê a planilha sem presumir que a primeira linha seja o cabeçalho. Isso
    permite planilhas com título mesclado, como 'CONTAS DO ANO'.
    Retorna (tipo, índice_da_linha_de_cabeçalho, colunas_de_mês).
    """
    limite = min(len(df_bruto), 50)
    for indice_linha in range(limite):
        valores = df_bruto.iloc[indice_linha].tolist()
        nomes = [normalizar_celula_excel(valor) for valor in valores]

        tem_data = any(
            nome == "data" or nome.startswith("data ") or nome == "date" or nome.startswith("date ")
            for nome in nomes
        )
        tem_valor = any(
            nome == "valor" or nome.startswith("valor ") or
            nome in {"value", "amount", "preco", "importe", "montante"} or
            nome.startswith(("value ", "amount ", "preco ", "importe ", "montante "))
            for nome in nomes
        )
        if tem_data and tem_valor:
            return "transacoes", indice_linha, []

        colunas_mes = []
        meses_vistos = set()
        for indice_coluna, valor in enumerate(valores):
            mes = detectar_mes_excel(valor)
            if mes is not None:
                colunas_mes.append((mes, indice_coluna, valor))
                meses_vistos.add(mes)

        # Duas ou mais colunas mensais distinguem a matriz anual de gastos.
        if len(colunas_mes) >= 2 and len(meses_vistos) >= 2:
            return "mensal", indice_linha, colunas_mes

    return None, None, []


# ==========================================
# 📋 SCRIPTS DISPONÍVEIS
# ==========================================
SCRIPTS = {
    "organizar_pdfs": {
        "titulo": "Organizar PDFs por Data",
        "descricao": "Lista todos os PDFs de uma pasta e organiza por data"
    },
    "buscar_email": {
        "titulo": "Buscar E-mails",
        "descricao": "Busca e filtra e-mails por remetente ou assunto"
    },
    "processar_excel": {
        "titulo": "Consolidar Planilhas Excel",
        "descricao": "Junta múltiplos arquivos em um só e calcula totais"
    },
    "analise_gastos": {
        "titulo": "📊 Análise de Gastos Mensais",
        "descricao": "Calcula totais por categoria, por mês e total geral"
    }
}


def chave_categoria_numerica(nome):
    """Ordena categorias pelo número inicial do nome (1, 2, 3... 10)."""
    texto = str(nome).strip()
    encontrado = re.match(r"^(\d+)", texto)
    if encontrado:
        return (0, int(encontrado.group(1)))
    return (1, 0)


FORMATO_EURO_EXCEL = '"€" #,##0.00;[Red]-"€" #,##0.00;"€" 0.00'
FORMATO_PERCENTUAL_EXCEL = '0.00"%"'


def adicionar_graficos_resumo(wb, ws_mensal, ws_anual, coluna_total_anual):
    """Adiciona gráficos mensais e anuais ao arquivo Excel.

    ws_mensal: aba com Mês na coluna A e Total do Mês na coluna B.
    ws_anual: aba com Categoria na coluna A e Total Anual na coluna informada.
    As linhas TOTAL GERAL são excluídas dos gráficos para evitar duplicidade.
    """
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    ultima_linha_mensal = ws_mensal.max_row - 1
    ultima_linha_anual = ws_anual.max_row - 1

    if ultima_linha_mensal < 2 or ultima_linha_anual < 2:
        return False

    # Cria uma aba separada para os gráficos, sem esconder os dados originais.
    if "Gráficos" in wb.sheetnames:
        del wb["Gráficos"]
    ws_graficos = wb.create_sheet("Gráficos")
    ws_graficos.sheet_view.showGridLines = False

    # Gráfico 1: total gasto em cada mês, em barras horizontais.
    # O formato horizontal segue o mesmo padrão visual do gráfico anual.
    grafico_mensal = BarChart()
    grafico_mensal.type = "bar"
    grafico_mensal.style = 11
    grafico_mensal.title = "Gastos por mês (€)"
    grafico_mensal.x_axis.title = "Total gasto (€)"
    grafico_mensal.y_axis.title = "Mês"
    # Os gráficos ficam empilhados verticalmente em uma aba exclusiva.
    # Isso impede que um gráfico cubra o outro ou qualquer tabela do relatório.
    grafico_mensal.height = 10
    grafico_mensal.width = 22
    grafico_mensal.legend = None
    dados_mensais = Reference(
        ws_mensal, min_col=2, min_row=1, max_row=ultima_linha_mensal
    )
    categorias_mensais = Reference(
        ws_mensal, min_col=1, min_row=2, max_row=ultima_linha_mensal
    )
    grafico_mensal.add_data(dados_mensais, titles_from_data=True)
    grafico_mensal.set_categories(categorias_mensais)
    grafico_mensal.x_axis.numFmt = '"€" #,##0.00'
    # Em barras horizontais, maxMin coloca JAN no topo e DEZ no final.
    grafico_mensal.y_axis.scaling.orientation = "maxMin"
    grafico_mensal.dLbls = DataLabelList()
    grafico_mensal.dLbls.showVal = True
    grafico_mensal.dLbls.showCatName = False
    grafico_mensal.dLbls.showSerName = False
    grafico_mensal.dLbls.showLegendKey = False
    grafico_mensal.dLbls.showPercent = False
    grafico_mensal.dLbls.showLeaderLines = False
    ws_graficos.add_chart(grafico_mensal, "A1")

    # Gráfico 2: total anual por categoria, sem a linha TOTAL GERAL.
    grafico_anual = BarChart()
    grafico_anual.type = "bar"
    grafico_anual.style = 11
    grafico_anual.title = "Gasto anual por categoria (€)"
    # Em gráficos horizontais, o openpyxl usa o eixo X para categorias
    # e o eixo Y para valores; os títulos são definidos nessa orientação.
    grafico_anual.x_axis.title = "Categoria"
    grafico_anual.y_axis.title = "Total anual (€)"
    grafico_anual.height = 10
    grafico_anual.width = 22
    grafico_anual.legend = None
    dados_anuais = Reference(
        ws_anual,
        min_col=coluna_total_anual,
        min_row=1,
        max_row=ultima_linha_anual,
    )
    categorias_anuais = Reference(
        ws_anual, min_col=1, min_row=2, max_row=ultima_linha_anual
    )
    grafico_anual.add_data(dados_anuais, titles_from_data=True)
    grafico_anual.set_categories(categorias_anuais)
    grafico_anual.x_axis.numFmt = '"€" #,##0.00'
    grafico_anual.dLbls = DataLabelList()
    grafico_anual.dLbls.showVal = True
    grafico_anual.dLbls.showCatName = False
    grafico_anual.dLbls.showSerName = False
    grafico_anual.dLbls.showLegendKey = False
    grafico_anual.dLbls.showPercent = False
    grafico_anual.dLbls.showLeaderLines = False
    # A altura de 10 cm ocupa aproximadamente 19-20 linhas. A segunda figura
    # começa bem abaixo, deixando uma margem confortável entre os gráficos.
    ws_graficos.add_chart(grafico_anual, "A28")

    ws_graficos.column_dimensions["A"].width = 22
    ws_graficos.sheet_properties.pageSetUpPr.fitToPage = True
    ws_graficos.page_setup.fitToWidth = 1
    ws_graficos.page_setup.fitToHeight = 0
    return True

# Cofre local criptografado: não depende do Chaveiro do GNOME.
ARQUIVO_COFRE_EMAIL = "credenciais_email.enc"
PBKDF2_ITERACOES = 600_000


# ==========================================
# 🏠 CLASSE PRINCIPAL DA APLICAÇÃO
# ==========================================


TRADUCOES_INTERFACE = {'📊 Sistema de Automação': ('📊 Automation System', '📊 Sistema de Automatización'), 'Sobre / Créditos': ('About / Credits', 'Acerca de / Créditos'), '📋 Log de Execução': ('📋 Execution Log', '📋 Registro de Ejecución'), '📋 Tarefas': ('📋 Tasks', '📋 Tareas'), '📜 Histórico': ('📜 History', '📜 Historial'), 'ℹ️ Selecione uma tarefa no menu à esquerda': ('ℹ️ Select a task from the menu on the left', 'ℹ️ Seleccione una tarea en el menú de la izquierda'), 'Organizar PDFs por Data': ('Organize PDFs by Date', 'Organizar PDF por fecha'), 'Lista todos os PDFs de uma pasta e organiza por data': ('Lists all PDFs in a folder and sorts them by date', 'Lista todos los PDF de una carpeta y los organiza por fecha'), 'Buscar E-mails': ('Search Emails', 'Buscar correos electrónicos'), 'Busca e filtra e-mails por remetente ou assunto': ('Searches and filters emails by sender or subject', 'Busca y filtra correos por remitente o asunto'), 'Consolidar Planilhas Excel': ('Consolidate Excel Spreadsheets', 'Consolidar hojas de cálculo de Excel'), 'Junta múltiplos arquivos em um só e calcula totais': ('Combines multiple files into one and calculates totals', 'Combina varios archivos en uno y calcula los totales'), '📊 Análise de Gastos Mensais': ('📊 Monthly Expense Analysis', '📊 Análisis de gastos mensuales'), 'Calcula totais por categoria, por mês e total geral': ('Calculates totals by category, month, and grand total', 'Calcula totales por categoría, mes y total general'), '📂 Organizar PDFs': ('📂 Organize PDFs', '📂 Organizar PDF'), 'Pasta:': ('Folder:', 'Carpeta:'), 'Nenhuma pasta selecionada': ('No folder selected', 'Ninguna carpeta seleccionada'), '📁 Escolher Pasta': ('📁 Choose Folder', '📁 Elegir carpeta'), '▶️ Listar PDFs': ('▶️ List PDFs', '▶️ Listar PDF'), '⏹️ Parar': ('⏹️ Stop', '⏹️ Detener'), '🔎 Localizar PDF:': ('🔎 Find PDF:', '🔎 Buscar PDF:'), 'Limpar': ('Clear', 'Limpiar'), 'Data': ('Date', 'Fecha'), 'Nome do Arquivo': ('File Name', 'Nombre del archivo'), 'Caminho Completo': ('Full Path', 'Ruta completa'), 'Nenhum PDF listado': ('No PDFs listed', 'No hay PDF en la lista'), '📄 Abrir PDF': ('📄 Open PDF', '📄 Abrir PDF'), '📂 Abrir Pasta': ('📂 Open Folder', '📂 Abrir carpeta'), '📧 Buscar E-mails': ('📧 Search Emails', '📧 Buscar correos'), '🔐 Dados de Acesso': ('🔐 Login Details', '🔐 Datos de acceso'), 'E-mail:': ('Email:', 'Correo electrónico:'), 'Senha de App:': ('App Password:', 'Contraseña de aplicación:'), '💾 Salvar acesso': ('💾 Save account', '💾 Guardar acceso'), '📋 Escolher acesso salvo': ('📋 Choose saved account', '📋 Elegir acceso guardado'), '❓ Como gerar senha de app?': ('❓ How to create an app password?', '❓ ¿Cómo crear una contraseña de aplicación?'), 'Use a senha de app do Google (16 caracteres), não a senha normal da conta.': ('Use your Google app password (16 characters), not your regular account password.', 'Use la contraseña de aplicación de Google (16 caracteres), no la contraseña normal de la cuenta.'), '🔍 Filtros de Busca': ('🔍 Search Filters', '🔍 Filtros de búsqueda'), 'Nome/Remetente:': ('Name/Sender:', 'Nombre/Remitente:'), 'Tipo de Documento:': ('Document Type:', 'Tipo de documento:'), 'Últimos:': ('Period:', 'Período:'), '▶️ Buscar': ('▶️ Search', '▶️ Buscar'), '👁️ Visualizar e-mail': ('👁️ View email', '👁️ Ver correo'), '💾 Salvar e-mail': ('💾 Save email', '💾 Guardar correo'), 'Remetente': ('Sender', 'Remitente'), 'Assunto': ('Subject', 'Asunto'), '📊 Consolidar Planilhas Excel': ('📊 Consolidate Excel Spreadsheets', '📊 Consolidar hojas de Excel'), 'Selecione um ou mais arquivos para juntar em um único relatório.': ('Select one or more files to combine into a single report.', 'Seleccione uno o más archivos para unirlos en un solo informe.'), 'Arquivos:': ('Files:', 'Archivos:'), 'Nenhum arquivo selecionado': ('No files selected', 'Ningún archivo seleccionado'), '📂 Selecionar': ('📂 Select', '📂 Seleccionar'), '▶️ Processar': ('▶️ Process', '▶️ Procesar'), '📂 Abrir Resultado': ('📂 Open Result', '📂 Abrir resultado'), 'Selecione sua planilha para visualizar os valores em euros e os gráficos. Depois, clique em Calcular Análise para salvar o relatório XLSX.': ('Select your spreadsheet to view values in euros and charts. Then click Calculate Analysis to save the XLSX report.', 'Seleccione su hoja de cálculo para ver los valores en euros y los gráficos. Después, haga clic en Calcular análisis para guardar el informe XLSX.'), 'Planilha:': ('Spreadsheet:', 'Hoja de cálculo:'), '📂 Selecionar XLSX': ('📂 Select XLSX', '📂 Seleccionar XLSX'), '▶️ Calcular Análise e Salvar XLSX': ('▶️ Calculate Analysis and Save XLSX', '▶️ Calcular análisis y guardar XLSX'), '📂 Abrir Relatório': ('📂 Open Report', '📂 Abrir informe'), '📋 Histórico': ('📋 History', '📋 Historial')}

class Aplicacao:
    def __init__(self, janela_principal):
        self.root = janela_principal
        self.root.title(f"📊 {APP_NAME} — v{APP_VERSION}")
        self.root.geometry("1320x720")
        self.root.minsize(1100, 620)

        self.script_atual = None
        self.idioma = "pt"
        self._widgets_traduziveis = []
        self.pasta_selecionada = ""
        self.caminhos_selecionados = []
        self.caminho_ultimo_excel = ""
        self.caminho_planilha_gastos = ""
        self.caminho_relatorio_gastos = ""
        self.dados_gastos_preview = None
        self.gastos_tree = None
        self.gastos_canvas_widget = None
        self.gastos_figure = None
        self.emails_carregados = []
        self.lista_pdfs_encontrados = []
        self.lista_pdfs_exibidos = []
        self.buscando = False
        self._cofre_email = None  # Desbloqueado somente durante esta sessão do programa.

        self.var_tipo_doc = tk.StringVar(value="todos")
        self.var_periodo_texto = tk.StringVar(value="Todos os e-mails")
        self.opcoes_periodo = {
            "0": "Todos os e-mails",
            "7": "Últimos 7 dias",
            "30": "Últimos 30 dias",
            "90": "Últimos 90 dias",
            "365": "Último ano"
        }

        self.btn_executar = None
        self.btn_parar = None

        self.criar_interface()
        self.log("✅ Sistema pronto! Selecione uma tarefa à esquerda.")

    def seguro_config(self, widget, **kwargs):
        try:
            if widget and widget.winfo_exists():
                widget.config(**kwargs)
        except:
            pass

    def log(self, mensagem):
        self.saida_log.config(state=tk.NORMAL)
        self.saida_log.insert(tk.END, mensagem + "\n")
        self.saida_log.see(tk.END)
        self.saida_log.config(state=tk.DISABLED)

    def decodificar_cabecalho(self, valor):
        if not valor: return ""
        try:
            partes = email.header.decode_header(valor)
            decodificado = []
            for texto, codificacao in partes:
                if isinstance(texto, bytes):
                    decodificado.append(texto.decode(codificacao or 'utf-8', errors='replace'))
                else:
                    decodificado.append(str(texto))
            return "".join(decodificado)
        except:
            return str(valor)

    def criar_interface(self):
        # Barra superior de identificação e créditos do projeto.
        # A informação sobre IA é apresentada de forma transparente:
        # o projeto é idealizado pelo autor e desenvolvido com apoio de IA.
        self.barra_ferramentas = ttk.Frame(self.root, padding=(10, 6))
        self.barra_ferramentas.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(
            self.barra_ferramentas,
            text="📊 Sistema de Automação",
            font=("Segoe UI", 11, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(
            self.barra_ferramentas,
            text=f"Versão {APP_VERSION}",
            foreground="#555555",
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Separator(
            self.barra_ferramentas, orient=tk.VERTICAL
        ).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Label(
            self.barra_ferramentas,
            text=f"Idealizado por {APP_AUTHOR}",
            foreground="#444444",
        ).pack(side=tk.LEFT, padx=(8, 12))

        ttk.Label(
            self.barra_ferramentas,
            text="Desenvolvido com auxílio de IA",
            foreground="#444444",
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(self.barra_ferramentas, text="Idioma:").pack(side=tk.RIGHT, padx=(8, 4))
        self.var_idioma = tk.StringVar(value="Português")
        self.cmb_idioma = ttk.Combobox(
            self.barra_ferramentas, textvariable=self.var_idioma,
            values=("Português", "English", "Español"),
            state="readonly", width=12
        )
        self.cmb_idioma.pack(side=tk.RIGHT, padx=4)
        self.cmb_idioma.bind("<<ComboboxSelected>>", self.alterar_idioma)

        ttk.Button(
            self.barra_ferramentas,
            text="Sobre / Créditos",
            command=self.mostrar_sobre,
        ).pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(
            side=tk.TOP, fill=tk.X
        )

        # Menu lateral: tarefas e histórico na parte superior; log abaixo.
        # O log fica na coluna esquerda para liberar altura para a área principal.
        frame_menu = ttk.Frame(self.root, padding=10)
        frame_menu.pack(side=tk.LEFT, fill=tk.Y, anchor="nw")

        ttk.Label(frame_menu, text="📋 Tarefas", font=("Segoe UI", 11, "bold")).pack(pady=(0, 10))
        for chave, info in SCRIPTS.items():
            btn = ttk.Button(frame_menu, text=info["titulo"], width=28,
                            command=lambda k=chave: self.selecionar_tarefa(k))
            btn.pack(pady=3)

        ttk.Separator(frame_menu, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        ttk.Button(frame_menu, text="📜 Histórico", command=self.mostrar_historico).pack(fill=tk.X, pady=5)

        # Log compacto na área marcada pelo usuário, com rolagem independente.
        frame_log = ttk.LabelFrame(
            frame_menu,
            text="📋 Log de Execução",
            padding=4
        )
        frame_log.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        self.saida_log = scrolledtext.ScrolledText(
            frame_log,
            width=30,
            height=12,
            font=("Consolas", 9),
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        self.saida_log.pack(fill=tk.BOTH, expand=True)

        self.area_conteudo = ttk.Frame(self.root, padding=15)
        self.area_conteudo.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(self.area_conteudo, text="ℹ️ Selecione uma tarefa no menu à esquerda",
                font=("Segoe UI", 12)).pack(pady=50)
        self.aplicar_idioma_interface()

    def traduzir_texto(self, texto):
        if not isinstance(texto, str) or self.idioma == "pt":
            return texto
        traducao = TRADUCOES_INTERFACE.get(texto)
        if not traducao:
            return texto
        return traducao[0] if self.idioma == "en" else traducao[1]

    def aplicar_idioma_interface(self):
        """Traduz os textos visíveis sem alterar valores internos nem a lógica."""
        def percorrer(widget):
            try:
                texto = widget.cget("text")
                if texto:
                    if not hasattr(widget, "_texto_original_pt"):
                        widget._texto_original_pt = texto
                    original = widget._texto_original_pt
                    novo = self.traduzir_texto(original)
                    if novo != texto:
                        widget.configure(text=novo)
            except (tk.TclError, AttributeError):
                pass
            for filho in widget.winfo_children():
                percorrer(filho)
        percorrer(self.root)
        # Cabeçalhos das tabelas são configurados separadamente dos widgets.
        if getattr(self, "tree_pdfs", None) and self.tree_pdfs.winfo_exists():
            for coluna, original in (("data", "Data"), ("nome", "Nome do Arquivo"), ("caminho", "Caminho Completo")):
                self.tree_pdfs.heading(coluna, text=self.traduzir_texto(original))
        if getattr(self, "tree_emails", None) and self.tree_emails.winfo_exists():
            for coluna, original in (("data", "Data"), ("remetente", "Remetente"), ("assunto", "Assunto")):
                self.tree_emails.heading(coluna, text=self.traduzir_texto(original))

    def alterar_idioma(self, evento=None):
        escolha = self.var_idioma.get()
        self.idioma = {"Português": "pt", "English": "en", "Español": "es"}.get(escolha, "pt")
        tarefa = self.script_atual
        # Reconstrói apenas a tela atual para atualizar também títulos e tabelas.
        if tarefa:
            self.selecionar_tarefa(tarefa)
        self.aplicar_idioma_interface()
        self.log({"pt": "🌐 Idioma alterado para Português.", "en": "🌐 Language changed to English.", "es": "🌐 Idioma cambiado a Español."}[self.idioma])

    def mostrar_sobre(self):
        """Exibe os créditos e a declaração transparente de uso de IA."""
        texto = (
            f"{APP_NAME}\n"
            f"Versão {APP_VERSION}\n\n"
            f"Idealização e direção do projeto: {APP_AUTHOR}.\n\n"
            "Desenvolvimento realizado com auxílio de Inteligência Artificial "
            "generativa (ChatGPT, da OpenAI).\n\n"
            "A IA foi utilizada como ferramenta de apoio. A ideia do produto, "
            "os requisitos e as decisões do projeto são de responsabilidade "
            f"de {APP_AUTHOR}.\n\n"
            "O programa reúne ferramentas para organizar PDFs, buscar e-mails, "
            "consolidar planilhas Excel e analisar gastos.\n\n"
            "Os arquivos PDF e Excel são processados localmente no computador. "
            "A busca de e-mails requer conexão com o serviço de e-mail escolhido."
        )
        messagebox.showinfo("Sobre o Sistema de Automação", texto, parent=self.root)

    def selecionar_tarefa(self, chave):
        self.script_atual = chave
        self.btn_executar = None
        self.btn_parar = None
        for widget in self.area_conteudo.winfo_children():
            widget.destroy()

        if chave == "organizar_pdfs":
            self.tela_pdfs()
        elif chave == "buscar_email":
            self.tela_emails()
        elif chave == "processar_excel":
            self.tela_processar_excel()
        elif chave == "analise_gastos":
            self.tela_analise_gastos()

    # ==========================================
    # 📂 TELA: ORGANIZAR PDFs
    # ==========================================
    def tela_pdfs(self):
        ttk.Label(self.area_conteudo, text="📂 Organizar PDFs", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))

        frame_pasta = ttk.Frame(self.area_conteudo)
        frame_pasta.pack(fill=tk.X, pady=5)
        ttk.Label(frame_pasta, text="Pasta:").pack(side=tk.LEFT)
        self.lbl_pasta = ttk.Label(frame_pasta, text="Nenhuma pasta selecionada", foreground="#777")
        self.lbl_pasta.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_pasta, text="📁 Escolher Pasta", command=self.escolher_pasta).pack(side=tk.LEFT)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=10)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Listar PDFs", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_parar = ttk.Button(frame_botoes, text="⏹️ Parar", command=self.parar_busca, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=2)

        # Pesquisa local pelo nome dos PDFs já listados.
        frame_pesquisa = ttk.Frame(self.area_conteudo)
        frame_pesquisa.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(frame_pesquisa, text="🔎 Localizar PDF:").pack(side=tk.LEFT)
        self.var_busca_pdf = tk.StringVar(master=self.root)
        self.ent_busca_pdf = ttk.Entry(
            frame_pesquisa, textvariable=self.var_busca_pdf, width=45
        )
        self.ent_busca_pdf.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.var_busca_pdf.trace_add("write", self.filtrar_pdfs_na_interface)
        ttk.Button(
            frame_pesquisa, text="Limpar",
            command=lambda: self.var_busca_pdf.set("")
        ).pack(side=tk.LEFT)

        colunas = ("data", "nome", "caminho")
        self.tree_pdfs = ttk.Treeview(self.area_conteudo, columns=colunas, show="headings", height=10)
        self.tree_pdfs.heading("data", text="Data")
        self.tree_pdfs.heading("nome", text="Nome do Arquivo")
        self.tree_pdfs.heading("caminho", text="Caminho Completo")
        self.tree_pdfs.column("data", width=100)
        self.tree_pdfs.column("nome", width=300)
        self.tree_pdfs.column("caminho", width=350)
        self.tree_pdfs.pack(fill=tk.BOTH, expand=True, pady=10)
        self.tree_pdfs.bind("<<TreeviewSelect>>", self.quando_selecionar_pdf)
        self.lbl_resultado_pdfs = ttk.Label(
            self.area_conteudo, text="Nenhum PDF listado", foreground="#666"
        )
        self.lbl_resultado_pdfs.pack(anchor="w", pady=(0, 5))

        frame_acoes = ttk.Frame(self.area_conteudo)
        frame_acoes.pack(fill=tk.X)
        self.btn_abrir_pdf = ttk.Button(frame_acoes, text="📄 Abrir PDF", command=self.abrir_pdf_selecionado, state=tk.DISABLED)
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=2)
        self.btn_ver_pdf = ttk.Button(frame_acoes, text="📂 Abrir Pasta", command=self.abrir_pasta_do_pdf, state=tk.DISABLED)
        self.btn_ver_pdf.pack(side=tk.LEFT, padx=2)

    # ==========================================
    # 📧 TELA: BUSCAR E-MAILS
    # ==========================================
    def tela_emails(self):
        ttk.Label(self.area_conteudo, text="📧 Buscar E-mails", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))

        frame_login = ttk.LabelFrame(self.area_conteudo, text="🔐 Dados de Acesso", padding=10)
        frame_login.pack(fill=tk.X, pady=5)
        ttk.Label(frame_login, text="E-mail:").grid(row=0, column=0, sticky="w")
        self.ent_email = ttk.Entry(frame_login, width=40)
        self.ent_email.grid(row=0, column=1, padx=5, pady=3)
        ttk.Label(frame_login, text="Senha de App:").grid(row=1, column=0, sticky="w")
        self.ent_senha = ttk.Entry(frame_login, width=40, show="*")
        self.ent_senha.grid(row=1, column=1, padx=5, pady=3)

        # Os botões guardam/recuperam os dados no cofre seguro do sistema operacional.
        frame_acesso_salvo = ttk.Frame(frame_login)
        frame_acesso_salvo.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.btn_salvar_acesso = ttk.Button(
            frame_acesso_salvo,
            text="💾 Salvar acesso",
            command=self.salvar_acesso_email
        )
        self.btn_salvar_acesso.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_usar_acesso = ttk.Button(
            frame_acesso_salvo,
            text="📋 Escolher acesso salvo",
            command=self.usar_acesso_salvo
        )
        self.btn_usar_acesso.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_ajuda_senha_app = ttk.Button(
            frame_acesso_salvo,
            text="❓ Como gerar senha de app?",
            command=self.mostrar_ajuda_senha_app
        )
        self.btn_ajuda_senha_app.pack(side=tk.LEFT)

        ttk.Label(
            frame_login,
            text="Use a senha de app do Google (16 caracteres), não a senha normal da conta.",
            foreground="#666"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 0))

        frame_filtro = ttk.LabelFrame(self.area_conteudo, text="🔍 Filtros de Busca", padding=10)
        frame_filtro.pack(fill=tk.X, pady=10)
        ttk.Label(frame_filtro, text="Nome/Remetente:").grid(row=0, column=0, sticky="w")
        self.ent_nome_remetente = ttk.Entry(frame_filtro, width=40)
        self.ent_nome_remetente.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(frame_filtro, text="Tipo de Documento:").grid(row=1, column=0, sticky="w")
        cmb_tipo = ttk.Combobox(frame_filtro, textvariable=self.var_tipo_doc,
                                values=[("todos", "Todos"), ("boleto", "Boleto"), ("fatura", "Fatura"),
                                        ("contrato", "Contrato"), ("notafiscal", "Nota Fiscal"), ("pedido", "Pedido")],
                                state="readonly", width=37)
        cmb_tipo.grid(row=1, column=1, padx=5, pady=3)
        cmb_tipo.current(0)

        ttk.Label(frame_filtro, text="Últimos:").grid(row=2, column=0, sticky="w")
        cmb_periodo = ttk.Combobox(frame_filtro, textvariable=self.var_periodo_texto,
                                values=list(self.opcoes_periodo.values()),
                                state="readonly", width=37)
        cmb_periodo.grid(row=2, column=1, padx=5, pady=3)
        cmb_periodo.current(0)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=10)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Buscar", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_parar = ttk.Button(frame_botoes, text="⏹️ Parar", command=self.parar_busca, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=2)

        # Barra de ações posicionada logo abaixo dos filtros e antes da lista.
        # Assim os botões permanecem visíveis mesmo em janelas baixas e não
        # são empurrados para baixo pela Treeview expansível.
        frame_acoes = ttk.Frame(self.area_conteudo)
        frame_acoes.pack(fill=tk.X, pady=(5, 0))

        self.btn_ver = ttk.Button(
            frame_acoes,
            text="👁️ Visualizar e-mail",
            command=self.visualizar_email,
            state=tk.DISABLED
        )
        self.btn_ver.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_salvar = ttk.Button(
            frame_acoes,
            text="💾 Salvar e-mail",
            command=self.salvar_email,
            state=tk.DISABLED
        )
        self.btn_salvar.pack(side=tk.LEFT, padx=2)

        # A lista ocupa somente o espaço restante entre os filtros e a barra
        # fixa. As barras de rolagem permitem ver todos os e-mails e colunas.
        frame_lista_emails = ttk.Frame(self.area_conteudo)
        frame_lista_emails.pack(fill=tk.BOTH, expand=True, pady=(5, 0), ipady=45)
        frame_lista_emails.rowconfigure(0, weight=1)
        frame_lista_emails.columnconfigure(0, weight=1)

        colunas = ("data", "remetente", "assunto")
        self.tree_emails = ttk.Treeview(
            frame_lista_emails,
            columns=colunas,
            show="headings",
            height=10
        )
        self.tree_emails.heading("data", text="Data")
        self.tree_emails.heading("remetente", text="Remetente")
        self.tree_emails.heading("assunto", text="Assunto")
        self.tree_emails.column("data", width=150, minwidth=120, stretch=False)
        self.tree_emails.column("remetente", width=250, minwidth=180, stretch=False)
        self.tree_emails.column("assunto", width=400, minwidth=220, stretch=True)

        barra_vertical_emails = ttk.Scrollbar(
            frame_lista_emails,
            orient=tk.VERTICAL,
            command=self.tree_emails.yview
        )
        barra_horizontal_emails = ttk.Scrollbar(
            frame_lista_emails,
            orient=tk.HORIZONTAL,
            command=self.tree_emails.xview
        )
        self.tree_emails.configure(
            yscrollcommand=barra_vertical_emails.set,
            xscrollcommand=barra_horizontal_emails.set
        )

        self.tree_emails.grid(row=0, column=0, sticky="nsew")
        barra_vertical_emails.grid(row=0, column=1, sticky="ns")
        barra_horizontal_emails.grid(row=1, column=0, sticky="ew")
        self.tree_emails.bind("<<TreeviewSelect>>", self.quando_selecionar_email)

    # ==========================================
    # 📊 TELA: CONSOLIDAR PLANILHAS EXCEL
    # ==========================================
    def tela_processar_excel(self):
        ttk.Label(self.area_conteudo, text="📊 Consolidar Planilhas Excel", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))
        ttk.Label(self.area_conteudo, text="Selecione um ou mais arquivos para juntar em um único relatório.").pack(anchor="w", pady=(0,10))

        frame_arquivos = ttk.Frame(self.area_conteudo)
        frame_arquivos.pack(fill=tk.X, pady=5)
        ttk.Label(frame_arquivos, text="Arquivos:").pack(side=tk.LEFT)
        self.lbl_arquivos = ttk.Label(frame_arquivos, text="Nenhum arquivo selecionado", foreground="#777")
        self.lbl_arquivos.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_arquivos, text="📂 Selecionar", command=self.escolher_arquivos).pack(side=tk.LEFT)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=15)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Processar", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_abrir_excel = ttk.Button(frame_botoes, text="📂 Abrir Resultado", command=self.abrir_arquivo_excel, state=tk.DISABLED)
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=2)

    # ==========================================
    # 📊 TELA: ANÁLISE DE GASTOS MENSAIS
    # ==========================================
    def tela_analise_gastos(self):
        """Mostra seleção do XLSX, tabela de prévia e gráficos na própria janela."""
        cabecalho = ttk.Frame(self.area_conteudo)
        cabecalho.pack(fill=tk.X)

        ttk.Label(
            cabecalho,
            text="📊 Análise de Gastos Mensais",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", pady=(0, 6))

        ttk.Label(
            cabecalho,
            text=(
                "Selecione sua planilha para visualizar os valores em euros e os gráficos. "
                "Depois, clique em Calcular Análise para salvar o relatório XLSX."
            ),
            justify=tk.LEFT,
            wraplength=850
        ).pack(anchor="w", pady=(0, 8))

        frame_arquivos = ttk.Frame(cabecalho)
        frame_arquivos.pack(fill=tk.X, pady=4)
        ttk.Label(frame_arquivos, text="Planilha:").pack(side=tk.LEFT)
        self.lbl_planilha_gastos = ttk.Label(
            frame_arquivos,
            text="Nenhum arquivo selecionado",
            foreground="#777"
        )
        self.lbl_planilha_gastos.pack(side=tk.LEFT, padx=6)
        ttk.Button(
            frame_arquivos,
            text="📂 Selecionar XLSX",
            command=self.escolher_planilha_gastos
        ).pack(side=tk.LEFT, padx=4)

        frame_botoes = ttk.Frame(cabecalho)
        frame_botoes.pack(fill=tk.X, pady=(5, 8))
        self.btn_executar = ttk.Button(
            frame_botoes,
            text="▶️ Calcular Análise e Salvar XLSX",
            command=self.executar_tarefa
        )
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_abrir_gastos = ttk.Button(
            frame_botoes,
            text="📂 Abrir Relatório",
            command=self.abrir_relatorio_gastos,
            state=tk.DISABLED
        )
        self.btn_abrir_gastos.pack(side=tk.LEFT, padx=2)

        # Área com rolagem: a tabela e os gráficos ficam abaixo dos controles.
        self.gastos_preview_area = ttk.Frame(self.area_conteudo)
        self.gastos_preview_area.pack(fill=tk.BOTH, expand=True)

        self.gastos_preview_scroll = tk.Canvas(
            self.gastos_preview_area,
            highlightthickness=0,
            borderwidth=0
        )
        barra_vertical = ttk.Scrollbar(
            self.gastos_preview_area,
            orient=tk.VERTICAL,
            command=self.gastos_preview_scroll.yview
        )
        self.gastos_preview_scroll.configure(yscrollcommand=barra_vertical.set)
        barra_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        self.gastos_preview_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.gastos_preview_body = ttk.Frame(self.gastos_preview_scroll, padding=(2, 2, 10, 12))
        self.gastos_preview_window = self.gastos_preview_scroll.create_window(
            (0, 0),
            window=self.gastos_preview_body,
            anchor="nw"
        )
        self.gastos_preview_body.bind(
            "<Configure>",
            lambda _event: self.gastos_preview_scroll.configure(
                scrollregion=self.gastos_preview_scroll.bbox("all")
            )
        )
        self.gastos_preview_scroll.bind(
            "<Configure>",
            lambda event: self.gastos_preview_scroll.itemconfigure(
                self.gastos_preview_window,
                width=event.width
            )
        )
        self.gastos_preview_scroll.bind("<Button-4>", self._rolar_preview_gastos)
        self.gastos_preview_scroll.bind("<Button-5>", self._rolar_preview_gastos)

        self.lbl_preview_gastos = ttk.Label(
            self.gastos_preview_body,
            text="Selecione um arquivo Excel para mostrar a planilha e os gráficos aqui.",
            foreground="#666",
            font=("Segoe UI", 10)
        )
        self.lbl_preview_gastos.pack(anchor="w", pady=18)

    def _rolar_preview_gastos(self, evento):
        """Permite rolar a prévia com a roda do mouse no Linux."""
        if evento.num == 4:
            self.gastos_preview_scroll.yview_scroll(-3, "units")
        elif evento.num == 5:
            self.gastos_preview_scroll.yview_scroll(3, "units")

    def escolher_planilha_gastos(self):
        caminho = filedialog.askopenfilename(
            title="Selecione sua planilha de contas",
            filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
        )
        if not caminho:
            return

        self.caminho_planilha_gastos = caminho
        self.lbl_planilha_gastos.config(text=os.path.basename(caminho), foreground="#000")
        self.log(f"✅ Planilha selecionada: {os.path.basename(caminho)}")

        try:
            dados = self.ler_dados_planilha_gastos(caminho)
            self.dados_gastos_preview = dados
            self.mostrar_preview_gastos(dados)
            self.log(
                f"✅ Prévia carregada: {len(dados['categorias'])} categorias | "
                f"Total: € {dados['total_geral']:,.2f}"
            )
        except Exception as erro:
            self.dados_gastos_preview = None
            messagebox.showerror(
                "Erro ao abrir planilha",
                f"Não foi possível montar a prévia da planilha.\n\n{erro}"
            )
            self.log(f"❌ Erro na prévia do Excel: {erro}")

    def ler_dados_planilha_gastos(self, caminho):
        """Lê a matriz Categoria x Meses e calcula totais uma única vez."""
        import pandas as pd

        engine = "xlrd" if caminho.lower().endswith(".xls") else None
        bruto = pd.read_excel(caminho, header=None, engine=engine)
        tipo_layout, linha_cabecalho, colunas_mes_detectadas = detectar_layout_excel(bruto)

        if tipo_layout != "mensal" or linha_cabecalho is None:
            raise ValueError(
                "Não identifiquei a tabela de gastos por mês. "
                "Preciso encontrar uma linha com os nomes dos meses (JAN a DEZ)."
            )

        meses_detectados = sorted(
            colunas_mes_detectadas,
            key=lambda item: (item[0], item[1])
        )
        if not meses_detectados:
            raise ValueError("Nenhuma coluna de mês foi encontrada na planilha.")

        meses = [str(item[2]).strip().upper() for item in meses_detectados]
        indices_meses = [item[1] for item in meses_detectados]
        dados = bruto.iloc[linha_cabecalho + 1:].copy()

        categorias = []
        valores_por_mes = [0.0] * len(meses)
        total_geral = 0.0

        for _, linha in dados.iterrows():
            categoria_bruta = linha.iloc[0] if len(linha) else None
            if pd.isna(categoria_bruta):
                continue

            nome_categoria = str(categoria_bruta).strip()
            nome_normalizado = normalizar_celula_excel(nome_categoria)
            if not nome_categoria or "total" in nome_normalizado:
                continue

            valores = []
            total_categoria = 0.0
            for posicao, indice_coluna in enumerate(indices_meses):
                valor_bruto = linha.iloc[indice_coluna] if indice_coluna < len(linha) else 0
                valor = self.converter_para_numero(valor_bruto)
                valores.append(valor)
                total_categoria += valor
                valores_por_mes[posicao] += valor

            categorias.append({
                "nome": nome_categoria,
                "valores": valores,
                "total": total_categoria
            })
            total_geral += total_categoria

        if not categorias:
            raise ValueError("Não encontrei categorias com valores para calcular.")

        categorias.sort(key=lambda item: chave_categoria_numerica(item["nome"]))
        totais_percentuais = [
            (item["total"] / total_geral * 100) if total_geral else 0.0
            for item in categorias
        ]

        return {
            "meses": meses,
            "categorias": categorias,
            "valores_por_mes": valores_por_mes,
            "total_geral": total_geral,
            "percentuais": totais_percentuais,
            "linha_cabecalho": linha_cabecalho + 1,
        }

    @staticmethod
    def _formatar_euro_tela(valor):
        texto = f"{float(valor):,.2f}"
        texto = texto.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
        return f"€ {texto}"

    def mostrar_preview_gastos(self, dados):
        """Desenha a tabela da planilha e os dois gráficos abaixo dela."""
        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.ticker import FuncFormatter
        except ImportError as erro:
            raise RuntimeError(
                "Para mostrar os gráficos dentro do programa, instale matplotlib com: "
                "python -m pip install matplotlib"
            ) from erro

        for widget in self.gastos_preview_body.winfo_children():
            widget.destroy()
        self.gastos_canvas_widget = None
        self.gastos_figure = None
        self.gastos_tree = None

        meses = dados["meses"]
        categorias = dados["categorias"]
        total_geral = dados["total_geral"]
        valores_por_mes = dados["valores_por_mes"]

        ttk.Label(
            self.gastos_preview_body,
            text=(
                f"Prévia da planilha • {len(categorias)} categorias • "
                f"Total anual: {self._formatar_euro_tela(total_geral)}"
            ),
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(2, 8))

        quadro_tabela = ttk.LabelFrame(
            self.gastos_preview_body,
            text="📋 Planilha — valores por categoria e mês",
            padding=6
        )
        quadro_tabela.pack(fill=tk.X, expand=False, pady=(0, 10))

        nomes_colunas = ["categoria"] + [f"mes_{i}" for i in range(len(meses))] + ["total_anual", "percentual"]
        self.gastos_tree = ttk.Treeview(
            quadro_tabela,
            columns=nomes_colunas,
            show="headings",
            height=min(max(len(categorias) + 1, 5), 11)
        )
        self.gastos_tree.heading("categoria", text="Categoria")
        self.gastos_tree.column("categoria", width=155, minwidth=130, stretch=False, anchor="w")

        for indice, mes in enumerate(meses):
            coluna = f"mes_{indice}"
            self.gastos_tree.heading(coluna, text=mes)
            self.gastos_tree.column(coluna, width=105, minwidth=90, stretch=False, anchor="e")

        self.gastos_tree.heading("total_anual", text="Total Anual")
        self.gastos_tree.column("total_anual", width=125, minwidth=110, stretch=False, anchor="e")
        self.gastos_tree.heading("percentual", text="% do Total")
        self.gastos_tree.column("percentual", width=100, minwidth=90, stretch=False, anchor="e")

        barra_horizontal = ttk.Scrollbar(
            quadro_tabela,
            orient=tk.HORIZONTAL,
            command=self.gastos_tree.xview
        )
        self.gastos_tree.configure(xscrollcommand=barra_horizontal.set)
        self.gastos_tree.pack(fill=tk.X, expand=True)
        barra_horizontal.pack(fill=tk.X)

        for categoria in categorias:
            percentual = (categoria["total"] / total_geral * 100) if total_geral else 0.0
            valores_linha = [categoria["nome"]]
            valores_linha.extend(self._formatar_euro_tela(valor) for valor in categoria["valores"])
            valores_linha.append(self._formatar_euro_tela(categoria["total"]))
            valores_linha.append(f"{percentual:.2f}%".replace(".", ","))
            self.gastos_tree.insert("", tk.END, values=valores_linha)

        valores_total = ["TOTAL GERAL"]
        valores_total.extend(self._formatar_euro_tela(valor) for valor in valores_por_mes)
        valores_total.append(self._formatar_euro_tela(total_geral))
        valores_total.append("100,00%" if total_geral else "0,00%")
        self.gastos_tree.insert("", tk.END, values=valores_total, tags=("total",))
        self.gastos_tree.tag_configure("total", font=("Segoe UI", 9, "bold"))

        quadro_graficos = ttk.LabelFrame(
            self.gastos_preview_body,
            text="📈 Gráficos mensal e anual",
            padding=6
        )
        quadro_graficos.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Os dois gráficos usam barras horizontais para manter o mesmo padrão
        # visual e evitar que os valores dos 12 meses fiquem sobrepostos.
        figura = Figure(figsize=(9.2, 8.2), dpi=90)
        ax_mensal = figura.add_subplot(2, 1, 1)
        ax_anual = figura.add_subplot(2, 1, 2)

        formatador_euro = FuncFormatter(
            lambda valor, _pos: "€ " + f"{valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        # Inverte a lista para que JAN apareça no topo e DEZ no final.
        meses_grafico = list(meses)[::-1]
        valores_mensais_grafico = list(valores_por_mes)[::-1]
        barras_mensais = ax_mensal.barh(meses_grafico, valores_mensais_grafico)
        ax_mensal.set_title("Gastos por mês", fontsize=11, fontweight="bold", pad=12)
        ax_mensal.set_xlabel("Total gasto (€)")
        ax_mensal.xaxis.set_major_formatter(formatador_euro)
        ax_mensal.grid(axis="x", alpha=0.25)
        maior_valor_mensal = max(valores_por_mes, default=0)
        ax_mensal.set_xlim(0, maior_valor_mensal * 1.25 if maior_valor_mensal > 0 else 1)
        rotulos_mensais = [self._formatar_euro_tela(valor) for valor in valores_mensais_grafico]
        ax_mensal.bar_label(
            barras_mensais,
            labels=rotulos_mensais,
            padding=4,
            fontsize=8,
        )
        ax_mensal.tick_params(axis="y", labelsize=8)

        categorias_ordenadas = sorted(categorias, key=lambda item: item["total"], reverse=True)
        nomes = [item["nome"] for item in categorias_ordenadas]
        totais = [item["total"] for item in categorias_ordenadas]
        barras_anuais = ax_anual.barh(nomes[::-1], totais[::-1])
        ax_anual.set_title("Gasto anual por categoria", fontsize=11, fontweight="bold")
        ax_anual.set_xlabel("Total anual (€)")
        ax_anual.xaxis.set_major_formatter(formatador_euro)
        ax_anual.grid(axis="x", alpha=0.25)
        ax_anual.bar_label(
            barras_anuais,
            labels=[self._formatar_euro_tela(v) for v in totais[::-1]],
            padding=3,
            fontsize=8
        )
        ax_anual.tick_params(axis="y", labelsize=8)

        figura.tight_layout(pad=2.0, h_pad=2.6)
        self.gastos_figure = figura
        self.gastos_canvas_widget = FigureCanvasTkAgg(figura, master=quadro_graficos)
        self.gastos_canvas_widget.draw()
        self.gastos_canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    
    def abrir_relatorio_gastos(self):
        """Abre a pasta onde o relatório de gastos foi salvo."""
        caminho = getattr(self, "caminho_relatorio_gastos", "")

        if not caminho:
            messagebox.showinfo(
                "Informação",
                "Nenhum relatório foi gerado ainda.\nExecute a análise primeiro!"
            )
            return

        caminho = os.path.abspath(os.path.expanduser(caminho))

        if not os.path.isfile(caminho):
            messagebox.showerror(
                "Relatório não encontrado",
                f"O relatório não foi encontrado:\n{caminho}"
            )
            return

        pasta = os.path.dirname(caminho)

        try:
            if os.name == "nt":
                os.startfile(pasta)

            elif sys.platform == "darwin":
                subprocess.Popen(["open", pasta])

            else:
                subprocess.Popen(
                    ["xdg-open", pasta],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )

            self.log(f"📂 Pasta do relatório de gastos aberta: {pasta}")

        except Exception as e:
            mensagem = (
                f"Não foi possível abrir a pasta:\n{pasta}\n\n"
                f"Detalhes: {e}"
            )
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir relatório", mensagem)

    def analise_gastos_calcular(self):
        """Gera um relatório financeiro sem somar linhas de total novamente."""
        if not self.caminho_planilha_gastos:
            self.log("❌ Selecione uma planilha primeiro!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"Analise_Gastos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not caminho_saida:
            self.log("⚠️ Operação cancelada.\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log("📊 Lendo planilha de gastos...")

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment

            dados = self.ler_dados_planilha_gastos(self.caminho_planilha_gastos)
            meses = dados["meses"]
            categorias = dados["categorias"]
            valores_por_mes = dados["valores_por_mes"]
            total_geral = dados["total_geral"]

            self.log(f"✅ Cabeçalho encontrado na linha {dados['linha_cabecalho']}")
            self.log(f"✅ {len(categorias)} categorias identificadas")

            wb = Workbook()
            ws_cat = wb.active
            ws_cat.title = "Resumo por Categoria"
            cabecalho = ["Categoria"] + meses + ["Total Anual", "% do Total"]
            ws_cat.append(cabecalho)

            for celula in ws_cat[1]:
                celula.font = Font(bold=True, color="FFFFFF")
                celula.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                celula.alignment = Alignment(horizontal="center")

            for categoria in categorias:
                percentual = (categoria["total"] / total_geral * 100) if total_geral else 0
                ws_cat.append(
                    [categoria["nome"]] + categoria["valores"] +
                    [round(categoria["total"], 2), round(percentual, 2)]
                )

            linha_total = ws_cat.max_row + 1
            ws_cat.append(["TOTAL GERAL"] + valores_por_mes + [round(total_geral, 2), 100.0 if total_geral else 0.0])
            for celula in ws_cat[linha_total]:
                celula.font = Font(bold=True)
                celula.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

            ws_mes = wb.create_sheet(title="Resumo Mensal")
            ws_mes.append(["Mês", "Total do Mês (€)", "% do Total Geral"])
            for celula in ws_mes[1]:
                celula.font = Font(bold=True, color="FFFFFF")
                celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                celula.alignment = Alignment(horizontal="center")

            for mes, valor in zip(meses, valores_por_mes):
                percentual = (valor / total_geral * 100) if total_geral else 0
                ws_mes.append([mes, round(valor, 2), round(percentual, 2)])

            linha_total_mes = ws_mes.max_row + 1
            ws_mes.append(["TOTAL GERAL", round(total_geral, 2), 100.0 if total_geral else 0.0])
            for celula in ws_mes[linha_total_mes]:
                celula.font = Font(bold=True)
                celula.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

            ws_cat.freeze_panes = "A2"
            for linha in ws_cat.iter_rows(min_row=2, max_col=len(meses) + 3):
                for celula in linha[1:len(meses) + 2]:
                    celula.number_format = FORMATO_EURO_EXCEL
                ws_cat.cell(row=linha[0].row, column=len(meses) + 3).number_format = FORMATO_PERCENTUAL_EXCEL

            ws_mes.freeze_panes = "A2"
            for linha in range(2, ws_mes.max_row + 1):
                ws_mes.cell(row=linha, column=2).number_format = FORMATO_EURO_EXCEL
                ws_mes.cell(row=linha, column=3).number_format = FORMATO_PERCENTUAL_EXCEL

            for ws in wb.worksheets:
                for coluna in ws.columns:
                    largura = max((len(str(c.value)) for c in coluna if c.value is not None), default=0)
                    ws.column_dimensions[coluna[0].column_letter].width = max(largura + 3, 12)

            # Larguras adequadas para o símbolo € e duas casas decimais.
            ws_cat.column_dimensions["A"].width = 22
            for indice_coluna in range(2, len(meses) + 2):
                letra = ws_cat.cell(row=1, column=indice_coluna).column_letter
                ws_cat.column_dimensions[letra].width = 15
            ws_cat.column_dimensions[
                ws_cat.cell(row=1, column=len(meses) + 2).column_letter
            ].width = 18
            ws_cat.column_dimensions[
                ws_cat.cell(row=1, column=len(meses) + 3).column_letter
            ].width = 16
            ws_mes.column_dimensions["A"].width = 16
            ws_mes.column_dimensions["B"].width = 20
            ws_mes.column_dimensions["C"].width = 20

            adicionar_graficos_resumo(
                wb,
                ws_mes,
                ws_cat,
                coluna_total_anual=len(meses) + 2,
            )

            wb.save(caminho_saida)
            self.caminho_relatorio_gastos = os.path.abspath(caminho_saida)
            self.root.after(0, lambda: self.seguro_config(self.btn_abrir_gastos, state=tk.NORMAL))

            self.log("\n" + "=" * 50)
            self.log("✅ ANÁLISE DE GASTOS CONCLUÍDA")
            self.log(f"💰 TOTAL GERAL: € {total_geral:,.2f}")
            self.log(f"📂 Relatório salvo: {caminho_saida}")
            self.log("📊 Maiores categorias por valor gasto:")
            maiores_categorias = sorted(categorias, key=lambda item: item["total"], reverse=True)
            for indice, categoria in enumerate(maiores_categorias[:5], start=1):
                self.log(f"   {indice}. {categoria['nome']}: € {categoria['total']:,.2f}")
            salvar_no_historico("Análise de Gastos", caminho_saida, len(categorias))

        except Exception as e:
            self.log(f"\n❌ Erro na análise de gastos: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    # ==========================================
    # 📁 FUNÇÕES GERAIS
    # ==========================================
    def escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta")
        if pasta:
            self.pasta_selecionada = pasta
            self.lbl_pasta.config(text=pasta, foreground="#000")

    def escolher_arquivos(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione os arquivos",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivos:
            self.caminhos_selecionados = list(arquivos)
            self.lbl_arquivos.config(text=f"{len(arquivos)} arquivo(s) selecionado(s)", foreground="#000")

    def quando_selecionar_pdf(self, evento):
        tem = bool(self.tree_pdfs.selection())
        self.seguro_config(self.btn_abrir_pdf, state=tk.NORMAL if tem else tk.DISABLED)
        self.seguro_config(self.btn_ver_pdf, state=tk.NORMAL if tem else tk.DISABLED)

    def quando_selecionar_email(self, evento):
        tem = bool(self.tree_emails.selection())
        self.seguro_config(self.btn_ver, state=tk.NORMAL if tem else tk.DISABLED)
        self.seguro_config(self.btn_salvar, state=tk.NORMAL if tem else tk.DISABLED)

    def abrir_pasta_do_pdf(self):
        sel = self.tree_pdfs.selection()
        if sel and self.lista_pdfs_exibidos:
            idx = self.tree_pdfs.index(sel[0])
            caminho = os.path.dirname(self.lista_pdfs_exibidos[idx]["caminho"])
            if os.name == "nt":
                os.startfile(caminho)
            else:
                subprocess.Popen(["xdg-open", caminho])

    def abrir_pdf_selecionado(self):
        sel = self.tree_pdfs.selection()
        if sel and self.lista_pdfs_exibidos:
            idx = self.tree_pdfs.index(sel[0])
            caminho = self.lista_pdfs_exibidos[idx]["caminho"]
            if os.name == "nt":
                os.startfile(caminho)
            else:
                subprocess.Popen(["xdg-open", caminho])

    def _caminho_cofre_email(self):
        """Retorna a pasta privada de dados do usuário em Linux ou Windows."""
        if os.name == "nt":
            base = os.environ.get("APPDATA") or os.path.expanduser("~")
        else:
            base = os.environ.get("XDG_DATA_HOME") or os.path.join(
                os.path.expanduser("~"), ".local", "share"
            )
        pasta = os.path.join(base, "SistemaAutomacao")
        os.makedirs(pasta, mode=0o700, exist_ok=True)
        if os.name != "nt":
            try:
                os.chmod(pasta, 0o700)
            except OSError:
                pass
        return os.path.join(pasta, ARQUIVO_COFRE_EMAIL)

    def _obter_cofre_credenciais(self, criar_se_necessario=False):
        """Abre o cofre criptografado sem chamar o Chaveiro do GNOME.

        A senha mestra é solicitada uma vez por execução do programa e a chave
        fica apenas em memória até o aplicativo ser fechado.
        """
        if self._cofre_email is not None and self._cofre_email.get("desbloqueado"):
            return self._cofre_email

        caminho = self._caminho_cofre_email()
        existe = os.path.isfile(caminho)
        if not existe and not criar_se_necessario:
            return None

        try:
            from cryptography.fernet import Fernet, InvalidToken
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
            import base64
            import secrets
            import tempfile

            def derivar_chave(senha_mestra, salt):
                kdf = PBKDF2HMAC(
                    algorithm=hashes.SHA256(),
                    length=32,
                    salt=salt,
                    iterations=PBKDF2_ITERACOES,
                )
                return base64.urlsafe_b64encode(kdf.derive(senha_mestra.encode("utf-8")))

            if existe:
                senha_mestra = simpledialog.askstring(
                    "Desbloquear cofre do aplicativo",
                    "Digite a senha mestra do Sistema de Automação:",
                    show="*",
                    parent=self.root,
                )
                if senha_mestra is None:
                    return None
                if not senha_mestra:
                    messagebox.showwarning(
                        "Senha mestra vazia",
                        "Digite a senha mestra para abrir os acessos salvos.",
                        parent=self.root,
                    )
                    return None

                with open(caminho, "r", encoding="utf-8") as arquivo:
                    estrutura = json.load(arquivo)
                salt = base64.b64decode(estrutura["salt"])
                chave = derivar_chave(senha_mestra, salt)
                fernet = Fernet(chave)
                try:
                    dados = json.loads(fernet.decrypt(estrutura["dados"].encode("ascii")).decode("utf-8"))
                except InvalidToken as exc:
                    raise ValueError("Senha mestra incorreta. Tente novamente.") from exc
                if not isinstance(dados, dict) or not isinstance(dados.get("contas", {}), dict):
                    raise ValueError("O arquivo do cofre está em formato inválido.")
            else:
                senha_mestra = simpledialog.askstring(
                    "Criar cofre de senhas",
                    "Crie uma senha mestra forte para proteger os e-mails salvos (mínimo 12 caracteres):",
                    show="*",
                    parent=self.root,
                )
                if senha_mestra is None:
                    return None
                if len(senha_mestra) < 12:
                    messagebox.showwarning(
                        "Senha muito curta",
                        "A senha mestra precisa ter pelo menos 12 caracteres. Use uma combinação longa e exclusiva.",
                        parent=self.root,
                    )
                    return None
                confirmacao = simpledialog.askstring(
                    "Confirmar senha mestra",
                    "Digite novamente a senha mestra:",
                    show="*",
                    parent=self.root,
                )
                if confirmacao is None:
                    return None
                if senha_mestra != confirmacao:
                    messagebox.showerror(
                        "Senhas diferentes",
                        "As senhas mestras não coincidem. Nenhum cofre foi criado.",
                        parent=self.root,
                    )
                    return None
                salt = secrets.token_bytes(16)
                fernet = Fernet(derivar_chave(senha_mestra, salt))
                dados = {"versao": 1, "contas": {}}

                # Cria o cofre vazio imediatamente, com permissões privadas.
                estrutura = {
                    "versao": 1,
                    "salt": base64.b64encode(salt).decode("ascii"),
                    "dados": fernet.encrypt(json.dumps(dados).encode("utf-8")).decode("ascii"),
                }
                fd, temporario = tempfile.mkstemp(prefix="cofre_", suffix=".tmp", dir=os.path.dirname(caminho))
                try:
                    with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                        json.dump(estrutura, arquivo, ensure_ascii=False)
                        arquivo.flush()
                        os.fsync(arquivo.fileno())
                    os.replace(temporario, caminho)
                    if os.name != "nt":
                        os.chmod(caminho, 0o600)
                finally:
                    if os.path.exists(temporario):
                        os.remove(temporario)

            self._cofre_email = {
                "fernet": fernet,
                "dados": dados,
                "caminho": caminho,
                "desbloqueado": True,
            }
            return self._cofre_email

        except ImportError as erro:
            messagebox.showerror(
                "Dependência ausente",
                "A biblioteca de criptografia não está instalada.\n\n"
                "Instale com: python -m pip install cryptography\n\n" + str(erro),
                parent=self.root,
            )
            return None
        except (ValueError, OSError, KeyError, json.JSONDecodeError) as erro:
            messagebox.showerror(
                "Não foi possível abrir o cofre",
                str(erro),
                parent=self.root,
            )
            return None

    def _salvar_cofre_credenciais(self, cofre):
        """Grava o conteúdo cifrado de forma atômica e com permissão privada."""
        import tempfile

        estrutura = {
            "versao": 1,
            "salt": self._ler_salt_cofre(cofre),
            "dados": cofre["fernet"].encrypt(
                json.dumps(cofre["dados"], ensure_ascii=False).encode("utf-8")
            ).decode("ascii"),
        }
        fd, temporario = tempfile.mkstemp(
            prefix="cofre_", suffix=".tmp", dir=os.path.dirname(cofre["caminho"])
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                json.dump(estrutura, arquivo, ensure_ascii=False)
                arquivo.flush()
                os.fsync(arquivo.fileno())
            os.replace(temporario, cofre["caminho"])
            if os.name != "nt":
                os.chmod(cofre["caminho"], 0o600)
        finally:
            if os.path.exists(temporario):
                os.remove(temporario)

    def _ler_salt_cofre(self, cofre):
        import base64
        with open(cofre["caminho"], "r", encoding="utf-8") as arquivo:
            estrutura = json.load(arquivo)
        return estrutura["salt"]

    def _ler_lista_acessos_salvos(self, cofre):
        """Lista os e-mails salvos no cofre local criptografado."""
        contas = cofre["dados"].setdefault("contas", {})
        return list(contas.keys())

    def mostrar_ajuda_senha_app(self):
        """Explica como criar uma senha de app do Google para a conexão IMAP."""
        textos = {
            "pt": {
                "titulo": "Como gerar a senha de app do Gmail",
                "corpo": (
                    "O Google normalmente bloqueia a senha comum da conta em programas que acessam o Gmail por IMAP. "
                    "Para conectar este sistema, você precisa criar uma SENHA DE APP, diferente da senha usada para entrar no Gmail.\n\n"
                    "PASSO A PASSO\n"
                    "1. Abra sua Conta Google: https://myaccount.google.com/\n"
                    "2. Entre em Segurança.\n"
                    "3. Ative a Verificação em duas etapas, caso ainda não esteja ativa.\n"
                    "4. Abra a página de Senhas de app: https://myaccount.google.com/apppasswords\n"
                    "5. Faça login novamente, se o Google solicitar.\n"
                    "6. Na criação da senha, dê um nome para identificar o aplicativo, por exemplo: Correio ou Sistema de Automação.\n"
                    "7. Clique em Criar. O Google vai gerar uma senha de 16 caracteres.\n"
                    "8. Copie o código gerado e cole no campo Senha de App deste programa.\n\n"
                    "IMPORTANTE\n"
                    "• O nome Correio é apenas uma identificação; o código é gerado pelo Google.\n"
                    "• A senha pode aparecer dividida em grupos por espaços. Você pode colar com ou sem espaços; o sistema os remove automaticamente.\n"
                    "• Não informe sua senha normal do Gmail neste programa.\n"
                    "• A senha de app é privada. Não a compartilhe com ninguém; digite-a diretamente no programa.\n"
                    "• Se a opção Senhas de app não aparecer, consulte as orientações de segurança da própria Conta Google para verificar como habilitá-la na sua conta."
                ), "conta": "Abrir Conta Google", "senha": "Abrir Senhas de app", "fechar": "Fechar"
            },
            "en": {
                "titulo": "How to create a Gmail app password",
                "corpo": (
                    "Google normally blocks your regular account password in programs that access Gmail through IMAP. "
                    "To connect this system, create an APP PASSWORD. It is different from the password you use to sign in to Gmail.\n\n"
                    "STEP BY STEP\n"
                    "1. Open your Google Account: https://myaccount.google.com/\n"
                    "2. Go to Security.\n"
                    "3. Turn on 2-Step Verification if it is not already enabled.\n"
                    "4. Open App passwords: https://myaccount.google.com/apppasswords\n"
                    "5. Sign in again if Google asks you to.\n"
                    "6. Give the password a name to identify the app, for example: Mail or Automation System.\n"
                    "7. Click Create. Google will generate a 16-character password.\n"
                    "8. Copy the generated code and paste it into this program's App Password field.\n\n"
                    "IMPORTANT\n"
                    "• The name Mail is only an identifier; Google generates the code.\n"
                    "• The password may be displayed in groups separated by spaces. Paste it with or without spaces; the system removes spaces automatically.\n"
                    "• Do not enter your regular Gmail password in this program.\n"
                    "• Your app password is private. Never share it; enter it directly into the program.\n"
                    "• If App passwords is not available, consult the security guidance in your Google Account to check how to enable it for your account."
                ), "conta": "Open Google Account", "senha": "Open App passwords", "fechar": "Close"
            },
            "es": {
                "titulo": "Cómo crear una contraseña de aplicación de Gmail",
                "corpo": (
                    "Google normalmente bloquea la contraseña habitual de la cuenta en programas que acceden a Gmail por IMAP. "
                    "Para conectar este sistema, debe crear una CONTRASEÑA DE APLICACIÓN, distinta de la contraseña que usa para entrar en Gmail.\n\n"
                    "PASO A PASO\n"
                    "1. Abra su Cuenta de Google: https://myaccount.google.com/\n"
                    "2. Entre en Seguridad.\n"
                    "3. Active la verificación en dos pasos si todavía no está activa.\n"
                    "4. Abra la página de Contraseñas de aplicación: https://myaccount.google.com/apppasswords\n"
                    "5. Inicie sesión de nuevo si Google se lo solicita.\n"
                    "6. Asigne un nombre para identificar la aplicación, por ejemplo: Correo o Sistema de Automatización.\n"
                    "7. Haga clic en Crear. Google generará una contraseña de 16 caracteres.\n"
                    "8. Copie el código generado y péguelo en el campo Contraseña de aplicación del programa.\n\n"
                    "IMPORTANTE\n"
                    "• El nombre Correo es solo una identificación; Google genera el código.\n"
                    "• La contraseña puede aparecer en grupos separados por espacios. Puede pegarla con o sin espacios; el sistema los elimina automáticamente.\n"
                    "• No introduzca aquí su contraseña habitual de Gmail.\n"
                    "• La contraseña de aplicación es privada. No la comparta con nadie; introdúzcala directamente en el programa.\n"
                    "• Si no aparece la opción, consulte las indicaciones de seguridad de su propia Cuenta de Google para comprobar cómo habilitarla."
                ), "conta": "Abrir Cuenta de Google", "senha": "Abrir Contraseñas de aplicación", "fechar": "Cerrar"
            }
        }
        idioma = textos[self.idioma]
        janela = tk.Toplevel(self.root)
        janela.title(idioma["titulo"])
        janela.transient(self.root)
        janela.grab_set()
        janela.geometry("700x590")
        janela.minsize(580, 480)
        principal = ttk.Frame(janela, padding=16)
        principal.pack(fill=tk.BOTH, expand=True)
        ttk.Label(principal, text=idioma["titulo"], font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        caixa = tk.Text(principal, wrap=tk.WORD, height=22, font=("Segoe UI", 10))
        caixa.pack(fill=tk.BOTH, expand=True)
        caixa.insert("1.0", idioma["corpo"])
        caixa.configure(state=tk.DISABLED)
        acoes = ttk.Frame(principal)
        acoes.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(acoes, text=idioma["conta"], command=lambda: webbrowser.open("https://myaccount.google.com/")).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(acoes, text=idioma["senha"], command=lambda: webbrowser.open("https://myaccount.google.com/apppasswords")).pack(side=tk.LEFT)
        ttk.Button(acoes, text=idioma["fechar"], command=janela.destroy).pack(side=tk.RIGHT)

    def salvar_acesso_email(self):
        """Salva vários acessos no cofre local sem usar o Chaveiro do GNOME."""
        email_usuario = self.ent_email.get().strip()
        senha = "".join(self.ent_senha.get().split())

        if not email_usuario or not senha:
            messagebox.showwarning(
                "Dados incompletos",
                "Preencha o e-mail e a senha de app antes de salvar."
            )
            return

        if "@" not in email_usuario or " " in email_usuario:
            messagebox.showwarning(
                "E-mail inválido",
                "Confira o endereço de e-mail informado."
            )
            return

        cofre = self._obter_cofre_credenciais(criar_se_necessario=True)
        if cofre is None:
            return

        contas = cofre["dados"].setdefault("contas", {})
        contas[email_usuario] = senha
        try:
            self._salvar_cofre_credenciais(cofre)
        except Exception as erro:
            self.log(f"❌ Não foi possível salvar o acesso: {erro}")
            messagebox.showerror(
                "Não foi possível salvar",
                f"O acesso não foi salvo corretamente.\n\n{erro}"
            )
            return

        self.log(f"🔐 Acesso de {email_usuario} salvo no cofre criptografado local.")
        messagebox.showinfo(
            "Acesso salvo",
            f"O acesso de {email_usuario} foi salvo com criptografia.\n\n"
            f"Você possui {len(contas)} acesso(s) salvo(s).\n\n"
            "A senha mestra será solicitada somente na primeira utilização do cofre em cada execução. "
            "Guarde essa senha: não existe recuperação automática."
        )

    def usar_acesso_salvo(self):
        """Abre a lista de e-mails salvos e preenche a conta escolhida."""
        cofre = self._obter_cofre_credenciais(criar_se_necessario=False)
        if cofre is None:
            messagebox.showinfo(
                "Nenhum acesso salvo",
                "Ainda não existe um cofre com acessos salvos.\n\n"
                "Preencha o e-mail e a senha de app e clique em 'Salvar acesso'."
            )
            return

        emails_salvos = self._ler_lista_acessos_salvos(cofre)
        if not emails_salvos:
            messagebox.showinfo(
                "Nenhum acesso salvo",
                "Ainda não existe acesso salvo. Preencha e-mail e senha de app e salve."
            )
            return

        janela = tk.Toplevel(self.root)
        janela.title("📧 Escolher acesso salvo")
        janela.geometry("430x300")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(
            janela,
            text="Selecione a conta que deseja utilizar:",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 7))

        frame_lista = ttk.Frame(janela, padding=(12, 0, 12, 8))
        frame_lista.pack(fill=tk.BOTH, expand=True)

        lista = tk.Listbox(frame_lista, height=8, exportselection=False)
        barra = ttk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=lista.yview)
        lista.configure(yscrollcommand=barra.set)
        lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        barra.pack(side=tk.RIGHT, fill=tk.Y)

        for endereco in emails_salvos:
            lista.insert(tk.END, endereco)

        lista.selection_set(0)
        lista.activate(0)

        frame_botoes = ttk.Frame(janela, padding=12)
        frame_botoes.pack(fill=tk.X)

        def carregar_selecionado(event=None):
            selecao = lista.curselection()
            if not selecao:
                messagebox.showwarning(
                    "Selecione uma conta",
                    "Escolha um e-mail da lista primeiro.",
                    parent=janela,
                )
                return

            endereco = lista.get(selecao[0])
            senha_salva = cofre["dados"].get("contas", {}).get(endereco)
            if not senha_salva:
                messagebox.showwarning(
                    "Senha não encontrada",
                    f"A senha de {endereco} não foi encontrada no cofre.",
                    parent=janela,
                )
                return

            self.ent_email.delete(0, tk.END)
            self.ent_email.insert(0, endereco)
            self.ent_senha.delete(0, tk.END)
            self.ent_senha.insert(0, senha_salva)
            self.log(f"✅ Acesso salvo carregado para {endereco}.")
            janela.destroy()

        ttk.Button(
            frame_botoes,
            text="Usar conta selecionada",
            command=carregar_selecionado
        ).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(
            frame_botoes,
            text="Cancelar",
            command=janela.destroy
        ).pack(side=tk.RIGHT)

        lista.bind("<Double-Button-1>", carregar_selecionado)
        lista.focus_set()

    def visualizar_email(self):
        sel = self.tree_emails.selection()
        if not sel: return
        idx = self.tree_emails.index(sel[0])
        em = self.emails_carregados[idx]
        janela = tk.Toplevel(self.root)
        janela.title(f"📧 {em['assunto'][:50]}")
        janela.geometry("600x450")
        ttk.Label(janela, text=f"De: {em['remetente']}\nData: {em['data']}", font=("Segoe UI", 10, "bold")).pack(pady=10, padx=10)
        txt = scrolledtext.ScrolledText(janela, wrap=tk.WORD)
        txt.pack(padx=10, pady=(0,10), fill=tk.BOTH, expand=True)
        txt.insert(tk.END, em["corpo"])
        txt.config(state=tk.DISABLED)

    def salvar_email(self):
        sel = self.tree_emails.selection()
        if not sel: return
        idx = self.tree_emails.index(sel[0])
        email_atual = self.emails_carregados[idx]

        caminho = filedialog.asksaveasfilename(
            defaultextension=".eml",
            filetypes=[("Arquivo de E-mail", "*.eml"), ("Texto", "*.txt"), ("Todos os arquivos", "*.*")],
            initialfile=email_atual["assunto"][:50].replace("/", "_").replace("\\", "_")
        )
        if caminho:
            try:
                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(f"De: {email_atual['remetente']}\n")
                    f.write(f"Assunto: {email_atual['assunto']}\n")
                    f.write(f"Data: {email_atual['data']}\n")
                    f.write("-" * 60 + "\n\n")
                    f.write(email_atual.get("corpo", ""))
                self.log(f"✅ E-mail salvo em: {caminho}\n")
            except Exception as e:
                self.log(f"❌ Erro ao salvar: {str(e)}\n")

    def converter_para_numero(self, valor):
        """Converte números do Excel e formatos monetários BR/EU para float."""
        import math

        if valor is None:
            return 0.0

        try:
            # Valores já numéricos vindos do pandas/openpyxl.
            if isinstance(valor, (int, float)):
                numero = float(valor)
                return numero if math.isfinite(numero) else 0.0

            texto = str(valor).strip()
            if not texto or texto.lower() in {"nan", "none", "nat", "-", "—", "n/a", "na"}:
                return 0.0

            negativo = texto.startswith("(") and texto.endswith(")")
            if negativo:
                texto = texto[1:-1].strip()

            # Remove moedas, espaços comuns e espaços não separáveis.
            texto = (texto.replace("€", "").replace("R$", "")
                           .replace("EUR", "").replace("BRL", "")
                           .replace("$", "").replace("\xa0", "")
                           .replace(" ", "").replace("'", ""))

            # Se vírgula e ponto coexistem, o último separador é o decimal.
            if "," in texto and "." in texto:
                if texto.rfind(",") > texto.rfind("."):
                    texto = texto.replace(".", "").replace(",", ".")
                else:
                    texto = texto.replace(",", "")
            elif "," in texto:
                texto = texto.replace(",", ".")

            numero = float(texto)
            if negativo:
                numero = -abs(numero)
            return numero if math.isfinite(numero) else 0.0

        except (ValueError, TypeError, OverflowError):
            return 0.0

    def executar_tarefa(self):
        if not self.script_atual:
            self.log("❌ Nenhuma tarefa selecionada!")
            return
        self.log(f"\n{'='*60}")
        self.log(f"▶️ EXECUTANDO: {SCRIPTS[self.script_atual]['titulo']}")
        self.log(f"{'='*60}\n")

        if self.script_atual == "organizar_pdfs":
            threading.Thread(target=self.organizar_pdfs_por_data, daemon=True).start()
        elif self.script_atual == "buscar_email":
            threading.Thread(target=self.executar_busca_emails, daemon=True).start()
        elif self.script_atual == "processar_excel":
            threading.Thread(target=self.processar_excel, daemon=True).start()
        elif self.script_atual == "analise_gastos":
            threading.Thread(target=self.analise_gastos_calcular, daemon=True).start()

        self.seguro_config(self.btn_executar, state=tk.DISABLED)
        self.seguro_config(self.btn_parar, state=tk.NORMAL)

    def parar_busca(self):
        self.buscando = False
        self.log(" ⏹ Solicitação de parada enviada...\n")
        self.seguro_config(self.btn_parar, state=tk.DISABLED)
        self.seguro_config(self.btn_executar, state=tk.NORMAL)

    # ==========================================
    # 📂 PDFs
    # ==========================================
    def buscar_pdfs_no_computador(self, pasta_alvo=None):
        if pasta_alvo is None:
            pasta_alvo = self.pasta_selecionada
        self.log(f" 🔍 Procurando PDFs em: {pasta_alvo}")
        encontrados = []
        if not os.path.exists(pasta_alvo):
            self.log(f" ⚠️ ERRO: A pasta não existe — {pasta_alvo}")
            return encontrados

        for raiz, pastas, arquivos in os.walk(pasta_alvo):
            for arq in arquivos:
                if arq.lower().endswith(".pdf"):
                    caminho_completo = os.path.join(raiz, arq)
                    try:
                        tempo = os.path.getmtime(caminho_completo)
                        data_criacao = datetime.fromtimestamp(tempo)
                    except:
                        data_criacao = datetime.now()
                    encontrados.append({
                        "nome": arq,
                        "caminho": caminho_completo,
                        "data_criacao": data_criacao
                    })
                    self.log(f" 📄 {arq} — {data_criacao.strftime('%d/%m/%Y')}")

        encontrados.sort(key=lambda x: x["data_criacao"], reverse=True)
        self.log(f" ✅ Total: {len(encontrados)} PDF(s)\n")
        return encontrados

    def listar_pdfs_na_interface(self, lista_pdfs):
        """Atualiza a lista completa e reaplica o filtro digitado, se houver."""
        self.lista_pdfs_encontrados = lista_pdfs
        self.filtrar_pdfs_na_interface()

    def filtrar_pdfs_na_interface(self, *_):
        """Filtra os PDFs pelo nome, sem alterar a lista original encontrada."""
        if not hasattr(self, "tree_pdfs") or not self.tree_pdfs.winfo_exists():
            return

        termo = self.var_busca_pdf.get().strip().casefold()
        if termo:
            filtrados = [
                item for item in self.lista_pdfs_encontrados
                if termo in item["nome"].casefold()
            ]
        else:
            filtrados = list(self.lista_pdfs_encontrados)

        self.lista_pdfs_exibidos = filtrados
        self.tree_pdfs.delete(*self.tree_pdfs.get_children())
        for item in filtrados:
            self.tree_pdfs.insert("", tk.END, values=(
                item["data_criacao"].strftime("%d/%m/%Y"),
                item["nome"],
                item["caminho"]
            ))

        if hasattr(self, "lbl_resultado_pdfs") and self.lbl_resultado_pdfs.winfo_exists():
            self.lbl_resultado_pdfs.config(
                text=f"Exibindo {len(filtrados)} de {len(self.lista_pdfs_encontrados)} PDF(s)"
            )

    def organizar_pdfs_por_data(self):
        if not self.pasta_selecionada:
            self.log("❌ Nenhuma pasta selecionada!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log("=" * 50)
        self.log("📂 LISTA DE PDFs ENCONTRADOS")
        self.log(f" 📁 Pasta: {self.pasta_selecionada}")
        self.log("=" * 50)

        self.lista_pdfs_encontrados = self.buscar_pdfs_no_computador()
        self.root.after(0, lambda: self.listar_pdfs_na_interface(self.lista_pdfs_encontrados))
        salvar_no_historico("Organizar PDFs", self.pasta_selecionada, len(self.lista_pdfs_encontrados))
        self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
        self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    def atualizar_lista_emails(self):
        """Atualiza a tabela de e-mails na thread principal do Tkinter."""
        if not hasattr(self, "tree_emails") or not self.tree_emails.winfo_exists():
            return

        self.tree_emails.delete(*self.tree_emails.get_children())
        for item in self.emails_carregados:
            self.tree_emails.insert(
                "",
                tk.END,
                values=(item["data"], item["remetente"], item["assunto"])
            )

        estado = tk.NORMAL if self.emails_carregados else tk.DISABLED
        self.seguro_config(self.btn_ver, state=estado)
        self.seguro_config(self.btn_salvar, state=estado)

    # ==========================================
    # 📧 E-MAILS
    # ==========================================
    def executar_busca_emails(self):
        email_usuario = self.ent_email.get().strip()
        # O Google pode exibir a senha de app em grupos separados por espaços.
        senha = "".join(self.ent_senha.get().split())
        remetente_filtro = self.ent_nome_remetente.get().strip() if hasattr(self, 'ent_nome_remetente') else ""
        tipo_doc = self.var_tipo_doc.get()

        texto_selecionado = self.var_periodo_texto.get()
        periodo = "0"
        for chave, valor in self.opcoes_periodo.items():
            if valor == texto_selecionado:
                periodo = chave
                break

        if not email_usuario or not senha:
            self.log("❌ Preencha e-mail e senha primeiro!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.emails_carregados = []
        self.root.after(0, lambda: self.tree_emails.delete(*self.tree_emails.get_children()))
        self.buscando = True

        imap = None
        try:
            self.log(" 🔌 Conectando ao Gmail com TLS...")
            # Timeout evita deixar a interface presa indefinidamente em uma rede instável.
            # IMAP4_SSL usa o contexto TLS padrão do Python (validação de certificado ativa).
            imap = imaplib.IMAP4_SSL("imap.gmail.com", 993, timeout=30)
            imap.login(email_usuario, senha)
            status_selecao, _ = imap.select("INBOX", readonly=True)
            if status_selecao != "OK":
                raise RuntimeError("Não foi possível acessar a caixa de entrada (INBOX).")
            self.log(" ✅ Conectado! Buscando e-mails...")

            filtro_data = ""
            if periodo != "0":
                dias = int(periodo)
                data_inicio = datetime.now() - timedelta(days=dias)
                filtro_data = f'SINCE "{data_inicio.strftime("%d-%b-%Y")}"'

            if remetente_filtro:
                # Evita que aspas ou quebras de linha alterem a sintaxe do comando IMAP.
                termo = remetente_filtro.strip().replace("\r", " ").replace("\n", " ")
                termo = termo.replace("\\", "\\\\").replace('"', '\\"')
                filtro_busca = f'OR FROM "{termo}" SUBJECT "{termo}"'
            else:
                filtro_busca = "ALL"

            if filtro_data and filtro_busca != "ALL":
                filtro_busca = f"({filtro_busca}) {filtro_data}"
            elif filtro_data:
                filtro_busca = filtro_data

            self.log(f" 🔍 Filtro: {filtro_busca}")
            status, dados = imap.search(None, filtro_busca)
            if status != "OK" or not dados:
                raise RuntimeError("O Gmail não conseguiu concluir a pesquisa de e-mails.")
            ids = dados[0].split()
            self.log(f" 📨 Encontrados {len(ids)} e-mail(s)")

            palavras_chave = []
            if tipo_doc != "todos":
                palavras = {
                    "boleto": ["boleto"],
                    "fatura": ["fatura"],
                    "contrato": ["contrato"],
                    "notafiscal": ["nota fiscal", "nfe", "nf-e"],
                    "pedido": ["pedido"]
                }
                palavras_chave = palavras.get(tipo_doc, [])

            limite = min(len(ids), 100)
            for id_email in reversed(ids[-limite:]):
                if not self.buscando:
                    self.log(" ⏹ Interrompido")
                    break

                status, dados = imap.fetch(id_email, "(RFC822)")
                if status != "OK" or not dados or not isinstance(dados[0], tuple):
                    continue
                mensagem = email.message_from_bytes(dados[0][1])

                assunto = self.decodificar_cabecalho(mensagem["Subject"]) or "(Sem assunto)"
                remetente = self.decodificar_cabecalho(mensagem["From"])
                data_raw = mensagem["Date"]
                try:
                    data_dt = parsedate_to_datetime(data_raw)
                    data_formatada = data_dt.strftime("%d/%m/%Y %H:%M")
                except:
                    data_formatada = data_raw or "Sem data"

                corpo = ""
                if mensagem.is_multipart():
                    for parte in mensagem.walk():
                        tipo = parte.get_content_type()
                        try:
                            if tipo == "text/plain":
                                texto = parte.get_payload(decode=True).decode("utf-8", errors="replace")
                                corpo += texto + "\n"
                            elif tipo == "text/html" and not corpo.strip():
                                html = parte.get_payload(decode=True).decode("utf-8", errors="replace")
                                corpo = extrair_texto_do_html(html)
                        except:
                            continue
                else:
                    try:
                        bruto = mensagem.get_payload(decode=True).decode("utf-8", errors="replace")
                        if "<html" in bruto.lower():
                            corpo = extrair_texto_do_html(bruto)
                        else:
                            corpo = bruto
                    except:
                        corpo = "[Não foi possível ler]"

                if palavras_chave:
                    texto_busca = (assunto + " " + corpo).lower()
                    if not any(p in texto_busca for p in palavras_chave):
                        continue

                email_dict = {
                    "data": data_formatada,
                    "remetente": remetente,
                    "assunto": assunto,
                    "corpo": corpo.strip() or "(Sem conteúdo)"
                }
                self.emails_carregados.append(email_dict)


            self.root.after(0, self.atualizar_lista_emails)
            self.log(f"✅ Pronto! {len(self.emails_carregados)} e-mail(s) carregado(s)\n")
            salvar_no_historico("Buscar E-mail", remetente_filtro or tipo_doc, len(self.emails_carregados))

        except imaplib.IMAP4.error as e:
            self.log(f"❌ Erro de login: {str(e)}")
            self.log("💡 Use SENHA DE APP — não sua senha normal\n")
        except Exception as e:
            self.log(f"❌ Erro: {str(e)}\n")
            import traceback
            self.log(traceback.format_exc() + "\n")
        finally:
            # Fecha a sessão mesmo quando login, busca ou leitura falhar.
            if imap is not None:
                try:
                    imap.logout()
                except Exception:
                    pass
            self.buscando = False
            self.root.after(0, self.atualizar_lista_emails)
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))

    # ==========================================
    # 📊 EXCEL — CONSOLIDAR PLANILHAS
    # ==========================================
    def processar_excel(self):
        """Consolida planilhas de transações ou matrizes anuais de contas.

        Transações: soma valores, resume por categoria e ordena pela data
        mais nova primeiro.
        Matriz anual: reconhece uma linha de meses mesmo quando existe um
        título mesclado acima, calcula totais por categoria e por mês.
        """
        if not self.caminhos_selecionados:
            self.log("❌ Nenhum arquivo Excel selecionado!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"Processado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not caminho_saida:
            self.log("⚠️ Operação cancelada.\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log(f"📂 Processando {len(self.caminhos_selecionados)} arquivo(s)...")
        self.log(f"💾 Salvando em: {caminho_saida}")

        try:
            import pandas as pd
            from datetime import date
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment

            registros = []
            totais_categoria_transacoes = {}
            total_transacoes = 0.0
            encontrou_data = False

            # Dados da matriz anual: {categoria: {numero_mes: valor}}
            totais_mensais_categoria = {}
            totais_por_mes = {}
            nomes_meses = {}
            total_matriz = 0.0
            arquivos_lidos = 0

            def ausente(valor):
                try:
                    return pd.isna(valor)
                except (TypeError, ValueError):
                    return False

            def localizar_coluna_transacao(cabecalhos, tipo):
                nomes = [normalizar_celula_excel(item) for item in cabecalhos]
                grupos = {
                    "data": ("data", "date", "dia"),
                    "descricao": ("descricao", "description", "desc", "nome", "produto", "item", "historico"),
                    "categoria": ("categoria", "category", "tipo", "classe"),
                    "valor": ("valor", "value", "preco", "importe", "montante", "amount", "total"),
                }
                sinonimos = grupos[tipo]
                for sinonimo in sinonimos:
                    for indice, nome in enumerate(nomes):
                        if nome == sinonimo:
                            return indice
                for sinonimo in sinonimos:
                    for indice, nome in enumerate(nomes):
                        if re.search(rf"\b{re.escape(sinonimo)}\b", nome):
                            return indice
                return None

            def converter_data_excel(valor):
                if valor is None or ausente(valor):
                    return None
                if isinstance(valor, datetime):
                    return valor.replace(tzinfo=None)
                if isinstance(valor, date):
                    return datetime.combine(valor, datetime.min.time())
                if isinstance(valor, (int, float)) and 1 <= float(valor) <= 80000:
                    return datetime(1899, 12, 30) + timedelta(days=float(valor))
                convertido = pd.to_datetime(valor, errors="coerce", dayfirst=True)
                if pd.isna(convertido):
                    return None
                return convertido.to_pydatetime().replace(tzinfo=None)

            for caminho in self.caminhos_selecionados:
                self.log(f"   📄 Lendo: {os.path.basename(caminho)}")
                try:
                    engine = "xlrd" if caminho.lower().endswith(".xls") else None
                    bruto = pd.read_excel(caminho, header=None, engine=engine)
                    if bruto.empty:
                        self.log(f"   ⚠️ Planilha vazia: {os.path.basename(caminho)}")
                        continue

                    tipo_layout, linha_cabecalho, colunas_mes = detectar_layout_excel(bruto)
                    if tipo_layout is None:
                        self.log(f"   ⚠️ Não reconheci o formato da planilha {os.path.basename(caminho)}.")
                        self.log("      Para transações, use cabeçalhos Data e Valor. Para contas anuais, use meses JAN a DEZ.")
                        continue

                    arquivos_lidos += 1
                    if tipo_layout == "mensal":
                        self.log(f"   ✅ Matriz de contas anuais detectada; cabeçalho na linha {linha_cabecalho + 1}.")
                        indices_meses = {indice for _, indice, _ in colunas_mes}
                        colunas_categoria = [i for i in range(bruto.shape[1]) if i not in indices_meses]
                        if not colunas_categoria:
                            self.log("   ⚠️ Não encontrei a coluna com os nomes das contas.")
                            continue
                        indice_categoria = colunas_categoria[0]

                        for _, linha in bruto.iloc[linha_cabecalho + 1:].iterrows():
                            nome_bruto = linha.iloc[indice_categoria]
                            if ausente(nome_bruto):
                                continue
                            categoria = str(nome_bruto).strip()
                            categoria_normalizada = normalizar_celula_excel(categoria)
                            if not categoria or categoria_normalizada.startswith("total"):
                                continue

                            valores_linha = {}
                            for numero_mes, indice_coluna, cabecalho_mes in colunas_mes:
                                valor = self.converter_para_numero(linha.iloc[indice_coluna])
                                valores_linha[numero_mes] = valores_linha.get(numero_mes, 0.0) + valor
                                nomes_meses[numero_mes] = str(cabecalho_mes).strip().upper()
                                totais_por_mes[numero_mes] = totais_por_mes.get(numero_mes, 0.0) + valor

                            soma_linha = sum(valores_linha.values())
                            if not any(abs(v) > 0 for v in valores_linha.values()):
                                continue
                            if categoria not in totais_mensais_categoria:
                                totais_mensais_categoria[categoria] = {}
                            for numero_mes, valor in valores_linha.items():
                                totais_mensais_categoria[categoria][numero_mes] = (
                                    totais_mensais_categoria[categoria].get(numero_mes, 0.0) + valor
                                )
                            total_matriz += soma_linha

                        self.log(f"   ✅ Contas mensais lidas de {os.path.basename(caminho)}")
                        continue

                    # Layout de transações: os nomes das colunas podem estar
                    # depois de uma ou mais linhas de título.
                    cabecalhos = bruto.iloc[linha_cabecalho].tolist()
                    dados = bruto.iloc[linha_cabecalho + 1:].copy()
                    mapa = {
                        "data": localizar_coluna_transacao(cabecalhos, "data"),
                        "descricao": localizar_coluna_transacao(cabecalhos, "descricao"),
                        "categoria": localizar_coluna_transacao(cabecalhos, "categoria"),
                        "valor": localizar_coluna_transacao(cabecalhos, "valor"),
                    }
                    if mapa["valor"] is None:
                        self.log(f"   ⚠️ Coluna de valor não encontrada em {os.path.basename(caminho)}.")
                        continue
                    if mapa["data"] is not None:
                        encontrou_data = True

                    quantidade = 0
                    for _, linha in dados.iterrows():
                        valor_bruto = linha.iloc[mapa["valor"]]
                        if ausente(valor_bruto) or str(valor_bruto).strip() == "":
                            continue
                        valor = self.converter_para_numero(valor_bruto)

                        data_bruta = linha.iloc[mapa["data"]] if mapa["data"] is not None else None
                        data_convertida = converter_data_excel(data_bruta)
                        descricao_bruta = linha.iloc[mapa["descricao"]] if mapa["descricao"] is not None else ""
                        categoria_bruta = linha.iloc[mapa["categoria"]] if mapa["categoria"] is not None else ""
                        descricao = "" if ausente(descricao_bruta) else str(descricao_bruta).strip()
                        categoria = "" if ausente(categoria_bruta) else str(categoria_bruta).strip()

                        registros.append({
                            "data_ordem": data_convertida,
                            "data_saida": data_convertida if data_convertida else ("" if ausente(data_bruta) else str(data_bruta).strip()),
                            "descricao": descricao,
                            "categoria": categoria,
                            "valor": valor,
                            "origem": os.path.basename(caminho),
                        })
                        total_transacoes += valor
                        quantidade += 1
                        if categoria:
                            totais_categoria_transacoes[categoria] = totais_categoria_transacoes.get(categoria, 0.0) + valor

                    self.log(f"   ✅ {quantidade} registro(s) de transação lido(s).")

                except Exception as erro_arquivo:
                    self.log(f"   ❌ Erro ao ler {os.path.basename(caminho)}: {erro_arquivo}")

            if not registros and not totais_mensais_categoria:
                raise ValueError(
                    "Nenhum dado foi reconhecido. Sua planilha pode ser uma matriz anual; "
                    "confira se existe uma linha com os meses JAN, FEV, MAR... DEZ e os nomes das contas na primeira coluna."
                )

            wb = Workbook()
            folha_inicial = wb.active
            folha_inicial.title = "Dados Consolidados"
            folha_inicial_foi_usada = False

            # Relatório de transações com ordenação de data decrescente.
            if registros:
                if encontrou_data:
                    registros.sort(
                        key=lambda item: (item["data_ordem"] is not None,
                                          item["data_ordem"] or datetime.min),
                        reverse=True
                    )
                    self.log("📅 Transações organizadas da mais nova para a mais antiga.")
                else:
                    self.log("⚠️ As planilhas de transações não possuem coluna Data; mantive a ordem original.")

                ws = folha_inicial
                folha_inicial_foi_usada = True
                cabecalho = ["Data", "Descrição", "Categoria", "Valor (€)", "Origem"]
                ws.append(cabecalho)
                for celula in ws[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    celula.alignment = Alignment(horizontal="center")

                for registro in registros:
                    ws.append([
                        registro["data_saida"], registro["descricao"], registro["categoria"],
                        round(registro["valor"], 2), registro["origem"]
                    ])
                    if isinstance(registro["data_saida"], datetime):
                        ws.cell(row=ws.max_row, column=1).number_format = "DD/MM/YYYY"
                    ws.cell(row=ws.max_row, column=4).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

                ultima_linha = ws.max_row
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = f"A1:E{ultima_linha}"
                linha_total = ws.max_row + 2
                ws.cell(row=linha_total, column=3, value="TOTAL GERAL").font = Font(bold=True)
                ws.cell(row=linha_total, column=4, value=round(total_transacoes, 2)).font = Font(bold=True)
                ws.cell(row=linha_total, column=4).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

                if totais_categoria_transacoes:
                    ws_resumo = wb.create_sheet("Resumo por Categoria")
                    ws_resumo.append(["Categoria", "Total (€)"])
                    for celula in ws_resumo[1]:
                        celula.font = Font(bold=True, color="FFFFFF")
                        celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                    for categoria, valor in sorted(totais_categoria_transacoes.items(), key=lambda item: item[1], reverse=True):
                        ws_resumo.append([categoria, round(valor, 2)])
                        ws_resumo.cell(row=ws_resumo.max_row, column=2).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

            # Relatório de matriz anual (CONTAS DO ANO, categorias x meses).
            if totais_mensais_categoria:
                ordem_meses = sorted(totais_por_mes.keys())  # janeiro a dezembro
                nomes = [nomes_meses.get(numero, str(numero)) for numero in ordem_meses]
                if not registros and not folha_inicial_foi_usada:
                    ws_anual = folha_inicial
                    ws_anual.title = "Resumo Anual por Categoria"
                    folha_inicial_foi_usada = True
                else:
                    ws_anual = wb.create_sheet("Resumo Anual por Categoria")
                ws_anual.append(["Categoria"] + nomes + ["Total Anual", "% do Total"])
                for celula in ws_anual[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    celula.alignment = Alignment(horizontal="center")

                totais_categoria_matriz = {
                    categoria: sum(valores.values())
                    for categoria, valores in totais_mensais_categoria.items()
                }
                total_matriz = sum(totais_categoria_matriz.values())
                # A coluna Categoria começa com um código numérico (ex.: "1 CARTÃO").
                # Ordena pelo código numérico, não pelo total gasto.
                categorias_em_ordem = sorted(
                    totais_categoria_matriz.items(),
                    key=lambda item: chave_categoria_numerica(item[0])
                )
                for categoria, total_cat in categorias_em_ordem:
                    valores = [totais_mensais_categoria[categoria].get(mes, 0.0) for mes in ordem_meses]
                    percentual = (total_cat / total_matriz * 100) if total_matriz else 0.0
                    ws_anual.append([categoria] + [round(v, 2) for v in valores] + [round(total_cat, 2), round(percentual, 2)])

                linha_total_anual = ws_anual.max_row + 1
                ws_anual.append(["TOTAL GERAL"] + [round(totais_por_mes.get(mes, 0.0), 2) for mes in ordem_meses] + [round(total_matriz, 2), 100.0 if total_matriz else 0.0])
                for celula in ws_anual[linha_total_anual]:
                    celula.font = Font(bold=True)
                    celula.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

                # Exibe todos os valores financeiros em euros, mantendo-os numéricos.
                formato_euro = FORMATO_EURO_EXCEL
                formato_percentual = FORMATO_PERCENTUAL_EXCEL
                ultima_coluna_valor = len(ordem_meses) + 2  # meses + Total Anual
                for linha in ws_anual.iter_rows(min_row=2, max_row=ws_anual.max_row):
                    for celula in linha[1:ultima_coluna_valor]:
                        celula.number_format = formato_euro
                    linha[-1].number_format = formato_percentual

                ws_meses = wb.create_sheet("Resumo Mensal")
                ws_meses.append(["Mês", "Total do Mês (€)", "% do Total Geral"])
                for celula in ws_meses[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                for mes in ordem_meses:
                    valor_mes = totais_por_mes.get(mes, 0.0)
                    percentual = (valor_mes / total_matriz * 100) if total_matriz else 0.0
                    ws_meses.append([nomes_meses.get(mes, str(mes)), round(valor_mes, 2), round(percentual, 2)])
                linha_total_mes = ws_meses.max_row + 1
                ws_meses.append(["TOTAL GERAL", round(total_matriz, 2), 100.0 if total_matriz else 0.0])
                for celula in ws_meses[linha_total_mes]:
                    celula.font = Font(bold=True)
                    celula.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                # Na aba mensal, a coluna B é dinheiro e a coluna C é percentual.
                for linha in range(2, ws_meses.max_row + 1):
                    ws_meses.cell(row=linha, column=2).number_format = formato_euro
                    ws_meses.cell(row=linha, column=3).number_format = formato_percentual

            for ws_final in wb.worksheets:
                for coluna in ws_final.columns:
                    largura = max((len(str(celula.value)) for celula in coluna if celula.value is not None), default=0)
                    ws_final.column_dimensions[coluna[0].column_letter].width = min(max(largura + 3, 12), 45)
                ws_final.freeze_panes = "A2"

            # Evita que o LibreOffice mostre ### nas células com moeda.
            if totais_mensais_categoria:
                ws_anual.column_dimensions["A"].width = 22
                for indice_coluna in range(2, len(ordem_meses) + 2):
                    letra = ws_anual.cell(row=1, column=indice_coluna).column_letter
                    ws_anual.column_dimensions[letra].width = 15
                coluna_total_anual = len(ordem_meses) + 2
                coluna_percentual = len(ordem_meses) + 3
                ws_anual.column_dimensions[
                    ws_anual.cell(row=1, column=coluna_total_anual).column_letter
                ].width = 18
                ws_anual.column_dimensions[
                    ws_anual.cell(row=1, column=coluna_percentual).column_letter
                ].width = 16

                ws_meses.column_dimensions["A"].width = 16
                ws_meses.column_dimensions["B"].width = 20
                ws_meses.column_dimensions["C"].width = 20

            if totais_mensais_categoria:
                adicionar_graficos_resumo(
                    wb,
                    ws_meses,
                    ws_anual,
                    coluna_total_anual=len(ordem_meses) + 2,
                )

            wb.save(caminho_saida)
            self.caminho_ultimo_excel = os.path.abspath(caminho_saida)
            self.root.after(0, lambda: self.seguro_config(self.btn_abrir_excel, state=tk.NORMAL))

            self.log("\n" + "=" * 50)
            self.log("✅ CONSOLIDAÇÃO CONCLUÍDA")
            if registros:
                self.log(f"📄 Transações: {len(registros)}")
                self.log(f"💰 TOTAL DAS TRANSAÇÕES: € {total_transacoes:,.2f}")
            if totais_mensais_categoria:
                self.log(f"📊 Categorias anuais: {len(totais_mensais_categoria)}")
                self.log(f"💰 TOTAL DO ANO: € {total_matriz:,.2f}")
                self.log("📅 Meses: janeiro a dezembro.")
                self.log("🔢 Categorias: ordenadas pelo número inicial (ordem numérica).")
            self.log(f"💾 Arquivo salvo: {caminho_saida}")
            self.log("=" * 50 + "\n")
            salvar_no_historico("Processar Excel", caminho_saida, len(registros) + len(totais_mensais_categoria))

        except Exception as e:
            self.log(f"\n❌ ERRO AO PROCESSAR EXCEL: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    def abrir_arquivo(self, caminho):
        """Abre o arquivo salvo usando o aplicativo adequado em cada sistema."""
        if not caminho:
            messagebox.showinfo("Abrir arquivo", "Nenhum arquivo foi informado.")
            return False

        caminho = os.path.abspath(os.path.expanduser(os.fspath(caminho)))
        if not os.path.isfile(caminho):
            mensagem = f"O arquivo não foi encontrado:\n{caminho}"
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Arquivo não encontrado", mensagem)
            return False

        try:
            if os.name == "nt":
                # Usa o aplicativo associado ao XLSX no Windows.
                os.startfile(caminho)
            elif sys.platform == "darwin":
                subprocess.Popen(
                    ["open", caminho],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True,
                )
            else:
                # Em Linux, abre XLSX diretamente no LibreOffice Calc quando
                # disponível. Isso evita falhas de associação MIME do xdg-open.
                comando = None
                extensao = os.path.splitext(caminho)[1].lower()
                if extensao in (".xlsx", ".xls", ".ods", ".csv"):
                    for nome_programa in ("libreoffice", "soffice"):
                        executavel = shutil.which(nome_programa)
                        if executavel:
                            comando = [executavel, "--calc", caminho]
                            break

                # Fallback para o aplicativo padrão do ambiente Linux.
                if comando is None:
                    xdg_open = shutil.which("xdg-open")
                    gio = shutil.which("gio")
                    if xdg_open:
                        comando = [xdg_open, caminho]
                    elif gio:
                        comando = [gio, "open", caminho]
                    else:
                        raise FileNotFoundError(
                            "Não encontrei LibreOffice, xdg-open ou gio para abrir o arquivo."
                        )

                subprocess.Popen(
                    comando,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True,
                    close_fds=True,
                )

            self.log(f"📂 Abrindo arquivo: {caminho}")
            return True
        except Exception as e:
            mensagem = (
                f"Não foi possível abrir o arquivo automaticamente.\n\n"
                f"O arquivo continua salvo em:\n{caminho}\n\n"
                f"Detalhes: {e}"
            )
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir arquivo", mensagem)
            return False

    
    def abrir_arquivo_excel(self):
        """Abre a pasta onde o último resultado do Excel foi salvo."""
        caminho = getattr(self, "caminho_ultimo_excel", "")

        if not caminho:
            messagebox.showinfo(
                "Aviso",
                "Nenhum resultado foi gerado ainda.\nExecute o processamento primeiro."
            )
            return

        caminho = os.path.abspath(os.path.expanduser(caminho))

        if not os.path.isfile(caminho):
            messagebox.showerror(
                "Arquivo não encontrado",
                f"O arquivo de resultado não foi encontrado:\n{caminho}"
            )
            return

        pasta = os.path.dirname(caminho)

        try:
            if os.name == "nt":
                os.startfile(pasta)

            elif sys.platform == "darwin":
                subprocess.Popen(["open", pasta])

            else:
                xdg_open = shutil.which("xdg-open")
                gio = shutil.which("gio")

                if xdg_open:
                    subprocess.Popen(
                        [xdg_open, pasta],
                        start_new_session=True
                    )
                elif gio:
                    subprocess.Popen(
                        [gio, "open", pasta],
                        start_new_session=True
                    )
                else:
                    raise FileNotFoundError(
                        "Não encontrei xdg-open nem gio."
                    )

            self.log(f"📂 Pasta do resultado aberta: {pasta}")

        except Exception as e:
            mensagem = f"Não foi possível abrir a pasta:\n{pasta}\n\nDetalhes: {e}"
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir pasta", mensagem)

    def mostrar_historico(self):
        janela = tk.Toplevel(self.root)
        janela.title("📜 Histórico de Ações")
        janela.geometry("600x400")

        texto = scrolledtext.ScrolledText(janela, wrap=tk.WORD, font=("Consolas", 10))
        texto.pack(fill=tk.BOTH, expand=True)

        arquivo_hist = caminho_arquivo_historico()
        if os.path.exists(arquivo_hist):
            try:
                with open(arquivo_hist, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                for item in reversed(dados):
                    texto.insert(tk.END, f"📅 {item.get('data', '')}\n")
                    texto.insert(tk.END, f"   Ação: {item.get('tipo', '')}\n")
                    texto.insert(tk.END, f"   Detalhe: {item.get('detalhe', '')}\n")
                    texto.insert(tk.END, f"   Registros: {item.get('quantidade', 0)}\n")
                    texto.insert(tk.END, "-" * 50 + "\n")
            except:
                texto.insert(tk.END, "Erro ao ler histórico.\n")
        else:
            texto.insert(tk.END, "Nenhum histórico ainda.\n")

        texto.config(state=tk.DISABLED)


# ==========================================
# ▶️ EXECUTAR O PROGRAMA
# ==========================================
if __name__ == "__main__":
    import sys
    import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext, simpledialog
import os
import re
import sys
import shutil
import json
import imaplib
import webbrowser
import email
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
import threading
import queue
import subprocess
from html.parser import HTMLParser


APP_VERSION = "1.1.5"
APP_NAME = "Sistema de Automação"
APP_AUTHOR = "Flávio Galera"
APP_AI_DISCLOSURE = "Desenvolvido com auxílio de Inteligência Artificial (ChatGPT)."


# ==========================================
# 🛠️ FUNÇÕES AUXILIARES
# ==========================================
class ExtrairTextoHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.texto = []
        self.dentro_script = False
        self.dentro_estilo = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            if tag == "script":
                self.dentro_script = True
            else:
                self.dentro_estilo = True
        elif tag in ("p", "br", "div", "h1", "h2", "h3", "tr"):
            self.texto.append("\n")

    def handle_endtag(self, tag):
        if tag == "script":
            self.dentro_script = False
        elif tag == "style":
            self.dentro_estilo = False

    def handle_data(self, dados):
        if not self.dentro_script and not self.dentro_estilo:
            self.texto.append(dados.strip() + " ")

    def pegar_texto(self):
        texto = "".join(self.texto)
        linhas = [linha.strip() for linha in texto.split("\n") if linha.strip()]
        return "\n".join(linhas)

def extrair_texto_do_html(html):
    if not html:
        return "(Sem conteúdo)"
    if "<" not in html and ">" not in html:
        return html.strip()
    parser = ExtrairTextoHTML()
    try:
        parser.feed(html)
        resultado = parser.pegar_texto()
        return resultado if resultado else "(Sem conteúdo de texto)"
    except Exception as e:
        return html.strip() if html else "(Sem conteúdo)"


def diretorio_dados_aplicativo():
    """Diretório gravável do usuário, também quando o app estiver empacotado."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.join(
            os.path.expanduser("~"), ".local", "share"
        )
    pasta = os.path.join(base, "SistemaAutomacao")
    os.makedirs(pasta, exist_ok=True)
    if os.name != "nt":
        try:
            os.chmod(pasta, 0o700)
        except OSError:
            pass
    return pasta


def caminho_arquivo_historico():
    """Mantém o histórico ao lado do .py em desenvolvimento e no perfil do usuário no app instalado."""
    if getattr(sys, "frozen", False):
        return os.path.join(diretorio_dados_aplicativo(), "historico.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico.json")


def salvar_no_historico(tipo, detalhe, registros):
    """Salva no máximo 100 ações, com gravação atômica e acesso privado no Linux."""
    import tempfile

    arquivo = caminho_arquivo_historico()
    historico = []
    if os.path.exists(arquivo):
        try:
            with open(arquivo, "r", encoding="utf-8") as f:
                dados_existentes = json.load(f)
            if isinstance(dados_existentes, list):
                historico = dados_existentes
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            # Um histórico inválido não deve impedir o uso do restante do sistema.
            historico = []

    entrada = {
        "data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "tipo": str(tipo)[:120],
        "detalhe": str(detalhe)[:1000],
        "quantidade": registros if isinstance(registros, int) else 0,
    }
    historico.insert(0, entrada)
    historico = historico[:100]

    pasta = os.path.dirname(arquivo)
    os.makedirs(pasta, exist_ok=True)
    fd, temporario = tempfile.mkstemp(prefix="historico_", suffix=".tmp", dir=pasta)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporario, arquivo)
        if os.name != "nt":
            os.chmod(arquivo, 0o600)
    finally:
        if os.path.exists(temporario):
            os.remove(temporario)


# ==========================================
# 📊 DETECÇÃO DO FORMATO DAS PLANILHAS EXCEL
# ==========================================
MESES_EXCEL = {
    1: ("JAN", "JANEIRO", "JANUARY"),
    2: ("FEV", "FEB", "FEVEREIRO", "FEBRUARY"),
    3: ("MAR", "MARCO", "MARCH"),
    4: ("ABR", "APR", "ABRIL", "APRIL"),
    5: ("MAI", "MAY", "MAIO"),
    6: ("JUN", "JUNE", "JUNHO"),
    7: ("JUL", "JULY", "JULHO"),
    8: ("AGO", "AUG", "AGOSTO", "AUGUST"),
    9: ("SET", "SEP", "SEPT", "SETEMBRO", "SEPTEMBER"),
    10: ("OUT", "OCT", "OUTUBRO", "OCTOBER"),
    11: ("NOV", "NOVEMBRO", "NOVEMBER"),
    12: ("DEZ", "DEC", "DEZEMBRO", "DECEMBER"),
}


def normalizar_celula_excel(valor):
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", texto.lower()).strip()


def detectar_mes_excel(valor):
    """Retorna o número do mês quando o cabeçalho contém JAN, JANEIRO etc."""
    texto = normalizar_celula_excel(valor)
    if not texto or texto in {"nan", "none"}:
        return None
    partes = texto.split()
    for numero, aliases in MESES_EXCEL.items():
        if any(parte.upper() in aliases for parte in partes):
            return numero
    return None


def detectar_layout_excel(df_bruto):
    """Detecta cabeçalho de transações ou cabeçalho de categorias x meses.

    Lê a planilha sem presumir que a primeira linha seja o cabeçalho. Isso
    permite planilhas com título mesclado, como 'CONTAS DO ANO'.
    Retorna (tipo, índice_da_linha_de_cabeçalho, colunas_de_mês).
    """
    limite = min(len(df_bruto), 50)
    for indice_linha in range(limite):
        valores = df_bruto.iloc[indice_linha].tolist()
        nomes = [normalizar_celula_excel(valor) for valor in valores]

        tem_data = any(
            nome == "data" or nome.startswith("data ") or nome == "date" or nome.startswith("date ")
            for nome in nomes
        )
        tem_valor = any(
            nome == "valor" or nome.startswith("valor ") or
            nome in {"value", "amount", "preco", "importe", "montante"} or
            nome.startswith(("value ", "amount ", "preco ", "importe ", "montante "))
            for nome in nomes
        )
        if tem_data and tem_valor:
            return "transacoes", indice_linha, []

        colunas_mes = []
        meses_vistos = set()
        for indice_coluna, valor in enumerate(valores):
            mes = detectar_mes_excel(valor)
            if mes is not None:
                colunas_mes.append((mes, indice_coluna, valor))
                meses_vistos.add(mes)

        # Duas ou mais colunas mensais distinguem a matriz anual de gastos.
        if len(colunas_mes) >= 2 and len(meses_vistos) >= 2:
            return "mensal", indice_linha, colunas_mes

    return None, None, []


# ==========================================
# 📋 SCRIPTS DISPONÍVEIS
# ==========================================
SCRIPTS = {
    "organizar_pdfs": {
        "titulo": "Organizar PDFs por Data",
        "descricao": "Lista todos os PDFs de uma pasta e organiza por data"
    },
    "buscar_email": {
        "titulo": "Buscar E-mails",
        "descricao": "Busca e filtra e-mails por remetente ou assunto"
    },
    "processar_excel": {
        "titulo": "Consolidar Planilhas Excel",
        "descricao": "Junta múltiplos arquivos em um só e calcula totais"
    },
    "analise_gastos": {
        "titulo": "📊 Análise de Gastos Mensais",
        "descricao": "Calcula totais por categoria, por mês e total geral"
    }
}


def chave_categoria_numerica(nome):
    """Ordena categorias pelo número inicial do nome (1, 2, 3... 10)."""
    texto = str(nome).strip()
    encontrado = re.match(r"^(\d+)", texto)
    if encontrado:
        return (0, int(encontrado.group(1)))
    return (1, 0)


FORMATO_EURO_EXCEL = '"€" #,##0.00;[Red]-"€" #,##0.00;"€" 0.00'
FORMATO_PERCENTUAL_EXCEL = '0.00"%"'


def adicionar_graficos_resumo(wb, ws_mensal, ws_anual, coluna_total_anual):
    """Adiciona gráficos mensais e anuais ao arquivo Excel.

    ws_mensal: aba com Mês na coluna A e Total do Mês na coluna B.
    ws_anual: aba com Categoria na coluna A e Total Anual na coluna informada.
    As linhas TOTAL GERAL são excluídas dos gráficos para evitar duplicidade.
    """
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.label import DataLabelList

    ultima_linha_mensal = ws_mensal.max_row - 1
    ultima_linha_anual = ws_anual.max_row - 1

    if ultima_linha_mensal < 2 or ultima_linha_anual < 2:
        return False

    # Cria uma aba separada para os gráficos, sem esconder os dados originais.
    if "Gráficos" in wb.sheetnames:
        del wb["Gráficos"]
    ws_graficos = wb.create_sheet("Gráficos")
    ws_graficos.sheet_view.showGridLines = False

    # Gráfico 1: total gasto em cada mês, em barras horizontais.
    # O formato horizontal segue o mesmo padrão visual do gráfico anual.
    grafico_mensal = BarChart()
    grafico_mensal.type = "bar"
    grafico_mensal.style = 11
    grafico_mensal.title = "Gastos por mês (€)"
    grafico_mensal.x_axis.title = "Total gasto (€)"
    grafico_mensal.y_axis.title = "Mês"
    # Os gráficos ficam empilhados verticalmente em uma aba exclusiva.
    # Isso impede que um gráfico cubra o outro ou qualquer tabela do relatório.
    grafico_mensal.height = 10
    grafico_mensal.width = 22
    grafico_mensal.legend = None
    dados_mensais = Reference(
        ws_mensal, min_col=2, min_row=1, max_row=ultima_linha_mensal
    )
    categorias_mensais = Reference(
        ws_mensal, min_col=1, min_row=2, max_row=ultima_linha_mensal
    )
    grafico_mensal.add_data(dados_mensais, titles_from_data=True)
    grafico_mensal.set_categories(categorias_mensais)
    grafico_mensal.x_axis.numFmt = '"€" #,##0.00'
    # Em barras horizontais, maxMin coloca JAN no topo e DEZ no final.
    grafico_mensal.y_axis.scaling.orientation = "maxMin"
    grafico_mensal.dLbls = DataLabelList()
    grafico_mensal.dLbls.showVal = True
    grafico_mensal.dLbls.showCatName = False
    grafico_mensal.dLbls.showSerName = False
    grafico_mensal.dLbls.showLegendKey = False
    grafico_mensal.dLbls.showPercent = False
    grafico_mensal.dLbls.showLeaderLines = False
    ws_graficos.add_chart(grafico_mensal, "A1")

    # Gráfico 2: total anual por categoria, sem a linha TOTAL GERAL.
    grafico_anual = BarChart()
    grafico_anual.type = "bar"
    grafico_anual.style = 11
    grafico_anual.title = "Gasto anual por categoria (€)"
    # Em gráficos horizontais, o openpyxl usa o eixo X para categorias
    # e o eixo Y para valores; os títulos são definidos nessa orientação.
    grafico_anual.x_axis.title = "Categoria"
    grafico_anual.y_axis.title = "Total anual (€)"
    grafico_anual.height = 10
    grafico_anual.width = 22
    grafico_anual.legend = None
    dados_anuais = Reference(
        ws_anual,
        min_col=coluna_total_anual,
        min_row=1,
        max_row=ultima_linha_anual,
    )
    categorias_anuais = Reference(
        ws_anual, min_col=1, min_row=2, max_row=ultima_linha_anual
    )
    grafico_anual.add_data(dados_anuais, titles_from_data=True)
    grafico_anual.set_categories(categorias_anuais)
    grafico_anual.x_axis.numFmt = '"€" #,##0.00'
    grafico_anual.dLbls = DataLabelList()
    grafico_anual.dLbls.showVal = True
    grafico_anual.dLbls.showCatName = False
    grafico_anual.dLbls.showSerName = False
    grafico_anual.dLbls.showLegendKey = False
    grafico_anual.dLbls.showPercent = False
    grafico_anual.dLbls.showLeaderLines = False
    # A altura de 10 cm ocupa aproximadamente 19-20 linhas. A segunda figura
    # começa bem abaixo, deixando uma margem confortável entre os gráficos.
    ws_graficos.add_chart(grafico_anual, "A28")

    ws_graficos.column_dimensions["A"].width = 22
    ws_graficos.sheet_properties.pageSetUpPr.fitToPage = True
    ws_graficos.page_setup.fitToWidth = 1
    ws_graficos.page_setup.fitToHeight = 0
    return True

# Cofre local criptografado: não depende do Chaveiro do GNOME.
ARQUIVO_COFRE_EMAIL = "credenciais_email.enc"
PBKDF2_ITERACOES = 600_000


# ==========================================
# 🏠 CLASSE PRINCIPAL DA APLICAÇÃO
# ==========================================


TRADUCOES_INTERFACE = {'📊 Sistema de Automação': ('📊 Automation System', '📊 Sistema de Automatización'), 'Sobre / Créditos': ('About / Credits', 'Acerca de / Créditos'), '📋 Log de Execução': ('📋 Execution Log', '📋 Registro de Ejecución'), '📋 Tarefas': ('📋 Tasks', '📋 Tareas'), '📜 Histórico': ('📜 History', '📜 Historial'), 'ℹ️ Selecione uma tarefa no menu à esquerda': ('ℹ️ Select a task from the menu on the left', 'ℹ️ Seleccione una tarea en el menú de la izquierda'), 'Organizar PDFs por Data': ('Organize PDFs by Date', 'Organizar PDF por fecha'), 'Lista todos os PDFs de uma pasta e organiza por data': ('Lists all PDFs in a folder and sorts them by date', 'Lista todos los PDF de una carpeta y los organiza por fecha'), 'Buscar E-mails': ('Search Emails', 'Buscar correos electrónicos'), 'Busca e filtra e-mails por remetente ou assunto': ('Searches and filters emails by sender or subject', 'Busca y filtra correos por remitente o asunto'), 'Consolidar Planilhas Excel': ('Consolidate Excel Spreadsheets', 'Consolidar hojas de cálculo de Excel'), 'Junta múltiplos arquivos em um só e calcula totais': ('Combines multiple files into one and calculates totals', 'Combina varios archivos en uno y calcula los totales'), '📊 Análise de Gastos Mensais': ('📊 Monthly Expense Analysis', '📊 Análisis de gastos mensuales'), 'Calcula totais por categoria, por mês e total geral': ('Calculates totals by category, month, and grand total', 'Calcula totales por categoría, mes y total general'), '📂 Organizar PDFs': ('📂 Organize PDFs', '📂 Organizar PDF'), 'Pasta:': ('Folder:', 'Carpeta:'), 'Nenhuma pasta selecionada': ('No folder selected', 'Ninguna carpeta seleccionada'), '📁 Escolher Pasta': ('📁 Choose Folder', '📁 Elegir carpeta'), '▶️ Listar PDFs': ('▶️ List PDFs', '▶️ Listar PDF'), '⏹️ Parar': ('⏹️ Stop', '⏹️ Detener'), '🔎 Localizar PDF:': ('🔎 Find PDF:', '🔎 Buscar PDF:'), 'Limpar': ('Clear', 'Limpiar'), 'Data': ('Date', 'Fecha'), 'Nome do Arquivo': ('File Name', 'Nombre del archivo'), 'Caminho Completo': ('Full Path', 'Ruta completa'), 'Nenhum PDF listado': ('No PDFs listed', 'No hay PDF en la lista'), '📄 Abrir PDF': ('📄 Open PDF', '📄 Abrir PDF'), '📂 Abrir Pasta': ('📂 Open Folder', '📂 Abrir carpeta'), '📧 Buscar E-mails': ('📧 Search Emails', '📧 Buscar correos'), '🔐 Dados de Acesso': ('🔐 Login Details', '🔐 Datos de acceso'), 'E-mail:': ('Email:', 'Correo electrónico:'), 'Senha de App:': ('App Password:', 'Contraseña de aplicación:'), '💾 Salvar acesso': ('💾 Save account', '💾 Guardar acceso'), '📋 Escolher acesso salvo': ('📋 Choose saved account', '📋 Elegir acceso guardado'), '❓ Como gerar senha de app?': ('❓ How to create an app password?', '❓ ¿Cómo crear una contraseña de aplicación?'), 'Use a senha de app do Google (16 caracteres), não a senha normal da conta.': ('Use your Google app password (16 characters), not your regular account password.', 'Use la contraseña de aplicación de Google (16 caracteres), no la contraseña normal de la cuenta.'), '🔍 Filtros de Busca': ('🔍 Search Filters', '🔍 Filtros de búsqueda'), 'Nome/Remetente:': ('Name/Sender:', 'Nombre/Remitente:'), 'Tipo de Documento:': ('Document Type:', 'Tipo de documento:'), 'Últimos:': ('Period:', 'Período:'), '▶️ Buscar': ('▶️ Search', '▶️ Buscar'), '👁️ Visualizar e-mail': ('👁️ View email', '👁️ Ver correo'), '💾 Salvar e-mail': ('💾 Save email', '💾 Guardar correo'), 'Remetente': ('Sender', 'Remitente'), 'Assunto': ('Subject', 'Asunto'), '📊 Consolidar Planilhas Excel': ('📊 Consolidate Excel Spreadsheets', '📊 Consolidar hojas de Excel'), 'Selecione um ou mais arquivos para juntar em um único relatório.': ('Select one or more files to combine into a single report.', 'Seleccione uno o más archivos para unirlos en un solo informe.'), 'Arquivos:': ('Files:', 'Archivos:'), 'Nenhum arquivo selecionado': ('No files selected', 'Ningún archivo seleccionado'), '📂 Selecionar': ('📂 Select', '📂 Seleccionar'), '▶️ Processar': ('▶️ Process', '▶️ Procesar'), '📂 Abrir Resultado': ('📂 Open Result', '📂 Abrir resultado'), 'Selecione sua planilha para visualizar os valores em euros e os gráficos. Depois, clique em Calcular Análise para salvar o relatório XLSX.': ('Select your spreadsheet to view values in euros and charts. Then click Calculate Analysis to save the XLSX report.', 'Seleccione su hoja de cálculo para ver los valores en euros y los gráficos. Después, haga clic en Calcular análisis para guardar el informe XLSX.'), 'Planilha:': ('Spreadsheet:', 'Hoja de cálculo:'), '📂 Selecionar XLSX': ('📂 Select XLSX', '📂 Seleccionar XLSX'), '▶️ Calcular Análise e Salvar XLSX': ('▶️ Calculate Analysis and Save XLSX', '▶️ Calcular análisis y guardar XLSX'), '📂 Abrir Relatório': ('📂 Open Report', '📂 Abrir informe'), '📋 Histórico': ('📋 History', '📋 Historial')}

# Complemento das traduções da interface.
# Não altera o conteúdo dos arquivos nem o log de execução.
TRADUCOES_INTERFACE.update({
    "Lista todos os PDFs de uma pasta e organiza por data": (
        "Lists all PDFs in a folder and sorts them by date",
        "Lista todos los PDF de una carpeta y los organiza por fecha"
    ),
    "Calcula totais por categoria, por mês e total geral": (
        "Calculates totals by category, month, and grand total",
        "Calcula los totales por categoría, mes y total general"
    ),
    "▶️ Listar PDFs": ("▶️ List PDFs", "▶️ Listar PDF"),
    "📂 Abrir Pasta": ("📂 Open Folder", "📂 Abrir carpeta"),
    "🔍 Filtros de Busca": ("🔍 Search Filters", "🔍 Filtros de búsqueda"),
    "E-mail:": ("Email:", "Correo electrónico:"),
    "Todos os e-mails": ("All emails", "Todos los correos"),
    "Últimos 7 dias": ("Last 7 days", "Últimos 7 días"),
    "Últimos 30 dias": ("Last 30 days", "Últimos 30 días"),
    "Últimos 90 dias": ("Last 90 days", "Últimos 90 días"),
    "Último ano": ("Last year", "Último año"),
    "Todos": ("All", "Todos"),
    "Boleto": ("Payment slip", "Comprobante de pago"),
    "Fatura": ("Invoice", "Factura"),
    "Contrato": ("Contract", "Contrato"),
    "Nota Fiscal": ("Tax invoice", "Factura fiscal"),
    "Pedido": ("Order", "Pedido"),
    "Selecione um ou mais arquivos para juntar em um único relatório.": (
        "Select one or more files to combine into a single report.",
        "Seleccione uno o más archivos para unirlos en un único informe."
    ),
    "Nenhum arquivo selecionado": (
        "No file selected",
        "Ningún archivo seleccionado"
    ),
    "Selecione sua planilha para visualizar os valores em euros e os gráficos. Depois, clique em Calcular Análise para salvar o relatório XLSX.": (
        "Select your spreadsheet to view the values in euros and the charts. Then click Calculate Analysis to save the XLSX report.",
        "Seleccione su hoja de cálculo para ver los valores en euros y los gráficos. Después, haga clic en Calcular análisis para guardar el informe XLSX."
    ),
    "Planilha:": ("Spreadsheet:", "Hoja de cálculo:"),
    "Nenhum arquivo selecionado": (
        "No file selected",
        "Ningún archivo seleccionado"
    ),
    "Selecione um arquivo Excel para mostrar a planilha e os gráficos aqui.": (
        "Select an Excel file to display the spreadsheet and charts here.",
        "Seleccione un archivo Excel para mostrar aquí la hoja de cálculo y los gráficos."
    ),
    "Categoria": ("Category", "Categoría"),
    "Total Anual": ("Annual Total", "Total anual"),
    "% do Total": ("% of Total", "% del total"),
    "TOTAL GERAL": ("GRAND TOTAL", "TOTAL GENERAL"),
    "📋 Planilha — valores por categoria e mês": (
        "📋 Spreadsheet — values by category and month",
        "📋 Hoja de cálculo — valores por categoría y mes"
    ),
    "📈 Gráficos mensal e anual": (
        "📈 Monthly and annual charts",
        "📈 Gráficos mensual y anual"
    ),
    "Gastos por mês": ("Expenses by month", "Gastos por mes"),
    "Total gasto (€)": ("Total expenses (€)", "Gastos totales (€)"),
    "Gasto anual por categoria": (
        "Annual expenses by category",
        "Gastos anuales por categoría"
    ),
    "Total anual (€)": ("Annual total (€)", "Total anual (€)"),
})

TRADUCOES_LOG = {
    "Sistema pronto! Selecione uma tarefa à esquerda.": (
        "System ready! Select a task on the left.",
        "¡Sistema listo! Seleccione una tarea a la izquierda."
    ),
    "Planilha selecionada:": ("Spreadsheet selected:", "Hoja de cálculo seleccionada:"),
    "Erro na prévia do Excel:": ("Excel preview error:", "Error en la vista previa de Excel:"),
    "Pasta do relatório de gastos aberta:": ("Expense report folder opened:", "Carpeta del informe de gastos abierta:"),
    "Selecione uma planilha primeiro!": ("Select a spreadsheet first!", "¡Seleccione una hoja de cálculo primero!"),
    "Operação cancelada.": ("Operation cancelled.", "Operación cancelada."),
    "Lendo planilha de gastos...": ("Reading expense spreadsheet...", "Leyendo hoja de gastos..."),
    "Cabeçalho encontrado na linha": ("Header found on row", "Encabezado encontrado en la fila"),
    "categorias identificadas": ("categories identified", "categorías identificadas"),
    "ANÁLISE DE GASTOS CONCLUÍDA": ("EXPENSE ANALYSIS COMPLETED", "ANÁLISIS DE GASTOS COMPLETADO"),
    "TOTAL GERAL:": ("GRAND TOTAL:", "TOTAL GENERAL:"),
    "Relatório salvo:": ("Report saved:", "Informe guardado:"),
    "Maiores categorias por valor gasto:": ("Largest categories by amount spent:", "Mayores categorías por importe gastado:"),
    "Erro na análise de gastos:": ("Expense analysis error:", "Error en el análisis de gastos:"),
    "Não foi possível salvar o acesso:": ("Could not save account:", "No se pudo guardar el acceso:"),
    "salvo no cofre criptografado local.": ("saved in the local encrypted vault.", "guardado en la bóveda cifrada local."),
    "Acesso salvo carregado para": ("Saved account loaded for", "Cuenta guardada cargada para"),
    "E-mail salvo em:": ("Email saved to:", "Correo guardado en:"),
    "Erro ao salvar:": ("Error saving:", "Error al guardar:"),
    "Nenhuma tarefa selecionada!": ("No task selected!", "¡No hay ninguna tarea seleccionada!"),
    "EXECUTANDO:": ("RUNNING:", "EJECUTANDO:"),
    "Solicitação de parada enviada...": ("Stop request sent...", "Solicitud de detención enviada..."),
    "Procurando PDFs em:": ("Searching for PDFs in:", "Buscando PDF en:"),
    "ERRO: A pasta não existe": ("ERROR: The folder does not exist", "ERROR: La carpeta no existe"),
    "Total:": ("Total:", "Total:"),
    "LISTA DE PDFs ENCONTRADOS": ("LIST OF FOUND PDFs", "LISTA DE PDF ENCONTRADOS"),
    "Nenhuma pasta selecionada!": ("No folder selected!", "¡No hay ninguna carpeta seleccionada!"),
    "Preencha e-mail e senha primeiro!": ("Enter your email and password first!", "¡Introduzca primero su correo y contraseña!"),
    "Conectando ao Gmail com TLS...": ("Connecting to Gmail using TLS...", "Conectando a Gmail mediante TLS..."),
    "Conectado! Buscando e-mails...": ("Connected! Searching emails...", "¡Conectado! Buscando correos..."),
    "Encontrados": ("Found", "Encontrados"),
    "Interrompido": ("Interrupted", "Interrumpido"),
    "Erro de login:": ("Login error:", "Error de inicio de sesión:"),
    "Use SENHA DE APP — não sua senha normal": ("Use an APP PASSWORD — not your regular password", "Use una CONTRASEÑA DE APLICACIÓN — no su contraseña habitual"),
    "Nenhum arquivo Excel selecionado!": ("No Excel file selected!", "¡No se ha seleccionado ningún archivo Excel!"),
    "Processando": ("Processing", "Procesando"),
    "Salvando em:": ("Saving to:", "Guardando en:"),
    "Lendo:": ("Reading:", "Leyendo:"),
    "Planilha vazia:": ("Empty spreadsheet:", "Hoja de cálculo vacía:"),
    "Não reconheci o formato da planilha": ("Spreadsheet format not recognized", "No se reconoció el formato de la hoja de cálculo"),
    "Para transações, use cabeçalhos Data e Valor. Para contas anuais, use meses JAN a DEZ.": (
        "For transactions, use Date and Amount headers. For annual accounts, use months JAN to DEC.",
        "Para transacciones, use los encabezados Fecha e Importe. Para cuentas anuales, use los meses ENE a DIC."
    ),
    "Matriz de contas anuais detectada; cabeçalho na linha": (
        "Annual accounts matrix detected; header on row",
        "Matriz de cuentas anuales detectada; encabezado en la fila"
    ),
    "Não encontrei a coluna com os nomes das contas.": (
        "Could not find the column containing account names.",
        "No se encontró la columna con los nombres de las cuentas."
    ),
    "Contas mensais lidas de": ("Monthly accounts read from", "Cuentas mensuales leídas de"),
    "Coluna de valor não encontrada em": ("Amount column not found in", "No se encontró la columna de importes en"),
    "registro(s) de transação lido(s).": ("transaction record(s) read.", "registro(s) de transacción leído(s)."),
    "Transações organizadas da mais nova para a mais antiga.": (
        "Transactions sorted from newest to oldest.",
        "Transacciones ordenadas de la más reciente a la más antigua."
    ),
    "As planilhas de transações não possuem coluna Data; mantive a ordem original.": (
        "The transaction spreadsheets have no Date column; the original order was kept.",
        "Las hojas de transacciones no tienen columna Fecha; se mantuvo el orden original."
    ),
    "CONSOLIDAÇÃO CONCLUÍDA": ("CONSOLIDATION COMPLETED", "CONSOLIDACIÓN COMPLETADA"),
    "TOTAL DAS TRANSAÇÕES:": ("TOTAL TRANSACTIONS:", "TOTAL DE LAS TRANSACCIONES:"),
    "Categorias anuais:": ("Annual categories:", "Categorías anuales:"),
    "TOTAL DO ANO:": ("YEAR TOTAL:", "TOTAL DEL AÑO:"),
    "Meses: janeiro a dezembro.": ("Months: January to December.", "Meses: enero a diciembre."),
    "Categorias: ordenadas pelo número inicial (ordem numérica).": (
        "Categories: sorted by their starting number (numerical order).",
        "Categorías: ordenadas por el número inicial (orden numérico)."
    ),
    "Arquivo salvo:": ("File saved:", "Archivo guardado:"),
    "ERRO AO PROCESSAR EXCEL:": ("ERROR PROCESSING EXCEL:", "ERROR AL PROCESAR EXCEL:"),
    "Abrindo pasta do resultado:": ("Opening result folder:", "Abriendo carpeta del resultado:"),
}
class Aplicacao:
    def __init__(self, janela_principal):
        self.root = janela_principal
        self.root.title(f"📊 {APP_NAME} — v{APP_VERSION}")
        self.root.geometry("1320x720")
        self.root.minsize(1100, 620)

        self.script_atual = None
        self.idioma = "pt"
        self._widgets_traduziveis = []
        self.pasta_selecionada = ""
        self.caminhos_selecionados = []
        self.caminho_ultimo_excel = ""
        self.caminho_planilha_gastos = ""
        self.caminho_relatorio_gastos = ""
        self.dados_gastos_preview = None
        self.gastos_tree = None
        self.gastos_canvas_widget = None
        self.gastos_figure = None
        self.emails_carregados = []
        self.lista_pdfs_encontrados = []
        self.lista_pdfs_exibidos = []
        self.buscando = False
        self._cofre_email = None  # Desbloqueado somente durante esta sessão do programa.

        self.var_tipo_doc = tk.StringVar(value="todos")
        self.var_periodo_texto = tk.StringVar(value="Todos os e-mails")
        self.opcoes_periodo = {
            "0": "Todos os e-mails",
            "7": "Últimos 7 dias",
            "30": "Últimos 30 dias",
            "90": "Últimos 90 dias",
            "365": "Último ano"
        }

        self.btn_executar = None
        self.btn_parar = None

        # Mensagens de log podem vir de tarefas em segundo plano.
        # A fila garante que somente a thread principal atualize o Tkinter.
        self._fila_log = queue.Queue()
        self._historico_log = []
        self._fechando = False

        self.criar_interface()
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        self.root.after(100, self._processar_fila_log)
        self.log("✅ Sistema pronto! Selecione uma tarefa à esquerda.")

    def seguro_config(self, widget, **kwargs):
        try:
            if widget and widget.winfo_exists():
                widget.config(**kwargs)
        except:
            pass

    def log(self, mensagem):
        """Envia mensagens para a fila; a interface é atualizada na thread principal."""
        try:
            self._fila_log.put(str(mensagem))
        except Exception:
            pass

    def _traduzir_log(self, mensagem):
        """Traduz somente textos conhecidos do sistema, preservando dados dinâmicos."""
        mensagem = str(mensagem)

        # Não traduzir rastreamentos técnicos de erros.
        if "Traceback (most recent call last)" in mensagem:
            return mensagem

        indice = {"pt": None, "en": 0, "es": 1}.get(self.idioma)

        if indice is None:
            return mensagem

        for original, traducoes in TRADUCOES_LOG.items():
            mensagem = mensagem.replace(original, traducoes[indice])

        return mensagem

    def _atualizar_log_exibicao(self):
        """Redesenha o histórico no idioma atualmente selecionado."""
        try:
            self.saida_log.config(state=tk.NORMAL)
            self.saida_log.delete("1.0", tk.END)

            for mensagem in self._historico_log:
                self.saida_log.insert(
                    tk.END,
                    self._traduzir_log(mensagem) + "\n"
                )

            self.saida_log.see(tk.END)
            self.saida_log.config(state=tk.DISABLED)

        except tk.TclError:
            pass

    def _processar_fila_log(self):
        """Atualiza o widget de log exclusivamente na thread principal."""
        if self._fechando:
            return

        try:
            while True:
                mensagem = self._fila_log.get_nowait()
                self._historico_log.append(mensagem)

                self.saida_log.config(state=tk.NORMAL)
                self.saida_log.insert(
                    tk.END,
                    self._traduzir_log(mensagem) + "\n"
                )
                self.saida_log.see(tk.END)
                self.saida_log.config(state=tk.DISABLED)

        except queue.Empty:
            pass
        except tk.TclError:
            return

        try:
            self.root.after(100, self._processar_fila_log)
        except tk.TclError:
            self._fechando = True
    def _ao_fechar(self):
        """Fecha a janela sem deixar o agendador do log tentar usar widgets destruídos."""
        self._fechando = True
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def decodificar_cabecalho(self, valor):
        if not valor: return ""
        try:
            partes = email.header.decode_header(valor)
            decodificado = []
            for texto, codificacao in partes:
                if isinstance(texto, bytes):
                    decodificado.append(texto.decode(codificacao or 'utf-8', errors='replace'))
                else:
                    decodificado.append(str(texto))
            return "".join(decodificado)
        except:
            return str(valor)

    def criar_interface(self):
        # Barra superior de identificação e créditos do projeto.
        # A informação sobre IA é apresentada de forma transparente:
        # o projeto é idealizado pelo autor e desenvolvido com apoio de IA.
        self.barra_ferramentas = ttk.Frame(self.root, padding=(10, 6))
        self.barra_ferramentas.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(
            self.barra_ferramentas,
            text="📊 Sistema de Automação",
            font=("Segoe UI", 11, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(
            self.barra_ferramentas,
            text=f"Versão {APP_VERSION}",
            foreground="#555555",
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Separator(
            self.barra_ferramentas, orient=tk.VERTICAL
        ).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Label(
            self.barra_ferramentas,
            text=f"Idealizado por {APP_AUTHOR}",
            foreground="#444444",
        ).pack(side=tk.LEFT, padx=(8, 12))

        ttk.Label(
            self.barra_ferramentas,
            text="Desenvolvido com auxílio de IA",
            foreground="#444444",
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(self.barra_ferramentas, text="Idioma:").pack(side=tk.RIGHT, padx=(8, 4))
        self.var_idioma = tk.StringVar(value="Português")
        self.cmb_idioma = ttk.Combobox(
            self.barra_ferramentas, textvariable=self.var_idioma,
            values=("Português", "English", "Español"),
            state="readonly", width=12
        )
        self.cmb_idioma.pack(side=tk.RIGHT, padx=4)
        self.cmb_idioma.bind("<<ComboboxSelected>>", self.alterar_idioma)

        ttk.Button(
            self.barra_ferramentas,
            text="Sobre / Créditos",
            command=self.mostrar_sobre,
        ).pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(
            side=tk.TOP, fill=tk.X
        )

        # Menu lateral: tarefas e histórico na parte superior; log abaixo.
        # O log fica na coluna esquerda para liberar altura para a área principal.
        frame_menu = ttk.Frame(self.root, padding=10)
        frame_menu.pack(side=tk.LEFT, fill=tk.Y, anchor="nw")

        ttk.Label(frame_menu, text="📋 Tarefas", font=("Segoe UI", 11, "bold")).pack(pady=(0, 10))
        for chave, info in SCRIPTS.items():
            btn = ttk.Button(frame_menu, text=info["titulo"], width=28,
                            command=lambda k=chave: self.selecionar_tarefa(k))
            btn.pack(pady=3)

        ttk.Separator(frame_menu, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        ttk.Button(frame_menu, text="📜 Histórico", command=self.mostrar_historico).pack(fill=tk.X, pady=5)

        # Log compacto na área marcada pelo usuário, com rolagem independente.
        frame_log = ttk.LabelFrame(
            frame_menu,
            text="📋 Log de Execução",
            padding=4
        )
        frame_log.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        self.saida_log = scrolledtext.ScrolledText(
            frame_log,
            width=30,
            height=12,
            font=("Consolas", 9),
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        self.saida_log.pack(fill=tk.BOTH, expand=True)

        self.area_conteudo = ttk.Frame(self.root, padding=15)
        self.area_conteudo.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(self.area_conteudo, text="ℹ️ Selecione uma tarefa no menu à esquerda",
                font=("Segoe UI", 12)).pack(pady=50)
        self.aplicar_idioma_interface()

    def traduzir_texto(self, texto):
        """Traduz reconhecendo o texto original ou uma tradução já aplicada."""
        if not isinstance(texto, str):
            return texto

        for original, traducoes in TRADUCOES_INTERFACE.items():
            ingles, espanhol = traducoes

            if texto == original or texto == ingles or texto == espanhol:
                if self.idioma == "en":
                    return ingles
                if self.idioma == "es":
                    return espanhol
                return original

        return texto

    def aplicar_idioma_interface(self):
        """Atualiza os textos visuais sem modificar dados dos usuários."""
        def percorrer(widget):
            try:
                texto_atual = widget.cget("text")

                if isinstance(texto_atual, str) and texto_atual:
                    novo_texto = self.traduzir_texto(texto_atual)

                    if novo_texto != texto_atual:
                        widget.configure(text=novo_texto)

            except (tk.TclError, AttributeError):
                pass

            try:
                filhos = widget.winfo_children()
            except (tk.TclError, AttributeError):
                filhos = []

            for filho in filhos:
                percorrer(filho)

        percorrer(self.root)

        # Cabeçalhos da tabela de PDFs.
        if getattr(self, "tree_pdfs", None):
            if self.tree_pdfs.winfo_exists():
                for coluna, original in (
                    ("data", "Data"),
                    ("nome", "Nome do Arquivo"),
                    ("caminho", "Caminho Completo")
                ):
                    self.tree_pdfs.heading(
                        coluna,
                        text=self.traduzir_texto(original)
                    )

        # Cabeçalhos da tabela de e-mails.
        if getattr(self, "tree_emails", None):
            if self.tree_emails.winfo_exists():
                for coluna, original in (
                    ("data", "Data"),
                    ("remetente", "Remetente"),
                    ("assunto", "Assunto")
                ):
                    self.tree_emails.heading(
                        coluna,
                        text=self.traduzir_texto(original)
                    )

    def alterar_idioma(self, evento=None):
        escolha = self.var_idioma.get()

        self.idioma = {
            "Português": "pt",
            "English": "en",
            "Español": "es"
        }.get(escolha, "pt")

        tarefa = self.script_atual

        if tarefa:
            self.selecionar_tarefa(tarefa)

        # Aplica também quando nenhuma tarefa está aberta.
        self.aplicar_idioma_interface()
        self._atualizar_log_exibicao()

        self.log({
            "pt": "🌐 Idioma alterado para Português.",
            "en": "🌐 Language changed to English.",
            "es": "🌐 Idioma cambiado a Español."
        }[self.idioma])

    def mostrar_sobre(self):
        """Exibe os créditos e a declaração transparente de uso de IA."""
        texto = (
            f"{APP_NAME}\n"
            f"Versão {APP_VERSION}\n\n"
            f"Idealização e direção do projeto: {APP_AUTHOR}.\n\n"
            "Desenvolvimento realizado com auxílio de Inteligência Artificial "
            "generativa (ChatGPT, da OpenAI).\n\n"
            "A IA foi utilizada como ferramenta de apoio. A ideia do produto, "
            "os requisitos e as decisões do projeto são de responsabilidade "
            f"de {APP_AUTHOR}.\n\n"
            "O programa reúne ferramentas para organizar PDFs, buscar e-mails, "
            "consolidar planilhas Excel e analisar gastos.\n\n"
            "Os arquivos PDF e Excel são processados localmente no computador. "
            "A busca de e-mails requer conexão com o serviço de e-mail escolhido."
        )
        messagebox.showinfo("Sobre o Sistema de Automação", texto, parent=self.root)

    def selecionar_tarefa(self, chave):
        self.script_atual = chave
        self.btn_executar = None
        self.btn_parar = None
        for widget in self.area_conteudo.winfo_children():
            widget.destroy()

        if chave == "organizar_pdfs":
            self.tela_pdfs()
        elif chave == "buscar_email":
            self.tela_emails()
        elif chave == "processar_excel":
            self.tela_processar_excel()
        elif chave == "analise_gastos":
            self.tela_analise_gastos()

    # ==========================================
    # 📂 TELA: ORGANIZAR PDFs
    # ==========================================
    def tela_pdfs(self):
        ttk.Label(self.area_conteudo, text="📂 Organizar PDFs", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))

        frame_pasta = ttk.Frame(self.area_conteudo)
        frame_pasta.pack(fill=tk.X, pady=5)
        ttk.Label(frame_pasta, text="Pasta:").pack(side=tk.LEFT)
        self.lbl_pasta = ttk.Label(frame_pasta, text="Nenhuma pasta selecionada", foreground="#777")
        self.lbl_pasta.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_pasta, text="📁 Escolher Pasta", command=self.escolher_pasta).pack(side=tk.LEFT)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=10)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Listar PDFs", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_parar = ttk.Button(frame_botoes, text="⏹️ Parar", command=self.parar_busca, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=2)

        # Pesquisa local pelo nome dos PDFs já listados.
        frame_pesquisa = ttk.Frame(self.area_conteudo)
        frame_pesquisa.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(frame_pesquisa, text="🔎 Localizar PDF:").pack(side=tk.LEFT)
        self.var_busca_pdf = tk.StringVar(master=self.root)
        self.ent_busca_pdf = ttk.Entry(
            frame_pesquisa, textvariable=self.var_busca_pdf, width=45
        )
        self.ent_busca_pdf.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.var_busca_pdf.trace_add("write", self.filtrar_pdfs_na_interface)
        ttk.Button(
            frame_pesquisa, text="Limpar",
            command=lambda: self.var_busca_pdf.set("")
        ).pack(side=tk.LEFT)

        colunas = ("data", "nome", "caminho")
        self.tree_pdfs = ttk.Treeview(self.area_conteudo, columns=colunas, show="headings", height=10)
        self.tree_pdfs.heading("data", text="Data")
        self.tree_pdfs.heading("nome", text="Nome do Arquivo")
        self.tree_pdfs.heading("caminho", text="Caminho Completo")
        self.tree_pdfs.column("data", width=100)
        self.tree_pdfs.column("nome", width=300)
        self.tree_pdfs.column("caminho", width=350)
        self.tree_pdfs.pack(fill=tk.BOTH, expand=True, pady=10)
        self.tree_pdfs.bind("<<TreeviewSelect>>", self.quando_selecionar_pdf)
        self.lbl_resultado_pdfs = ttk.Label(
            self.area_conteudo, text="Nenhum PDF listado", foreground="#666"
        )
        self.lbl_resultado_pdfs.pack(anchor="w", pady=(0, 5))

        frame_acoes = ttk.Frame(self.area_conteudo)
        frame_acoes.pack(fill=tk.X)
        self.btn_abrir_pdf = ttk.Button(frame_acoes, text="📄 Abrir PDF", command=self.abrir_pdf_selecionado, state=tk.DISABLED)
        self.btn_abrir_pdf.pack(side=tk.LEFT, padx=2)
        self.btn_ver_pdf = ttk.Button(frame_acoes, text="📂 Abrir Pasta", command=self.abrir_pasta_do_pdf, state=tk.DISABLED)
        self.btn_ver_pdf.pack(side=tk.LEFT, padx=2)

    # ==========================================
    # 📧 TELA: BUSCAR E-MAILS
    # ==========================================
    def tela_emails(self):
        ttk.Label(self.area_conteudo, text="📧 Buscar E-mails", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))

        frame_login = ttk.LabelFrame(self.area_conteudo, text="🔐 Dados de Acesso", padding=10)
        frame_login.pack(fill=tk.X, pady=5)
        ttk.Label(frame_login, text="E-mail:").grid(row=0, column=0, sticky="w")
        self.ent_email = ttk.Entry(frame_login, width=40)
        self.ent_email.grid(row=0, column=1, padx=5, pady=3)
        ttk.Label(frame_login, text="Senha de App:").grid(row=1, column=0, sticky="w")
        self.ent_senha = ttk.Entry(frame_login, width=40, show="*")
        self.ent_senha.grid(row=1, column=1, padx=5, pady=3)

        # Os botões guardam/recuperam os dados no cofre seguro do sistema operacional.
        frame_acesso_salvo = ttk.Frame(frame_login)
        frame_acesso_salvo.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.btn_salvar_acesso = ttk.Button(
            frame_acesso_salvo,
            text="💾 Salvar acesso",
            command=self.salvar_acesso_email
        )
        self.btn_salvar_acesso.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_usar_acesso = ttk.Button(
            frame_acesso_salvo,
            text="📋 Escolher acesso salvo",
            command=self.usar_acesso_salvo
        )
        self.btn_usar_acesso.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_ajuda_senha_app = ttk.Button(
            frame_acesso_salvo,
            text="❓ Como gerar senha de app?",
            command=self.mostrar_ajuda_senha_app
        )
        self.btn_ajuda_senha_app.pack(side=tk.LEFT)

        ttk.Label(
            frame_login,
            text="Use a senha de app do Google (16 caracteres), não a senha normal da conta.",
            foreground="#666"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 0))

        frame_filtro = ttk.LabelFrame(self.area_conteudo, text="🔍 Filtros de Busca", padding=10)
        frame_filtro.pack(fill=tk.X, pady=10)
        ttk.Label(frame_filtro, text="Nome/Remetente:").grid(row=0, column=0, sticky="w")
        self.ent_nome_remetente = ttk.Entry(frame_filtro, width=40)
        self.ent_nome_remetente.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(frame_filtro, text="Tipo de Documento:").grid(row=1, column=0, sticky="w")
        cmb_tipo = ttk.Combobox(frame_filtro, textvariable=self.var_tipo_doc,
                                values=[("todos", "Todos"), ("boleto", "Boleto"), ("fatura", "Fatura"),
                                        ("contrato", "Contrato"), ("notafiscal", "Nota Fiscal"), ("pedido", "Pedido")],
                                state="readonly", width=37)
        cmb_tipo.grid(row=1, column=1, padx=5, pady=3)
        cmb_tipo.current(0)

        ttk.Label(frame_filtro, text="Últimos:").grid(row=2, column=0, sticky="w")
        cmb_periodo = ttk.Combobox(frame_filtro, textvariable=self.var_periodo_texto,
                                values=list(self.opcoes_periodo.values()),
                                state="readonly", width=37)
        cmb_periodo.grid(row=2, column=1, padx=5, pady=3)
        cmb_periodo.current(0)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=10)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Buscar", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_parar = ttk.Button(frame_botoes, text="⏹️ Parar", command=self.parar_busca, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=2)

        # Barra de ações posicionada logo abaixo dos filtros e antes da lista.
        # Assim os botões permanecem visíveis mesmo em janelas baixas e não
        # são empurrados para baixo pela Treeview expansível.
        frame_acoes = ttk.Frame(self.area_conteudo)
        frame_acoes.pack(fill=tk.X, pady=(5, 0))

        self.btn_ver = ttk.Button(
            frame_acoes,
            text="👁️ Visualizar e-mail",
            command=self.visualizar_email,
            state=tk.DISABLED
        )
        self.btn_ver.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_salvar = ttk.Button(
            frame_acoes,
            text="💾 Salvar e-mail",
            command=self.salvar_email,
            state=tk.DISABLED
        )
        self.btn_salvar.pack(side=tk.LEFT, padx=2)

        # A lista ocupa somente o espaço restante entre os filtros e a barra
        # fixa. As barras de rolagem permitem ver todos os e-mails e colunas.
        frame_lista_emails = ttk.Frame(self.area_conteudo)
        frame_lista_emails.pack(fill=tk.BOTH, expand=True, pady=(5, 0), ipady=45)
        frame_lista_emails.rowconfigure(0, weight=1)
        frame_lista_emails.columnconfigure(0, weight=1)

        colunas = ("data", "remetente", "assunto")
        self.tree_emails = ttk.Treeview(
            frame_lista_emails,
            columns=colunas,
            show="headings",
            height=10
        )
        self.tree_emails.heading("data", text="Data")
        self.tree_emails.heading("remetente", text="Remetente")
        self.tree_emails.heading("assunto", text="Assunto")
        self.tree_emails.column("data", width=150, minwidth=120, stretch=False)
        self.tree_emails.column("remetente", width=250, minwidth=180, stretch=False)
        self.tree_emails.column("assunto", width=400, minwidth=220, stretch=True)

        barra_vertical_emails = ttk.Scrollbar(
            frame_lista_emails,
            orient=tk.VERTICAL,
            command=self.tree_emails.yview
        )
        barra_horizontal_emails = ttk.Scrollbar(
            frame_lista_emails,
            orient=tk.HORIZONTAL,
            command=self.tree_emails.xview
        )
        self.tree_emails.configure(
            yscrollcommand=barra_vertical_emails.set,
            xscrollcommand=barra_horizontal_emails.set
        )

        self.tree_emails.grid(row=0, column=0, sticky="nsew")
        barra_vertical_emails.grid(row=0, column=1, sticky="ns")
        barra_horizontal_emails.grid(row=1, column=0, sticky="ew")
        self.tree_emails.bind("<<TreeviewSelect>>", self.quando_selecionar_email)

    # ==========================================
    # 📊 TELA: CONSOLIDAR PLANILHAS EXCEL
    # ==========================================
    def tela_processar_excel(self):
        ttk.Label(self.area_conteudo, text="📊 Consolidar Planilhas Excel", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0,15))
        ttk.Label(self.area_conteudo, text="Selecione um ou mais arquivos para juntar em um único relatório.").pack(anchor="w", pady=(0,10))

        frame_arquivos = ttk.Frame(self.area_conteudo)
        frame_arquivos.pack(fill=tk.X, pady=5)
        ttk.Label(frame_arquivos, text="Arquivos:").pack(side=tk.LEFT)
        self.lbl_arquivos = ttk.Label(frame_arquivos, text="Nenhum arquivo selecionado", foreground="#777")
        self.lbl_arquivos.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_arquivos, text="📂 Selecionar", command=self.escolher_arquivos).pack(side=tk.LEFT)

        frame_botoes = ttk.Frame(self.area_conteudo)
        frame_botoes.pack(fill=tk.X, pady=15)
        self.btn_executar = ttk.Button(frame_botoes, text="▶️ Processar", command=self.executar_tarefa)
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_abrir_excel = ttk.Button(frame_botoes, text="📂 Abrir Resultado", command=self.abrir_arquivo_excel, state=tk.DISABLED)
        self.btn_abrir_excel.pack(side=tk.LEFT, padx=2)

    # ==========================================
    # 📊 TELA: ANÁLISE DE GASTOS MENSAIS
    # ==========================================
    def tela_analise_gastos(self):
        """Mostra seleção do XLSX, tabela de prévia e gráficos na própria janela."""
        cabecalho = ttk.Frame(self.area_conteudo)
        cabecalho.pack(fill=tk.X)

        ttk.Label(
            cabecalho,
            text="📊 Análise de Gastos Mensais",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", pady=(0, 6))

        ttk.Label(
            cabecalho,
            text=(
                "Selecione sua planilha para visualizar os valores em euros e os gráficos. "
                "Depois, clique em Calcular Análise para salvar o relatório XLSX."
            ),
            justify=tk.LEFT,
            wraplength=850
        ).pack(anchor="w", pady=(0, 8))

        frame_arquivos = ttk.Frame(cabecalho)
        frame_arquivos.pack(fill=tk.X, pady=4)
        ttk.Label(frame_arquivos, text="Planilha:").pack(side=tk.LEFT)
        self.lbl_planilha_gastos = ttk.Label(
            frame_arquivos,
            text="Nenhum arquivo selecionado",
            foreground="#777"
        )
        self.lbl_planilha_gastos.pack(side=tk.LEFT, padx=6)
        ttk.Button(
            frame_arquivos,
            text="📂 Selecionar XLSX",
            command=self.escolher_planilha_gastos
        ).pack(side=tk.LEFT, padx=4)

        frame_botoes = ttk.Frame(cabecalho)
        frame_botoes.pack(fill=tk.X, pady=(5, 8))
        self.btn_executar = ttk.Button(
            frame_botoes,
            text="▶️ Calcular Análise e Salvar XLSX",
            command=self.executar_tarefa
        )
        self.btn_executar.pack(side=tk.LEFT, padx=2)
        self.btn_abrir_gastos = ttk.Button(
            frame_botoes,
            text="📂 Abrir Relatório",
            command=self.abrir_relatorio_gastos,
            state=tk.DISABLED
        )
        self.btn_abrir_gastos.pack(side=tk.LEFT, padx=2)

        # Área com rolagem: a tabela e os gráficos ficam abaixo dos controles.
        self.gastos_preview_area = ttk.Frame(self.area_conteudo)
        self.gastos_preview_area.pack(fill=tk.BOTH, expand=True)

        self.gastos_preview_scroll = tk.Canvas(
            self.gastos_preview_area,
            highlightthickness=0,
            borderwidth=0
        )
        barra_vertical = ttk.Scrollbar(
            self.gastos_preview_area,
            orient=tk.VERTICAL,
            command=self.gastos_preview_scroll.yview
        )
        self.gastos_preview_scroll.configure(yscrollcommand=barra_vertical.set)
        barra_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        self.gastos_preview_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.gastos_preview_body = ttk.Frame(self.gastos_preview_scroll, padding=(2, 2, 10, 12))
        self.gastos_preview_window = self.gastos_preview_scroll.create_window(
            (0, 0),
            window=self.gastos_preview_body,
            anchor="nw"
        )
        self.gastos_preview_body.bind(
            "<Configure>",
            lambda _event: self.gastos_preview_scroll.configure(
                scrollregion=self.gastos_preview_scroll.bbox("all")
            )
        )
        self.gastos_preview_scroll.bind(
            "<Configure>",
            lambda event: self.gastos_preview_scroll.itemconfigure(
                self.gastos_preview_window,
                width=event.width
            )
        )
        self.gastos_preview_scroll.bind("<Button-4>", self._rolar_preview_gastos)
        self.gastos_preview_scroll.bind("<Button-5>", self._rolar_preview_gastos)

        self.lbl_preview_gastos = ttk.Label(
            self.gastos_preview_body,
            text="Selecione um arquivo Excel para mostrar a planilha e os gráficos aqui.",
            foreground="#666",
            font=("Segoe UI", 10)
        )
        self.lbl_preview_gastos.pack(anchor="w", pady=18)

    def _rolar_preview_gastos(self, evento):
        """Permite rolar a prévia com a roda do mouse no Linux."""
        if evento.num == 4:
            self.gastos_preview_scroll.yview_scroll(-3, "units")
        elif evento.num == 5:
            self.gastos_preview_scroll.yview_scroll(3, "units")

    def escolher_planilha_gastos(self):
        caminho = filedialog.askopenfilename(
            title="Selecione sua planilha de contas",
            filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
        )
        if not caminho:
            return

        self.caminho_planilha_gastos = caminho
        self.lbl_planilha_gastos.config(text=os.path.basename(caminho), foreground="#000")
        self.log(f"✅ Planilha selecionada: {os.path.basename(caminho)}")

        try:
            dados = self.ler_dados_planilha_gastos(caminho)
            self.dados_gastos_preview = dados
            self.mostrar_preview_gastos(dados)
            self.log(
                f"✅ Prévia carregada: {len(dados['categorias'])} categorias | "
                f"Total: € {dados['total_geral']:,.2f}"
            )
        except Exception as erro:
            self.dados_gastos_preview = None
            messagebox.showerror(
                "Erro ao abrir planilha",
                f"Não foi possível montar a prévia da planilha.\n\n{erro}"
            )
            self.log(f"❌ Erro na prévia do Excel: {erro}")

    def ler_dados_planilha_gastos(self, caminho):
        """Lê a matriz Categoria x Meses e calcula totais uma única vez."""
        import pandas as pd

        engine = "xlrd" if caminho.lower().endswith(".xls") else None
        bruto = pd.read_excel(caminho, header=None, engine=engine)
        tipo_layout, linha_cabecalho, colunas_mes_detectadas = detectar_layout_excel(bruto)

        if tipo_layout != "mensal" or linha_cabecalho is None:
            raise ValueError(
                "Não identifiquei a tabela de gastos por mês. "
                "Preciso encontrar uma linha com os nomes dos meses (JAN a DEZ)."
            )

        meses_detectados = sorted(
            colunas_mes_detectadas,
            key=lambda item: (item[0], item[1])
        )
        if not meses_detectados:
            raise ValueError("Nenhuma coluna de mês foi encontrada na planilha.")

        meses = [str(item[2]).strip().upper() for item in meses_detectados]
        indices_meses = [item[1] for item in meses_detectados]
        dados = bruto.iloc[linha_cabecalho + 1:].copy()

        categorias = []
        valores_por_mes = [0.0] * len(meses)
        total_geral = 0.0

        for _, linha in dados.iterrows():
            categoria_bruta = linha.iloc[0] if len(linha) else None
            if pd.isna(categoria_bruta):
                continue

            nome_categoria = str(categoria_bruta).strip()
            nome_normalizado = normalizar_celula_excel(nome_categoria)
            if not nome_categoria or "total" in nome_normalizado:
                continue

            valores = []
            total_categoria = 0.0
            for posicao, indice_coluna in enumerate(indices_meses):
                valor_bruto = linha.iloc[indice_coluna] if indice_coluna < len(linha) else 0
                valor = self.converter_para_numero(valor_bruto)
                valores.append(valor)
                total_categoria += valor
                valores_por_mes[posicao] += valor

            categorias.append({
                "nome": nome_categoria,
                "valores": valores,
                "total": total_categoria
            })
            total_geral += total_categoria

        if not categorias:
            raise ValueError("Não encontrei categorias com valores para calcular.")

        categorias.sort(key=lambda item: chave_categoria_numerica(item["nome"]))
        totais_percentuais = [
            (item["total"] / total_geral * 100) if total_geral else 0.0
            for item in categorias
        ]

        return {
            "meses": meses,
            "categorias": categorias,
            "valores_por_mes": valores_por_mes,
            "total_geral": total_geral,
            "percentuais": totais_percentuais,
            "linha_cabecalho": linha_cabecalho + 1,
        }

    @staticmethod
    def _formatar_euro_tela(valor):
        texto = f"{float(valor):,.2f}"
        texto = texto.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
        return f"€ {texto}"

    def mostrar_preview_gastos(self, dados):
        """Desenha a tabela da planilha e os dois gráficos abaixo dela."""
        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.ticker import FuncFormatter
        except ImportError as erro:
            raise RuntimeError(
                "Para mostrar os gráficos dentro do programa, instale matplotlib com: "
                "python -m pip install matplotlib"
            ) from erro

        for widget in self.gastos_preview_body.winfo_children():
            widget.destroy()
        self.gastos_canvas_widget = None
        self.gastos_figure = None
        self.gastos_tree = None

        meses = dados["meses"]
        categorias = dados["categorias"]
        total_geral = dados["total_geral"]
        valores_por_mes = dados["valores_por_mes"]

        ttk.Label(
            self.gastos_preview_body,
            text=(
                f"Prévia da planilha • {len(categorias)} categorias • "
                f"Total anual: {self._formatar_euro_tela(total_geral)}"
            ),
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(2, 8))

        quadro_tabela = ttk.LabelFrame(
            self.gastos_preview_body,
            text="📋 Planilha — valores por categoria e mês",
            padding=6
        )
        quadro_tabela.pack(fill=tk.X, expand=False, pady=(0, 10))

        nomes_colunas = ["categoria"] + [f"mes_{i}" for i in range(len(meses))] + ["total_anual", "percentual"]
        self.gastos_tree = ttk.Treeview(
            quadro_tabela,
            columns=nomes_colunas,
            show="headings",
            height=min(max(len(categorias) + 1, 5), 11)
        )
        self.gastos_tree.heading("categoria", text="Categoria")
        self.gastos_tree.column("categoria", width=155, minwidth=130, stretch=False, anchor="w")

        for indice, mes in enumerate(meses):
            coluna = f"mes_{indice}"
            self.gastos_tree.heading(coluna, text=mes)
            self.gastos_tree.column(coluna, width=105, minwidth=90, stretch=False, anchor="e")

        self.gastos_tree.heading("total_anual", text="Total Anual")
        self.gastos_tree.column("total_anual", width=125, minwidth=110, stretch=False, anchor="e")
        self.gastos_tree.heading("percentual", text="% do Total")
        self.gastos_tree.column("percentual", width=100, minwidth=90, stretch=False, anchor="e")

        barra_horizontal = ttk.Scrollbar(
            quadro_tabela,
            orient=tk.HORIZONTAL,
            command=self.gastos_tree.xview
        )
        self.gastos_tree.configure(xscrollcommand=barra_horizontal.set)
        self.gastos_tree.pack(fill=tk.X, expand=True)
        barra_horizontal.pack(fill=tk.X)

        for categoria in categorias:
            percentual = (categoria["total"] / total_geral * 100) if total_geral else 0.0
            valores_linha = [categoria["nome"]]
            valores_linha.extend(self._formatar_euro_tela(valor) for valor in categoria["valores"])
            valores_linha.append(self._formatar_euro_tela(categoria["total"]))
            valores_linha.append(f"{percentual:.2f}%".replace(".", ","))
            self.gastos_tree.insert("", tk.END, values=valores_linha)

        valores_total = ["TOTAL GERAL"]
        valores_total.extend(self._formatar_euro_tela(valor) for valor in valores_por_mes)
        valores_total.append(self._formatar_euro_tela(total_geral))
        valores_total.append("100,00%" if total_geral else "0,00%")
        self.gastos_tree.insert("", tk.END, values=valores_total, tags=("total",))
        self.gastos_tree.tag_configure("total", font=("Segoe UI", 9, "bold"))

        quadro_graficos = ttk.LabelFrame(
            self.gastos_preview_body,
            text="📈 Gráficos mensal e anual",
            padding=6
        )
        quadro_graficos.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Os dois gráficos usam barras horizontais para manter o mesmo padrão
        # visual e evitar que os valores dos 12 meses fiquem sobrepostos.
        figura = Figure(figsize=(9.2, 8.2), dpi=90)
        ax_mensal = figura.add_subplot(2, 1, 1)
        ax_anual = figura.add_subplot(2, 1, 2)

        formatador_euro = FuncFormatter(
            lambda valor, _pos: "€ " + f"{valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        # Inverte a lista para que JAN apareça no topo e DEZ no final.
        meses_grafico = list(meses)[::-1]
        valores_mensais_grafico = list(valores_por_mes)[::-1]
        barras_mensais = ax_mensal.barh(meses_grafico, valores_mensais_grafico)
        ax_mensal.set_title("Gastos por mês", fontsize=11, fontweight="bold", pad=12)
        ax_mensal.set_xlabel("Total gasto (€)")
        ax_mensal.xaxis.set_major_formatter(formatador_euro)
        ax_mensal.grid(axis="x", alpha=0.25)
        maior_valor_mensal = max(valores_por_mes, default=0)
        ax_mensal.set_xlim(0, maior_valor_mensal * 1.25 if maior_valor_mensal > 0 else 1)
        rotulos_mensais = [self._formatar_euro_tela(valor) for valor in valores_mensais_grafico]
        ax_mensal.bar_label(
            barras_mensais,
            labels=rotulos_mensais,
            padding=4,
            fontsize=8,
        )
        ax_mensal.tick_params(axis="y", labelsize=8)

        categorias_ordenadas = sorted(categorias, key=lambda item: item["total"], reverse=True)
        nomes = [item["nome"] for item in categorias_ordenadas]
        totais = [item["total"] for item in categorias_ordenadas]
        barras_anuais = ax_anual.barh(nomes[::-1], totais[::-1])
        ax_anual.set_title("Gasto anual por categoria", fontsize=11, fontweight="bold")
        ax_anual.set_xlabel("Total anual (€)")
        ax_anual.xaxis.set_major_formatter(formatador_euro)
        ax_anual.grid(axis="x", alpha=0.25)
        ax_anual.bar_label(
            barras_anuais,
            labels=[self._formatar_euro_tela(v) for v in totais[::-1]],
            padding=3,
            fontsize=8
        )
        ax_anual.tick_params(axis="y", labelsize=8)

        figura.tight_layout(pad=2.0, h_pad=2.6)
        self.gastos_figure = figura
        self.gastos_canvas_widget = FigureCanvasTkAgg(figura, master=quadro_graficos)
        self.gastos_canvas_widget.draw()
        self.gastos_canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    
    def abrir_relatorio_gastos(self):
        """Abre a pasta onde o relatório de gastos foi salvo."""
        caminho = getattr(self, "caminho_relatorio_gastos", "")

        if not caminho:
            messagebox.showinfo(
                "Informação",
                "Nenhum relatório foi gerado ainda.\nExecute a análise primeiro!"
            )
            return

        caminho = os.path.abspath(os.path.expanduser(caminho))

        if not os.path.isfile(caminho):
            messagebox.showerror(
                "Relatório não encontrado",
                f"O relatório não foi encontrado:\n{caminho}"
            )
            return

        pasta = os.path.dirname(caminho)

        try:
            if os.name == "nt":
                os.startfile(pasta)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", pasta])
            else:
                subprocess.Popen(
                    ["xdg-open", pasta],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )

            self.log(f"📂 Pasta do relatório de gastos aberta: {pasta}")

        except Exception as e:
            mensagem = f"Erro ao abrir a pasta:\n{pasta}\n\nDetalhes: {e}"
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir relatório", mensagem)

    def analise_gastos_calcular(self):
        """Gera um relatório financeiro sem somar linhas de total novamente."""
        if not self.caminho_planilha_gastos:
            self.log("❌ Selecione uma planilha primeiro!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"Analise_Gastos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not caminho_saida:
            self.log("⚠️ Operação cancelada.\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log("📊 Lendo planilha de gastos...")

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment

            dados = self.ler_dados_planilha_gastos(self.caminho_planilha_gastos)
            meses = dados["meses"]
            categorias = dados["categorias"]
            valores_por_mes = dados["valores_por_mes"]
            total_geral = dados["total_geral"]

            self.log(f"✅ Cabeçalho encontrado na linha {dados['linha_cabecalho']}")
            self.log(f"✅ {len(categorias)} categorias identificadas")

            wb = Workbook()
            ws_cat = wb.active
            ws_cat.title = "Resumo por Categoria"
            cabecalho = ["Categoria"] + meses + ["Total Anual", "% do Total"]
            ws_cat.append(cabecalho)

            for celula in ws_cat[1]:
                celula.font = Font(bold=True, color="FFFFFF")
                celula.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                celula.alignment = Alignment(horizontal="center")

            for categoria in categorias:
                percentual = (categoria["total"] / total_geral * 100) if total_geral else 0
                ws_cat.append(
                    [categoria["nome"]] + categoria["valores"] +
                    [round(categoria["total"], 2), round(percentual, 2)]
                )

            linha_total = ws_cat.max_row + 1
            ws_cat.append(["TOTAL GERAL"] + valores_por_mes + [round(total_geral, 2), 100.0 if total_geral else 0.0])
            for celula in ws_cat[linha_total]:
                celula.font = Font(bold=True)
                celula.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

            ws_mes = wb.create_sheet(title="Resumo Mensal")
            ws_mes.append(["Mês", "Total do Mês (€)", "% do Total Geral"])
            for celula in ws_mes[1]:
                celula.font = Font(bold=True, color="FFFFFF")
                celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                celula.alignment = Alignment(horizontal="center")

            for mes, valor in zip(meses, valores_por_mes):
                percentual = (valor / total_geral * 100) if total_geral else 0
                ws_mes.append([mes, round(valor, 2), round(percentual, 2)])

            linha_total_mes = ws_mes.max_row + 1
            ws_mes.append(["TOTAL GERAL", round(total_geral, 2), 100.0 if total_geral else 0.0])
            for celula in ws_mes[linha_total_mes]:
                celula.font = Font(bold=True)
                celula.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

            ws_cat.freeze_panes = "A2"
            for linha in ws_cat.iter_rows(min_row=2, max_col=len(meses) + 3):
                for celula in linha[1:len(meses) + 2]:
                    celula.number_format = FORMATO_EURO_EXCEL
                ws_cat.cell(row=linha[0].row, column=len(meses) + 3).number_format = FORMATO_PERCENTUAL_EXCEL

            ws_mes.freeze_panes = "A2"
            for linha in range(2, ws_mes.max_row + 1):
                ws_mes.cell(row=linha, column=2).number_format = FORMATO_EURO_EXCEL
                ws_mes.cell(row=linha, column=3).number_format = FORMATO_PERCENTUAL_EXCEL

            for ws in wb.worksheets:
                for coluna in ws.columns:
                    largura = max((len(str(c.value)) for c in coluna if c.value is not None), default=0)
                    ws.column_dimensions[coluna[0].column_letter].width = max(largura + 3, 12)

            # Larguras adequadas para o símbolo € e duas casas decimais.
            ws_cat.column_dimensions["A"].width = 22
            for indice_coluna in range(2, len(meses) + 2):
                letra = ws_cat.cell(row=1, column=indice_coluna).column_letter
                ws_cat.column_dimensions[letra].width = 15
            ws_cat.column_dimensions[
                ws_cat.cell(row=1, column=len(meses) + 2).column_letter
            ].width = 18
            ws_cat.column_dimensions[
                ws_cat.cell(row=1, column=len(meses) + 3).column_letter
            ].width = 16
            ws_mes.column_dimensions["A"].width = 16
            ws_mes.column_dimensions["B"].width = 20
            ws_mes.column_dimensions["C"].width = 20

            adicionar_graficos_resumo(
                wb,
                ws_mes,
                ws_cat,
                coluna_total_anual=len(meses) + 2,
            )

            wb.save(caminho_saida)
            self.caminho_relatorio_gastos = os.path.abspath(caminho_saida)
            self.root.after(0, lambda: self.seguro_config(self.btn_abrir_gastos, state=tk.NORMAL))

            self.log("\n" + "=" * 50)
            self.log("✅ ANÁLISE DE GASTOS CONCLUÍDA")
            self.log(f"💰 TOTAL GERAL: € {total_geral:,.2f}")
            self.log(f"📂 Relatório salvo: {caminho_saida}")
            self.log("📊 Maiores categorias por valor gasto:")
            maiores_categorias = sorted(categorias, key=lambda item: item["total"], reverse=True)
            for indice, categoria in enumerate(maiores_categorias[:5], start=1):
                self.log(f"   {indice}. {categoria['nome']}: € {categoria['total']:,.2f}")
            salvar_no_historico("Análise de Gastos", caminho_saida, len(categorias))

        except Exception as e:
            self.log(f"\n❌ Erro na análise de gastos: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    # ==========================================
    # 📁 FUNÇÕES GERAIS
    # ==========================================
    def escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta")
        if pasta:
            self.pasta_selecionada = pasta
            self.lbl_pasta.config(text=pasta, foreground="#000")

    def escolher_arquivos(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione os arquivos",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivos:
            self.caminhos_selecionados = list(arquivos)
            self.lbl_arquivos.config(text=f"{len(arquivos)} arquivo(s) selecionado(s)", foreground="#000")

    def quando_selecionar_pdf(self, evento):
        tem = bool(self.tree_pdfs.selection())
        self.seguro_config(self.btn_abrir_pdf, state=tk.NORMAL if tem else tk.DISABLED)
        self.seguro_config(self.btn_ver_pdf, state=tk.NORMAL if tem else tk.DISABLED)

    def quando_selecionar_email(self, evento):
        tem = bool(self.tree_emails.selection())
        self.seguro_config(self.btn_ver, state=tk.NORMAL if tem else tk.DISABLED)
        self.seguro_config(self.btn_salvar, state=tk.NORMAL if tem else tk.DISABLED)

    def abrir_pasta_do_pdf(self):
        sel = self.tree_pdfs.selection()
        if sel and self.lista_pdfs_exibidos:
            idx = self.tree_pdfs.index(sel[0])
            caminho = os.path.dirname(self.lista_pdfs_exibidos[idx]["caminho"])
            if os.name == "nt":
                os.startfile(caminho)
            else:
                subprocess.Popen(["xdg-open", caminho])

    def abrir_pdf_selecionado(self):
        sel = self.tree_pdfs.selection()
        if sel and self.lista_pdfs_exibidos:
            idx = self.tree_pdfs.index(sel[0])
            caminho = self.lista_pdfs_exibidos[idx]["caminho"]
            if os.name == "nt":
                os.startfile(caminho)
            else:
                subprocess.Popen(["xdg-open", caminho])

    def _caminho_cofre_email(self):
        """Retorna a pasta privada de dados do usuário em Linux ou Windows."""
        if os.name == "nt":
            base = os.environ.get("APPDATA") or os.path.expanduser("~")
        else:
            base = os.environ.get("XDG_DATA_HOME") or os.path.join(
                os.path.expanduser("~"), ".local", "share"
            )
        pasta = os.path.join(base, "SistemaAutomacao")
        os.makedirs(pasta, mode=0o700, exist_ok=True)
        if os.name != "nt":
            try:
                os.chmod(pasta, 0o700)
            except OSError:
                pass
        return os.path.join(pasta, ARQUIVO_COFRE_EMAIL)

    def _obter_cofre_credenciais(self, criar_se_necessario=False):
        """Abre o cofre criptografado sem chamar o Chaveiro do GNOME.

        A senha mestra é solicitada uma vez por execução do programa e a chave
        fica apenas em memória até o aplicativo ser fechado.
        """
        if self._cofre_email is not None and self._cofre_email.get("desbloqueado"):
            return self._cofre_email

        caminho = self._caminho_cofre_email()
        existe = os.path.isfile(caminho)
        if not existe and not criar_se_necessario:
            return None

        try:
            from cryptography.fernet import Fernet, InvalidToken
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
            import base64
            import secrets
            import tempfile

            def derivar_chave(senha_mestra, salt):
                kdf = PBKDF2HMAC(
                    algorithm=hashes.SHA256(),
                    length=32,
                    salt=salt,
                    iterations=PBKDF2_ITERACOES,
                )
                return base64.urlsafe_b64encode(kdf.derive(senha_mestra.encode("utf-8")))

            if existe:
                senha_mestra = simpledialog.askstring(
                    "Desbloquear cofre do aplicativo",
                    "Digite a senha mestra do Sistema de Automação:",
                    show="*",
                    parent=self.root,
                )
                if senha_mestra is None:
                    return None
                if not senha_mestra:
                    messagebox.showwarning(
                        "Senha mestra vazia",
                        "Digite a senha mestra para abrir os acessos salvos.",
                        parent=self.root,
                    )
                    return None

                with open(caminho, "r", encoding="utf-8") as arquivo:
                    estrutura = json.load(arquivo)
                salt = base64.b64decode(estrutura["salt"])
                chave = derivar_chave(senha_mestra, salt)
                fernet = Fernet(chave)
                try:
                    dados = json.loads(fernet.decrypt(estrutura["dados"].encode("ascii")).decode("utf-8"))
                except InvalidToken as exc:
                    raise ValueError("Senha mestra incorreta. Tente novamente.") from exc
                if not isinstance(dados, dict) or not isinstance(dados.get("contas", {}), dict):
                    raise ValueError("O arquivo do cofre está em formato inválido.")
            else:
                senha_mestra = simpledialog.askstring(
                    "Criar cofre de senhas",
                    "Crie uma senha mestra forte para proteger os e-mails salvos (mínimo 12 caracteres):",
                    show="*",
                    parent=self.root,
                )
                if senha_mestra is None:
                    return None
                if len(senha_mestra) < 12:
                    messagebox.showwarning(
                        "Senha muito curta",
                        "A senha mestra precisa ter pelo menos 12 caracteres. Use uma combinação longa e exclusiva.",
                        parent=self.root,
                    )
                    return None
                confirmacao = simpledialog.askstring(
                    "Confirmar senha mestra",
                    "Digite novamente a senha mestra:",
                    show="*",
                    parent=self.root,
                )
                if confirmacao is None:
                    return None
                if senha_mestra != confirmacao:
                    messagebox.showerror(
                        "Senhas diferentes",
                        "As senhas mestras não coincidem. Nenhum cofre foi criado.",
                        parent=self.root,
                    )
                    return None
                salt = secrets.token_bytes(16)
                fernet = Fernet(derivar_chave(senha_mestra, salt))
                dados = {"versao": 1, "contas": {}}

                # Cria o cofre vazio imediatamente, com permissões privadas.
                estrutura = {
                    "versao": 1,
                    "salt": base64.b64encode(salt).decode("ascii"),
                    "dados": fernet.encrypt(json.dumps(dados).encode("utf-8")).decode("ascii"),
                }
                fd, temporario = tempfile.mkstemp(prefix="cofre_", suffix=".tmp", dir=os.path.dirname(caminho))
                try:
                    with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                        json.dump(estrutura, arquivo, ensure_ascii=False)
                        arquivo.flush()
                        os.fsync(arquivo.fileno())
                    os.replace(temporario, caminho)
                    if os.name != "nt":
                        os.chmod(caminho, 0o600)
                finally:
                    if os.path.exists(temporario):
                        os.remove(temporario)

            self._cofre_email = {
                "fernet": fernet,
                "dados": dados,
                "caminho": caminho,
                "desbloqueado": True,
            }
            return self._cofre_email

        except ImportError as erro:
            messagebox.showerror(
                "Dependência ausente",
                "A biblioteca de criptografia não está instalada.\n\n"
                "Instale com: python -m pip install cryptography\n\n" + str(erro),
                parent=self.root,
            )
            return None
        except (ValueError, OSError, KeyError, json.JSONDecodeError) as erro:
            messagebox.showerror(
                "Não foi possível abrir o cofre",
                str(erro),
                parent=self.root,
            )
            return None

    def _salvar_cofre_credenciais(self, cofre):
        """Grava o conteúdo cifrado de forma atômica e com permissão privada."""
        import tempfile

        estrutura = {
            "versao": 1,
            "salt": self._ler_salt_cofre(cofre),
            "dados": cofre["fernet"].encrypt(
                json.dumps(cofre["dados"], ensure_ascii=False).encode("utf-8")
            ).decode("ascii"),
        }
        fd, temporario = tempfile.mkstemp(
            prefix="cofre_", suffix=".tmp", dir=os.path.dirname(cofre["caminho"])
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as arquivo:
                json.dump(estrutura, arquivo, ensure_ascii=False)
                arquivo.flush()
                os.fsync(arquivo.fileno())
            os.replace(temporario, cofre["caminho"])
            if os.name != "nt":
                os.chmod(cofre["caminho"], 0o600)
        finally:
            if os.path.exists(temporario):
                os.remove(temporario)

    def _ler_salt_cofre(self, cofre):
        import base64
        with open(cofre["caminho"], "r", encoding="utf-8") as arquivo:
            estrutura = json.load(arquivo)
        return estrutura["salt"]

    def _ler_lista_acessos_salvos(self, cofre):
        """Lista os e-mails salvos no cofre local criptografado."""
        contas = cofre["dados"].setdefault("contas", {})
        return list(contas.keys())

    def mostrar_ajuda_senha_app(self):
        """Explica como criar uma senha de app do Google para a conexão IMAP."""
        textos = {
            "pt": {
                "titulo": "Como gerar a senha de app do Gmail",
                "corpo": (
                    "O Google normalmente bloqueia a senha comum da conta em programas que acessam o Gmail por IMAP. "
                    "Para conectar este sistema, você precisa criar uma SENHA DE APP, diferente da senha usada para entrar no Gmail.\n\n"
                    "PASSO A PASSO\n"
                    "1. Abra sua Conta Google: https://myaccount.google.com/\n"
                    "2. Entre em Segurança.\n"
                    "3. Ative a Verificação em duas etapas, caso ainda não esteja ativa.\n"
                    "4. Abra a página de Senhas de app: https://myaccount.google.com/apppasswords\n"
                    "5. Faça login novamente, se o Google solicitar.\n"
                    "6. Na criação da senha, dê um nome para identificar o aplicativo, por exemplo: Correio ou Sistema de Automação.\n"
                    "7. Clique em Criar. O Google vai gerar uma senha de 16 caracteres.\n"
                    "8. Copie o código gerado e cole no campo Senha de App deste programa.\n\n"
                    "IMPORTANTE\n"
                    "• O nome Correio é apenas uma identificação; o código é gerado pelo Google.\n"
                    "• A senha pode aparecer dividida em grupos por espaços. Você pode colar com ou sem espaços; o sistema os remove automaticamente.\n"
                    "• Não informe sua senha normal do Gmail neste programa.\n"
                    "• A senha de app é privada. Não a compartilhe com ninguém; digite-a diretamente no programa.\n"
                    "• Se a opção Senhas de app não aparecer, consulte as orientações de segurança da própria Conta Google para verificar como habilitá-la na sua conta."
                ), "conta": "Abrir Conta Google", "senha": "Abrir Senhas de app", "fechar": "Fechar"
            },
            "en": {
                "titulo": "How to create a Gmail app password",
                "corpo": (
                    "Google normally blocks your regular account password in programs that access Gmail through IMAP. "
                    "To connect this system, create an APP PASSWORD. It is different from the password you use to sign in to Gmail.\n\n"
                    "STEP BY STEP\n"
                    "1. Open your Google Account: https://myaccount.google.com/\n"
                    "2. Go to Security.\n"
                    "3. Turn on 2-Step Verification if it is not already enabled.\n"
                    "4. Open App passwords: https://myaccount.google.com/apppasswords\n"
                    "5. Sign in again if Google asks you to.\n"
                    "6. Give the password a name to identify the app, for example: Mail or Automation System.\n"
                    "7. Click Create. Google will generate a 16-character password.\n"
                    "8. Copy the generated code and paste it into this program's App Password field.\n\n"
                    "IMPORTANT\n"
                    "• The name Mail is only an identifier; Google generates the code.\n"
                    "• The password may be displayed in groups separated by spaces. Paste it with or without spaces; the system removes spaces automatically.\n"
                    "• Do not enter your regular Gmail password in this program.\n"
                    "• Your app password is private. Never share it; enter it directly into the program.\n"
                    "• If App passwords is not available, consult the security guidance in your Google Account to check how to enable it for your account."
                ), "conta": "Open Google Account", "senha": "Open App passwords", "fechar": "Close"
            },
            "es": {
                "titulo": "Cómo crear una contraseña de aplicación de Gmail",
                "corpo": (
                    "Google normalmente bloquea la contraseña habitual de la cuenta en programas que acceden a Gmail por IMAP. "
                    "Para conectar este sistema, debe crear una CONTRASEÑA DE APLICACIÓN, distinta de la contraseña que usa para entrar en Gmail.\n\n"
                    "PASO A PASO\n"
                    "1. Abra su Cuenta de Google: https://myaccount.google.com/\n"
                    "2. Entre en Seguridad.\n"
                    "3. Active la verificación en dos pasos si todavía no está activa.\n"
                    "4. Abra la página de Contraseñas de aplicación: https://myaccount.google.com/apppasswords\n"
                    "5. Inicie sesión de nuevo si Google se lo solicita.\n"
                    "6. Asigne un nombre para identificar la aplicación, por ejemplo: Correo o Sistema de Automatización.\n"
                    "7. Haga clic en Crear. Google generará una contraseña de 16 caracteres.\n"
                    "8. Copie el código generado y péguelo en el campo Contraseña de aplicación del programa.\n\n"
                    "IMPORTANTE\n"
                    "• El nombre Correo es solo una identificación; Google genera el código.\n"
                    "• La contraseña puede aparecer en grupos separados por espacios. Puede pegarla con o sin espacios; el sistema los elimina automáticamente.\n"
                    "• No introduzca aquí su contraseña habitual de Gmail.\n"
                    "• La contraseña de aplicación es privada. No la comparta con nadie; introdúzcala directamente en el programa.\n"
                    "• Si no aparece la opción, consulte las indicaciones de seguridad de su propia Cuenta de Google para comprobar cómo habilitarla."
                ), "conta": "Abrir Cuenta de Google", "senha": "Abrir Contraseñas de aplicación", "fechar": "Cerrar"
            }
        }
        idioma = textos[self.idioma]
        janela = tk.Toplevel(self.root)
        janela.title(idioma["titulo"])
        janela.transient(self.root)
        janela.grab_set()
        janela.geometry("700x590")
        janela.minsize(580, 480)
        principal = ttk.Frame(janela, padding=16)
        principal.pack(fill=tk.BOTH, expand=True)
        ttk.Label(principal, text=idioma["titulo"], font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        caixa = tk.Text(principal, wrap=tk.WORD, height=22, font=("Segoe UI", 10))
        caixa.pack(fill=tk.BOTH, expand=True)
        caixa.insert("1.0", idioma["corpo"])
        caixa.configure(state=tk.DISABLED)
        acoes = ttk.Frame(principal)
        acoes.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(acoes, text=idioma["conta"], command=lambda: webbrowser.open("https://myaccount.google.com/")).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(acoes, text=idioma["senha"], command=lambda: webbrowser.open("https://myaccount.google.com/apppasswords")).pack(side=tk.LEFT)
        ttk.Button(acoes, text=idioma["fechar"], command=janela.destroy).pack(side=tk.RIGHT)

    def salvar_acesso_email(self):
        """Salva vários acessos no cofre local sem usar o Chaveiro do GNOME."""
        email_usuario = self.ent_email.get().strip()
        senha = "".join(self.ent_senha.get().split())

        if not email_usuario or not senha:
            messagebox.showwarning(
                "Dados incompletos",
                "Preencha o e-mail e a senha de app antes de salvar."
            )
            return

        if "@" not in email_usuario or " " in email_usuario:
            messagebox.showwarning(
                "E-mail inválido",
                "Confira o endereço de e-mail informado."
            )
            return

        cofre = self._obter_cofre_credenciais(criar_se_necessario=True)
        if cofre is None:
            return

        contas = cofre["dados"].setdefault("contas", {})
        contas[email_usuario] = senha
        try:
            self._salvar_cofre_credenciais(cofre)
        except Exception as erro:
            self.log(f"❌ Não foi possível salvar o acesso: {erro}")
            messagebox.showerror(
                "Não foi possível salvar",
                f"O acesso não foi salvo corretamente.\n\n{erro}"
            )
            return

        self.log(f"🔐 Acesso de {email_usuario} salvo no cofre criptografado local.")
        messagebox.showinfo(
            "Acesso salvo",
            f"O acesso de {email_usuario} foi salvo com criptografia.\n\n"
            f"Você possui {len(contas)} acesso(s) salvo(s).\n\n"
            "A senha mestra será solicitada somente na primeira utilização do cofre em cada execução. "
            "Guarde essa senha: não existe recuperação automática."
        )

    def usar_acesso_salvo(self):
        """Abre a lista de e-mails salvos e preenche a conta escolhida."""
        cofre = self._obter_cofre_credenciais(criar_se_necessario=False)
        if cofre is None:
            messagebox.showinfo(
                "Nenhum acesso salvo",
                "Ainda não existe um cofre com acessos salvos.\n\n"
                "Preencha o e-mail e a senha de app e clique em 'Salvar acesso'."
            )
            return

        emails_salvos = self._ler_lista_acessos_salvos(cofre)
        if not emails_salvos:
            messagebox.showinfo(
                "Nenhum acesso salvo",
                "Ainda não existe acesso salvo. Preencha e-mail e senha de app e salve."
            )
            return

        janela = tk.Toplevel(self.root)
        janela.title("📧 Escolher acesso salvo")
        janela.geometry("430x300")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(
            janela,
            text="Selecione a conta que deseja utilizar:",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 7))

        frame_lista = ttk.Frame(janela, padding=(12, 0, 12, 8))
        frame_lista.pack(fill=tk.BOTH, expand=True)

        lista = tk.Listbox(frame_lista, height=8, exportselection=False)
        barra = ttk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=lista.yview)
        lista.configure(yscrollcommand=barra.set)
        lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        barra.pack(side=tk.RIGHT, fill=tk.Y)

        for endereco in emails_salvos:
            lista.insert(tk.END, endereco)

        lista.selection_set(0)
        lista.activate(0)

        frame_botoes = ttk.Frame(janela, padding=12)
        frame_botoes.pack(fill=tk.X)

        def carregar_selecionado(event=None):
            selecao = lista.curselection()
            if not selecao:
                messagebox.showwarning(
                    "Selecione uma conta",
                    "Escolha um e-mail da lista primeiro.",
                    parent=janela,
                )
                return

            endereco = lista.get(selecao[0])
            senha_salva = cofre["dados"].get("contas", {}).get(endereco)
            if not senha_salva:
                messagebox.showwarning(
                    "Senha não encontrada",
                    f"A senha de {endereco} não foi encontrada no cofre.",
                    parent=janela,
                )
                return

            self.ent_email.delete(0, tk.END)
            self.ent_email.insert(0, endereco)
            self.ent_senha.delete(0, tk.END)
            self.ent_senha.insert(0, senha_salva)
            self.log(f"✅ Acesso salvo carregado para {endereco}.")
            janela.destroy()

        ttk.Button(
            frame_botoes,
            text="Usar conta selecionada",
            command=carregar_selecionado
        ).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(
            frame_botoes,
            text="Cancelar",
            command=janela.destroy
        ).pack(side=tk.RIGHT)

        lista.bind("<Double-Button-1>", carregar_selecionado)
        lista.focus_set()

    def visualizar_email(self):
        sel = self.tree_emails.selection()
        if not sel: return
        idx = self.tree_emails.index(sel[0])
        em = self.emails_carregados[idx]
        janela = tk.Toplevel(self.root)
        janela.title(f"📧 {em['assunto'][:50]}")
        janela.geometry("600x450")
        ttk.Label(janela, text=f"De: {em['remetente']}\nData: {em['data']}", font=("Segoe UI", 10, "bold")).pack(pady=10, padx=10)
        txt = scrolledtext.ScrolledText(janela, wrap=tk.WORD)
        txt.pack(padx=10, pady=(0,10), fill=tk.BOTH, expand=True)
        txt.insert(tk.END, em["corpo"])
        txt.config(state=tk.DISABLED)

    def salvar_email(self):
        sel = self.tree_emails.selection()
        if not sel: return
        idx = self.tree_emails.index(sel[0])
        email_atual = self.emails_carregados[idx]

        caminho = filedialog.asksaveasfilename(
            defaultextension=".eml",
            filetypes=[("Arquivo de E-mail", "*.eml"), ("Texto", "*.txt"), ("Todos os arquivos", "*.*")],
            initialfile=email_atual["assunto"][:50].replace("/", "_").replace("\\", "_")
        )
        if caminho:
            try:
                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(f"De: {email_atual['remetente']}\n")
                    f.write(f"Assunto: {email_atual['assunto']}\n")
                    f.write(f"Data: {email_atual['data']}\n")
                    f.write("-" * 60 + "\n\n")
                    f.write(email_atual.get("corpo", ""))
                self.log(f"✅ E-mail salvo em: {caminho}\n")
            except Exception as e:
                self.log(f"❌ Erro ao salvar: {str(e)}\n")

    def converter_para_numero(self, valor):
        """Converte números do Excel e formatos monetários BR/EU para float."""
        import math

        if valor is None:
            return 0.0

        try:
            # Valores já numéricos vindos do pandas/openpyxl.
            if isinstance(valor, (int, float)):
                numero = float(valor)
                return numero if math.isfinite(numero) else 0.0

            texto = str(valor).strip()
            if not texto or texto.lower() in {"nan", "none", "nat", "-", "—", "n/a", "na"}:
                return 0.0

            negativo = texto.startswith("(") and texto.endswith(")")
            if negativo:
                texto = texto[1:-1].strip()

            # Remove moedas, espaços comuns e espaços não separáveis.
            texto = (texto.replace("€", "").replace("R$", "")
                           .replace("EUR", "").replace("BRL", "")
                           .replace("$", "").replace("\xa0", "")
                           .replace(" ", "").replace("'", ""))

            # Se vírgula e ponto coexistem, o último separador é o decimal.
            if "," in texto and "." in texto:
                if texto.rfind(",") > texto.rfind("."):
                    texto = texto.replace(".", "").replace(",", ".")
                else:
                    texto = texto.replace(",", "")
            elif "," in texto:
                texto = texto.replace(",", ".")

            numero = float(texto)
            if negativo:
                numero = -abs(numero)
            return numero if math.isfinite(numero) else 0.0

        except (ValueError, TypeError, OverflowError):
            return 0.0

    def executar_tarefa(self):
        if not self.script_atual:
            self.log("❌ Nenhuma tarefa selecionada!")
            return
        self.log(f"\n{'='*60}")
        self.log(f"▶️ EXECUTANDO: {SCRIPTS[self.script_atual]['titulo']}")
        self.log(f"{'='*60}\n")

        if self.script_atual == "organizar_pdfs":
            threading.Thread(target=self.organizar_pdfs_por_data, daemon=True).start()
        elif self.script_atual == "buscar_email":
            threading.Thread(target=self.executar_busca_emails, daemon=True).start()
        elif self.script_atual == "processar_excel":
            threading.Thread(target=self.processar_excel, daemon=True).start()
        elif self.script_atual == "analise_gastos":
            threading.Thread(target=self.analise_gastos_calcular, daemon=True).start()

        self.seguro_config(self.btn_executar, state=tk.DISABLED)
        self.seguro_config(self.btn_parar, state=tk.NORMAL)

    def parar_busca(self):
        self.buscando = False
        self.log(" ⏹ Solicitação de parada enviada...\n")
        self.seguro_config(self.btn_parar, state=tk.DISABLED)
        self.seguro_config(self.btn_executar, state=tk.NORMAL)

    # ==========================================
    # 📂 PDFs
    # ==========================================
    def buscar_pdfs_no_computador(self, pasta_alvo=None):
        if pasta_alvo is None:
            pasta_alvo = self.pasta_selecionada
        self.log(f" 🔍 Procurando PDFs em: {pasta_alvo}")
        encontrados = []
        if not os.path.exists(pasta_alvo):
            self.log(f" ⚠️ ERRO: A pasta não existe — {pasta_alvo}")
            return encontrados

        for raiz, pastas, arquivos in os.walk(pasta_alvo):
            for arq in arquivos:
                if arq.lower().endswith(".pdf"):
                    caminho_completo = os.path.join(raiz, arq)
                    try:
                        tempo = os.path.getmtime(caminho_completo)
                        data_criacao = datetime.fromtimestamp(tempo)
                    except:
                        data_criacao = datetime.now()
                    encontrados.append({
                        "nome": arq,
                        "caminho": caminho_completo,
                        "data_criacao": data_criacao
                    })
                    self.log(f" 📄 {arq} — {data_criacao.strftime('%d/%m/%Y')}")

        encontrados.sort(key=lambda x: x["data_criacao"], reverse=True)
        self.log(f" ✅ Total: {len(encontrados)} PDF(s)\n")
        return encontrados

    def listar_pdfs_na_interface(self, lista_pdfs):
        """Atualiza a lista completa e reaplica o filtro digitado, se houver."""
        self.lista_pdfs_encontrados = lista_pdfs
        self.filtrar_pdfs_na_interface()

    def filtrar_pdfs_na_interface(self, *_):
        """Filtra os PDFs pelo nome, sem alterar a lista original encontrada."""
        if not hasattr(self, "tree_pdfs") or not self.tree_pdfs.winfo_exists():
            return

        termo = self.var_busca_pdf.get().strip().casefold()
        if termo:
            filtrados = [
                item for item in self.lista_pdfs_encontrados
                if termo in item["nome"].casefold()
            ]
        else:
            filtrados = list(self.lista_pdfs_encontrados)

        self.lista_pdfs_exibidos = filtrados
        self.tree_pdfs.delete(*self.tree_pdfs.get_children())
        for item in filtrados:
            self.tree_pdfs.insert("", tk.END, values=(
                item["data_criacao"].strftime("%d/%m/%Y"),
                item["nome"],
                item["caminho"]
            ))

        if hasattr(self, "lbl_resultado_pdfs") and self.lbl_resultado_pdfs.winfo_exists():
            self.lbl_resultado_pdfs.config(
                text=f"Exibindo {len(filtrados)} de {len(self.lista_pdfs_encontrados)} PDF(s)"
            )

    def organizar_pdfs_por_data(self):
        if not self.pasta_selecionada:
            self.log("❌ Nenhuma pasta selecionada!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log("=" * 50)
        self.log("📂 LISTA DE PDFs ENCONTRADOS")
        self.log(f" 📁 Pasta: {self.pasta_selecionada}")
        self.log("=" * 50)

        self.lista_pdfs_encontrados = self.buscar_pdfs_no_computador()
        self.root.after(0, lambda: self.listar_pdfs_na_interface(self.lista_pdfs_encontrados))
        salvar_no_historico("Organizar PDFs", self.pasta_selecionada, len(self.lista_pdfs_encontrados))
        self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
        self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    def atualizar_lista_emails(self):
        """Atualiza a tabela de e-mails na thread principal do Tkinter."""
        if not hasattr(self, "tree_emails") or not self.tree_emails.winfo_exists():
            return

        self.tree_emails.delete(*self.tree_emails.get_children())
        for item in self.emails_carregados:
            self.tree_emails.insert(
                "",
                tk.END,
                values=(item["data"], item["remetente"], item["assunto"])
            )

        estado = tk.NORMAL if self.emails_carregados else tk.DISABLED
        self.seguro_config(self.btn_ver, state=estado)
        self.seguro_config(self.btn_salvar, state=estado)

    # ==========================================
    # 📧 E-MAILS
    # ==========================================
    def executar_busca_emails(self):
        email_usuario = self.ent_email.get().strip()
        # O Google pode exibir a senha de app em grupos separados por espaços.
        senha = "".join(self.ent_senha.get().split())
        remetente_filtro = self.ent_nome_remetente.get().strip() if hasattr(self, 'ent_nome_remetente') else ""
        tipo_doc = self.var_tipo_doc.get()

        texto_selecionado = self.var_periodo_texto.get()
        periodo = "0"
        for chave, valor in self.opcoes_periodo.items():
            if valor == texto_selecionado:
                periodo = chave
                break

        if not email_usuario or not senha:
            self.log("❌ Preencha e-mail e senha primeiro!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.emails_carregados = []
        self.root.after(0, lambda: self.tree_emails.delete(*self.tree_emails.get_children()))
        self.buscando = True

        imap = None
        try:
            self.log(" 🔌 Conectando ao Gmail com TLS...")
            # Timeout evita deixar a interface presa indefinidamente em uma rede instável.
            # IMAP4_SSL usa o contexto TLS padrão do Python (validação de certificado ativa).
            imap = imaplib.IMAP4_SSL("imap.gmail.com", 993, timeout=30)
            imap.login(email_usuario, senha)
            status_selecao, _ = imap.select("INBOX", readonly=True)
            if status_selecao != "OK":
                raise RuntimeError("Não foi possível acessar a caixa de entrada (INBOX).")
            self.log(" ✅ Conectado! Buscando e-mails...")

            filtro_data = ""
            if periodo != "0":
                dias = int(periodo)
                data_inicio = datetime.now() - timedelta(days=dias)
                filtro_data = f'SINCE "{data_inicio.strftime("%d-%b-%Y")}"'

            if remetente_filtro:
                # Evita que aspas ou quebras de linha alterem a sintaxe do comando IMAP.
                termo = remetente_filtro.strip().replace("\r", " ").replace("\n", " ")
                termo = termo.replace("\\", "\\\\").replace('"', '\\"')
                filtro_busca = f'OR FROM "{termo}" SUBJECT "{termo}"'
            else:
                filtro_busca = "ALL"

            if filtro_data and filtro_busca != "ALL":
                filtro_busca = f"({filtro_busca}) {filtro_data}"
            elif filtro_data:
                filtro_busca = filtro_data

            self.log(f" 🔍 Filtro: {filtro_busca}")
            status, dados = imap.search(None, filtro_busca)
            if status != "OK" or not dados:
                raise RuntimeError("O Gmail não conseguiu concluir a pesquisa de e-mails.")
            ids = dados[0].split()
            self.log(f" 📨 Encontrados {len(ids)} e-mail(s)")

            palavras_chave = []
            if tipo_doc != "todos":
                palavras = {
                    "boleto": ["boleto"],
                    "fatura": ["fatura"],
                    "contrato": ["contrato"],
                    "notafiscal": ["nota fiscal", "nfe", "nf-e"],
                    "pedido": ["pedido"]
                }
                palavras_chave = palavras.get(tipo_doc, [])

            limite = min(len(ids), 100)
            for id_email in reversed(ids[-limite:]):
                if not self.buscando:
                    self.log(" ⏹ Interrompido")
                    break

                status, dados = imap.fetch(id_email, "(RFC822)")
                if status != "OK" or not dados or not isinstance(dados[0], tuple):
                    continue
                mensagem = email.message_from_bytes(dados[0][1])

                assunto = self.decodificar_cabecalho(mensagem["Subject"]) or "(Sem assunto)"
                remetente = self.decodificar_cabecalho(mensagem["From"])
                data_raw = mensagem["Date"]
                try:
                    data_dt = parsedate_to_datetime(data_raw)
                    data_formatada = data_dt.strftime("%d/%m/%Y %H:%M")
                except:
                    data_formatada = data_raw or "Sem data"

                corpo = ""
                if mensagem.is_multipart():
                    for parte in mensagem.walk():
                        tipo = parte.get_content_type()
                        try:
                            if tipo == "text/plain":
                                texto = parte.get_payload(decode=True).decode("utf-8", errors="replace")
                                corpo += texto + "\n"
                            elif tipo == "text/html" and not corpo.strip():
                                html = parte.get_payload(decode=True).decode("utf-8", errors="replace")
                                corpo = extrair_texto_do_html(html)
                        except:
                            continue
                else:
                    try:
                        bruto = mensagem.get_payload(decode=True).decode("utf-8", errors="replace")
                        if "<html" in bruto.lower():
                            corpo = extrair_texto_do_html(bruto)
                        else:
                            corpo = bruto
                    except:
                        corpo = "[Não foi possível ler]"

                if palavras_chave:
                    texto_busca = (assunto + " " + corpo).lower()
                    if not any(p in texto_busca for p in palavras_chave):
                        continue

                email_dict = {
                    "data": data_formatada,
                    "remetente": remetente,
                    "assunto": assunto,
                    "corpo": corpo.strip() or "(Sem conteúdo)"
                }
                self.emails_carregados.append(email_dict)


            self.root.after(0, self.atualizar_lista_emails)
            self.log(f"✅ Pronto! {len(self.emails_carregados)} e-mail(s) carregado(s)\n")
            salvar_no_historico("Buscar E-mail", remetente_filtro or tipo_doc, len(self.emails_carregados))

        except imaplib.IMAP4.error as e:
            self.log(f"❌ Erro de login: {str(e)}")
            self.log("💡 Use SENHA DE APP — não sua senha normal\n")
        except Exception as e:
            self.log(f"❌ Erro: {str(e)}\n")
            import traceback
            self.log(traceback.format_exc() + "\n")
        finally:
            # Fecha a sessão mesmo quando login, busca ou leitura falhar.
            if imap is not None:
                try:
                    imap.logout()
                except Exception:
                    pass
            self.buscando = False
            self.root.after(0, self.atualizar_lista_emails)
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))

    # ==========================================
    # 📊 EXCEL — CONSOLIDAR PLANILHAS
    # ==========================================
    def processar_excel(self):
        """Consolida planilhas de transações ou matrizes anuais de contas.

        Transações: soma valores, resume por categoria e ordena pela data
        mais nova primeiro.
        Matriz anual: reconhece uma linha de meses mesmo quando existe um
        título mesclado acima, calcula totais por categoria e por mês.
        """
        if not self.caminhos_selecionados:
            self.log("❌ Nenhum arquivo Excel selecionado!\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"Processado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not caminho_saida:
            self.log("⚠️ Operação cancelada.\n")
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))
            return

        self.log(f"📂 Processando {len(self.caminhos_selecionados)} arquivo(s)...")
        self.log(f"💾 Salvando em: {caminho_saida}")

        try:
            import pandas as pd
            from datetime import date
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment

            registros = []
            totais_categoria_transacoes = {}
            total_transacoes = 0.0
            encontrou_data = False

            # Dados da matriz anual: {categoria: {numero_mes: valor}}
            totais_mensais_categoria = {}
            totais_por_mes = {}
            nomes_meses = {}
            total_matriz = 0.0
            arquivos_lidos = 0

            def ausente(valor):
                try:
                    return pd.isna(valor)
                except (TypeError, ValueError):
                    return False

            def localizar_coluna_transacao(cabecalhos, tipo):
                nomes = [normalizar_celula_excel(item) for item in cabecalhos]
                grupos = {
                    "data": ("data", "date", "dia"),
                    "descricao": ("descricao", "description", "desc", "nome", "produto", "item", "historico"),
                    "categoria": ("categoria", "category", "tipo", "classe"),
                    "valor": ("valor", "value", "preco", "importe", "montante", "amount", "total"),
                }
                sinonimos = grupos[tipo]
                for sinonimo in sinonimos:
                    for indice, nome in enumerate(nomes):
                        if nome == sinonimo:
                            return indice
                for sinonimo in sinonimos:
                    for indice, nome in enumerate(nomes):
                        if re.search(rf"\b{re.escape(sinonimo)}\b", nome):
                            return indice
                return None

            def converter_data_excel(valor):
                if valor is None or ausente(valor):
                    return None
                if isinstance(valor, datetime):
                    return valor.replace(tzinfo=None)
                if isinstance(valor, date):
                    return datetime.combine(valor, datetime.min.time())
                if isinstance(valor, (int, float)) and 1 <= float(valor) <= 80000:
                    return datetime(1899, 12, 30) + timedelta(days=float(valor))
                convertido = pd.to_datetime(valor, errors="coerce", dayfirst=True)
                if pd.isna(convertido):
                    return None
                return convertido.to_pydatetime().replace(tzinfo=None)

            for caminho in self.caminhos_selecionados:
                self.log(f"   📄 Lendo: {os.path.basename(caminho)}")
                try:
                    engine = "xlrd" if caminho.lower().endswith(".xls") else None
                    bruto = pd.read_excel(caminho, header=None, engine=engine)
                    if bruto.empty:
                        self.log(f"   ⚠️ Planilha vazia: {os.path.basename(caminho)}")
                        continue

                    tipo_layout, linha_cabecalho, colunas_mes = detectar_layout_excel(bruto)
                    if tipo_layout is None:
                        self.log(f"   ⚠️ Não reconheci o formato da planilha {os.path.basename(caminho)}.")
                        self.log("      Para transações, use cabeçalhos Data e Valor. Para contas anuais, use meses JAN a DEZ.")
                        continue

                    arquivos_lidos += 1
                    if tipo_layout == "mensal":
                        self.log(f"   ✅ Matriz de contas anuais detectada; cabeçalho na linha {linha_cabecalho + 1}.")
                        indices_meses = {indice for _, indice, _ in colunas_mes}
                        colunas_categoria = [i for i in range(bruto.shape[1]) if i not in indices_meses]
                        if not colunas_categoria:
                            self.log("   ⚠️ Não encontrei a coluna com os nomes das contas.")
                            continue
                        indice_categoria = colunas_categoria[0]

                        for _, linha in bruto.iloc[linha_cabecalho + 1:].iterrows():
                            nome_bruto = linha.iloc[indice_categoria]
                            if ausente(nome_bruto):
                                continue
                            categoria = str(nome_bruto).strip()
                            categoria_normalizada = normalizar_celula_excel(categoria)
                            if not categoria or categoria_normalizada.startswith("total"):
                                continue

                            valores_linha = {}
                            for numero_mes, indice_coluna, cabecalho_mes in colunas_mes:
                                valor = self.converter_para_numero(linha.iloc[indice_coluna])
                                valores_linha[numero_mes] = valores_linha.get(numero_mes, 0.0) + valor
                                nomes_meses[numero_mes] = str(cabecalho_mes).strip().upper()
                                totais_por_mes[numero_mes] = totais_por_mes.get(numero_mes, 0.0) + valor

                            soma_linha = sum(valores_linha.values())
                            if not any(abs(v) > 0 for v in valores_linha.values()):
                                continue
                            if categoria not in totais_mensais_categoria:
                                totais_mensais_categoria[categoria] = {}
                            for numero_mes, valor in valores_linha.items():
                                totais_mensais_categoria[categoria][numero_mes] = (
                                    totais_mensais_categoria[categoria].get(numero_mes, 0.0) + valor
                                )
                            total_matriz += soma_linha

                        self.log(f"   ✅ Contas mensais lidas de {os.path.basename(caminho)}")
                        continue

                    # Layout de transações: os nomes das colunas podem estar
                    # depois de uma ou mais linhas de título.
                    cabecalhos = bruto.iloc[linha_cabecalho].tolist()
                    dados = bruto.iloc[linha_cabecalho + 1:].copy()
                    mapa = {
                        "data": localizar_coluna_transacao(cabecalhos, "data"),
                        "descricao": localizar_coluna_transacao(cabecalhos, "descricao"),
                        "categoria": localizar_coluna_transacao(cabecalhos, "categoria"),
                        "valor": localizar_coluna_transacao(cabecalhos, "valor"),
                    }
                    if mapa["valor"] is None:
                        self.log(f"   ⚠️ Coluna de valor não encontrada em {os.path.basename(caminho)}.")
                        continue
                    if mapa["data"] is not None:
                        encontrou_data = True

                    quantidade = 0
                    for _, linha in dados.iterrows():
                        valor_bruto = linha.iloc[mapa["valor"]]
                        if ausente(valor_bruto) or str(valor_bruto).strip() == "":
                            continue
                        valor = self.converter_para_numero(valor_bruto)

                        data_bruta = linha.iloc[mapa["data"]] if mapa["data"] is not None else None
                        data_convertida = converter_data_excel(data_bruta)
                        descricao_bruta = linha.iloc[mapa["descricao"]] if mapa["descricao"] is not None else ""
                        categoria_bruta = linha.iloc[mapa["categoria"]] if mapa["categoria"] is not None else ""
                        descricao = "" if ausente(descricao_bruta) else str(descricao_bruta).strip()
                        categoria = "" if ausente(categoria_bruta) else str(categoria_bruta).strip()

                        registros.append({
                            "data_ordem": data_convertida,
                            "data_saida": data_convertida if data_convertida else ("" if ausente(data_bruta) else str(data_bruta).strip()),
                            "descricao": descricao,
                            "categoria": categoria,
                            "valor": valor,
                            "origem": os.path.basename(caminho),
                        })
                        total_transacoes += valor
                        quantidade += 1
                        if categoria:
                            totais_categoria_transacoes[categoria] = totais_categoria_transacoes.get(categoria, 0.0) + valor

                    self.log(f"   ✅ {quantidade} registro(s) de transação lido(s).")

                except Exception as erro_arquivo:
                    self.log(f"   ❌ Erro ao ler {os.path.basename(caminho)}: {erro_arquivo}")

            if not registros and not totais_mensais_categoria:
                raise ValueError(
                    "Nenhum dado foi reconhecido. Sua planilha pode ser uma matriz anual; "
                    "confira se existe uma linha com os meses JAN, FEV, MAR... DEZ e os nomes das contas na primeira coluna."
                )

            wb = Workbook()
            folha_inicial = wb.active
            folha_inicial.title = "Dados Consolidados"
            folha_inicial_foi_usada = False

            # Relatório de transações com ordenação de data decrescente.
            if registros:
                if encontrou_data:
                    registros.sort(
                        key=lambda item: (item["data_ordem"] is not None,
                                          item["data_ordem"] or datetime.min),
                        reverse=True
                    )
                    self.log("📅 Transações organizadas da mais nova para a mais antiga.")
                else:
                    self.log("⚠️ As planilhas de transações não possuem coluna Data; mantive a ordem original.")

                ws = folha_inicial
                folha_inicial_foi_usada = True
                cabecalho = ["Data", "Descrição", "Categoria", "Valor (€)", "Origem"]
                ws.append(cabecalho)
                for celula in ws[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    celula.alignment = Alignment(horizontal="center")

                for registro in registros:
                    ws.append([
                        registro["data_saida"], registro["descricao"], registro["categoria"],
                        round(registro["valor"], 2), registro["origem"]
                    ])
                    if isinstance(registro["data_saida"], datetime):
                        ws.cell(row=ws.max_row, column=1).number_format = "DD/MM/YYYY"
                    ws.cell(row=ws.max_row, column=4).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

                ultima_linha = ws.max_row
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = f"A1:E{ultima_linha}"
                linha_total = ws.max_row + 2
                ws.cell(row=linha_total, column=3, value="TOTAL GERAL").font = Font(bold=True)
                ws.cell(row=linha_total, column=4, value=round(total_transacoes, 2)).font = Font(bold=True)
                ws.cell(row=linha_total, column=4).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

                if totais_categoria_transacoes:
                    ws_resumo = wb.create_sheet("Resumo por Categoria")
                    ws_resumo.append(["Categoria", "Total (€)"])
                    for celula in ws_resumo[1]:
                        celula.font = Font(bold=True, color="FFFFFF")
                        celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                    for categoria, valor in sorted(totais_categoria_transacoes.items(), key=lambda item: item[1], reverse=True):
                        ws_resumo.append([categoria, round(valor, 2)])
                        ws_resumo.cell(row=ws_resumo.max_row, column=2).number_format = '€ #,##0.00;[Red]-€ #,##0.00'

            # Relatório de matriz anual (CONTAS DO ANO, categorias x meses).
            if totais_mensais_categoria:
                ordem_meses = sorted(totais_por_mes.keys())  # janeiro a dezembro
                nomes = [nomes_meses.get(numero, str(numero)) for numero in ordem_meses]
                if not registros and not folha_inicial_foi_usada:
                    ws_anual = folha_inicial
                    ws_anual.title = "Resumo Anual por Categoria"
                    folha_inicial_foi_usada = True
                else:
                    ws_anual = wb.create_sheet("Resumo Anual por Categoria")
                ws_anual.append(["Categoria"] + nomes + ["Total Anual", "% do Total"])
                for celula in ws_anual[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    celula.alignment = Alignment(horizontal="center")

                totais_categoria_matriz = {
                    categoria: sum(valores.values())
                    for categoria, valores in totais_mensais_categoria.items()
                }
                total_matriz = sum(totais_categoria_matriz.values())
                # A coluna Categoria começa com um código numérico (ex.: "1 CARTÃO").
                # Ordena pelo código numérico, não pelo total gasto.
                categorias_em_ordem = sorted(
                    totais_categoria_matriz.items(),
                    key=lambda item: chave_categoria_numerica(item[0])
                )
                for categoria, total_cat in categorias_em_ordem:
                    valores = [totais_mensais_categoria[categoria].get(mes, 0.0) for mes in ordem_meses]
                    percentual = (total_cat / total_matriz * 100) if total_matriz else 0.0
                    ws_anual.append([categoria] + [round(v, 2) for v in valores] + [round(total_cat, 2), round(percentual, 2)])

                linha_total_anual = ws_anual.max_row + 1
                ws_anual.append(["TOTAL GERAL"] + [round(totais_por_mes.get(mes, 0.0), 2) for mes in ordem_meses] + [round(total_matriz, 2), 100.0 if total_matriz else 0.0])
                for celula in ws_anual[linha_total_anual]:
                    celula.font = Font(bold=True)
                    celula.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

                # Exibe todos os valores financeiros em euros, mantendo-os numéricos.
                formato_euro = FORMATO_EURO_EXCEL
                formato_percentual = FORMATO_PERCENTUAL_EXCEL
                ultima_coluna_valor = len(ordem_meses) + 2  # meses + Total Anual
                for linha in ws_anual.iter_rows(min_row=2, max_row=ws_anual.max_row):
                    for celula in linha[1:ultima_coluna_valor]:
                        celula.number_format = formato_euro
                    linha[-1].number_format = formato_percentual

                ws_meses = wb.create_sheet("Resumo Mensal")
                ws_meses.append(["Mês", "Total do Mês (€)", "% do Total Geral"])
                for celula in ws_meses[1]:
                    celula.font = Font(bold=True, color="FFFFFF")
                    celula.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                for mes in ordem_meses:
                    valor_mes = totais_por_mes.get(mes, 0.0)
                    percentual = (valor_mes / total_matriz * 100) if total_matriz else 0.0
                    ws_meses.append([nomes_meses.get(mes, str(mes)), round(valor_mes, 2), round(percentual, 2)])
                linha_total_mes = ws_meses.max_row + 1
                ws_meses.append(["TOTAL GERAL", round(total_matriz, 2), 100.0 if total_matriz else 0.0])
                for celula in ws_meses[linha_total_mes]:
                    celula.font = Font(bold=True)
                    celula.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                # Na aba mensal, a coluna B é dinheiro e a coluna C é percentual.
                for linha in range(2, ws_meses.max_row + 1):
                    ws_meses.cell(row=linha, column=2).number_format = formato_euro
                    ws_meses.cell(row=linha, column=3).number_format = formato_percentual

            for ws_final in wb.worksheets:
                for coluna in ws_final.columns:
                    largura = max((len(str(celula.value)) for celula in coluna if celula.value is not None), default=0)
                    ws_final.column_dimensions[coluna[0].column_letter].width = min(max(largura + 3, 12), 45)
                ws_final.freeze_panes = "A2"

            # Evita que o LibreOffice mostre ### nas células com moeda.
            if totais_mensais_categoria:
                ws_anual.column_dimensions["A"].width = 22
                for indice_coluna in range(2, len(ordem_meses) + 2):
                    letra = ws_anual.cell(row=1, column=indice_coluna).column_letter
                    ws_anual.column_dimensions[letra].width = 15
                coluna_total_anual = len(ordem_meses) + 2
                coluna_percentual = len(ordem_meses) + 3
                ws_anual.column_dimensions[
                    ws_anual.cell(row=1, column=coluna_total_anual).column_letter
                ].width = 18
                ws_anual.column_dimensions[
                    ws_anual.cell(row=1, column=coluna_percentual).column_letter
                ].width = 16

                ws_meses.column_dimensions["A"].width = 16
                ws_meses.column_dimensions["B"].width = 20
                ws_meses.column_dimensions["C"].width = 20

            if totais_mensais_categoria:
                adicionar_graficos_resumo(
                    wb,
                    ws_meses,
                    ws_anual,
                    coluna_total_anual=len(ordem_meses) + 2,
                )

            wb.save(caminho_saida)
            self.caminho_ultimo_excel = os.path.abspath(caminho_saida)
            self.root.after(0, lambda: self.seguro_config(self.btn_abrir_excel, state=tk.NORMAL))

            self.log("\n" + "=" * 50)
            self.log("✅ CONSOLIDAÇÃO CONCLUÍDA")
            if registros:
                self.log(f"📄 Transações: {len(registros)}")
                self.log(f"💰 TOTAL DAS TRANSAÇÕES: € {total_transacoes:,.2f}")
            if totais_mensais_categoria:
                self.log(f"📊 Categorias anuais: {len(totais_mensais_categoria)}")
                self.log(f"💰 TOTAL DO ANO: € {total_matriz:,.2f}")
                self.log("📅 Meses: janeiro a dezembro.")
                self.log("🔢 Categorias: ordenadas pelo número inicial (ordem numérica).")
            self.log(f"💾 Arquivo salvo: {caminho_saida}")
            self.log("=" * 50 + "\n")
            salvar_no_historico("Processar Excel", caminho_saida, len(registros) + len(totais_mensais_categoria))

        except Exception as e:
            self.log(f"\n❌ ERRO AO PROCESSAR EXCEL: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.seguro_config(self.btn_executar, state=tk.NORMAL))
            self.root.after(0, lambda: self.seguro_config(self.btn_parar, state=tk.DISABLED))

    def abrir_arquivo(self, caminho):
        """Abre o arquivo salvo usando o aplicativo adequado em cada sistema."""
        if not caminho:
            messagebox.showinfo("Abrir arquivo", "Nenhum arquivo foi informado.")
            return False

        caminho = os.path.abspath(os.path.expanduser(os.fspath(caminho)))
        if not os.path.isfile(caminho):
            mensagem = f"O arquivo não foi encontrado:\n{caminho}"
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Arquivo não encontrado", mensagem)
            return False

        try:
            if os.name == "nt":
                # Usa o aplicativo associado ao XLSX no Windows.
                os.startfile(caminho)
            elif sys.platform == "darwin":
                subprocess.Popen(
                    ["open", caminho],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True,
                )
            else:
                # Em Linux, abre XLSX diretamente no LibreOffice Calc quando
                # disponível. Isso evita falhas de associação MIME do xdg-open.
                comando = None
                extensao = os.path.splitext(caminho)[1].lower()
                if extensao in (".xlsx", ".xls", ".ods", ".csv"):
                    for nome_programa in ("libreoffice", "soffice"):
                        executavel = shutil.which(nome_programa)
                        if executavel:
                            comando = [executavel, "--calc", caminho]
                            break

                # Fallback para o aplicativo padrão do ambiente Linux.
                if comando is None:
                    xdg_open = shutil.which("xdg-open")
                    gio = shutil.which("gio")
                    if xdg_open:
                        comando = [xdg_open, caminho]
                    elif gio:
                        comando = [gio, "open", caminho]
                    else:
                        raise FileNotFoundError(
                            "Não encontrei LibreOffice, xdg-open ou gio para abrir o arquivo."
                        )

                subprocess.Popen(
                    comando,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True,
                    close_fds=True,
                )

            self.log(f"📂 Abrindo arquivo: {caminho}")
            return True
        except Exception as e:
            mensagem = (
                f"Não foi possível abrir o arquivo automaticamente.\n\n"
                f"O arquivo continua salvo em:\n{caminho}\n\n"
                f"Detalhes: {e}"
            )
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir arquivo", mensagem)
            return False

    
    def abrir_arquivo_excel(self):
        """Abre a pasta onde o último resultado do Excel foi salvo."""
        caminho = getattr(self, "caminho_ultimo_excel", "")

        if not caminho:
            messagebox.showinfo(
                "Aviso",
                "Nenhum resultado foi gerado ainda.\nExecute o processamento primeiro."
            )
            return

        caminho = os.path.abspath(os.path.expanduser(caminho))

        if not os.path.isfile(caminho):
            messagebox.showerror(
                "Arquivo não encontrado",
                f"O arquivo não foi encontrado:\n{caminho}"
            )
            return

        pasta = os.path.dirname(caminho)

        try:
            if os.name == "nt":
                os.startfile(pasta)

            elif sys.platform == "darwin":
                subprocess.Popen(["open", pasta])

            else:
                xdg_open = shutil.which("xdg-open")
                gio = shutil.which("gio")

                xdg_open = shutil.which("xdg-open")
                gio = shutil.which("gio")
                
                if xdg_open:
                    subprocess.Popen([xdg_open, pasta])
                elif gio:
                    subprocess.Popen([gio, "open", pasta])
                else:
                    raise FileNotFoundError(
                        "Não encontrei xdg-open nem gio."
                    )
                    

            self.log(f"📂 Abrindo pasta do resultado: {pasta}")

        except Exception as e:
            mensagem = (
                f"Não foi possível abrir a pasta:\n{pasta}\n\n"
                f"Detalhes: {e}"
            )
            self.log(f"❌ {mensagem}")
            messagebox.showerror("Erro ao abrir pasta", mensagem)

    def mostrar_historico(self):
        janela = tk.Toplevel(self.root)
        janela.title("📜 Histórico de Ações")
        janela.geometry("600x400")

        texto = scrolledtext.ScrolledText(janela, wrap=tk.WORD, font=("Consolas", 10))
        texto.pack(fill=tk.BOTH, expand=True)

        arquivo_hist = caminho_arquivo_historico()
        if os.path.exists(arquivo_hist):
            try:
                with open(arquivo_hist, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                for item in reversed(dados):
                    texto.insert(tk.END, f"📅 {item.get('data', '')}\n")
                    texto.insert(tk.END, f"   Ação: {item.get('tipo', '')}\n")
                    texto.insert(tk.END, f"   Detalhe: {item.get('detalhe', '')}\n")
                    texto.insert(tk.END, f"   Registros: {item.get('quantidade', 0)}\n")
                    texto.insert(tk.END, "-" * 50 + "\n")
            except:
                texto.insert(tk.END, "Erro ao ler histórico.\n")
        else:
            texto.insert(tk.END, "Nenhum histórico ainda.\n")

        texto.config(state=tk.DISABLED)


# ==========================================
# ▶ EXECUTAR O PROGRAMA
# ==========================================

if __name__ == "__main__":
    raiz = tk.Tk()
    app = Aplicacao(raiz)
    raiz.mainloop()