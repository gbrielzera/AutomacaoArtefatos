import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import copy
from pygments import lex
from pygments.lexers import PythonLexer
from pygments.token import Token
import os
import re
import sys
import logging
from datetime import datetime


def get_base_dir():
    """Retorna a pasta do executável (.exe) ou do script .py, para localizar a pasta de logs."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def limpar_nome_arquivo(nome):
    """Remove caracteres que o Windows não aceita em nomes de arquivo."""
    for c in '<>:"/\\|?*':
        nome = nome.replace(c, "_")
    return nome.strip() or "Supravizio"


class BufferLogHandler(logging.Handler):
    """
    Acumula as mensagens em memória para que cada artefato gerado receba o seu próprio
    arquivo de log, contendo tudo desde a geração anterior (inclusive a seleção do XML).
    """

    def __init__(self):
        super().__init__()
        self.linhas = []

    def emit(self, record):
        self.linhas.append(self.format(record))

    def drenar(self):
        linhas, self.linhas = self.linhas, []
        return linhas


def setup_logger():
    """Configura o logger da aplicação com um buffer em memória."""
    logger = logging.getLogger("SupravizioDocApp")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    handler = BufferLogHandler()
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)

    return logger, handler

class SupravizioDocApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Artefato Supravizio")
        self.root.geometry("650x750")
        self.root.resizable(True, True)

        self.xml_path = ""
        self.template_path = ""
        self.output_dir = ""

        self.logger, self.log_handler = setup_logger()

        # Guarda o que foi preenchido automaticamente, para poder substituir esses valores
        # ao trocar de XML sem sobrescrever o que o usuário digitou à mão.
        self.macro_auto = ""
        self.proc_auto = ""

        # =========================
        # TÍTULO E INPUTS
        # =========================
        tk.Label(root, text="Automação de Artefatos - Supravizio", font=("Arial", 14, "bold")).pack(pady=10)

        frame_inputs = tk.Frame(root)
        frame_inputs.pack(pady=5, fill="x", padx=20)

        tk.Label(frame_inputs, text="Macroprocesso:", anchor="w").pack(fill="x")
        self.entry_macro = tk.Entry(frame_inputs)
        self.entry_macro.pack(pady=2, fill="x")

        tk.Label(frame_inputs, text="Processo:", anchor="w").pack(fill="x")
        self.entry_proc = tk.Entry(frame_inputs)
        self.entry_proc.pack(pady=2, fill="x")

        self.lbl_deteccao = tk.Label(
            frame_inputs, text="", fg="gray", anchor="w", font=("Arial", 8, "italic"), wraplength=580, justify="left"
        )
        self.lbl_deteccao.pack(fill="x", pady=(0, 4))

        tk.Label(frame_inputs, text="Descrição da Alteração/Criação:", anchor="w").pack(fill="x")
        self.text_descricao = tk.Text(frame_inputs, height=4, width=50)
        self.text_descricao.pack(pady=2, fill="x")

        tk.Label(frame_inputs, text="Evidências em Homologação (Nº Chamado):", anchor="w").pack(fill="x")
        self.entry_evidencia = tk.Entry(frame_inputs)
        self.entry_evidencia.pack(pady=2, fill="x")

        # =========================
        # SELETORES DE DIRETÓRIOS E ARQUIVOS
        # =========================
        frame_files = tk.Frame(root)
        frame_files.pack(pady=10, fill="x", padx=20)

        self.btn_xml = tk.Button(frame_files, text="1. Selecionar XML do Fluxo", command=self.load_xml)
        self.btn_xml.pack(fill="x", pady=2)
        self.lbl_xml = tk.Label(frame_files, text="Nenhum arquivo XML selecionado", fg="gray", anchor="w")
        self.lbl_xml.pack(fill="x", pady=(0, 10))

        self.btn_template = tk.Button(frame_files, text="2. Selecionar Template DOCX", command=self.load_template)
        self.btn_template.pack(fill="x", pady=2)
        self.lbl_template = tk.Label(frame_files, text="Nenhum template DOCX selecionado", fg="gray", anchor="w")
        self.lbl_template.pack(fill="x", pady=(0, 10))

        self.btn_dir = tk.Button(frame_files, text="3. Selecionar Pasta de Destino", command=self.load_dir)
        self.btn_dir.pack(fill="x", pady=2)
        self.lbl_dir = tk.Label(frame_files, text="Nenhuma pasta selecionada", fg="gray", anchor="w")
        self.lbl_dir.pack(fill="x")

        self.btn_gerar = tk.Button(
            root, text="Gerar Artefato DOCX", font=("Arial", 12, "bold"), 
            bg="#4CAF50", fg="white", command=self.gerar_documento
        )
        self.btn_gerar.pack(pady=15, fill="x", padx=50)

    # ==========================================================
    # LOG
    # ==========================================================
    def gravar_log(self, nome_fluxo=""):
        """Grava um arquivo de log para a geração recém-concluída e esvazia o buffer."""
        linhas = self.log_handler.drenar()
        if not linhas:
            return ""
        try:
            logs_dir = os.path.join(get_base_dir(), "logs")
            os.makedirs(logs_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            sufixo = limpar_nome_arquivo(nome_fluxo) if nome_fluxo else "Supravizio"
            log_path = os.path.join(logs_dir, f"log_{sufixo}_{timestamp}.log")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("\n".join(linhas) + "\n")
            return log_path
        except Exception as e:
            print(f"Não foi possível gravar o arquivo de log: {e}")
            return ""

    # ==========================================================
    # FUNÇÕES DE INTERFACE
    # ==========================================================
    def load_xml(self):
        self.xml_path = filedialog.askopenfilename(title="Selecionar XML", filetypes=[("XML", "*.xml")])
        if self.xml_path:
            self.lbl_xml.config(text=os.path.basename(self.xml_path), fg="black")
            self.logger.info(f"XML selecionado: {self.xml_path}")
            self.tentar_preencher_macro_processo()

    # ==========================================================
    # DETECÇÃO AUTOMÁTICA DE PROCESSO / SUBPROCESSO
    # ==========================================================
    def extrair_processo_do_nome_arquivo(self, caminho_xml):
        """
        O Supravizio exporta o XML com o padrão:
            "{Processo}_Versão_{N}_{NomeSubProcesso}"
        O Processo pode conter hífens (ex.: "Financeiro - Serviços Gerais"), por isso
        tudo que vem antes de "_Versão_{N}_" é considerado o Processo inteiro.
        O Macroprocesso NÃO consta nem no nome do arquivo nem no XML.
        Retorna (processo, subprocesso_do_nome); cada item pode ser None.
        """
        nome_base = os.path.splitext(os.path.basename(caminho_xml))[0]
        nome_normalizado = re.sub(r"\s+", " ", nome_base.replace("_", " ")).strip()

        match = re.match(
            r"^(?P<processo>.+?)\s+Vers[aã]o\s+\d+\s*(?P<sub>.*)$",
            nome_normalizado,
            flags=re.IGNORECASE
        )
        if not match:
            return None, None

        processo = match.group("processo").strip(" -")
        sub = match.group("sub").strip()
        return (processo or None), (sub or None)

    def extrair_macroprocesso_do_xml(self, root):
        """
        Procura o Macroprocesso na árvore do XML. O Supravizio não exporta o nome do
        macroprocesso no XML de fluxo (apenas a classe vazia
        'Venki.Supravizio.Processo.MacroProcesso'), então na prática isto quase sempre
        retorna None e o campo precisa ser preenchido manualmente.
        """
        for tag in ("MacroProcesso", "Macroprocesso", "NomeMacroProcesso", "ClasseMacroProcesso"):
            for node in root.iter(tag):
                if node.text and node.text.strip():
                    return node.text.strip()
                desc = node.find("Descricao")
                if desc is not None and desc.text and desc.text.strip():
                    return desc.text.strip()
        return None

    def aplicar_valor_automatico(self, entry, novo_valor, valor_auto_anterior, rotulo):
        """
        Atualiza um campo preenchido automaticamente ao trocar de XML.
        Substitui o valor quando o campo está vazio ou ainda contém o que foi detectado
        para o XML anterior; se o usuário digitou algo à mão, o texto dele é preservado.
        Retorna o novo valor automático a ser lembrado.
        """
        atual = entry.get().strip()
        if atual and atual != valor_auto_anterior:
            self.logger.info(
                f"Campo '{rotulo}' foi editado manualmente ('{atual}'); "
                f"o valor detectado ('{novo_valor}') não foi aplicado."
            )
            return valor_auto_anterior

        entry.delete(0, tk.END)
        if novo_valor:
            entry.insert(0, novo_valor)
        return novo_valor

    def tentar_preencher_macro_processo(self):
        processo, sub_do_nome = self.extrair_processo_do_nome_arquivo(self.xml_path)

        macro = None
        try:
            macro = self.extrair_macroprocesso_do_xml(self.ler_xml_root())
        except Exception as e:
            self.logger.warning(f"Não foi possível ler o XML para buscar o Macroprocesso: {e}")

        avisos = []

        if processo:
            self.logger.info(f"Processo detectado pelo nome do arquivo: '{processo}'")
            self.proc_auto = self.aplicar_valor_automatico(self.entry_proc, processo, self.proc_auto, "Processo")
        else:
            self.logger.warning(
                f"Não foi possível detectar o Processo pelo nome do arquivo "
                f"'{os.path.basename(self.xml_path)}' (padrão esperado: "
                f"'Processo_Versão_N_Subprocesso'). Preencha manualmente."
            )
            self.proc_auto = self.aplicar_valor_automatico(self.entry_proc, "", self.proc_auto, "Processo")
            avisos.append("Processo")

        if sub_do_nome:
            self.logger.info(f"Subprocesso indicado pelo nome do arquivo: '{sub_do_nome}'")

        if macro:
            self.logger.info(f"Macroprocesso encontrado no XML: '{macro}'")
            self.macro_auto = self.aplicar_valor_automatico(self.entry_macro, macro, self.macro_auto, "Macroprocesso")
        else:
            self.logger.warning(
                "Macroprocesso não consta no XML exportado pelo Supravizio nem no nome do "
                "arquivo. Preencha manualmente."
            )
            self.macro_auto = self.aplicar_valor_automatico(self.entry_macro, "", self.macro_auto, "Macroprocesso")
            avisos.append("Macroprocesso")

        if avisos:
            self.lbl_deteccao.config(
                text="Preencha manualmente: " + ", ".join(avisos)
                     + ". O Subprocesso é lido direto do XML.",
                fg="#B71C1C"
            )
        else:
            self.lbl_deteccao.config(
                text="Macroprocesso e Processo preenchidos automaticamente (confira antes de gerar).",
                fg="#2E7D32"
            )

    def load_template(self):
        self.template_path = filedialog.askopenfilename(title="Selecionar DOCX", filetypes=[("Word", "*.docx")])
        if self.template_path: self.lbl_template.config(text=os.path.basename(self.template_path), fg="black")

    def load_dir(self):
        self.output_dir = filedialog.askdirectory(title="Selecionar Destino")
        if self.output_dir: self.lbl_dir.config(text=self.output_dir, fg="black")

    # ==========================================================
    # LÓGICA DE EXTRAÇÃO DE DADOS
    # ==========================================================
    def get_texto(self, node, tag, default=""):
        child = node.find(tag)
        if child is not None and child.text:
            return child.text.strip()
        return default

    def extrair_propriedades_campos(self, root):
        propriedades = {}

        for prop in root.findall(".//CustomProperty"):
            nome = self.get_texto(prop, "Name")
            tabela = self.get_texto(prop, "TableName")

            if not nome or not tabela:
                continue

            if nome not in propriedades:
                propriedades[nome] = {
                    "rotulo": self.get_texto(prop, "Text") or self.get_texto(prop, "Description") or nome,
                    "descricao": self.get_texto(prop, "Description") or self.get_texto(prop, "Text") or nome,
                    "tipo": self.get_texto(prop, "Type", "String"),
                    "tabela": tabela,
                    "coluna": self.get_texto(prop, "TableColumn"),
                }

        return propriedades

    def ler_xml_root(self):
        try:
            return ET.parse(self.xml_path).getroot()
        except ET.ParseError:
            with open(self.xml_path, 'r', encoding='utf-8-sig') as file:
                xml_content = file.read().strip()
            if not xml_content:
                raise Exception("O arquivo XML selecionado está vazio.")
            return ET.fromstring(xml_content)

    def extrair_dados_xml(self):
        self.logger.info(f"Iniciando extração de dados do XML: {self.xml_path}")
        root = self.ler_xml_root()

        dados = {"nome_fluxo": "", "servicos": [], "campos": [], "anexos": [], "scripts": []}
        propriedades_campos = self.extrair_propriedades_campos(root)

        node_nome = root.find(".//NomeSubProcesso")
        if node_nome is not None and node_nome.text:
            dados["nome_fluxo"] = node_nome.text.strip()
        else:
            self.logger.warning("Não foi possível localizar a tag <NomeSubProcesso> no XML.")

        # Dedup por (local, código): evita duplicar o mesmo script quando a MESMA atividade
        # aparece repetida no diagrama (ex: um LinkInicial compartilhado entre várias páginas),
        # mas preserva scripts com código idêntico que estejam em locais distintos (ex: o mesmo
        # preenchimento de campo replicado no Evento Inicial e no Link Inicial).
        scripts_vistos = set()

        def registrar_script(local, sc_code):
            chave = (local, sc_code)
            if chave in scripts_vistos:
                self.logger.debug(f"Script duplicado ignorado (mesma atividade repetida no diagrama): {local}")
                return
            scripts_vistos.add(chave)
            dados["scripts"].append({"local": local, "codigo": sc_code})

        servicos_vistos = set()

        def coletar_servicos(restricoes):
            """Lê pares (tipo, serviço) de nós <RestricaoServico>, ignorando entradas vazias."""
            for rest in restricoes:
                srv = rest.find("Servico")
                tipo_txt = (
                    self.get_texto(rest, "ClasseServico/Descricao")
                    or (self.get_texto(srv, "ClasseServico/Descricao") if srv is not None else "")
                )
                nome_txt = self.get_texto(srv, "Descricao") if srv is not None else ""

                if not tipo_txt and not nome_txt:
                    self.logger.warning(
                        "Encontrada uma <RestricaoServico> sem Tipo e sem Descrição do serviço; "
                        "entrada ignorada (o serviço não veio expandido no XML)."
                    )
                    continue

                chave = f"{tipo_txt} - {nome_txt}"
                if chave not in servicos_vistos:
                    servicos_vistos.add(chave)
                    dados["servicos"].append({"tipo": tipo_txt, "nome": nome_txt})

        # Fonte autoritativa: os serviços do PRÓPRIO subprocesso (ClasseSubProcesso).
        # Isto evita capturar serviços de subprocessos secundários apenas engatilhados,
        # e funciona mesmo em fluxos sem LinkInicial.
        coletar_servicos(root.findall(".//SubProcesso/ClasseSubProcesso/RestricoesServicos/RestricaoServico"))

        if not dados["servicos"]:
            self.logger.warning(
                "Nenhum serviço em <SubProcesso/ClasseSubProcesso>; tentando o LinkInicial como alternativa."
            )
            for link_inicial in root.findall(".//Atividade[Tipo='LinkInicial']"):
                for alvo in link_inicial.findall(".//ClasseAlvo"):
                    coletar_servicos(alvo.findall(".//RestricoesServicos/RestricaoServico"))

        mapa_scripts = {
            "ScriptModificado": "modificado",
            "ScriptValidacao": "de validação",
            "ScriptFormCarregado": "de formulário carregado",
            "ScriptInicio": "de início",
            "ScriptFim": "de fim",
            "ScriptVolta": "de volta",
            "ScriptEvento": "de evento",
        }

        # Scripts que ficam aninhados sob outro nó da atividade, e não como filho direto.
        mapa_scripts_aninhados = {
            "PapelResponsavel/PapelClasseNegocio/ScriptSelecaoAtores": "de seleção de atores (papel responsável)",
            "PapelDestinatario/PapelClasseNegocio/ScriptSelecaoAtores": "de seleção de atores (papel destinatário)",
        }

        campos_vistos = set()
        anexos_vistos = set()

        for figura in root.findall(".//Figura"):
            atividade = figura.find("Atividade")

            if atividade is not None:
                tag_tipo_atv = atividade.find("Tipo")
                tipo_atv = tag_tipo_atv.text.strip() if (tag_tipo_atv is not None and tag_tipo_atv.text) else "Atividade"

                # Nome real da tarefa (ex.: "Anexar comprovante de pagamento"), quando existir.
                # Tipos sem rótulo próprio (LinkInicial, Gateway, etc.) caem no fallback do Tipo.
                tag_desc_atv = atividade.find("Descricao")
                descricao_atv = tag_desc_atv.text.strip() if (tag_desc_atv is not None and tag_desc_atv.text) else ""
                if descricao_atv and descricao_atv != tipo_atv:
                    nome_atv = f"{descricao_atv} ({tipo_atv})"
                else:
                    nome_atv = tipo_atv

                # As operações do LinkInicial (evento de inicialização do fluxo) ficam em
                # <Figura><OperacaoAtividade>, IRMÃO de <Atividade> e não dentro dela.
                # Ler só atividade//OperacaoAtividade perdia esses scripts e campos.
                operacoes = atividade.findall(".//OperacaoAtividade") + figura.findall("OperacaoAtividade")

                for op in operacoes:
                    for campo in op.findall(".//CampoPreenchimento"):
                        nome_custom = campo.find("./NomeCustomizado")
                        rotulo = campo.find("./Rotulo")
                        
                        # Extrai a Tabela da Tag 'FormaEdicaoWeb' no nível do Campo (Corrigido)
                        tabela_nome = "Formulário"
                        form_edicao = campo.find("./FormaEdicaoWeb")
                        if form_edicao is not None and form_edicao.text:
                            if form_edicao.text.strip().lower() != "formulario":
                                tabela_nome = form_edicao.text.strip()
                        
                        if nome_custom is not None and nome_custom.text:
                            nome_c = nome_custom.text.strip()
                            prop_campo = propriedades_campos.get(nome_c, {})
                            rot_txt = rotulo.text.strip() if (rotulo is not None and rotulo.text) else prop_campo.get("rotulo", nome_c)
                            desc_txt = prop_campo.get("descricao", rot_txt)
                            tipo_txt = prop_campo.get("tipo", "String")
                            tabela_nome = prop_campo.get("tabela", tabela_nome)
                            
                            if nome_c and nome_c not in campos_vistos:
                                dados["campos"].append({"nome": nome_c, "rotulo": rot_txt, "descricao": desc_txt, "tipo": tipo_txt, "tabela": tabela_nome})
                                campos_vistos.add(nome_c)
                                
                            for tag_xml, nome_amigavel in mapa_scripts.items():
                                s_node = campo.find(tag_xml)
                                if s_node is not None and s_node.text and s_node.text.strip():
                                    sc_code = s_node.text.strip()
                                    registrar_script(
                                        f"Atividade: {nome_atv} | Script {nome_amigavel} no campo: {nome_c}",
                                        sc_code
                                    )

                    for tag_xml, nome_amigavel in mapa_scripts.items():
                        s_node = op.find(tag_xml)
                        if s_node is not None and s_node.text and s_node.text.strip():
                            sc_code = s_node.text.strip()
                            registrar_script(
                                f"Atividade: {nome_atv} | Script {nome_amigavel} da operação",
                                sc_code
                            )

                for tag_xml, nome_amigavel in list(mapa_scripts.items()) + list(mapa_scripts_aninhados.items()):
                    s_node = atividade.find(tag_xml)
                    if s_node is not None and s_node.text and s_node.text.strip():
                        sc_code = s_node.text.strip()
                        registrar_script(
                            f"Atividade: {nome_atv} | Script {nome_amigavel} da atividade",
                            sc_code
                        )

                for vi in atividade.findall(".//ValoresInputs/ValorInput"):
                    expr_in = vi.find("ExpressaoValor")
                    if expr_in is not None and expr_in.text and expr_in.text.strip():
                        nome_input = self.get_texto(vi, "CustomProperty/Name") or self.get_texto(vi, "Nome") or "input"
                        registrar_script(
                            f"Atividade: {nome_atv} | Expressão de valor do input: {nome_input}",
                            expr_in.text.strip()
                        )

                for escopo in atividade.findall(".//EscopoClasseAnexo/ClasseConfiguracao"):
                    desc = escopo.find("./Descricao")
                    sigla = escopo.find("./Sigla")
                    if desc is not None and sigla is not None:
                        d_txt = desc.text.strip() if desc.text else ""
                        s_txt = sigla.text.strip() if sigla.text else ""
                        if d_txt and s_txt and s_txt not in anexos_vistos:
                            dados["anexos"].append({"nome": d_txt, "sigla": s_txt})
                            anexos_vistos.add(s_txt)
                            
            gateway = figura.find("Gateway")
            if gateway is not None:
                tag_desc_gw = gateway.find("Descricao")
                desc_gw = tag_desc_gw.text.strip() if (tag_desc_gw is not None and tag_desc_gw.text) else "Gateway"
                expr = gateway.find("ExpressaoComparacaoDecision")
                
                if expr is not None and expr.text and expr.text.strip():
                    sc_code = expr.text.strip()
                    registrar_script(
                        f"Gateway de Decisão: {desc_gw} | Expressão de Decisão",
                        sc_code
                    )

        self.logger.info(
            f"Extração concluída. nome_fluxo='{dados['nome_fluxo']}' | "
            f"servicos={len(dados['servicos'])} | campos={len(dados['campos'])} | "
            f"anexos={len(dados['anexos'])} | scripts={len(dados['scripts'])}"
        )
        if not dados["servicos"]:
            self.logger.warning("Nenhum serviço associado foi encontrado no Link Inicial do fluxo.")
        if not dados["campos"]:
            self.logger.warning("Nenhum campo de preenchimento foi encontrado no fluxo.")
        if not dados["scripts"]:
            self.logger.warning("Nenhum script (IronPython) ou expressão de Gateway foi encontrado no fluxo.")

        return dados

    def add_code_block(self, paragraph, code_text):
        """Cria um bloco de código com aparência de IDE, incluindo sintaxe colorida."""

        # ------------------------------------------------------------------
        # Aparência do bloco
        # ------------------------------------------------------------------
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.left_indent = Cm(0.15)
        paragraph.paragraph_format.right_indent = Cm(0.15)

        pPr = paragraph._p.get_or_add_pPr()

        # Fundo cinza nativo XML
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F4F4F4') 
        pPr.append(shd)

        # Borda discreta ao redor do código
        pBdr = OxmlElement('w:pBdr')
        for side in ["top", "left", "bottom", "right"]:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '4')
            border.set(qn('w:color'), 'CCCCCC')
            pBdr.append(border)
        pPr.append(pBdr)

        # ------------------------------------------------------------------
        # Highlight de sintaxe
        # ------------------------------------------------------------------
        cores = {
            Token.Keyword: "0000FF",
            Token.Keyword.Constant: "0000FF",
            Token.Keyword.Namespace: "0000FF",
            Token.Name.Builtin: "267F99",
            Token.Name.Function: "795E26",
            Token.Name.Class: "267F99",
            Token.Name.Decorator: "795E26",
            Token.String: "A31515",
            Token.String.Doc: "A31515",
            Token.Number: "098658",
            Token.Comment: "008000",
            Token.Operator: "000000",
            Token.Punctuation: "000000",
            Token.Name: "000000",
            Token.Text: "000000",
        }

        linhas_do_codigo = code_text.split('\n')
        
        for k, linha_bruta in enumerate(linhas_do_codigo):
            if not linha_bruta.strip():
                if k < len(linhas_do_codigo) - 1:
                    paragraph.add_run().add_break()
                continue
                
            espacos_iniciais = len(linha_bruta) - len(linha_bruta.lstrip(' '))
            if espacos_iniciais > 0:
                 run_espaco = paragraph.add_run(" " * espacos_iniciais)
                 run_espaco.font.name = "Consolas"
                 run_espaco.font.size = Pt(9)
                 
            for token_type, token_text in lex(linha_bruta.lstrip(' '), PythonLexer()):
                cor = "000000"
                tipo = token_type
                while tipo is not Token:
                    if tipo in cores:
                        cor = cores[tipo]
                        break
                    tipo = tipo.parent

                run = paragraph.add_run(token_text)
                run.font.name = "Consolas"
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor.from_string(cor)

                # Força compatibilidade da fonte Consolas
                rPr = run._r.get_or_add_rPr()
                rFonts = rPr.rFonts
                if rFonts is None:
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                rFonts.set(qn("w:ascii"), "Consolas")
                rFonts.set(qn("w:hAnsi"), "Consolas")
                rFonts.set(qn("w:cs"), "Consolas")

            if k < len(linhas_do_codigo) - 1:
                paragraph.add_run().add_break()

    def inserir_paragrafo_depois(self, paragraph):
        """Cria um parágrafo vazio logo após `paragraph`, herdando sua formatação."""
        novo_p = OxmlElement("w:p")
        pPr_origem = paragraph._p.find(qn("w:pPr"))
        if pPr_origem is not None:
            novo_p.append(copy.deepcopy(pPr_origem))
        paragraph._p.addnext(novo_p)
        return Paragraph(novo_p, paragraph._parent)

    def localizar_placeholder(self, doc, idx_titulo):
        """
        Devolve o índice do parágrafo vazio que o template deixa logo abaixo de um título,
        para que o conteúdo seja escrito NELE em vez de numa linha nova depois dele
        (o que deixava uma linha em branco entre o título e o texto).
        """
        if idx_titulo == -1:
            return -1
        idx = idx_titulo + 1
        if idx < len(doc.paragraphs) and not doc.paragraphs[idx].text.strip():
            # O placeholder pode conter runs vazios com quebras de linha, que empurrariam
            # o texto para baixo do título; remove-os mantendo a formatação do parágrafo.
            for run in list(doc.paragraphs[idx].runs):
                run._r.getparent().remove(run._r)
            return idx
        return -1

    def gerar_documento(self):
        if not self.xml_path or not self.template_path or not self.output_dir:
            messagebox.showerror("Erro", "Preencha todos os caminhos (XML, Template e Pasta).")
            return

        self.logger.info("=== Iniciando geração do artefato ===")
        self.logger.info(f"XML: {self.xml_path}")
        self.logger.info(f"Template: {self.template_path}")
        self.logger.info(f"Pasta de destino: {self.output_dir}")

        nome_fluxo = ""
        try:
            dados = self.extrair_dados_xml()
            nome_fluxo = dados["nome_fluxo"]

            macro = self.entry_macro.get().strip()
            proc = self.entry_proc.get().strip()
            desc = self.text_descricao.get("1.0", tk.END).strip()
            evid = self.entry_evidencia.get().strip()

            if not macro:
                self.logger.warning("Campo 'Macroprocesso' não foi preenchido (nem manual, nem automaticamente).")
            if not proc:
                self.logger.warning("Campo 'Processo' não foi preenchido (nem manual, nem automaticamente).")
            if not evid:
                self.logger.warning("Campo 'Evidências em Homologação' não foi preenchido.")

            doc = Document(self.template_path)

            idx_titulo_desc = -1
            idx_titulo_servico = -1
            idx_evid = -1

            for i, p in enumerate(doc.paragraphs):
                texto = p.text.strip()
                if texto.startswith("Macroprocesso:"):
                    p.text = ""
                    p.add_run("Macroprocesso: ").bold = True
                    p.add_run(macro)
                elif texto.startswith("Processo:"):
                    p.text = ""
                    p.add_run("Processo: ").bold = True
                    p.add_run(proc)
                elif texto.startswith("Sub-processos:"):
                    p.text = ""
                    p.add_run("Sub-processos: ").bold = True
                    p.add_run(dados["nome_fluxo"] if dados["nome_fluxo"] else "")
                elif texto == "Descrição":
                    idx_titulo_desc = i
                elif texto == "Serviço":
                    idx_titulo_servico = i
                elif texto.startswith("Evidências em Homologação:"):
                    idx_evid = i
                    p.text = ""
                    p.add_run("Evidências em Homologação: ").bold = True
                    p.add_run(evid)

            idx_desc = self.localizar_placeholder(doc, idx_titulo_desc)
            idx_servico = self.localizar_placeholder(doc, idx_titulo_servico)

            if idx_desc == -1:
                self.logger.warning("Parágrafo em branco após o título 'Descrição' não encontrado; descrição não inserida.")
            if idx_servico == -1:
                self.logger.warning("Parágrafo em branco após o título 'Serviço' não encontrado; serviços não inseridos.")
            if idx_evid == -1:
                self.logger.warning("Âncora 'Evidências em Homologação:' não encontrada no template DOCX; scripts não inseridos.")
            if len(doc.tables) == 0:
                self.logger.warning("Template DOCX não possui tabelas; tabela de Campos não preenchida.")
            if len(doc.tables) <= 2:
                self.logger.warning("Template DOCX não possui a 3ª tabela esperada; tabela de Anexos não preenchida.")

            if idx_evid != -1 and dados["scripts"]:
                estilo_ancora = doc.paragraphs[idx_evid].style
                for s_dict in reversed(dados["scripts"]):
                    doc.paragraphs[idx_evid].insert_paragraph_before("", style=estilo_ancora)
                    
                    p_code = doc.paragraphs[idx_evid].insert_paragraph_before("")
                    self.add_code_block(p_code, s_dict["codigo"])
                    
                    p_local = doc.paragraphs[idx_evid].insert_paragraph_before("", style=estilo_ancora)
                    p_local.add_run(s_dict["local"]).bold = True

            # Escreve no parágrafo-placeholder do template (preservando o recuo dele) e só
            # cria linhas extras para os itens seguintes, para o texto colar logo abaixo do título.
            if idx_servico != -1 and dados["servicos"]:
                p_atual = doc.paragraphs[idx_servico]
                p_atual.add_run(f"- Tipo: {dados['servicos'][0]['tipo']} | Serviço: {dados['servicos'][0]['nome']}")
                for srv in dados["servicos"][1:]:
                    p_atual = self.inserir_paragrafo_depois(p_atual)
                    p_atual.add_run(f"- Tipo: {srv['tipo']} | Serviço: {srv['nome']}")

            if idx_desc != -1 and desc:
                linhas_desc = desc.split("\n")
                p_atual = doc.paragraphs[idx_desc]
                p_atual.add_run(linhas_desc[0])
                for linha in linhas_desc[1:]:
                    p_atual = self.inserir_paragrafo_depois(p_atual)
                    p_atual.add_run(linha)

            if len(doc.tables) > 0:
                t_campos = doc.tables[0]
                for row in t_campos.rows[1:]: t_campos._element.remove(row._tr)
                for c in dados["campos"]:
                    row = t_campos.add_row()
                    row.cells[0].text = c["nome"]
                    row.cells[1].text = c.get("descricao", c["rotulo"])
                    row.cells[2].text = c.get("tipo", "String")
                    row.cells[3].text = c["tabela"]
                    row.cells[4].text = "Sim"

            if len(doc.tables) > 2:
                t_anexos = doc.tables[2]
                for row in t_anexos.rows[1:]: t_anexos._element.remove(row._tr)
                for a in dados["anexos"]:
                    row = t_anexos.add_row()
                    row.cells[0].text = a["nome"]
                    row.cells[1].text = a["sigla"]
                    row.cells[2].text = "Sim"

            nome_arquivo = limpar_nome_arquivo(dados["nome_fluxo"])
            caminho_completo = os.path.join(self.output_dir, f"Artefato_{nome_arquivo}.docx")

            doc.save(caminho_completo)
            self.logger.info(f"Artefato salvo com sucesso em: {caminho_completo}")
            self.logger.info("=== Geração concluída ===")

            log_path = self.gravar_log(nome_fluxo)
            messagebox.showinfo(
                "Sucesso",
                f"Artefato Injetado com formatação visual de código!\n\nSalvo em:\n{caminho_completo}"
                f"\n\nLog desta execução:\n{log_path}"
            )

        except Exception as e:
            self.logger.exception(f"Falha na geração do artefato: {e}")
            log_path = self.gravar_log(nome_fluxo)
            messagebox.showerror(
                "Erro",
                f"Falha na manipulação do DOCX:\n{e}\n\nDetalhes em:\n{log_path}"
            )

if __name__ == "__main__":
    root = tk.Tk()
    app = SupravizioDocApp(root)
    root.mainloop()
