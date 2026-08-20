import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pygments import lex
from pygments.lexers import PythonLexer
from pygments.token import Token
import os

class SupravizioDocApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Artefato Supravizio")
        self.root.geometry("650x750")
        self.root.resizable(True, True)

        self.xml_path = ""
        self.template_path = ""
        self.output_dir = ""

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
    # FUNÇÕES DE INTERFACE
    # ==========================================================
    def load_xml(self):
        self.xml_path = filedialog.askopenfilename(title="Selecionar XML", filetypes=[("XML", "*.xml")])
        if self.xml_path: self.lbl_xml.config(text=os.path.basename(self.xml_path), fg="black")

    def load_template(self):
        self.template_path = filedialog.askopenfilename(title="Selecionar DOCX", filetypes=[("Word", "*.docx")])
        if self.template_path: self.lbl_template.config(text=os.path.basename(self.template_path), fg="black")

    def load_dir(self):
        self.output_dir = filedialog.askdirectory(title="Selecionar Destino")
        if self.output_dir: self.lbl_dir.config(text=self.output_dir, fg="black")

    # ==========================================================
    # LÓGICA DE EXTRAÇÃO DE DADOS
    # ==========================================================
    def extrair_dados_xml(self):
        try:
            tree = ET.parse(self.xml_path)
            root = tree.getroot()
        except ET.ParseError:
            with open(self.xml_path, 'r', encoding='utf-8-sig') as file:
                xml_content = file.read().strip()
            if not xml_content:
                raise Exception("O arquivo XML selecionado está vazio.")
            root = ET.fromstring(xml_content)

        dados = {"nome_fluxo": "", "servicos": [], "campos": [], "anexos": [], "scripts": []}

        node_nome = root.find(".//NomeSubProcesso")
        if node_nome is not None and node_nome.text:
            dados["nome_fluxo"] = node_nome.text.strip()

        servicos_vistos = set()
        for link_inicial in root.findall(".//Atividade[Tipo='LinkInicial']"):
            for alvo in link_inicial.findall(".//ClasseAlvo"):
                for srv in alvo.findall(".//RestricoesServicos/RestricaoServico/Servico"):
                    classe_srv = srv.find(".//ClasseServico/Descricao")
                    desc_srv = srv.find("Descricao")
                    if classe_srv is not None and desc_srv is not None:
                        tipo_txt = classe_srv.text.strip() if classe_srv.text else ""
                        nome_txt = desc_srv.text.strip() if desc_srv.text else ""
                        srv_str = f"{tipo_txt} - {nome_txt}"
                        if srv_str not in servicos_vistos:
                            dados["servicos"].append({"tipo": tipo_txt, "nome": nome_txt})
                            servicos_vistos.add(srv_str)

        mapa_scripts = {
            "ScriptModificado": "modificado",
            "ScriptValidacao": "de validação",
            "ScriptFormularioCarregado": "de formulário carregado",
            "ScriptInicio": "de início",
            "ScriptFim": "de fim",
            "ScriptVolta": "de volta"
        }

        campos_vistos = set()
        anexos_vistos = set()
        scripts_vistos = set()
        
        for figura in root.findall(".//Figura"):
            atividade = figura.find("Atividade")
            
            if atividade is not None:
                tag_tipo_atv = atividade.find("Tipo")
                tipo_atv = tag_tipo_atv.text.strip() if (tag_tipo_atv is not None and tag_tipo_atv.text) else "Atividade"
                
                for op in atividade.findall(".//OperacaoAtividade"):
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
                            rot_txt = rotulo.text.strip() if (rotulo is not None and rotulo.text) else nome_c
                            
                            if nome_c and nome_c not in campos_vistos:
                                dados["campos"].append({"nome": nome_c, "rotulo": rot_txt, "tabela": tabela_nome})
                                campos_vistos.add(nome_c)
                                
                            for tag_xml, nome_amigavel in mapa_scripts.items():
                                s_node = campo.find(tag_xml)
                                if s_node is not None and s_node.text and s_node.text.strip():
                                    sc_code = s_node.text.strip()
                                    if sc_code not in scripts_vistos:
                                        dados["scripts"].append({
                                            "local": f"Atividade: {tipo_atv} | Script {nome_amigavel} no campo: {nome_c}",
                                            "codigo": sc_code
                                        })
                                        scripts_vistos.add(sc_code)

                    for tag_xml, nome_amigavel in mapa_scripts.items():
                        s_node = op.find(tag_xml)
                        if s_node is not None and s_node.text and s_node.text.strip():
                            sc_code = s_node.text.strip()
                            if sc_code not in scripts_vistos:
                                dados["scripts"].append({
                                    "local": f"Atividade: {tipo_atv} | Script {nome_amigavel} da operação",
                                    "codigo": sc_code
                                })
                                scripts_vistos.add(sc_code)

                for tag_xml, nome_amigavel in mapa_scripts.items():
                    s_node = atividade.find(tag_xml)
                    if s_node is not None and s_node.text and s_node.text.strip():
                        sc_code = s_node.text.strip()
                        if sc_code not in scripts_vistos:
                            dados["scripts"].append({
                                "local": f"Atividade: {tipo_atv} | Script {nome_amigavel} da atividade",
                                "codigo": sc_code
                            })
                            scripts_vistos.add(sc_code)

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
                    if sc_code not in scripts_vistos:
                        dados["scripts"].append({
                            "local": f"Gateway de Decisão: {desc_gw} | Expressão de Decisão",
                            "codigo": sc_code
                        })
                        scripts_vistos.add(sc_code)

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

    def gerar_documento(self):
        if not self.xml_path or not self.template_path or not self.output_dir:
            messagebox.showerror("Erro", "Preencha todos os caminhos (XML, Template e Pasta).")
            return

        try:
            dados = self.extrair_dados_xml()
            
            macro = self.entry_macro.get().strip()
            proc = self.entry_proc.get().strip()
            desc = self.text_descricao.get("1.0", tk.END).strip()
            evid = self.entry_evidencia.get().strip()

            doc = Document(self.template_path)

            idx_desc = -1
            idx_servico = -1
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
                elif texto == "Pesquisa de Mercado":
                    idx_desc = i
                elif texto == "Campos":
                    idx_servico = i
                elif texto.startswith("Evidências em Homologação:"):
                    idx_evid = i
                    p.text = ""
                    p.add_run("Evidências em Homologação: ").bold = True
                    p.add_run(evid)

            if idx_evid != -1 and dados["scripts"]:
                estilo_ancora = doc.paragraphs[idx_evid].style
                for s_dict in reversed(dados["scripts"]):
                    doc.paragraphs[idx_evid].insert_paragraph_before("", style=estilo_ancora)
                    
                    p_code = doc.paragraphs[idx_evid].insert_paragraph_before("")
                    self.add_code_block(p_code, s_dict["codigo"])
                    
                    p_local = doc.paragraphs[idx_evid].insert_paragraph_before("", style=estilo_ancora)
                    p_local.add_run(s_dict["local"]).bold = True

            if idx_servico != -1 and dados["servicos"]:
                estilo_ancora_srv = doc.paragraphs[idx_servico].style
                for srv in reversed(dados["servicos"]):
                    doc.paragraphs[idx_servico].insert_paragraph_before(
                        f"• Tipo: {srv['tipo']} | Serviço: {srv['nome']}", 
                        style=estilo_ancora_srv
                    )

            if idx_desc != -1 and desc:
                estilo_ancora_desc = doc.paragraphs[idx_desc].style
                doc.paragraphs[idx_desc].insert_paragraph_before(desc, style=estilo_ancora_desc)

            if len(doc.tables) > 0:
                t_campos = doc.tables[0]
                for row in t_campos.rows[1:]: t_campos._element.remove(row._tr)
                for c in dados["campos"]:
                    row = t_campos.add_row()
                    row.cells[0].text = c["nome"]
                    row.cells[1].text = c["rotulo"]
                    row.cells[2].text = "String"
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

            nome_arquivo = dados["nome_fluxo"] if dados["nome_fluxo"] else "Supravizio"
            for c in '<>:"/\\|?*': nome_arquivo = nome_arquivo.replace(c, "_")
            caminho_completo = os.path.join(self.output_dir, f"Artefato_{nome_arquivo}.docx")
            
            doc.save(caminho_completo)
            messagebox.showinfo("Sucesso", f"Artefato Injetado com formatação visual de código!\n\nSalvo em:\n{caminho_completo}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha na manipulação do DOCX:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SupravizioDocApp(root)
    root.mainloop()