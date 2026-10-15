"""
Testes da comparação entre duas versões do mesmo fluxo (aba 2 / v2).

As fixtures comparacao_antes.xml e comparacao_depois.xml representam o mesmo fluxo em
duas versões, com um exemplo de cada tipo de mudança que o artefato precisa documentar:
serviço/anexo/campo incluído e removido, campo cuja lista de combobox mudou, e script
incluído, removido e alterado.
"""

import pytest


@pytest.fixture
def comp(mod, logger, fixture_path):
    antes = mod.extrair_dados_xml(fixture_path("comparacao_antes.xml"), logger)
    depois = mod.extrair_dados_xml(fixture_path("comparacao_depois.xml"), logger)
    return mod.comparar_dados(antes, depois, logger)


# --------------------------------------------------------------------------
# Serviços e itens de configuração: só entram ou saem
# --------------------------------------------------------------------------
def test_servicos_incluidos_e_removidos(comp):
    incluidos = [s["nome"] for s in comp["servicos"]["incluidos"]]
    removidos = [s["nome"] for s in comp["servicos"]["removidos"]]

    assert incluidos == ["Serviço Que Entra"]
    assert removidos == ["Serviço Que Sai"]


def test_servicos_trazem_a_lista_completa_da_versao_nova(comp):
    """
    Serviços são a exceção: o artefato lista TODOS os do fluxo, não só o delta.
    'Serviço Mantido' não mudou, mas ainda assim precisa constar.
    """
    todos = [s["nome"] for s in comp["servicos"]["todos"]]

    assert "Serviço Mantido" in todos
    assert "Serviço Que Entra" in todos
    # o removido não faz parte da versão nova; ele é listado à parte, marcado
    assert "Serviço Que Sai" not in todos
    assert ("Requisição", "Serviço Que Entra") in comp["servicos"]["chaves_incluidas"]
    assert ("Solicitação", "Serviço Mantido") not in comp["servicos"]["chaves_incluidas"]


def test_anexos_incluidos_e_removidos(comp):
    assert [a["sigla"] for a in comp["anexos"]["incluidos"]] == ["ANXENTRA"]
    assert [a["sigla"] for a in comp["anexos"]["removidos"]] == ["ANXSAI"]


# --------------------------------------------------------------------------
# Campos: incluídos/removidos vão para uma tabela, modificados para outra
# --------------------------------------------------------------------------
def test_campos_incluidos_e_removidos(comp):
    assert [c["nome"] for c in comp["campos"]["incluidos"]] == ["CAMPO_NOVO"]
    assert [c["nome"] for c in comp["campos"]["removidos"]] == ["CAMPO_REMOVIDO"]


def _por_nome(modificados):
    return {m["chave"]: m for m in modificados}


def test_campo_com_lista_alterada_entra_como_modificado(comp):
    alterado = _por_nome(comp["campos"]["modificados"])["TIPO_DEMANDA"]
    assert alterado["diferencas"] == ["lista_itens"]


def test_campo_que_mudou_de_tipo_entra_como_modificado(comp):
    """String -> Integer no mesmo campo: o caso central da tabela 'Campos alterados'."""
    alterado = _por_nome(comp["campos"]["modificados"])["TEXT"]

    assert "tipo" in alterado["diferencas"]
    assert alterado["antes"]["tipo"] == "String"
    assert alterado["depois"]["tipo"] == "Integer"


def test_coluna_tipo_mostra_a_transicao_quando_o_tipo_muda(mod, comp):
    alterado = _por_nome(comp["campos"]["modificados"])["TEXT"]
    assert mod.descrever_mudanca_tipo(alterado["antes"], alterado["depois"]) == "String → Integer"


def test_coluna_tipo_mostra_so_o_tipo_atual_quando_ele_nao_muda(mod, comp):
    # TIPO_DEMANDA mudou a lista, mas continua String (DropDownList)
    alterado = _por_nome(comp["campos"]["modificados"])["TIPO_DEMANDA"]
    assert mod.descrever_mudanca_tipo(alterado["antes"], alterado["depois"]) == "String (DropDownList)"


def test_campo_sem_mudanca_nao_aparece(comp):
    todos = (
        [c["nome"] for c in comp["campos"]["incluidos"]]
        + [c["nome"] for c in comp["campos"]["removidos"]]
        + [m["chave"] for m in comp["campos"]["modificados"]]
    )
    assert "CAMPO_ESTAVEL" not in todos


# --------------------------------------------------------------------------
# Coluna "Lista de Itens" da tabela de Campos alterados
# --------------------------------------------------------------------------
def test_descricao_da_lista_mostra_itens_atuais_e_o_que_mudou(mod, comp):
    alterado = _por_nome(comp["campos"]["modificados"])["TIPO_DEMANDA"]
    texto = mod.descrever_mudanca_lista(alterado["antes"], alterado["depois"])

    # primeira linha: a lista como está na versão nova
    assert texto.splitlines()[0] == "Mercadoria; Serviços; Transporte"
    assert "Incluídos: Transporte" in texto
    assert "Removidos: Obsoleto" in texto


def test_descricao_da_lista_vazia_para_campo_sem_lista(mod):
    assert mod.descrever_mudanca_lista({"lista_itens": ""}, {"lista_itens": ""}) == ""


def test_tipo_do_campo_indica_o_controle_de_lista(mod, comp):
    alterado = _por_nome(comp["campos"]["modificados"])["TIPO_DEMANDA"]
    assert mod.descrever_tipo_campo(alterado["depois"]) == "String (DropDownList)"
    # um TextBox não recebe sufixo
    assert mod.descrever_tipo_campo({"tipo": "String", "controle": "TextBox"}) == "String"


def test_separar_lista_itens_ignora_vazios(mod):
    assert mod.separar_lista_itens("A;B;;C ; ") == ["A", "B", "C"]
    assert mod.separar_lista_itens("") == []


# --------------------------------------------------------------------------
# Campos alterados: troca de tipo e origem da lista de itens
#
# comparacao_listas_*.xml traz um campo para cada situação. A lista de um combo pode ser
# fixa (itens separados por ';'), carregada por script (<LookupScript>, em geral um SELECT)
# ou vir de uma classe de negócio (<ClassReference>). Os itens de um script não dá para
# enumerar no XML, então o artefato só informa se o script mudou.
# --------------------------------------------------------------------------
@pytest.fixture
def comp_listas(mod, logger, fixture_path):
    antes = mod.extrair_dados_xml(fixture_path("comparacao_listas_antes.xml"), logger)
    depois = mod.extrair_dados_xml(fixture_path("comparacao_listas_depois.xml"), logger)
    return mod.comparar_dados(antes, depois, logger)


@pytest.fixture
def linhas_lista(mod, comp_listas):
    """Devolve as linhas da coluna 'Lista de Itens' de um campo alterado."""
    def _linhas(nome):
        m = _por_nome(comp_listas["campos"]["modificados"])[nome]
        return mod.descrever_mudanca_lista(m["antes"], m["depois"]).splitlines()
    return _linhas


def test_decimal_que_virou_string_aparece_com_a_transicao(mod, comp_listas):
    alterado = _por_nome(comp_listas["campos"]["modificados"])["DECIMAL_PARA_STRING"]

    assert alterado["diferencas"] == ["tipo"]
    assert mod.descrever_mudanca_tipo(alterado["antes"], alterado["depois"]) == "Decimal → String"
    # não é campo de lista: a coluna fica vazia
    assert mod.descrever_mudanca_lista(alterado["antes"], alterado["depois"]) == ""


def test_lista_fixa_que_passou_a_vir_de_script(comp_listas, linhas_lista):
    alterado = _por_nome(comp_listas["campos"]["modificados"])["FIXA_PARA_SCRIPT"]

    assert "lookup_script" in alterado["diferencas"]
    assert linhas_lista("FIXA_PARA_SCRIPT") == [
        "Itens carregados por script",
        "Antes: lista fixa (Contrato; Aditivo)",
    ]


def test_script_de_carga_alterado_e_detectado(comp_listas, linhas_lista):
    """Regressão: mudar o SELECT que alimenta o combo não aparecia em 'Campos alterados'."""
    alterado = _por_nome(comp_listas["campos"]["modificados"])["SCRIPT_ALTERADO"]

    assert alterado["diferencas"] == ["lookup_script"]
    assert linhas_lista("SCRIPT_ALTERADO") == [
        "Itens carregados por script",
        "Script de carga dos itens alterado",
    ]


def test_script_com_mesmo_codigo_nao_acusa_mudanca(comp_listas):
    """Espaço sobrando e fim de linha CRLF diferentes entre as exportações não são mudança."""
    nomes = [m["chave"] for m in comp_listas["campos"]["modificados"]]

    assert "SCRIPT_MESMO_CODIGO" not in nomes
    assert "SCRIPT_ESTAVEL" not in nomes


def test_script_so_com_comentario_nao_conta_como_lista(comp_listas):
    """
    Um SELECT comentado não carrega item algum: trocar o texto do comentário não é
    alteração de lista, e o campo nem chega a ter 'script de carga'.
    """
    nomes = [m["chave"] for m in comp_listas["campos"]["modificados"]]

    assert "SCRIPT_SO_COMENTARIO" not in nomes


def test_comentario_que_ganhou_codigo_passa_a_ser_lista_por_script(comp_listas, linhas_lista):
    assert linhas_lista("COMENTARIO_PARA_SCRIPT") == [
        "Itens carregados por script",
        "Antes: sem lista de itens",
    ]


def test_script_que_virou_lista_fixa(comp_listas, linhas_lista):
    alterado = _por_nome(comp_listas["campos"]["modificados"])["SCRIPT_PARA_FIXA"]

    assert set(alterado["diferencas"]) == {"lista_itens", "lookup_script"}
    assert linhas_lista("SCRIPT_PARA_FIXA") == [
        "Aberto; Fechado",
        "Antes: itens carregados por script",
    ]


def test_classe_de_origem_trocada(comp_listas, linhas_lista):
    alterado = _por_nome(comp_listas["campos"]["modificados"])["CLASSE_ALTERADA"]

    assert alterado["diferencas"] == ["classe_referencia"]
    assert linhas_lista("CLASSE_ALTERADA") == [
        "Itens da classe Filial",
        "Classe de origem alterada: Área → Filial",
    ]


def test_classe_sem_texto_usa_o_ultimo_trecho_do_nome_tecnico(comp_listas, mod, logger, fixture_path):
    depois = mod.extrair_dados_xml(fixture_path("comparacao_listas_depois.xml"), logger)
    campo = next(c for c in depois["campos"] if c["nome"] == "CLASSE_SEM_TEXTO")

    assert campo["classe_referencia"] == "Orgao"
    assert "CLASSE_SEM_TEXTO" not in [m["chave"] for m in comp_listas["campos"]["modificados"]]


def test_mesmos_itens_em_outra_ordem_sao_sinalizados(linhas_lista):
    """Sem isso a linha apareceria sem nenhuma indicação visível do que mudou."""
    assert linhas_lista("ORDEM_ALTERADA") == [
        "Baixa; Media; Alta",
        "Ordem dos itens alterada",
    ]


def test_lista_fixa_esvaziada(linhas_lista):
    assert linhas_lista("LISTA_ESVAZIADA") == [
        "(sem itens)",
        "Antes: lista fixa (Sim; Nao)",
    ]


def test_lista_fixa_continua_mostrando_incluidos_e_removidos(mod):
    """O comportamento antigo da lista fixa não pode mudar."""
    antes = {"lista_itens": "A;B;C"}
    depois = {"lista_itens": "A;C;D"}

    assert mod.descrever_mudanca_lista(antes, depois).splitlines() == [
        "A; C; D", "Incluídos: D", "Removidos: B",
    ]


def test_normalizar_script_ignora_fim_de_linha_e_espaco_a_direita(mod):
    assert mod.normalizar_script("a = 1  \r\nb = 2\t\r\n") == "a = 1\nb = 2"
    assert mod.normalizar_script("   ") == ""


def test_script_so_com_boilerplate_ou_comentario_nao_tem_codigo_ativo(mod):
    assert mod.script_tem_codigo_ativo("import clr\nfrom System import *\n# INICIO SCRIPT USUARIO") is False
    assert mod.script_tem_codigo_ativo("") is False
    assert mod.script_tem_codigo_ativo("import clr\nItens = DB.ExecuteDataTable('SELECT 1')") is True


# --------------------------------------------------------------------------
# Scripts
# --------------------------------------------------------------------------
def test_scripts_incluidos_removidos_e_alterados(comp):
    scripts = comp["scripts"]

    incluidos = [s["local"] for s in scripts["incluidos"]]
    removidos = [s["local"] for s in scripts["removidos"]]
    modificados = [m["chave"] for m in scripts["modificados"]]

    assert any("Novo gateway" in l for l in incluidos)
    assert any("Fim (EventoFinal)" in l for l in removidos)
    assert any("Preencher dados (Tarefa)" in l for l in modificados)


def test_script_alterado_carrega_as_duas_versoes_do_codigo(comp):
    alterado = comp["scripts"]["modificados"][0]
    assert "script que sera alterado" in alterado["antes"]["codigo"]
    assert "script JA alterado" in alterado["depois"]["codigo"]
    assert alterado["diferencas"] == ["codigo"]


# --------------------------------------------------------------------------
# Metadados da comparação
# --------------------------------------------------------------------------
def test_nome_fluxo_igual_nao_gera_alerta(comp):
    assert comp["nome_fluxo"] == "Fluxo Comparavel"
    assert comp["nome_fluxo_mudou"] is False


def test_nome_fluxo_diferente_e_sinalizado(mod, logger, fixture_path):
    antes = mod.extrair_dados_xml(fixture_path("comparacao_antes.xml"), logger)
    outro = mod.extrair_dados_xml(fixture_path("basico.xml"), logger)
    resultado = mod.comparar_dados(antes, outro, logger)

    assert resultado["nome_fluxo_mudou"] is True
    assert resultado["nome_fluxo_antes"] == "Fluxo Comparavel"
    assert resultado["nome_fluxo"] == "Processo Basico Teste"


def test_comparar_versoes_identicas_nao_acusa_mudanca(mod, logger, fixture_path):
    dados = mod.extrair_dados_xml(fixture_path("comparacao_antes.xml"), logger)
    resultado = mod.comparar_dados(dados, dados, logger)

    for secao in ("servicos", "campos", "anexos", "scripts"):
        assert resultado[secao]["incluidos"] == []
        assert resultado[secao]["removidos"] == []
        assert resultado[secao]["modificados"] == []


# --------------------------------------------------------------------------
# Versão no nome do arquivo
# --------------------------------------------------------------------------
def test_extrai_versao_do_nome_arquivo(mod):
    assert mod.extrair_versao_do_nome_arquivo(
        "Administração_-_Contratos_Versão_44_[LOTE]_Pré-Notificação.xml") == "44"
    assert mod.extrair_versao_do_nome_arquivo(
        "Serviços_Assistência_Técnica_Versão_15_Recebimento.xml") == "15"
    assert mod.extrair_versao_do_nome_arquivo("arquivo_sem_versao.xml") is None
