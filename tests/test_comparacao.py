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
