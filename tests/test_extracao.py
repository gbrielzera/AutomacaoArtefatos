"""
Testes de extrair_dados_xml (o núcleo da leitura do XML do Supravizio).

Cada teste usa uma fixture pequena e específica em tests/fixtures/. Os dois testes de
dedup (test_dedup_*) são testes de regressão de um bug real já corrigido: a
deduplicação de scripts passou a usar a chave (local, código) em vez de só código,
porque script idêntico em locais diferentes (ex.: Link Inicial x Evento Inicial) era
descartado por engano.
"""


def test_extrai_campo_e_script_basico(mod, logger, fixture_path):
    dados = mod.extrair_dados_xml(fixture_path("basico.xml"), logger)

    assert dados["nome_fluxo"] == "Processo Basico Teste"
    assert len(dados["campos"]) == 1
    assert dados["campos"][0]["nome"] == "CAMPO_TESTE"

    assert len(dados["scripts"]) == 1
    assert "CAMPO_TESTE" in dados["scripts"][0]["local"]
    assert "print('ok')" in dados["scripts"][0]["codigo"]

    assert len(dados["servicos"]) == 1
    assert dados["servicos"][0] == {"tipo": "Solicitação", "nome": "Serviço Teste"}

    assert len(dados["anexos"]) == 1
    assert dados["anexos"][0] == {"nome": "Comprovante", "sigla": "COMP"}


def test_dedup_preserva_script_em_locais_distintos(mod, logger, fixture_path):
    """
    Regressão: o mesmo script (mesmo código) no campo FORNECEDOR1 aparece tanto no
    LinkInicial quanto no EventoInicial do fluxo. Os dois devem ser preservados,
    porque são cópias em ATIVIDADES diferentes, não a mesma atividade duplicada.
    """
    dados = mod.extrair_dados_xml(fixture_path("link_evento_mesmo_script.xml"), logger)

    assert len(dados["scripts"]) == 2
    locais = [s["local"] for s in dados["scripts"]]
    assert any("LinkInicial" in l for l in locais)
    assert any("EventoInicial" in l for l in locais)


def test_dedup_colapsa_atividade_repetida_no_diagrama(mod, logger, fixture_path):
    """
    O MESMO nó de atividade pode aparecer redesenhado em mais de uma Figura do
    diagrama (mesmo Tipo, Descricao e script). Isso deve colapsar para 1 entrada,
    não duplicar.
    """
    dados = mod.extrair_dados_xml(fixture_path("atividade_repetida.xml"), logger)

    assert len(dados["scripts"]) == 1
    assert len(dados["campos"]) == 1  # dedup de campo por nome também não deve regredir


def test_gateway_expressao_decisao_extraida(mod, logger, fixture_path):
    dados = mod.extrair_dados_xml(fixture_path("gateway_decisao.xml"), logger)

    assert len(dados["scripts"]) == 1
    script = dados["scripts"][0]
    assert "Gateway de Decisão: Aprovado?" in script["local"]
    assert "return True" in script["codigo"]


def test_nome_fluxo_extraido_de_nome_subprocesso(mod, logger, fixture_path):
    dados = mod.extrair_dados_xml(fixture_path("gateway_decisao.xml"), logger)
    assert dados["nome_fluxo"] == "Fluxo com Gateway"


def test_xml_vazio_levanta_excecao(mod, fixture_path):
    import pytest as _pytest
    with _pytest.raises(Exception, match="vazio"):
        mod.ler_xml_root(fixture_path("vazio.xml"))


def _local_dos_scripts(dados):
    return [s["local"] for s in dados["scripts"]]


def test_extrai_script_de_coluna_de_grade(mod, logger, fixture_path):
    """
    Regressão: colunas de um campo de grade guardam suas abas de script em
    <CamposPreenchimentoRegistro>, irmão das abas do campo pai. A leitura só olhava
    o <CampoPreenchimento>, então todo script de coluna de grade sumia do artefato.
    """
    dados = mod.extrair_dados_xml(fixture_path("campo_grade.xml"), logger)
    codigos = {s["codigo"] for s in dados["scripts"]}

    assert "busca_patrimonio()" in codigos
    assert "valida_planta()" in codigos
    # O script do próprio campo-grade continua sendo lido.
    assert "script_do_campo_pai()" in codigos


def test_script_de_grade_usa_rotulo_da_coluna(mod, logger, fixture_path):
    """O local do script deve citar o rótulo de <RecordColumn>, não o nome técnico."""
    dados = mod.extrair_dados_xml(fixture_path("campo_grade.xml"), logger)
    por_codigo = {s["codigo"]: s["local"] for s in dados["scripts"]}

    local = por_codigo["busca_patrimonio()"]
    assert "Número de Patrimônio" in local
    assert "da GRID: DADOS_BEM" in local
    assert "Transferência de Bem" in local

    assert "Script de validação" in por_codigo["valida_planta()"]
    assert "Local - Planta" in por_codigo["valida_planta()"]


def test_script_de_grade_sem_recordcolumn_usa_nome_tecnico(mod, logger, fixture_path):
    """Coluna ausente de <RecordColumns> não pode ser descartada: cai no nome técnico."""
    dados = mod.extrair_dados_xml(fixture_path("campo_grade.xml"), logger)
    por_codigo = {s["codigo"]: s["local"] for s in dados["scripts"]}

    assert "COLUNA_ORFA" in por_codigo["script_coluna_orfa()"]


def test_script_de_grade_vazio_e_ignorado(mod, logger, fixture_path):
    """Uma aba de script em branco não deve virar entrada no artefato."""
    dados = mod.extrair_dados_xml(fixture_path("campo_grade.xml"), logger)

    assert not any("SEM_SCRIPT" in local for local in _local_dos_scripts(dados))


def test_extrai_abas_confirmado_e_adicionado(mod, logger, fixture_path):
    """ScriptConfirmado/ScriptAdicionado são abas do campo e também vão para o artefato."""
    dados = mod.extrair_dados_xml(fixture_path("campo_scripts_abas.xml"), logger)
    por_codigo = {s["codigo"]: s["local"] for s in dados["scripts"]}

    assert "Script confirmado no campo: ITEM_LISTA" in por_codigo["ao_confirmar()"]
    assert "Script adicionado no campo: ITEM_LISTA" in por_codigo["ao_adicionar()"]


def test_script_de_selecao_de_atores_nao_entra_no_artefato(mod, logger, fixture_path):
    """
    Decisão de conteúdo, não lacuna: <ScriptSelecaoAtores> é excluído de propósito.
    Chegava a um terço de tudo que era listado e afogava os scripts que interessam.
    Se este teste falhar, alguém reintroduziu a extração achando que era um bug.
    """
    dados = mod.extrair_dados_xml(fixture_path("selecao_atores.xml"), logger)
    codigos = {s["codigo"] for s in dados["scripts"]}

    assert codigos == {"script_de_inicio()"}
    assert not any("seleção de atores" in s["local"] for s in dados["scripts"])
