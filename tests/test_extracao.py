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
