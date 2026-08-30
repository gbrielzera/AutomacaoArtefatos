"""Testes de extrair_propriedades_campos (casamento de CampoPreenchimento com CustomProperty)."""


def test_extrai_propriedades_por_nome_customizado(mod, fixture_path):
    root = mod.ler_xml_root(fixture_path("custom_property.xml"))
    propriedades = mod.extrair_propriedades_campos(root)

    assert "CSC_TESTE" in propriedades
    prop = propriedades["CSC_TESTE"]
    assert prop["tabela"] == "CPE_CSC"
    assert prop["rotulo"] == "Rótulo CSC"
    assert prop["descricao"] == "Descrição CSC"
    assert prop["tipo"] == "Integer"
    assert prop["coluna"] == "COL_CSC"


def test_ignora_custom_property_sem_tabela(mod, fixture_path):
    root = mod.ler_xml_root(fixture_path("custom_property.xml"))
    propriedades = mod.extrair_propriedades_campos(root)

    assert "SEM_TABELA" not in propriedades


def test_extrai_dados_xml_usa_propriedades_do_custom_property(mod, logger, fixture_path):
    """O campo do fluxo deve herdar rótulo/descrição/tipo/tabela do CustomProperty."""
    dados = mod.extrair_dados_xml(fixture_path("custom_property.xml"), logger)

    assert len(dados["campos"]) == 1
    campo = dados["campos"][0]
    assert campo["nome"] == "CSC_TESTE"
    assert campo["tabela"] == "CPE_CSC"
    assert campo["rotulo"] == "Rótulo CSC"
