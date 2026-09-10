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


def test_extrai_colunas_de_campo_de_grade(mod, fixture_path):
    """Campos de grade (Type='RecordList') expõem suas colunas para rotular os scripts."""
    root = mod.ler_xml_root(fixture_path("campo_grade.xml"))
    propriedades = mod.extrair_propriedades_campos(root)

    colunas = propriedades["DADOS_BEM"]["colunas"]
    assert colunas == {"NUM_PATRIM": "Número de Patrimônio", "PLANTA": "Local - Planta"}


def test_campo_comum_tem_colunas_vazias(mod, fixture_path):
    """Campo que não é grade não pode inventar colunas."""
    root = mod.ler_xml_root(fixture_path("custom_property.xml"))
    propriedades = mod.extrair_propriedades_campos(root)

    assert propriedades["CSC_TESTE"]["colunas"] == {}
