"""Testes de parsing do nome de arquivo (Processo/Subprocesso) e de sanitização de nomes."""


def test_extrai_processo_padrao_valido(mod):
    processo, sub = mod.extrair_processo_do_nome_arquivo(
        "Financeiro - Serviços Gerais_Versão_38_Comprovantes de Valores Pagos"
    )
    assert processo == "Financeiro - Serviços Gerais"
    assert sub == "Comprovantes de Valores Pagos"


def test_extrai_processo_com_underscores(mod):
    processo, sub = mod.extrair_processo_do_nome_arquivo(
        "Suprimentos_-_Contratos_Versão_27_Registro_de_Notas_Fiscais_de_Fornecedores"
    )
    assert processo == "Suprimentos - Contratos"
    assert sub == "Registro de Notas Fiscais de Fornecedores"


def test_extrai_processo_padrao_invalido_retorna_none(mod):
    processo, sub = mod.extrair_processo_do_nome_arquivo("arquivo_sem_padrao_esperado")
    assert processo is None
    assert sub is None


def test_limpar_nome_arquivo_remove_caracteres_invalidos(mod):
    assert mod.limpar_nome_arquivo('a:b/c*d?e"f<g>h|i') == "a_b_c_d_e_f_g_h_i"


def test_limpar_nome_arquivo_vazio_retorna_fallback(mod):
    assert mod.limpar_nome_arquivo("") == "Supravizio"
    assert mod.limpar_nome_arquivo("   ") == "Supravizio"


def test_nome_artefato_criacao(mod):
    assert mod.montar_nome_artefato("Reembolso") == "Artefato_Reembolso.docx"


def test_nome_artefato_comparacao_nao_leva_versao_nem_prefixo(mod):
    """
    Regressão: a aba de comparação gerava "Artefato_Reembolso_v44_para_v45.docx".
    O comparativo perde o número da versão E o prefixo "Artefato_" — é o prefixo que
    passa a diferenciar os dois artefatos do mesmo fluxo na pasta de destino.
    """
    assert mod.montar_nome_artefato_comparacao("Reembolso") == "Reembolso.docx"


def test_nomes_das_duas_abas_nao_colidem(mod):
    """Gerar os dois artefatos do mesmo fluxo na mesma pasta não pode sobrescrever nada."""
    assert mod.montar_nome_artefato("Reembolso") != mod.montar_nome_artefato_comparacao("Reembolso")


def test_nome_artefato_sanitiza_subprocesso(mod):
    assert mod.montar_nome_artefato("Publicação no DOU: Extrato") == "Artefato_Publicação no DOU_ Extrato.docx"
    assert mod.montar_nome_artefato_comparacao("Publicação no DOU: Extrato") == "Publicação no DOU_ Extrato.docx"


def test_nome_artefato_sem_nome_de_fluxo_usa_fallback(mod):
    assert mod.montar_nome_artefato("") == "Artefato_Supravizio.docx"
    assert mod.montar_nome_artefato_comparacao("") == "Supravizio.docx"
