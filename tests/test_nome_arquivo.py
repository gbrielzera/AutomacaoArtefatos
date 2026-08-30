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
