"""
Configuração compartilhada dos testes.

Importa supra.py diretamente pelo caminho (em vez de depender de sys.path/instalação
como pacote), já que o projeto é um app de arquivo único, não uma lib empacotada.
Isso NÃO cria nenhuma janela/root do Tkinter — só o necessário para os testes,
que chamam as funções de extração de módulo (get_texto, extrair_dados_xml, etc.)
diretamente, sem instanciar SupravizioDocApp.
"""
import importlib.util
import logging
import sys
from pathlib import Path

import pytest

TESTS_DIR = Path(__file__).parent
FIXTURES_DIR = TESTS_DIR / "fixtures"
SUPRA_PATH = TESTS_DIR.parent / "supra.py"


def _carregar_supra():
    spec = importlib.util.spec_from_file_location("supra", SUPRA_PATH)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


# Carregado uma única vez por sessão de testes (o módulo não tem efeitos colaterais
# no import — só cria a janela Tkinter dentro de `if __name__ == "__main__":`).
supra = _carregar_supra()


@pytest.fixture(scope="session")
def mod():
    """O módulo supra.py já carregado, para acessar suas funções de módulo."""
    return supra


@pytest.fixture
def logger():
    """Logger simples que não escreve em lugar nenhum, só para satisfazer a assinatura."""
    log = logging.getLogger("supra-tests")
    log.addHandler(logging.NullHandler())
    log.setLevel(logging.DEBUG)
    return log


@pytest.fixture
def fixture_path():
    """Devolve uma função que resolve o caminho de um XML dentro de tests/fixtures/."""
    def _resolver(nome_arquivo):
        caminho = FIXTURES_DIR / nome_arquivo
        assert caminho.exists(), f"Fixture não encontrada: {caminho}"
        return str(caminho)
    return _resolver
