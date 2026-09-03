# Automação de Artefatos

> **Transformando processos complexos em documentação estruturada, rastreável e pronta para uso.**

Aplicação desktop desenvolvida em **Python** para automatizar a geração e atualização de documentos técnicos a partir de arquivos **XML de processos** e modelos **DOCX**.

O projeto combina **processamento de XML, comparação entre versões, análise contextual de scripts, geração de documentos, formatação de código e testes automatizados** em uma única aplicação.

A ideia central é eliminar uma classe de trabalho extremamente suscetível a erros:

> **se uma informação já está estruturada em um sistema, o usuário não deveria precisar encontrá-la, copiá-la e documentá-la manualmente.**

---

## Visão geral

A aplicação recebe informações estruturadas de um processo e transforma esses dados em documentação técnica.

Além da geração de um artefato a partir de uma única versão do fluxo, o sistema é capaz de **comparar duas versões do mesmo processo**, identificando o que foi:

* incluído;
* removido;
* modificado;
* mantido sem alterações.

Isso permite transformar a documentação de uma tarefa puramente manual em um processo **automatizado, rastreável e baseado em diferenças reais entre versões**.

---

## O problema

Processos empresariais podem possuir uma estrutura bastante complexa.

Uma única alteração pode envolver simultaneamente:

* serviços;
* campos;
* tabelas;
* controles de interface;
* listas de opções;
* anexos;
* scripts;
* validações;
* regras de negócio;
* gateways;
* diferentes atividades do fluxo.

Documentar essas mudanças manualmente exige navegar pela estrutura do processo, localizar cada alteração e posteriormente transcrever essas informações para um documento padronizado.

Isso cria três problemas principais:

### Produtividade

Grande quantidade de trabalho repetitivo de busca, conferência e transcrição.

### Confiabilidade

É fácil esquecer um elemento, duplicar informações ou documentar uma alteração no contexto errado.

### Rastreabilidade

Ao comparar duas versões manualmente, torna-se difícil responder rapidamente:

> **O que realmente mudou?**

A aplicação foi construída justamente para resolver esse problema.

---

# A solução

O sistema transforma:

```text
XML + Modelo DOCX
        │
        ▼
┌──────────────────────┐
│ Extração estruturada │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Análise / Comparação │
└──────────┬───────────┘
           │
           ├── Serviços
           ├── Campos
           ├── Anexos
           ├── Scripts
           └── Regras
           │
           ▼
┌──────────────────────┐
│ Geração do artefato  │
└──────────┬───────────┘
           │
           ▼
     Documento DOCX
```

O resultado é uma documentação construída a partir dos próprios dados do processo, reduzindo significativamente a necessidade de intervenção manual.

---

# Comparação entre versões

Uma das principais evoluções do projeto é a capacidade de comparar duas versões de um fluxo.

O sistema extrai os dados de cada XML e produz uma estrutura de diferenças.

### Exemplo conceitual

```text
VERSÃO ANTERIOR
       │
       ▼
   Extrator
       │
       ▼
┌───────────────┐
│ Dados antigos │
└───────┬───────┘
        │
        │ comparar
        │
        ▼
┌───────────────┐
│ Dados atuais  │
└───────┬───────┘
       ▲
       │
   Extrator
       ▲
       │
VERSÃO ATUAL
```

A comparação identifica alterações em diferentes categorias.

### Serviços

```text
Incluídos
Removidos
Mantidos
```

### Anexos

```text
Incluídos
Removidos
```

### Campos

```text
Incluídos
Removidos
Modificados
```

A análise de campos consegue identificar alterações como:

* tipo;
* controle;
* lista de itens;
* descrição;
* propriedades relacionadas.

Por exemplo:

```text
String → Integer
```

ou:

```text
Lista anterior:
Mercadoria; Serviços; Obsoleto

Lista atual:
Mercadoria; Serviços; Transporte
```

O sistema consegue representar a alteração como:

```text
Incluídos: Transporte
Removidos: Obsoleto
```

---

# Análise contextual de scripts

Scripts são tratados como elementos de primeira classe.

O sistema não considera apenas o conteúdo do código para identificar uma ocorrência.

O **contexto e o local de execução** também fazem parte da identificação.

Isso é importante porque o mesmo código pode aparecer legitimamente em diferentes pontos do fluxo.

Por exemplo:

```text
Atividade: LinkInicial
Script modificado no campo: FORNECEDOR1
```

ou:

```text
Gateway de Decisão: Validar solicitação
Expressão de Decisão
```

Durante a comparação entre versões, os scripts podem ser classificados como:

```text
┌───────────────┐
│ Script novo   │
└───────────────┘

┌───────────────┐
│ Script removido│
└───────────────┘

┌───────────────┐
│ Script alterado│
└───────────────┘
```

Quando um script é alterado, a estrutura da comparação mantém as duas versões do código:

```text
ANTES
    ↓
código original

DEPOIS
    ↓
código modificado
```

Isso permite que a documentação represente não apenas a existência da alteração, mas também **o que foi alterado**.

---

# Código com formatação de IDE

Um dos diferenciais do projeto está na forma como os scripts são inseridos no documento.

Em vez de transformar o código em texto simples, a aplicação utiliza **Pygments** para realizar a tokenização e aplicar destaque sintático.

O objetivo é preservar:

* indentação;
* quebras de linha;
* espaçamento;
* estrutura visual;
* fonte monoespaçada;
* destaque de sintaxe;
* contexto do script.

O resultado aproxima a apresentação do código da experiência de leitura de uma IDE.

Isso é especialmente importante quando o documento contém grandes quantidades de **Python/IronPython**.

---

# Extração resiliente de XML

A aplicação utiliza `xml.etree.ElementTree` para navegar pela estrutura dos arquivos exportados.

A extração possui mecanismos específicos para lidar com características dos XMLs utilizados pelo processo.

Entre eles:

* tratamento de encoding e BOM;
* identificação do processo a partir do nome do arquivo;
* extração de propriedades de campos;
* identificação de controles de interface;
* extração de listas de opções;
* extração de serviços;
* extração de anexos;
* identificação de atividades;
* identificação de gateways;
* extração de scripts em diferentes contextos;
* deduplicação contextual;
* fallback para estruturas alternativas.

### Deduplicação contextual

Um simples:

```python
if script in scripts:
    ignore()
```

não seria suficiente.

O projeto utiliza uma abordagem baseada no contexto:

```text
(local, código)
```

Dessa forma, um mesmo código pode aparecer mais de uma vez quando realmente representa ocorrências diferentes.

Ao mesmo tempo, ocorrências duplicadas provocadas pela própria estrutura do diagrama podem ser eliminadas.

---

# Informações extraídas

A aplicação trabalha com diferentes categorias de informações.

### Fluxo

* nome do subprocesso;
* processo;
* versão.

### Serviços

* tipo;
* nome;
* serviços associados ao fluxo principal.

### Campos

* nome;
* rótulo;
* descrição;
* tipo;
* tabela;
* coluna;
* controle;
* lista de itens.

### Configurações

* anexos;
* arquivos permitidos;
* propriedades relacionadas.

### Scripts

* script modificado;
* validação;
* formulário carregado;
* início;
* fim;
* eventos;
* retorno;
* seleção de atores;
* expressões;
* regras de decisão.

---

# Interface desktop

A aplicação possui uma interface gráfica desenvolvida com **Tkinter**.

A proposta é manter a complexidade da automação escondida do usuário.

O fluxo de utilização é essencialmente:

```text
Selecionar arquivos
       ↓
Preencher informações complementares
       ↓
Processar
       ↓
Gerar documentação
```

Também existe suporte opcional a **Drag & Drop** através de `tkinterdnd2`.

Caso a biblioteca não esteja disponível, a aplicação continua funcionando utilizando a interface convencional do Tkinter.

---

# Observabilidade e tratamento de erros

Automação em ambiente real precisa ser diagnosticável.

Por isso, o projeto possui uma camada própria de logging.

Durante a execução podem ser registrados:

* início da operação;
* arquivos utilizados;
* quantidade de elementos encontrados;
* elementos ignorados;
* avisos de inconsistências;
* etapas da extração;
* resultados da comparação;
* exceções.

Em caso de falha inesperada, o sistema registra um **crash log** antes de apresentar o erro ao usuário.

A estrutura também possui tratamento para falhas em recursos opcionais, evitando que funcionalidades auxiliares comprometam a execução principal.

---

# Persistência de configuração

Informações utilizadas frequentemente podem ser persistidas localmente.

A configuração é armazenada em:

```text
config/
└── settings.json
```

Isso permite recuperar informações da execução anterior e reduzir preenchimentos repetitivos.

Falhas na leitura ou gravação dessas configurações não impedem a execução principal da aplicação.

---

# Testes automatizados

A lógica de negócio foi estruturada de forma que as principais funções possam ser testadas **sem depender da interface gráfica**.

O projeto possui testes para:

* extração de dados;
* propriedades de campos;
* parsing do nome do arquivo;
* identificação de versão;
* comparação entre versões;
* serviços incluídos/removidos;
* anexos incluídos/removidos;
* campos incluídos/removidos;
* campos modificados;
* alterações de tipo;
* alterações de listas;
* scripts incluídos/removidos;
* scripts modificados;
* comparação de versões idênticas;
* alteração do nome do fluxo.

As fixtures utilizadas pelos testes simulam diferentes cenários de alteração, permitindo validar a lógica de comparação de forma determinística.

Executar:

```bash
pytest
```

---

# Arquitetura

Apesar de ser uma aplicação desktop compacta, existe uma preocupação clara em manter a lógica de processamento separada da interface.

A estrutura conceitual é:

```text
┌───────────────────────────┐
│        Interface GUI      │
│          Tkinter          │
└─────────────┬─────────────┘
              │
              ▼
┌───────────────────────────┐
│      Camada de domínio    │
│                           │
│  Extração                 │
│  Comparação               │
│  Transformação            │
│  Validação                │
└─────────────┬─────────────┘
              │
       ┌──────┴──────┐
       ▼             ▼
┌────────────┐ ┌────────────┐
│    XML     │ │    DOCX    │
│ ElementTree│ │python-docx │
└────────────┘ └────────────┘
```

Essa separação permite evoluir a lógica de processamento sem tornar a aplicação completamente dependente da GUI.

---

# Stack tecnológica

| Tecnologia      | Utilização                        |
| --------------- | --------------------------------- |
| **Python**      | Linguagem principal               |
| **Tkinter**     | Interface desktop                 |
| **ElementTree** | Parsing de XML                    |
| **python-docx** | Manipulação e geração de DOCX     |
| **Pygments**    | Tokenização e syntax highlighting |
| **lxml**        | Processamento XML                 |
| **pytest**      | Testes automatizados              |
| **tkinterdnd2** | Drag & Drop opcional              |
| **JSON**        | Persistência local                |
| **logging**     | Logs e diagnóstico                |

---

# Estrutura do projeto

```text
AutomacaoArtefatos/
│
├── supra.py
│
├── Modelo de Artefato.docx
│
├── requirements.txt
├── requirements-dev.txt
│
├── tests/
│   ├── conftest.py
│   ├── test_comparacao.py
│   ├── test_extracao.py
│   ├── test_nome_arquivo.py
│   ├── test_propriedades_campos.py
│   │
│   └── fixtures/
│       ├── ...
│
└── .gitignore
```

A suíte de testes está organizada por responsabilidade, com fixtures específicas para reproduzir cenários de extração e comparação.

---

# Instalação

## Pré-requisitos

* Python 3.x
* Windows recomendado para utilização da aplicação desktop

Clone o projeto:

```bash
git clone https://github.com/gbrielzera/AutomacaoArtefatos.git
cd AutomacaoArtefatos
```

Crie um ambiente virtual:

```bash
python -m venv .venv
```

Ative no Windows:

```bash
.venv\Scripts\activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Para desenvolvimento:

```bash
pip install -r requirements-dev.txt
```

---

# Executando

```bash
python supra.py
```

A interface gráfica será iniciada.

A partir dela é possível selecionar os arquivos necessários e executar o processamento.

---

# Distribuição

A aplicação foi projetada para também poder ser distribuída como um executável desktop.

A experiência esperada para o usuário final é:

```text
        Usuário
           │
           ▼
      Aplicação
           │
     ┌─────┴─────┐
     ▼           ▼
   XML       Modelo DOCX
     │           │
     └─────┬─────┘
           ▼
    Processamento
           │
           ▼
   Artefato final
```

O objetivo é que o usuário final não precise conhecer a implementação interna nem possuir conhecimento de Python.

---

# Impacto

O projeto foi criado para automatizar uma tarefa que possui três características:

```text
Dados estruturados
        +
Processo repetitivo
        +
Necessidade de precisão
```

Esse é exatamente o tipo de cenário onde automação pode gerar grande valor.

A solução reduz a necessidade de:

* copiar informações manualmente;
* procurar elementos em estruturas complexas;
* comparar versões visualmente;
* copiar scripts;
* identificar alterações uma a uma;
* formatar código manualmente;
* revisar documentos extensos em busca de inconsistências.

Mais importante:

**o sistema transforma uma operação manual em uma operação determinística e reproduzível.**

---

# Evolução do projeto

O projeto começou com uma proposta simples:

```text
XML → DOCX
```

e evoluiu para:

```text
                 ┌───────────────┐
                 │ Versão antiga │
                 └───────┬───────┘
                         │
                         ▼
                      Parser
                         │
                         ▼
                   Dados estruturados
                         │
                         │
                         ▼
                    COMPARAÇÃO
                         ▲
                         │
                   Dados estruturados
                         ▲
                         │
                      Parser
                         ▲
                         │
                 ┌───────┴───────┐
                 │ Versão atual  │
                 └───────────────┘
                         │
                         ▼
                 Diferenças detectadas
                         │
                         ▼
                   Documento DOCX
```

Essa evolução mudou o papel da aplicação:

> **de um simples gerador de documentos para uma ferramenta de análise e documentação de mudanças em processos.**

---

# Próximos passos

Algumas evoluções naturais do projeto incluem:

* interface dedicada para visualização das diferenças;
* diff visual entre scripts;
* comparação de múltiplas versões;
* suporte a diferentes modelos de documentação;
* validação automática do DOCX gerado;
* relatórios em outros formatos;
* histórico de artefatos;
* maior cobertura de testes;
* pipeline automatizado de build;
* distribuição simplificada do executável;
* separação ainda maior entre domínio, infraestrutura e apresentação.

---

# Sobre

Este projeto demonstra a aplicação prática de conceitos de **Engenharia de Software e Automação** para transformar uma atividade manual em um processo sistematizado.

Entre os conceitos aplicados estão:

* processamento de dados estruturados;
* parsing de XML;
* comparação de versões;
* análise contextual;
* geração automatizada de documentos;
* syntax highlighting;
* testes automatizados;
* tratamento de exceções;
* logging;
* persistência local;
* desenvolvimento desktop;
* preocupação com experiência do usuário.

A premissa é simples:

> ## **Automação não é apenas fazer uma tarefa mais rápido.**
>
> ## **É transformar uma tarefa repetitiva em um processo confiável, reproduzível e escalável.**

---

## Licença

Consulte o repositório para informações sobre licenciamento e uso do projeto.
