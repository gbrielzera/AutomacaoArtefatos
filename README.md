# Automação de Artefatos

> **Do dado estruturado ao documento pronto — sem copiar e colar.**

Aplicação desktop desenvolvida em **Python** para automatizar a geração de documentos técnicos a partir de arquivos **XML de processos** e de um **modelo DOCX**.

O projeto nasceu de um problema real de produtividade: transformar informações técnicas espalhadas em uma estrutura complexa de processo em um documento padronizado, confiável e pronto para validação.

A proposta vai além de simplesmente preencher campos. A aplicação interpreta a estrutura do XML, identifica elementos relevantes do fluxo, preserva o contexto de scripts e expressões e gera um documento final com formatação adequada para documentação técnica.

---

## O problema

Documentar alterações em processos pode parecer simples até chegar a hora de fazer isso manualmente.

Em um fluxo complexo, é necessário localizar e transcrever informações como:

- nome do fluxo e processo;
- serviços associados ao fluxo principal;
- campos utilizados pelo processo;
- rótulos, tipos e tabelas de persistência;
- configurações de anexos;
- scripts executados em diferentes etapas;
- expressões de regras de negócio e decisões;
- localização exata de cada script dentro do fluxo.

Além de consumir tempo, esse tipo de atividade cria pontos de falha: informações podem ser esquecidas, duplicadas ou associadas ao elemento errado.

**A solução deste projeto é transformar essa atividade manual em um pipeline automatizado e rastreável.**

---

## O que a aplicação faz

A aplicação funciona como uma ponte entre o **XML exportado de um processo** e um **modelo de documentação em Word**.

### Fluxo de execução

```text
┌────────────────────┐
│ XML do processo    │
└─────────┬──────────┘
          │
          ▼
┌────────────────────┐
│ Parser / Extração  │
│ de dados           │
└─────────┬──────────┘
          │
          ├── Fluxo
          ├── Serviços
          ├── Campos
          ├── Anexos
          ├── Scripts
          └── Expressões
          │
          ▼
┌────────────────────┐
│ Modelo DOCX        │
└─────────┬──────────┘
          │
          ▼
┌────────────────────┐
│ Documento final    │
│ padronizado        │
└────────────────────┘
```

O usuário fornece apenas as informações que dependem de contexto humano, enquanto os dados que já existem no XML são extraídos e estruturados automaticamente.

---

## Destaques técnicos

### Parser de XML resiliente

A extração utiliza `xml.etree.ElementTree` e foi estruturada para lidar com diferentes situações encontradas nos arquivos exportados.

Entre os cuidados implementados estão:

- leitura da árvore XML;
- fallback para conteúdo com BOM/encoding compatível;
- identificação do processo a partir do nome do arquivo;
- extração de propriedades de campos;
- deduplicação de serviços;
- deduplicação contextual de scripts;
- identificação de scripts em diferentes níveis da árvore;
- tratamento de atividades, operações e gateways;
- geração de logs durante a extração.

Um detalhe importante é a **extração contextual**: scripts iguais não são simplesmente descartados.

A aplicação considera o local em que o script aparece, evitando perder informações relevantes quando o mesmo código é utilizado em pontos diferentes do fluxo.

---

### Extração de scripts e regras de negócio

A aplicação identifica diferentes tipos de scripts, incluindo:

- scripts modificados;
- scripts de validação;
- scripts de formulário carregado;
- scripts de início e fim;
- scripts de volta e evento;
- scripts de seleção de atores;
- expressões de valor de inputs;
- expressões de decisão de gateways.

Cada ocorrência recebe uma descrição contextual, por exemplo:

```text
Atividade: LinkInicial | Script modificado no campo: FORNECEDOR1
```

ou:

```text
Gateway de Decisão: Validar solicitação | Expressão de Decisão
```

Isso permite que o documento responda não apenas:

> **"Qual é o código?"**

mas também:

> **"Onde esse código está sendo executado?"**

---

## Código formatado como código

Um dos pontos centrais do projeto é a documentação dos scripts.

Em vez de inserir o conteúdo como texto puro no Word, o projeto utiliza **Pygments** para realizar a tokenização do código Python e reproduzir uma apresentação semelhante à encontrada em uma IDE.

O resultado busca preservar:

- indentação;
- quebras de linha;
- estrutura visual do código;
- fonte monoespaçada;
- destaque sintático;
- contexto de execução.

Isso torna documentos extensos de scripts significativamente mais legíveis e úteis para revisão técnica.

---

## Interface

A aplicação possui uma interface desktop construída com **Tkinter**, mantendo o fluxo de uso simples.

O usuário pode:

1. informar ou confirmar o macroprocesso;
2. informar ou confirmar o processo;
3. descrever a alteração/criação;
4. informar a evidência/chamado;
5. selecionar o XML;
6. selecionar o modelo DOCX;
7. gerar o artefato.

Também existe suporte opcional a **arrastar e soltar arquivos**, quando `tkinterdnd2` está disponível.

---

## Tratamento de erros e observabilidade

Automação boa não é só aquela que funciona quando tudo dá certo.

O projeto possui uma camada de logging para registrar o processamento e facilitar o diagnóstico de problemas.

Os logs podem registrar:

- início e fim da extração;
- quantidade de serviços, campos, anexos e scripts encontrados;
- elementos ignorados ou incompletos;
- arquivos processados;
- exceções não tratadas;
- informações de execução necessárias para investigação.

Em caso de falha inesperada, a aplicação registra um **log de crash** antes de apresentar a mensagem ao usuário.

Também existe persistência local de configurações/histórico de execução para reduzir trabalho repetitivo em utilizações posteriores.

---

## Stack

| Tecnologia | Utilização |
|---|---|
| **Python** | Linguagem principal |
| **Tkinter** | Interface gráfica desktop |
| **ElementTree** | Parsing e navegação do XML |
| **python-docx** | Manipulação e geração de DOCX |
| **Pygments** | Tokenização e destaque sintático dos scripts |
| **lxml** | Suporte ao processamento XML |
| **pytest** | Testes automatizados |
| **tkinterdnd2** | Drag & Drop opcional |
| **JSON** | Persistência de configurações |
| **logging** | Observabilidade e diagnóstico |

---

## Estrutura do projeto

```text
AutomacaoArtefatos/
│
├── supra.py                  # Aplicação e lógica principal
├── Modelo de Artefato.docx   # Modelo utilizado para geração
├── requirements.txt          # Dependências de execução
├── requirements-dev.txt      # Dependências de desenvolvimento/testes
├── tests/                    # Testes automatizados
└── .gitignore
```

---

## Instalação

### Pré-requisitos

- Python 3.x
- Windows recomendado para execução da aplicação desktop
- Um arquivo XML exportado do sistema de processos
- Um modelo DOCX compatível com a estrutura esperada pela aplicação

### 1. Clone o repositório

```bash
git clone https://github.com/gbrielzera/AutomacaoArtefatos.git
cd AutomacaoArtefatos
```

### 2. Crie um ambiente virtual

```bash
python -m venv .venv
```

No Windows:

```bash
.venv\Scripts\activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

Para desenvolvimento e testes:

```bash
pip install -r requirements-dev.txt
```

---

## Execução

```bash
python supra.py
```

A interface gráfica será aberta e poderá ser utilizada para selecionar os arquivos e preencher os dados complementares.

---

## Testes

As funções responsáveis pela extração foram projetadas para serem testáveis independentemente da interface gráfica.

Para executar os testes:

```bash
pytest
```

---

## Distribuição

A arquitetura foi pensada para permitir a distribuição como aplicação desktop, sem exigir que o usuário final interaja diretamente com o código-fonte Python.

O objetivo é disponibilizar uma experiência no formato:

```text
Usuário
   │
   ▼
Executável
   │
   ├── XML
   └── Modelo DOCX
          │
          ▼
    Documento final
```

---

## Decisões de engenharia

### Separação entre interface e processamento

A lógica de extração do XML foi mantida em funções independentes da GUI.

Isso facilita:

- testes;
- manutenção;
- debugging;
- reutilização da lógica;
- evolução do parser sem acoplamento excessivo à interface.

---

### Deduplicação contextual

A aplicação não utiliza apenas o código do script como chave de deduplicação.

O **contexto e o local de execução** também são considerados.

Isso permite preservar ocorrências legítimas do mesmo código em pontos diferentes do processo.

---

### Foco em dados autoritativos

Quando existem múltiplas representações de uma informação dentro do XML, a extração prioriza a estrutura considerada mais confiável para aquele dado e utiliza alternativas apenas como fallback.

Isso reduz o risco de documentar informações provenientes de elementos secundários do processo.

---

### Tolerância a falhas

Falhas em configurações locais ou em recursos auxiliares, como Drag & Drop, não devem impedir a execução principal quando não forem essenciais ao processamento.

A aplicação também possui tratamento específico para exceções inesperadas, evitando que erros ocorram de forma silenciosa.

---

## Impacto

O projeto foi concebido para atacar uma classe comum de problemas em ambientes corporativos:

> **tarefas operacionais repetitivas que exigem atenção humana, mas que trabalham sobre dados que já estão estruturados.**

A automação reduz a necessidade de:

- navegar manualmente por estruturas complexas;
- copiar grandes blocos de código;
- localizar scripts individualmente;
- preencher repetidamente informações deriváveis;
- revisar documentos procurando inconsistências de transcrição.

Mais do que automatizar um documento, o projeto demonstra uma abordagem de **engenharia de automação orientada a dados**, combinando:

```text
Parsing
   +
Transformação
   +
Validação
   +
Geração de documentos
   +
Observabilidade
   +
Experiência do usuário
```

---

## Sobre o projeto

Este projeto representa a aplicação prática de conceitos de:

- **Engenharia de Software**
- **Python**
- **Automação de processos**
- **Processamento de dados estruturados**
- **Parsing de XML**
- **Geração e manipulação de documentos**
- **Testes automatizados**
- **Tratamento de exceções**
- **Observabilidade**
- **Desenvolvimento de aplicações desktop**

A ideia central é simples:

> ### **Se a informação já existe em formato estruturado, o usuário não deveria precisar digitá-la novamente.**

---

## 📄 Licença

Consulte o repositório para informações sobre licenciamento e uso do projeto.