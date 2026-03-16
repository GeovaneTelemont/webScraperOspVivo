# 🕸️ OSP Vivo Web Scraper

Uma aplicação de desktop com interface gráfica (GUI) para automatizar a extração de dados do portal OSP Vivo Control. A ferramenta utiliza Playwright para automação de navegador e PyQt6 para a interface, facilitando o processo de coleta de informações de múltiplos IDs.

## ✨ Recursos Principais

- **Interface Gráfica Amigável**: Construída com PyQt6 para uma experiência de usuário simples e intuitiva.
- **Múltiplos Modos de Extração**: Suporte para diferentes tipos de extração de dados:
  - 🧾 Draft
  - 📊 Medição
  - 🔎 ID Cancelados
  - 🧮 Memória de Cálculo
- **Gerenciamento de Sessão**: Salva e carrega o estado de login (`auth.json`) para evitar a necessidade de fazer login manualmente a cada execução.
- **Instalação Automatizada**: O Playwright e seus navegadores necessários são instalados automaticamente na primeira execução, caso não existam.
- **Entrada e Saída Flexíveis**:
  - Utiliza um arquivo `.csv` como entrada para a lista de IDs.
  - Salva os dados extraídos em arquivos `.xlsx` (Excel) na pasta de Downloads do usuário.
- **Feedback em Tempo Real**: Exibe logs detalhados, uma barra de progresso e mensagens de status durante a execução.
- **Gerenciamento de Credenciais**: Permite salvar o usuário e senha em um arquivo de configuração local (`config.json`) para preenchimento automático.

## 🔧 Pré-requisitos

Antes de começar, certifique-se de ter o seguinte instalado:

- **Python 3.9+**
- **Poetry** (para gerenciamento de dependências)

## 🚀 Instalação

1.  **Clone o repositório** (se estiver em um controle de versão como o Git):
    ```bash
    git clone <url-do-seu-repositorio>
    cd webscraper_osp
    ```

2.  **Instale as dependências** usando o Poetry:
    ```bash
    poetry install
    ```
    O Poetry criará um ambiente virtual e instalará todas as bibliotecas listadas no arquivo `poetry.lock`.

3.  **Instalação do Playwright**:
    Não é necessário fazer nada! O script verificará se o Playwright está instalado e, caso não esteja, fará a instalação e o download dos navegadores necessários automaticamente.

## 💻 Como Usar

1.  **Execute a aplicação** através do Poetry:
    ```bash
    poetry run python app/webScraperOsp.py
    ```

2.  **Configurações Iniciais**:
    - **Credenciais**: (Opcional) Insira seu usuário e senha do OSP Control e clique em `💾 Salvar Credenciais`. Isso criará um arquivo `config.json` para pré-preencher esses campos nas próximas vezes que usar a aplicação.
    - **Login**: Na primeira execução (ou se a sessão expirar), o navegador será aberto para que você faça o login manualmente, incluindo o preenchimento do CAPTCHA. Após o login, a sessão será salva no arquivo `auth.json`, permitindo que as próximas execuções sejam automáticas.

3.  **Selecione o Arquivo de Entrada**:
    - Clique em `📂 Selecionar CSV` e escolha o arquivo que contém os IDs a serem processados.
    - **Importante**: O arquivo CSV **deve** conter uma coluna chamada `ID`.

4.  **Escolha o Modo de Extração**:
    - Selecione uma das quatro opções disponíveis na seção `🎯 Modo de Extração`.

5.  **Inicie o Scraping**:
    - Clique no botão `🚀 Iniciar Scraping`.
    - Acompanhe o andamento na barra de progresso e veja os detalhes do que está acontecendo na área de `📝 Logs`.
    - Para interromper o processo a qualquer momento, clique em `⏹️ Parar`.

6.  **Arquivos de Saída**:
    - Os dados extraídos serão salvos em um arquivo `.xlsx` na sua pasta de **Downloads**. O nome do arquivo será padronizado de acordo com o modo de extração (ex: `osp_vivo_medicao.xlsx`).

## 🎯 Modos de Extração

- **🧾 Draft**: Extrai as tabelas de materiais e serviços da aba "Draft" para cada ID fornecido.
- **📊 Medição**: Extrai as tabelas de materiais e serviços da aba "Medição" para cada ID.
- **🔎 ID Cancelados**: Busca informações básicas como "Contrato" e "OSP" para os IDs informados. Útil para verificar status.
- **🧮 Memória de Cálculo**: Um modo de extração avançado que:
  - Navega até a seção "Memória de Cálculo" de cada serviço associado a um ID.
  - Percorre **todas as páginas** da tabela de memória de cálculo.
  - Extrai uma tabela detalhada contendo colunas como Classe, Código, Descrição, Pontos, Custos e Quantidades.

## 📁 Arquivos de Configuração

A aplicação gera dois arquivos no diretório raiz para facilitar o uso:

- `config.json`: Armazena o nome de usuário e a senha (em texto plano) quando você clica em "Salvar Credenciais".
- `auth.json`: Armazena o estado da sessão do navegador (cookies, local storage, etc.). É este arquivo que permite pular a tela de login nas execuções futuras. **Não compartilhe este arquivo**, pois ele contém sua sessão de login ativa.

---

*Este projeto foi desenvolvido para otimizar e automatizar tarefas repetitivas de extração de dados, aumentando a produtividade e reduzindo erros manuais.*