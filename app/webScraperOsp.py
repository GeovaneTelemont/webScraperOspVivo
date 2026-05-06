import json
import os
import subprocess
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from time import sleep

import pandas as pd
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *

# Verifica se Playwright está instalado
try:
    from playwright.sync_api import TimeoutError, sync_playwright
except ImportError:
    print("📦 Playwright não encontrado. Instalando automaticamente...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "playwright"])
    print("✅ Playwright instalado. Instalando navegadores...")
    subprocess.check_call([sys.executable, "-m", "playwright", "install"])

# ===========================================================
# 📁 MÓDULO DE WEBSCRAPING
# ===========================================================


class WebScraperWorker(QThread):
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    data_saved = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.csv_path = ""
        self.mode = 1  # 1=Draft, 2=Medição, 3=ID Cancelados, 4=Memória de Cálculo
        self.auth_file = "auth.json"
        self.login_url = "https://devopsredes.vivo.com.br/ospcontrol/home"
        self.logged_selector = 'xpath=//*[@id="ott-username"]'
        self.username = ""
        self.password = ""
        self.download_path = Path.home() / "Downloads"
        self.headless_enabled = False
        self._running = True

    def _normalize_text(self, s: str) -> str:
        """Normaliza texto para comparação"""
        if not s:
            return ""
        s = " ".join(s.split())
        s_norm = (
            unicodedata.normalize("NFKD", s)
            .encode("ASCII", "ignore")
            .decode("ASCII")
            .lower()
        )
        return s_norm

    def _determinar_tipo_registro(self, categoria, unidade="", descricao=""):
        """Determina o tipo de registro baseado na categoria, unidade e descrição"""
        cat_norm = self._normalize_text(categoria)

        if any(x in cat_norm for x in ("material", "materiais", "telefonica")):
            return "Material"
        elif any(x in cat_norm for x in ("custo", "custos")):
            return "Custo"
        elif any(x in cat_norm for x in ("servico", "servicos", "classe", "valor")):
            return "Serviço"
        else:
            unidade_lower = unidade.lower() if unidade else ""
            descricao_lower = descricao.lower() if descricao else ""

            if any(k in unidade_lower for k in ("m", "u", "cj", "un", "metro", "kg")):
                return "Material"
            elif any(
                k in descricao_lower
                for k in (
                    "cfo",
                    "chassi",
                    "conj",
                    "subduto",
                    "material",
                    "cabos",
                    "fibra",
                )
            ):
                return "Material"
            else:
                return "Serviço"

    def _extrair_categoria_tabela(self, tabela):
        """Extrai a categoria da tabela"""
        try:
            categoria = tabela.evaluate("""
                el => {
                    function findPrevText(e){
                        let node = e.previousElementSibling;
                        while(node){
                            const txt = node.innerText ? node.innerText.trim() : '';
                            if(txt) return txt.replace(/\\s+/g,' ');
                            node = node.previousElementSibling;
                        }
                        let parent = e.parentElement;
                        while(parent){
                            let prev = parent.previousElementSibling;
                            while(prev){
                                const txt = prev.innerText ? prev.innerText.trim() : '';
                                if(txt) return txt.replace(/\\s+/g,' ');
                                prev = prev.previousElementSibling;
                            }
                            parent = parent.parentElement;
                        }
                        return '';
                    }
                    return findPrevText(el);
                }
            """)

            if not isinstance(categoria, str):
                categoria = "" if categoria is None else str(categoria)

            categoria = " ".join(categoria.split())
            return categoria

        except Exception as e:
            self.message.emit(f"⚠️ Erro ao extrair categoria: {e}")
            return ""

    def _extrair_status_id(self, page, id_value):
        """Extrai o status do ID na tela de busca, antes de editar."""
        try:
            # Espera a tabela de resultados aparecer
            # Usa um seletor mais genérico para garantir que encontre a tabela mesmo se o ID do painel mudar
            page.wait_for_selector("table tbody tr", timeout=8000)

            # Tenta localizar a tabela correta (prefere a visível/ativa)
            tabela = page.locator("table").first
            if page.locator(".tab-pane.active table").count() > 0:
                tabela = page.locator(".tab-pane.active table").first

            # Encontra o índice da coluna "Status" pelo cabeçalho para ser mais robusto
            headers = tabela.locator("thead th").all()
            status_index = -1
            for i, header in enumerate(headers):
                header_text = self._normalize_text(header.text_content())
                if "status" in header_text or "situacao" in header_text:
                    status_index = i
                    break

            if status_index == -1:
                # Fallback para o método antigo se o cabeçalho "Status" não for encontrado.
                self.message.emit(
                    f"⚠️ ID {id_value}: Cabeçalho 'Status' não encontrado, tentando coluna 13."
                )
                status_locator = tabela.locator("tbody tr:first-child td:nth-child(13)")
            else:
                # Usa o índice da coluna "Status" para pegar o dado correto.
                col_index = status_index + 1
                status_locator = tabela.locator(
                    f"tbody tr:first-child td:nth-child({col_index})"
                )

            if status_locator.count() > 0:
                status = status_locator.text_content().strip()
            else:
                status = "CÉLULA VAZIA/NÃO ENCONTRADA"

            self.message.emit(f"ℹ️ ID {id_value}: Status encontrado: '{status}'")
            return status
        except Exception as e:
            self.message.emit(f"⚠️ Erro ao ler status para ID {id_value}: {e}")
            return "STATUS NÃO ENCONTRADO"

    def _recover_page_state(self, page):
        """Tenta recuperar o estado da página em caso de erro"""
        try:
            self.message.emit("🔄 Reiniciando a página para uma nova tentativa...")
            page.goto(self.login_url, wait_until="networkidle", timeout=60000)
        except Exception as e:
            self.message.emit(f"⚠️ Falha na recuperação da página: {str(e)}")

    def _scrap_memoria_calculo(self, page, id_value):
        """Extrai todos os dados da memória de cálculo com paginação"""
        # 1. Navegação inicial
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        page.wait_for_selector('xpath=//*[@id="filtroId"]', timeout=10000)
        page.fill('xpath=//*[@id="filtroId"]', str(id_value))
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click(
            timeout=10000
        )

        # Extrai o status antes de clicar em editar
        status = self._extrair_status_id(page, id_value)

        try:
            # Tenta clicar no botão "Editar"
            page.locator("span.badge.bg-primary:has-text('Editar')").click(
                timeout=10000
            )
            page.wait_for_selector('a[title="Serviços"]', timeout=15000)
        except TimeoutError:
            # Se o botão não for encontrado, registra o status e retorna
            self.message.emit(
                f"⚠️ ID {id_value}: Botão 'Editar' não encontrado. Status: '{status}'."
            )
            # Volta ao menu principal para o próximo ID
            try:
                page.click('//*[@id="ott-sidebar-collapse"]')
                page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception as nav_error:
                self.message.emit(
                    f"⚠️ Erro ao voltar ao menu para ID {id_value}: {nav_error}"
                )
            return [[id_value, "", "", "", "", "", "", "", "", "", status]]

        # 2. Clica em todos os serviços
        serviços = page.locator('a[title="Serviços"]').all()
        todos_dados = []

        for servico_idx, servico in enumerate(serviços):
            servico.click()

            # 3. Tenta encontrar "Memória de Cálculo"
            try:
                page.locator('//a[text()="Memória de Cálculo"]').click(timeout=15000)
                # esperar tabela carregar antes de continuar
                page.wait_for_timeout(2000)

                page.wait_for_selector(
                    "table.ott-table-sm.ott-table-nowrap", timeout=15000
                )
                self.message.emit(
                    f"✅ ID {id_value}: Serviço {servico_idx + 1} - Memória de Cálculo"
                )
            except:
                self.message.emit(
                    f"⚠️ ID {id_value}: Serviço {servico_idx + 1} - Não encontrou Memória de Cálculo"
                )
                if len(serviços) > 1:
                    page.go_back()
                    page.wait_for_load_state("networkidle", timeout=15000)
                continue

            # 4. Tenta mostrar 50 itens
            try:
                page.select_option(
                    "select.custom-select", label="50 itens", timeout=5000
                )
                page.wait_for_load_state("networkidle", timeout=15000)
                self.message.emit(f"📊 ID {id_value}: Selecionou '50 itens'")
            except:
                self.message.emit(f"ℹ️ ID {id_value}: Sem select ou já está em 50 itens")

            # 5. EXTRAI TODAS AS PÁGINAS
            pagina = 1
            while True:
                self.message.emit(f"   📄 ID {id_value}: Página {pagina}")

                # Procura a tabela específica
                tabela = page.locator("table.ott-table-sm.ott-table-nowrap").first

                if tabela.count() > 0:
                    # Extrai todas as linhas da tabela atual
                    linhas = tabela.locator("tbody tr").all()

                    for linha in linhas:
                        # Pega todas as células da linha
                        celulas = linha.locator("td").all()

                        # Verifica se tem pelo menos 9 colunas
                        if len(celulas) >= 9:
                            # Extrai texto de cada célula
                            valores = []
                            for celula in celulas:
                                texto = celula.text_content().strip()
                                # Remove R$ dos valores monetários
                                if "R$" in texto:
                                    texto = texto.replace("R$", "").strip()
                                valores.append(texto)

                            # Garante que temos exatamente 9 valores
                            while len(valores) < 9:
                                valores.append("")

                            # Monta a linha completa: ID + 9 valores da tabela
                            linha_completa = [id_value] + valores + [status]
                            todos_dados.append(linha_completa)

                    self.message.emit(
                        f"   ✓ Extraiu {len(linhas)} linhas da página {pagina}"
                    )
                else:
                    self.message.emit(f"   ⚠️ Tabela não encontrada na página {pagina}")
                    # Tenta buscar por outra classe
                    try:
                        tabela_alt = page.locator("table.table-bordered").first
                        if tabela_alt.count() > 0:
                            linhas = tabela_alt.locator("tbody tr").all()
                            for linha in linhas:
                                celulas = linha.locator("td").all()
                                if len(celulas) >= 9:
                                    valores = [
                                        c.text_content().strip().replace("R$", "")
                                        for c in celulas
                                    ]
                                    while len(valores) < 9:
                                        valores.append("")
                                    todos_dados.append([id_value] + valores + [status])
                            self.message.emit(
                                f"   ✓ Extraiu {len(linhas)} linhas (tabela alternativa)"
                            )
                    except:
                        pass

                # 6. VERIFICA SE TEM PRÓXIMA PÁGINA
                tem_proxima = False
                try:
                    next_btns = page.locator(
                        '//li[not(contains(@class, "disabled"))]//a[@aria-label="Next"]'
                    )

                    for i in range(next_btns.count()):
                        btn = next_btns.nth(i)
                        if btn.is_visible():
                            tem_proxima = True
                            break
                except:
                    tem_proxima = False

                # 7. SE NÃO TEM PRÓXIMA PÁGINA, PARA
                if not tem_proxima:
                    self.message.emit(f"   🏁 Última página ({pagina})")
                    break

                # 8. VAI PARA PRÓXIMA PÁGINA
                try:
                    page.locator(
                        '//li[not(contains(@class, "disabled"))]//a[@aria-label="Next"]'
                    ).first.click()
                    page.wait_for_load_state("networkidle", timeout=15000)
                    pagina += 1
                except Exception as e:
                    self.message.emit(f"   ❌ Erro ao mudar página: {str(e)}")
                    break

            # 9. VOLTA PARA LISTA DE SERVIÇOS (se houver mais de um)
            if len(serviços) > 1 and servico_idx < len(serviços) - 1:
                page.go_back()
                page.wait_for_load_state("networkidle", timeout=15000)

        # 10. VOLTA AO MENU PRINCIPAL
        page.click('//*[@id="ott-sidebar-collapse"]')
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        page.wait_for_load_state("networkidle", timeout=15000)

        self.message.emit(
            f"✅ ID {id_value}: Finalizado - {len(todos_dados)} linhas extraídas"
        )
        return todos_dados

    def run(self):
        try:
            self.message.emit("🔄 Iniciando web scraping...")
            self._running = True

            if not self.csv_path or not os.path.exists(self.csv_path):
                self.error.emit("❌ Arquivo CSV não encontrado!")
                return

            # Ler CSV
            self.message.emit("📋 Lendo arquivo CSV...")
            df = pd.read_csv(self.csv_path, sep=";", encoding="utf-8")

            # Executar no thread separado usando playwright
            self._run_with_playwright(df)

        except Exception as e:
            self.error.emit(f"❌ Erro no processo: {str(e)}")
        finally:
            self.finished.emit()

    def _run_with_playwright(self, df):
        try:
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                # Lógica de Headless Inteligente
                # Se estiver configurado para headless, mas não tiver arquivo de sessão,
                # forçamos o modo visível para permitir o login.
                start_headless = self.headless_enabled
                if start_headless and not os.path.exists(self.auth_file):
                    self.message.emit(
                        "ℹ️ Modo Headless ativo, mas sem sessão salva. Iniciando visível para login..."
                    )
                    start_headless = False

                self.message.emit(
                    f"🌐 Iniciando navegador (Headless: {'Sim' if start_headless else 'Não'})..."
                )

                browser = p.chromium.launch(
                    channel="chrome",
                    headless=start_headless,
                    args=["--ignore-certificate-errors"],
                )

                # Configurações para garantir que o site carregue igual ao modo normal
                # (Tamanho de tela Full HD e User Agent de navegador real)
                user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                viewport_config = {"width": 800, "height": 600}

                # Verificar se há sessão salva
                if os.path.exists(self.auth_file):
                    self.message.emit("🔑 Carregando sessão existente...")
                    context = browser.new_context(
                        storage_state=self.auth_file,
                        user_agent=user_agent,
                        viewport=viewport_config,
                    )
                else:
                    self.message.emit("🆕 Criando nova sessão...")
                    context = browser.new_context(
                        user_agent=user_agent, viewport=viewport_config
                    )

                page = context.new_page()

                # Loop infinito de tentativa de carregamento da página
                while self._running:
                    try:
                        self.message.emit(f"🌐 Acessando {self.login_url}...")
                        page.goto(
                            self.login_url, wait_until="networkidle", timeout=60000
                        )
                        self.message.emit("✅ Página carregada.")
                        break
                    except Exception as e:
                        self.message.emit(
                            f"⚠️ Erro ao carregar página: {str(e)[:100]}... Tentando novamente em 5s."
                        )
                        sleep(5)

                # Verificar login
                if not self._is_logged(page):
                    self.message.emit("🔐 Sessão expirada. Aguardando login manual...")

                    # SE estivermos em modo Headless e o login falhou (sessão expirada),
                    # precisamos reiniciar o navegador em modo VISÍVEL.
                    if start_headless:
                        self.message.emit(
                            "⚠️ Login necessário! Reiniciando navegador em modo visível..."
                        )
                        context.close()
                        browser.close()

                        # Reinicia visível
                        start_headless = False
                        browser = p.chromium.launch(
                            channel="chrome",
                            headless=False,
                            args=["--ignore-certificate-errors"],
                        )

                        # Recria contexto e página
                        self.message.emit("🆕 Criando nova sessão para login...")
                        context = browser.new_context(
                            user_agent=user_agent, viewport=viewport_config
                        )
                        page = context.new_page()
                        page.goto(
                            self.login_url, wait_until="networkidle", timeout=60000
                        )

                    # Injeta a mensagem inicial imediatamente se houver credenciais
                    if self.username and self.password:
                        page.evaluate("""
                            () => {
                                let div = document.getElementById('scraping-msg-overlay');
                                if (!div) {
                                    div = document.createElement('div');
                                    div.id = 'scraping-msg-overlay';
                                    div.style.position = 'fixed';
                                    div.style.top = '0';
                                    div.style.left = '0';
                                    div.style.width = '100%';
                                    div.style.textAlign = 'center';
                                    div.style.zIndex = '99999';
                                    div.style.padding = '15px';
                                    div.style.fontSize = '18px';
                                    div.style.fontWeight = 'bold';
                                    document.body.prepend(div);
                                }
                                div.style.backgroundColor = '#fff3cd';
                                div.style.color = '#856404';
                                div.style.borderBottom = '2px solid #ffeeba';
                                div.innerText = '🤖 O Robô irá preencher as credenciais... Por favor, aguarde!';
                            }
                        """)

                    # Loop até detectar login ou cancelamento
                    while self._running:
                        # 1. Verifica sucesso do login (timeout curto para não bloquear)
                        if self._is_logged(page, timeout=2000):
                            self.message.emit("✅ Login detectado! Salvando sessão...")
                            context.storage_state(path=self.auth_file)
                            break

                        # 2. Verifica erro de acesso inválido
                        try:
                            if page.locator(
                                "p.msg", has_text="Acesso inválido"
                            ).is_visible(timeout=500):
                                self.message.emit(
                                    "❌ Login inválido detectado. Recarregando página..."
                                )
                                page.goto(
                                    self.login_url,
                                    wait_until="networkidle",
                                    timeout=60000,
                                )
                                continue
                        except Exception:
                            pass

                        # 3. Se tiver credenciais, verifica se precisa preencher
                        if self.username and self.password:
                            try:
                                # Verifica se estamos na página de login (campo user visível)
                                if page.locator('//*[@id="username"]').is_visible():
                                    val_user = page.input_value('//*[@id="username"]')
                                    val_pass = page.input_value('//*[@id="password"]')

                                    # Se campos vazios (primeira vez ou reset pós-erro), preenche
                                    if not val_user and not val_pass:
                                        self.message.emit(
                                            "ℹ️ Preenchendo credenciais..."
                                        )

                                        # Re-injetar mensagem (garante que apareça após reload por erro)
                                        page.evaluate("""
                                            () => {
                                                let div = document.getElementById('scraping-msg-overlay');
                                                if (!div) {
                                                    div = document.createElement('div');
                                                    div.id = 'scraping-msg-overlay';
                                                    div.style.position = 'fixed';
                                                    div.style.top = '0';
                                                    div.style.left = '0';
                                                    div.style.width = '100%';
                                                    div.style.textAlign = 'center';
                                                    div.style.zIndex = '99999';
                                                    div.style.padding = '15px';
                                                    div.style.fontSize = '18px';
                                                    div.style.fontWeight = 'bold';
                                                    document.body.prepend(div);
                                                }
                                                div.style.backgroundColor = '#fff3cd';
                                                div.style.color = '#856404';
                                                div.style.borderBottom = '2px solid #ffeeba';
                                                div.innerText = '🤖 O Robô irá preencher as credenciais... Por favor, aguarde!';
                                            }
                                        """)

                                        # Delay para leitura da mensagem
                                        sleep(3)

                                        # Digitação humanizada
                                        page.click('//*[@id="username"]')
                                        page.fill('//*[@id="username"]', "")
                                        page.keyboard.type(self.username, delay=100)

                                        page.click('//*[@id="password"]')
                                        page.fill('//*[@id="password"]', "")
                                        page.keyboard.type(self.password, delay=100)

                                        # Atualizar mensagem na web
                                        page.evaluate("""
                                            () => {
                                                let div = document.getElementById('scraping-msg-overlay');
                                                if(div) {
                                                    div.style.backgroundColor = '#d4edda';
                                                    div.style.color = '#155724';
                                                    div.style.borderBottom = '2px solid #c3e6cb';
                                                    div.innerText = '⚠️ AÇÃO NECESSÁRIA: Preencha o CAPTCHA e clique em Entrar!';
                                                }
                                            }
                                        """)

                                        self.message.emit("✅ Credenciais preenchidas.")
                            except Exception:
                                pass

                        sleep(1)

                    # Se configurado para Headless e estava visível para login, reinicia em Headless após sucesso
                    if self.headless_enabled and not start_headless and self._running:
                        self.message.emit(
                            "🔄 Login concluído. Alternando para modo Headless..."
                        )
                        context.close()
                        browser.close()

                        start_headless = True
                        browser = p.chromium.launch(
                            channel="chrome",
                            headless=True,
                            args=["--ignore-certificate-errors"],
                        )
                        context = browser.new_context(
                            storage_state=self.auth_file,
                            user_agent=user_agent,
                            viewport=viewport_config,
                        )
                        page = context.new_page()
                        page.goto(
                            self.login_url, wait_until="networkidle", timeout=60000
                        )
                else:
                    self.message.emit("✅ Já está logado!")

                # Executar scraping baseado no modo
                if self.mode == 1:
                    self._scrap_draft(page, df)
                elif self.mode == 2:
                    self._scrap_medicao(page, df)
                elif self.mode == 3:
                    self._scrap_id_cancelado(page, df)
                elif self.mode == 4:
                    self._scrap_memoria_calculo_main(page, df)

                browser.close()
                self.message.emit("✅ Processo concluído!")

        except Exception as e:
            self.error.emit(f"❌ Erro no playwright: {str(e)}")

    def _is_logged(self, page, timeout=5000):
        try:
            page.wait_for_selector(self.logged_selector, timeout=timeout)
            return True
        except:
            return False

    def _scrap_draft(self, page, df):
        """Função de extração de Draft"""
        self.message.emit("🧾 Iniciando extração de Draft...")
        timestamp = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
        arquivo = self.download_path / f"osp_vivo_draft_{timestamp}.xlsx"
        colunas = [
            "ID",
            "TIPO DE REGISTRO",
            "CÓDIGO",
            "DESCRIÇÃO",
            "QUANTIDADE",
            "PREÇO UNITÁRIO",
            "UNIDADE",
            "PREÇO TOTAL",
            "CATEGORIA",
            "STATUS",
        ]
        resultados = []

        total_ids = len(df)
        for idx, row in df.iterrows():
            if not self._running:
                break

            id_value = int(row["ID"])
            progress = int((idx + 1) / total_ids * 100)
            self.progress.emit(progress)
            self.message.emit(
                f"📋 Processando ID {id_value} ({idx + 1}/{total_ids})..."
            )

            while self._running:
                try:
                    dados = self._pesquisar_id_draft(page, id_value)
                    if dados:
                        resultados.extend(dados)
                        self.message.emit(
                            f"✅ ID {id_value} extraído ({len(dados)} linhas)"
                        )
                    else:
                        self.message.emit(
                            f"⚠️ Nenhum dado encontrado para ID {id_value}"
                        )
                    break  # Sucesso, sai do loop de retry
                except Exception as e:
                    self.message.emit(
                        f"❌ Erro no ID {id_value}: {str(e)}. Tentando novamente..."
                    )
                    self._recover_page_state(page)

            # Salvar incremental
            try:
                df_parcial = pd.DataFrame(resultados, columns=colunas)
                df_parcial.to_excel(arquivo, index=False)
                self.data_saved.emit(str(arquivo))
            except Exception as e:
                self.message.emit(f"⚠️ Erro ao salvar: {e}")

        self.message.emit(f"✅ Arquivo final salvo: {arquivo}")

    def _scrap_medicao(self, page, df):
        """Função de extração de Medição"""
        self.message.emit("📊 Iniciando extração de Medição...")
        timestamp = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
        arquivo = self.download_path / f"osp_vivo_medicao_{timestamp}.xlsx"
        colunas = [
            "ID",
            "TIPO DE REGISTRO",
            "CÓDIGO",
            "DESCRIÇÃO",
            "QUANTIDADE",
            "PREÇO UNITÁRIO",
            "UNIDADE",
            "PREÇO TOTAL",
            "CATEGORIA",
            "STATUS",
        ]
        resultados = []

        total_ids = len(df)
        for idx, row in df.iterrows():
            if not self._running:
                break

            id_value = int(row["ID"])
            progress = int((idx + 1) / total_ids * 100)
            self.progress.emit(progress)
            self.message.emit(
                f"📋 Processando ID {id_value} ({idx + 1}/{total_ids})..."
            )

            while self._running:
                try:
                    dados = self._pesquisar_id_medicao(page, id_value)
                    if dados:
                        resultados.extend(dados)
                        self.message.emit(
                            f"✅ ID {id_value} extraído ({len(dados)} linhas)"
                        )
                    else:
                        self.message.emit(
                            f"⚠️ Nenhum dado encontrado para ID {id_value}"
                        )
                    break  # Sucesso
                except Exception as e:
                    self.message.emit(
                        f"❌ Erro no ID {id_value}: {str(e)}. Tentando novamente..."
                    )
                    self._recover_page_state(page)

            # Salvar incremental
            try:
                df_parcial = pd.DataFrame(resultados, columns=colunas)
                df_parcial.to_excel(arquivo, index=False)
                self.data_saved.emit(str(arquivo))
            except Exception as e:
                self.message.emit(f"⚠️ Erro ao salvar: {e}")

        self.message.emit(f"✅ Arquivo final salvo: {arquivo}")

    def _scrap_id_cancelado(self, page, df):
        """Função de extração de ID Cancelados"""
        self.message.emit("🔎 Iniciando extração de ID Cancelados...")
        timestamp = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
        arquivo = self.download_path / f"osp_id_cancelado_{timestamp}.xlsx"
        colunas = ["ID", "CONTRATO", "OSP", "STATUS"]
        resultados = []

        total_ids = len(df)
        for idx, row in df.iterrows():
            if not self._running:
                break

            id_value = int(row["ID"])
            progress = int((idx + 1) / total_ids * 100)
            self.progress.emit(progress)
            self.message.emit(
                f"📋 Processando ID {id_value} ({idx + 1}/{total_ids})..."
            )

            while self._running:
                try:
                    dados = self._pesquisar_id(page, id_value)
                    if dados:
                        resultados.extend(dados)
                        self.message.emit(
                            f"✅ ID {id_value}: Extraído ({len(dados)} registros)"
                        )
                    else:
                        resultados.append(
                            [id_value, "ERRO", "Nenhum dado retornado", ""]
                        )
                        self.message.emit(
                            f"⚠️ Nenhum dado encontrado para ID {id_value}"
                        )

                    break  # Sucesso
                except Exception as e:
                    # resultados.append([id_value, "ERRO", str(e), ""]) # Não salva erro, tenta de novo
                    self.message.emit(
                        f"❌ Erro no ID {id_value}: {str(e)}. Tentando novamente..."
                    )
                    self._recover_page_state(page)

            # Salvar incremental
            try:
                df_parcial = pd.DataFrame(resultados, columns=colunas)
                df_parcial.to_excel(arquivo, index=False)
                self.data_saved.emit(str(arquivo))
            except Exception as e:
                self.message.emit(f"⚠️ Erro ao salvar: {e}")

        self.message.emit(f"✅ Arquivo final salvo: {arquivo}")

    def _scrap_memoria_calculo_main(self, page, df):
        """Função principal que processa todos os IDs da memória de cálculo"""
        self.message.emit("🧮 Iniciando extração de Memória de Cálculo...")

        # Cria arquivo com timestamp
        timestamp = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
        arquivo = self.download_path / f"osp_memoria_calculo_{timestamp}.xlsx"

        # Define colunas
        colunas = [
            "ID",
            "CLASSE",
            "CODIGO",
            "DESCRIÇÃO DO SERVIÇO",
            "UNIDADE",
            "PONTOS",
            "CUSTO UNITÁRIO (R$)",
            "QUANTIDADE EXECUTADA",
            "PONTOS TOTAIS",
            "CUSTO TOTAL (R$)",
            "STATUS",
        ]

        # Lista para todos os resultados
        resultados = []
        total_ids = len(df)

        for idx, row in df.iterrows():
            if not self._running:
                break

            id_value = int(row["ID"])
            progress = int((idx + 1) / total_ids * 100)
            self.progress.emit(progress)
            self.message.emit(
                f"🔍 Processando ID {id_value} ({idx + 1}/{total_ids})..."
            )

            while self._running:
                try:
                    # Chama a função de extração
                    dados_id = self._scrap_memoria_calculo(page, id_value)

                    # Adiciona à lista principal
                    if dados_id:
                        resultados.extend(dados_id)
                        self.message.emit(
                            f"✅ ID {id_value}: Adicionou {len(dados_id)} linhas"
                        )
                    else:
                        # Adiciona linha vazia para manter o ID
                        resultados.append([id_value] + [""] * 10)
                        self.message.emit(f"⚠️ ID {id_value}: Nenhum dado encontrado")
                    break  # Sucesso
                except Exception as e:
                    self.message.emit(
                        f"❌ Erro no ID {id_value}: {str(e)}. Tentando novamente..."
                    )
                    self._recover_page_state(page)

            # Salva incrementalmente
            try:
                if resultados:
                    df_temp = pd.DataFrame(resultados, columns=colunas)
                    df_temp.to_excel(arquivo, index=False)
                    self.data_saved.emit(str(arquivo))
            except Exception as e:
                self.message.emit(f"⚠️ Erro ao salvar: {e}")

        # Finaliza
        self.message.emit(f"✅ Processo concluído! Total: {len(resultados)} linhas")
        self.message.emit(f"💾 Arquivo salvo: {arquivo}")

        return resultados

    def _pesquisar_id_cancelado(self, page, id_value):
        todos_dados = []
        status = ""
        # Navegação
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        page.wait_for_selector('xpath=//*[@id="filtroId"]', timeout=10000)
        page.fill('xpath=//*[@id="filtroId"]', str(id_value))
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()

        status = self._extrair_status_id(page, id_value)

        try:
            # Botão Editar com seletor mais robusto
            page.locator("span.badge.bg-primary", has_text="Editar").click(
                timeout=10000
            )

        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Botão 'Editar' não encontrado. Status: '{status}'."
            )
            return [[id_value, "", "", status]]

        try:
            page.wait_for_selector('a[title="Serviços"]', timeout=10000)

            # Loop por todos os serviços
            count_servicos = page.locator('a[title="Serviços"]').count()
        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Não encontrou link de 'Serviços'. Status: '{status}'."
            )
            return [[id_value, "", "", "Link de serviço não encontrado"]]

        for i in range(count_servicos):
            # Re-localiza o botão do serviço atual

            servicos_btn = page.locator('a[title="Serviços"]').nth(i)
            servicos_btn.click(timeout=10000)

            # Aguarda carregamento da rede para garantir que os campos carreguem
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except TimeoutError:
                pass
            page.wait_for_timeout(1000)

            page.wait_for_selector(
                "xpath=/html/body/app-root/app-requisicoes-servicos/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/span",
                timeout=15000,
            )

            # Extrai Contrato
            contrato_el = page.locator(
                "xpath=/html/body/app-root/app-requisicoes-servicos/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/span"
            )
            contrato = (
                contrato_el.text_content().strip() if contrato_el.count() > 0 else ""
            )

            # Extrai OSP
            osp_locator = page.locator(
                "xpath=/html/body/app-root/app-requisicoes-servicos/div/div/div/div/div[2]/div[3]/div/div[2]/div/strong"
            )
            osp = (
                osp_locator.text_content().strip()
                if osp_locator.count() > 0
                else "ATIVO"
            )

            todos_dados.append([id_value, contrato, osp, status])

            # Volta para a lista de serviços se houver mais de um
            if count_servicos > 1 and i < count_servicos - 1:
                page.go_back()
                page.wait_for_load_state("networkidle", timeout=15000)

        return todos_dados

    # Método alias para manter compatibilidade se chamado como _pesquisar_id
    def _pesquisar_id(self, page, id_value):
        return self._pesquisar_id_cancelado(page, id_value)

    def _pesquisar_id_draft(self, page, id_value):
        todos_dados = []
        status = ""
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        page.wait_for_selector('xpath=//*[@id="filtroId"]', timeout=10000)
        page.fill('xpath=//*[@id="filtroId"]', str(id_value))
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()

        status = self._extrair_status_id(page, id_value)

        try:
            # Botão Editar com seletor mais robusto
            page.locator("span.badge.bg-primary", has_text="Editar").first.click(
                timeout=10000
            )
        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Botão 'Editar' não encontrado. Status: '{status}'."
            )
            try:
                page.click('//*[@id="ott-sidebar-collapse"]')
                page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
                page.wait_for_load_state("networkidle", timeout=15000)
            except:
                pass
            return [[id_value, "", "", "", "", "", "", "", "", status]]

        try:
            page.wait_for_selector('a[title="Serviços"]', timeout=10000)

            # Loop por todos os serviços
            count_servicos = page.locator('a[title="Serviços"]').count()

        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Link de 'Serviços' não encontrado. Status: '{status}'."
            )
            return [
                [
                    id_value,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "Link de serviço não encontrado",
                ]
            ]

        for i in range(count_servicos):
            # Re-localiza o botão do serviço atual
            servicos_btn = page.locator('a[title="Serviços"]').nth(i)
            servicos_btn.click(timeout=10000)

            # Aguarda carregamento da rede e das tabelas
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except:
                pass
            page.wait_for_selector("table tbody tr", timeout=15000)
            # Pequena pausa extra para garantir renderização completa do Angular
            page.wait_for_timeout(1000)

            # Extração de tabelas com categoria correta
            tabelas = page.locator("table").all()

            for tabela in tabelas:
                # Extrai categoria da tabela
                categoria = self._extrair_categoria_tabela(tabela)

                # Extrai linhas da tabela
                linhas = tabela.locator("tbody tr").all()
                for linha in linhas:
                    tds = linha.locator("td").all()
                    # Usa text_content para garantir leitura mesmo se oculto/aninhado
                    valores = [td.text_content().strip() for td in tds]

                    if len(valores) >= 6:
                        # Determina tipo de registro
                        tipo_registro = self._determinar_tipo_registro(
                            categoria,
                            valores[4] if len(valores) > 4 else "",
                            valores[1] if len(valores) > 1 else "",
                        )

                        # Monta a linha de dados
                        dados_linha = [
                            id_value,  # ID
                            tipo_registro,  # TIPO DE REGISTRO
                            valores[0],  # CÓDIGO
                            valores[1],  # DESCRIÇÃO
                            valores[2],  # QUANTIDADE
                            valores[3],  # PREÇO UNITÁRIO
                            valores[4] if len(valores) > 4 else "",  # UNIDADE
                            valores[5] if len(valores) > 5 else "",  # PREÇO TOTAL
                            categoria,  # CATEGORIA (extraída da página)
                            status,  # STATUS
                        ]
                        todos_dados.append(dados_linha)
            # Volta para a lista de serviços se houver mais de um
            if count_servicos > 1 and i < count_servicos - 1:
                page.go_back()
                page.wait_for_load_state("networkidle", timeout=15000)

        # Volta ao menu
        page.click('//*[@id="ott-sidebar-collapse"]')
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        page.wait_for_load_state("networkidle", timeout=15000)

        return todos_dados

    def _pesquisar_id_medicao(self, page, id_value):
        todos_dados = []
        status = ""

        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        page.wait_for_selector('xpath=//*[@id="filtroId"]', timeout=10000)
        page.fill('xpath=//*[@id="filtroId"]', str(id_value))
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()

        try:
            status = self._extrair_status_id(page, id_value)
            page.locator("span.badge.bg-primary", has_text="Editar").click(
                timeout=10000
            )

            medicao = page.get_by_text("Medição", exact=True)

            medicao.click(timeout=2000)

            print(status)

            servico = page.locator('a[title="Serviços"]').count()

            for i in range(1, servico):
                btn_servico = page.locator('a[title="Serviços"]').nth(i)
                btn_servico.click(timeout=10000)

                page.wait_for_selector("table tbody tr", timeout=15000)
                page.wait_for_timeout(1000)

                # Extração de tabelas com categoria correta
                tabelas = page.locator("table").all()
                for tabela in tabelas:
                    # Extrai categoria da tabela
                    categoria = self._extrair_categoria_tabela(tabela)
                    # Extrai linhas da tabela
                    linhas = tabela.locator("tbody tr").all()
                    for linha in linhas:
                        tds = linha.locator("td").all()
                        valores = [td.text_content().strip() for td in tds]

                        if len(valores) >= 6:
                            # Determina tipo de registro
                            tipo_registro = self._determinar_tipo_registro(
                                categoria,
                                valores[4] if len(valores) > 4 else "",
                                valores[1] if len(valores) > 1 else "",
                            )

                            # Monta a linha de dados
                            dados_linha = [
                                id_value,  # ID
                                tipo_registro,  # TIPO DE REGISTRO
                                valores[0],  # CÓDIGO
                                valores[1],  # DESCRIÇÃO
                                valores[2],  # QUANTIDADE
                                valores[3],  # PREÇO UNITÁRIO
                                valores[4] if len(valores) > 4 else "",  # UNIDADE
                                valores[5] if len(valores) > 5 else "",  # PREÇO TOTAL
                                categoria,  # CATEGORIA (extraída da página)
                                status,  # STATUS
                            ]
                            todos_dados.append(dados_linha)

                page.go_back()
                medicao = page.get_by_text("Medição", exact=True)

                medicao.click(timeout=2000)
        except:
            return [
                [
                    id_value,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "Botão Editar não encontrado",
                ]
            ]

        return todos_dados

    def _pesquisar_id_medicao_antigo_backp(self, page, id_value):
        todos_dados = []
        status = ""
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        page.wait_for_selector('xpath=//*[@id="filtroId"]', timeout=10000)
        page.fill('xpath=//*[@id="filtroId"]', str(id_value))
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()

        status = self._extrair_status_id(page, id_value)

        try:
            # Botão Editar com seletor mais robusto
            page.locator("span.badge.bg-primary", has_text="Editar").click(
                timeout=10000
            )

        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Botão 'Editar' não encontrado. Status: '{status}'."
            )
            try:
                page.click('//*[@id="ott-sidebar-collapse"]')
                page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
                page.wait_for_load_state("networkidle", timeout=15000)
            except:
                pass
            return [[id_value, "", "", "", "", "", "", "", "", status]]

        # Navega até aba "Medição"
        # links = page.locator("a.nav-link")
        # total = int(links.count())
        # for i in range(total):
        #     texto = links.nth(i).text_content().strip()
        #     if texto == "Medição":
        #         links.nth(i).click()
        #         break

        # clicar no link de Medição usando seletor mais robusto
        try:
            page.get_by_role("tab", name="Medição", exact=True).click(timeout=10000)

        except TimeoutError:
            status = "Este ID não existe link de Medição"
            self.message.emit(
                f"⚠️ ID {id_value}: Botão 'Medição' não encontrado. Status: '{status}'."
            )
            return [[id_value, "", "", "", "", "", "", "", "", status]]

        page.wait_for_timeout(2000)

        try:
            # clicar no link de Serviços usando seletor
            page.wait_for_selector('a[title="Serviços"]', timeout=15000)
            count_servicos = page.locator('a[title="Serviços"]').count()
        except TimeoutError:
            self.message.emit(
                f"⚠️ ID {id_value}: Link de 'Serviços' não encontrado. Status: '{status}'."
            )
            return [
                [
                    id_value,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "Link de serviço não encontrado",
                ]
            ]

        for i in range(count_servicos):
            # Re-localiza o botão do serviço atual
            servico = page.locator('a[title="Serviços"]').nth(i)
            servico.click(timeout=10000)

            # Aguarda carregamento da rede e das tabelas (Mesma lógica do Draft)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except:
                pass
            page.wait_for_selector("table tbody tr", timeout=15000)
            page.wait_for_timeout(1000)

            # Extração de tabelas com categoria correta
            tabelas = page.locator("table").all()
            for tabela in tabelas:
                # Extrai categoria da tabela
                categoria = self._extrair_categoria_tabela(tabela)
                print(tabela)
                # Extrai linhas da tabela
                linhas = tabela.locator("tbody tr").all()
                for linha in linhas:
                    tds = linha.locator("td").all()
                    valores = [td.text_content().strip() for td in tds]

                    if len(valores) >= 6:
                        # Determina tipo de registro
                        tipo_registro = self._determinar_tipo_registro(
                            categoria,
                            valores[4] if len(valores) > 4 else "",
                            valores[1] if len(valores) > 1 else "",
                        )

                        # Monta a linha de dados
                        dados_linha = [
                            id_value,  # ID
                            tipo_registro,  # TIPO DE REGISTRO
                            valores[0],  # CÓDIGO
                            valores[1],  # DESCRIÇÃO
                            valores[2],  # QUANTIDADE
                            valores[3],  # PREÇO UNITÁRIO
                            valores[4] if len(valores) > 4 else "",  # UNIDADE
                            valores[5] if len(valores) > 5 else "",  # PREÇO TOTAL
                            categoria,  # CATEGORIA (extraída da página)
                            status,  # STATUS
                        ]
                        todos_dados.append(dados_linha)

            if count_servicos > 1 and i < count_servicos - 1:
                page.go_back()
                page.wait_for_load_state("networkidle", timeout=15000)

        # Volta ao menu
        page.click('//*[@id="ott-sidebar-collapse"]')
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        page.wait_for_load_state("networkidle", timeout=15000)

        return todos_dados

    def stop(self):
        self._running = False


# ===========================================================
# 🎨 INTERFACE GRÁFICA
# ===========================================================


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OSP Vivo Web Scraper")

        # Definir ícone da janela
        try:
            icon_path = resource_path(os.path.join("app", "img", "ico_osp.ico"))
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        self.setGeometry(100, 100, 800, 600)
        self.csv_path = ""
        self.worker = None
        self.last_generated_file = None

        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Header Layout (Título + Versão)
        header_layout = QHBoxLayout()

        # Título
        title_label = QLabel("🕸️ OSP Vivo Web Scraper")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # Versão
        version_label = QLabel("v1.0.0")
        version_label.setStyleSheet("font-size: 11px; color: gray; padding: 5px;")
        version_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop
        )

        header_layout.addWidget(title_label, 1)  # Stretch=1 para centralizar o título
        header_layout.addWidget(version_label, 0, Qt.AlignmentFlag.AlignTop)

        layout.addLayout(header_layout)

        # Seção de configurações
        config_group = QGroupBox("🔧 Configurações")
        config_layout = QFormLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Usuário do OSP Control")
        config_layout.addRow("👤 Usuário:", self.username_input)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setPlaceholderText("Senha do OSP Control")

        # Adicionar ação de 'olhinho' para revelar senha
        self.eye_action = self.password_input.addAction(
            self._create_eye_icon(False), QLineEdit.ActionPosition.TrailingPosition
        )
        self.eye_action.triggered.connect(self.toggle_password_visibility)
        self.eye_action.setToolTip("Mostrar senha")

        config_layout.addRow("🔒 Senha:", self.password_input)

        # Checkbox para modo Headless
        self.headless_checkbox = QCheckBox("👻 Executar em modo oculto (Headless)")
        config_layout.addRow(self.headless_checkbox)

        # Botão para salvar credenciais
        save_btn = QPushButton("💾 Salvar Credenciais")
        save_btn.clicked.connect(self.save_credentials)
        config_layout.addRow(save_btn)

        config_group.setLayout(config_layout)
        layout.addWidget(config_group)

        # Seção de arquivo CSV
        file_group = QGroupBox("📁 Arquivo CSV")
        file_layout = QVBoxLayout()

        self.file_label = QLabel("Nenhum arquivo selecionado")
        self.file_label.setStyleSheet(
            "padding: 5px; background-color: #f0f0f0; border-radius: 3px;"
        )

        file_btn_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("📂 Selecionar CSV")
        self.select_file_btn.clicked.connect(self.select_csv_file)
        self.clear_file_btn = QPushButton("🗑️ Limpar")
        self.clear_file_btn.clicked.connect(self.clear_csv_file)

        file_btn_layout.addWidget(self.select_file_btn)
        file_btn_layout.addWidget(self.clear_file_btn)

        file_layout.addWidget(self.file_label)
        file_layout.addLayout(file_btn_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Seção de modo de extração
        mode_group = QGroupBox("🎯 Modo de Extração")
        mode_layout = QVBoxLayout()

        self.mode_draft = QRadioButton("🧾 Draft")
        self.mode_draft.setChecked(True)
        self.mode_medicao = QRadioButton("📊 Medição")
        self.mode_cancelados = QRadioButton("🔎 ID Cancelados")
        self.mode_memoria = QRadioButton("🧮 Memória de Cálculo")

        mode_layout.addWidget(self.mode_draft)
        mode_layout.addWidget(self.mode_medicao)
        mode_layout.addWidget(self.mode_cancelados)
        mode_layout.addWidget(self.mode_memoria)

        # Tooltip explicativo para Memória de Cálculo
        self.mode_memoria.setToolTip(
            "Extrai tabela específica de memória de cálculo com:\n- Classe, Código, Descrição\n- Pontos, Custos Unitário e Total\n- Quantidade Executada"
        )

        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Área de logs
        log_group = QGroupBox("📝 Logs")
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Botões de ação
        action_layout = QHBoxLayout()

        self.start_btn = QPushButton("🚀 Iniciar Scraping")
        self.start_btn.clicked.connect(self.start_scraping)
        self.start_btn.setStyleSheet(
            "background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;"
        )

        self.stop_btn = QPushButton("⏹️ Parar")
        self.stop_btn.clicked.connect(self.stop_scraping)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet(
            "background-color: #f44336; color: white; padding: 10px;"
        )

        self.clear_logs_btn = QPushButton("🗑️ Limpar Logs")
        self.clear_logs_btn.clicked.connect(self.clear_logs)

        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.clear_logs_btn)

        layout.addLayout(action_layout)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Pronto")

    def _create_eye_icon(self, visible):
        """Cria um ícone de olho dinamicamente, sem usar emojis."""
        pixmap = QPixmap(24, 24)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        pen = QPen(QColor("#555555"))
        pen.setWidth(2)
        painter.setPen(pen)

        # Desenha o contorno do olho
        path = QPainterPath()
        path.moveTo(3, 12)
        path.quadTo(12, 5, 21, 12)
        path.quadTo(12, 19, 3, 12)
        painter.drawPath(path)

        # Desenha a íris
        painter.drawEllipse(QPoint(12, 12), 3, 3)

        if not visible:
            # Se a senha estiver oculta (não visível), desenha um risco sobre o olho
            painter.drawLine(5, 19, 19, 5)

        painter.end()
        return QIcon(pixmap)

    def toggle_password_visibility(self):
        """Alterna visibilidade da senha"""
        if self.password_input.echoMode() == QLineEdit.EchoMode.Password:
            self.password_input.setEchoMode(QLineEdit.EchoMode.Normal)
            self.eye_action.setIcon(self._create_eye_icon(True))
            self.eye_action.setToolTip("Ocultar senha")
        else:
            self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.eye_action.setIcon(self._create_eye_icon(False))
            self.eye_action.setToolTip("Mostrar senha")

    def load_config(self):
        """Carrega configurações salvas"""
        config_file = "config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, "r") as f:
                    config = json.load(f)
                    self.username_input.setText(config.get("username", ""))
                    self.password_input.setText(config.get("password", ""))
                    self.headless_checkbox.setChecked(config.get("headless", False))
            except:
                pass

    def save_credentials(self):
        """Salva credenciais em arquivo"""
        config = {
            "username": self.username_input.text(),
            "password": self.password_input.text(),
            "headless": self.headless_checkbox.isChecked(),
        }

        try:
            with open("config.json", "w") as f:
                json.dump(config, f)
            self.log_message("✅ Credenciais salvas com sucesso!")
        except Exception as e:
            self.log_message(f"❌ Erro ao salvar credenciais: {str(e)}")

    def select_csv_file(self):
        """Seleciona arquivo CSV"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo CSV", "", "Arquivos CSV (*.csv)"
        )

        if file_path:
            self.csv_path = file_path
            self.file_label.setText(f"📄 {os.path.basename(file_path)}")
            self.log_message(f"✅ Arquivo CSV selecionado: {file_path}")

    def clear_csv_file(self):
        """Limpa seleção de arquivo CSV"""
        self.csv_path = ""
        self.file_label.setText("Nenhum arquivo selecionado")
        self.log_message("🗑️ Seleção de arquivo removida")

    def get_selected_mode(self):
        """Retorna o modo selecionado"""
        if self.mode_draft.isChecked():
            return 1
        elif self.mode_medicao.isChecked():
            return 2
        elif self.mode_cancelados.isChecked():
            return 3
        elif self.mode_memoria.isChecked():
            return 4
        return 1

    def start_scraping(self):
        """Inicia o processo de scraping"""
        if not self.csv_path:
            QMessageBox.warning(self, "Atenção", "Selecione um arquivo CSV primeiro!")
            return

        # Aviso especial para Memória de Cálculo
        if self.mode_memoria.isChecked():
            reply = QMessageBox.information(
                self,
                "Modo Memória de Cálculo",
                "Este modo extrai a tabela específica de memória de cálculo com:\n"
                "- Classe, Código, Descrição do Serviço\n"
                "- Unidade, Pontos, Custo Unitário\n"
                "- Quantidade Executada, Pontos Totais, Custo Total\n\n"
                "O processo navegará por todas as páginas automaticamente.\n\n"
                "Deseja continuar?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return

        if not self.username_input.text() or not self.password_input.text():
            reply = QMessageBox.question(
                self,
                "Confirmação",
                "Credenciais não preenchidas. O navegador abrirá para login manual.\nDeseja continuar?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.No:
                return

        # Desabilitar controles durante execução
        self.start_btn.setEnabled(False)
        self.select_file_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.clear_logs_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.last_generated_file = None

        # Criar e configurar worker
        self.worker = WebScraperWorker()
        self.worker.csv_path = self.csv_path
        self.worker.mode = self.get_selected_mode()
        self.worker.username = self.username_input.text()
        self.worker.password = self.password_input.text()
        self.worker.headless_enabled = self.headless_checkbox.isChecked()

        # Conectar sinais
        self.worker.progress.connect(self.update_progress)
        self.worker.message.connect(self.log_message)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(self.on_finished)
        self.worker.data_saved.connect(self.on_data_saved)

        # Iniciar thread
        self.worker.start()

        mode_text = {
            1: "Draft",
            2: "Medição",
            3: "ID Cancelados",
            4: "Memória de Cálculo",
        }
        self.log_message(f"🚀 Iniciando extração de {mode_text[self.worker.mode]}...")

    def stop_scraping(self):
        """Para o scraping em execução"""
        reply = QMessageBox.question(
            self,
            "Confirmar Parada",
            "Tem certeza que deseja parar o processo?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.worker:
                self.worker.stop()
                self.log_message(
                    "🛑 Solicitada parada... Aguarde a finalização da tarefa atual."
                )
                QMessageBox.information(
                    self,
                    "Aguarde",
                    "O processo de parada foi iniciado.\nPor favor, aguarde enquanto a operação atual é finalizada com segurança.",
                )
                self.stop_btn.setEnabled(False)

    def update_progress(self, value):
        """Atualiza barra de progresso"""
        self.progress_bar.setValue(value)

    def log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
        self.log_text.append(f"[{timestamp}] {message}")
        self.status_bar.showMessage(message)

        # Rolagem automática para baixo
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_text.setTextCursor(cursor)

    def show_error(self, error_message):
        """Exibe mensagem de erro"""
        self.log_message(f"❌ {error_message}")
        QMessageBox.critical(self, "Erro", error_message)

    def on_finished(self):
        """Chamado quando o worker termina"""
        self.start_btn.setEnabled(True)
        self.select_file_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.clear_logs_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.log_message("✅ Processo finalizado!")

        # Perguntar se deseja abrir a pasta de downloads
        if self.csv_path:
            msg = "Processo finalizado!"
            if self.last_generated_file:
                msg += f"\nArquivo gerado: {os.path.basename(self.last_generated_file)}"
            msg += "\nDeseja abrir o local do arquivo?"

            reply = QMessageBox.question(
                self,
                "Concluído",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                if self.last_generated_file and os.path.exists(
                    self.last_generated_file
                ):
                    if sys.platform == "win32":
                        subprocess.run(
                            [
                                "explorer",
                                "/select,",
                                os.path.normpath(self.last_generated_file),
                            ]
                        )
                    else:
                        os.startfile(os.path.dirname(self.last_generated_file))
                else:
                    downloads_path = Path.home() / "Downloads"
                    os.startfile(downloads_path)

    def on_data_saved(self, file_path):
        """Chamado quando dados são salvos"""
        self.last_generated_file = file_path
        self.log_message(f"💾 Dados salvos em: {file_path}")

    def clear_logs(self):
        """Limpa os logs"""
        reply = QMessageBox.question(
            self,
            "Confirmar Limpeza",
            "Tem certeza que deseja limpar os logs?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.log_text.clear()
            self.log_message("🗑️ Logs limpos")

    def closeEvent(self, event):
        """Evento de fechamento da janela"""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirmar saída",
                "O scraping está em execução. Deseja realmente sair?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.worker.stop()
                self.worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


# ===========================================================
# 📦 FUNÇÃO PRINCIPAL
# ===========================================================


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Definir estilo CSS
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f5;
        }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #cccccc;
            border-radius: 5px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QPushButton {
            padding: 5px 15px;
            border-radius: 3px;
            border: 1px solid #cccccc;
        }
        QPushButton:hover {
            background-color: #e0e0e0;
        }
        QTextEdit {
            border: 1px solid #cccccc;
            border-radius: 3px;
            font-family: Consolas, monospace;
        }
        QProgressBar {
            border: 1px solid #cccccc;
            border-radius: 3px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #4CAF50;
            border-radius: 3px;
        }
        QRadioButton {
            padding: 5px;
        }
        QRadioButton:hover {
            background-color: #e8f4fd;
            border-radius: 3px;
        }
    """)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
