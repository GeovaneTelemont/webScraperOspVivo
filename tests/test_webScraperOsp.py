"""
Arquivo de Teste para o OSP Vivo Web Scraper.
Execute este arquivo para garantir que a lógica do sistema está funcionando corretamente.
Comando: python test_webscraperOsp.py
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, mock_open, patch

# Adiciona o diretório app ao path para poder importar o módulo webScraperOsp
current_dir = os.path.dirname(os.path.abspath(__file__))

# Ajusta o path dependendo de onde o teste está sendo executado (raiz ou tests/)
if os.path.basename(current_dir) == "tests":
    project_root = os.path.dirname(current_dir)
else:
    project_root = current_dir

app_dir = os.path.join(project_root, "app")
sys.path.insert(0, app_dir)
sys.path.insert(0, project_root)

# Importação condicional do PyQt6 para evitar erro em ambientes sem display (CI/CD)
try:
    from PyQt6.QtWidgets import QApplication

    # Garante que existe uma instância do QApplication antes de criar widgets
    if not QApplication.instance():
        app_instance = QApplication(sys.argv)
    else:
        app_instance = QApplication.instance()
except ImportError:
    app_instance = None

# Tenta importar o módulo principal
try:
    import webScraperOsp
except ImportError:
    # Fallback se a estrutura de pastas for diferente
    try:
        from app import webScraperOsp
    except ImportError:
        raise ImportError(
            "❌ Erro: Não foi possível importar 'webScraperOsp.py'. Verifique o caminho."
        )


class TestWebScraperLogic(unittest.TestCase):
    """Testes de unidade para a lógica de negócios do Web Scraper (Worker)"""

    def setUp(self):
        # Instancia o worker sem iniciar a thread (apenas a classe lógica)
        self.worker = webScraperOsp.WebScraperWorker()

    def test_normalize_text(self):
        """Testa se a normalização de texto remove acentos, espaços e coloca em minúsculo"""
        self.assertEqual(self.worker._normalize_text("Coração"), "coracao")
        self.assertEqual(
            self.worker._normalize_text("  Espaços   Extras  "), "espacos extras"
        )
        self.assertEqual(self.worker._normalize_text("Maçã"), "maca")
        self.assertEqual(self.worker._normalize_text("Água"), "agua")
        self.assertEqual(self.worker._normalize_text(None), "")
        self.assertEqual(self.worker._normalize_text(""), "")

    def test_determinar_tipo_registro_material(self):
        """Testa regras para identificar corretamente o tipo 'Material'"""
        # Regra 1: Palavra chave na categoria
        self.assertEqual(
            self.worker._determinar_tipo_registro("Material de Rede"), "Material"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Materiais Diversos"), "Material"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Telefonica"), "Material"
        )

        # Regra 2: Por unidade de medida
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", unidade="m"), "Material"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", unidade="un"), "Material"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", unidade="kg"), "Material"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", unidade="CJ"), "Material"
        )

        # Regra 3: Por palavra chave na descrição
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", descricao="Cabo de fibra"),
            "Material",
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro(
                "Outros", descricao="Subduto corrugado"
            ),
            "Material",
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Outros", descricao="Chassi"),
            "Material",
        )

    def test_determinar_tipo_registro_servico(self):
        """Testa regras para identificar corretamente o tipo 'Serviço'"""
        self.assertEqual(
            self.worker._determinar_tipo_registro("Serviços Técnicos"), "Serviço"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Mão de Obra"), "Serviço"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Classe de Serviço"), "Serviço"
        )

        # Teste de Fallback (padrão deve ser serviço se não cair em outras regras)
        self.assertEqual(
            self.worker._determinar_tipo_registro(
                "Categoria Desconhecida", unidade="", descricao=""
            ),
            "Serviço",
        )

    def test_determinar_tipo_registro_custo(self):
        """Testa regras para identificar corretamente o tipo 'Custo'"""
        self.assertEqual(
            self.worker._determinar_tipo_registro("Custos Adicionais"), "Custo"
        )
        self.assertEqual(
            self.worker._determinar_tipo_registro("Custo de Viagem"), "Custo"
        )

    def test_extrair_categoria_tabela_sucesso(self):
        """Testa a extração de categoria simulando um elemento do Playwright"""
        mock_element = MagicMock()
        # Simula o retorno do JavaScript evaluate que o código usa
        mock_element.evaluate.return_value = "  Categoria Teste  "

        resultado = self.worker._extrair_categoria_tabela(mock_element)
        self.assertEqual(resultado, "Categoria Teste")

    def test_extrair_categoria_tabela_erro(self):
        """Testa comportamento resiliente quando falha a extração da categoria"""
        mock_element = MagicMock()
        mock_element.evaluate.side_effect = Exception("Erro no JS")

        resultado = self.worker._extrair_categoria_tabela(mock_element)
        self.assertEqual(resultado, "")


class TestMainWindowInterface(unittest.TestCase):
    """Testes para a Interface Gráfica (PyQt6)"""

    @classmethod
    def setUpClass(cls):
        # Garante QApplication apenas uma vez
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)

    def setUp(self):
        # Cria a janela principal para cada teste

        # Patch os.path.exists para simular que não há arquivo de config
        # Isso evita que o estado da máquina local afete os testes (AssertionError em headless)
        self.patcher_exists = patch("os.path.exists", return_value=False)
        self.mock_exists = self.patcher_exists.start()
        self.addCleanup(self.patcher_exists.stop)

        self.window = webScraperOsp.MainWindow()

    def test_initial_ui_state(self):
        """Verifica se a janela inicia com os valores padrão corretos"""
        self.assertEqual(self.window.windowTitle(), "OSP Vivo Web Scraper")

        # Botões de controle
        self.assertTrue(self.window.start_btn.isEnabled())
        self.assertFalse(self.window.stop_btn.isEnabled())

        # Configurações padrão
        self.assertFalse(
            self.window.headless_checkbox.isChecked(),
            "Headless deve começar desligado por padrão se não houver config",
        )
        self.assertTrue(
            self.window.mode_draft.isChecked(), "Modo Draft deve ser o padrão"
        )

    @patch(
        "builtins.open",
        new_callable=mock_open,
        read_data='{"username": "user_test", "password": "123", "headless": true}',
    )
    @patch("json.load")
    def test_load_config(self, mock_json_load, mock_file):
        """Testa se as configurações (usuário, senha, headless) são carregadas do JSON"""
        # Simula que o arquivo de configuração existe para este teste específico
        self.mock_exists.return_value = True

        # Configura o mock do json.load para retornar dados fictícios
        mock_json_load.return_value = {
            "username": "user_test",
            "password": "123",
            "headless": True,
        }

        # Chama a função de carregar
        self.window.load_config()

        # Verifica se a UI foi atualizada
        self.assertEqual(self.window.username_input.text(), "user_test")
        self.assertEqual(self.window.password_input.text(), "123")
        self.assertTrue(self.window.headless_checkbox.isChecked())

    @patch("webScraperOsp.WebScraperWorker")
    @patch("PyQt6.QtWidgets.QMessageBox.warning")
    def test_start_scraping_validation_sem_csv(self, mock_msg_box, MockWorker):
        """Testa se o sistema impede iniciar sem selecionar CSV"""
        self.window.csv_path = ""  # Garante que está vazio

        self.window.start_scraping()

        mock_msg_box.assert_called_once()  # Deve exibir popup de aviso
        self.assertIsNone(self.window.worker)  # Não deve ter criado o worker

    @patch("webScraperOsp.WebScraperWorker")
    def test_start_scraping_success(self, MockWorker):
        """Testa o fluxo feliz de iniciar o scraping"""
        # Configura o ambiente
        self.window.csv_path = "c:/caminho/falso/teste.csv"
        self.window.username_input.setText("usuario")
        self.window.password_input.setText("senha")

        # Pega a instância mockada do worker
        mock_worker_instance = MockWorker.return_value

        # Ação
        self.window.start_scraping()

        # Verificações
        self.assertIsNotNone(self.window.worker)
        # Verifica se passou os dados corretamente para o worker
        self.assertEqual(self.window.worker.csv_path, "c:/caminho/falso/teste.csv")
        self.assertEqual(self.window.worker.username, "usuario")

        # Verifica se chamou o start() da thread
        mock_worker_instance.start.assert_called_once()

        # Verifica se a UI travou os botões corretamente
        self.assertFalse(self.window.start_btn.isEnabled())
        self.assertTrue(self.window.stop_btn.isEnabled())


if __name__ == "__main__":
    unittest.main()
