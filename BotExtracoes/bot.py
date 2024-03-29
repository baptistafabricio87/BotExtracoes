from botcity.core import DesktopBot
import pyautogui


class Bot(DesktopBot):
    def action(self, execution=None):
        pyautogui.alert(
            "INICIANDO PROCESSO DE EXTRAÇÕES \
                \nInforme login e senha SAP e aguarde."
        )

        self.login_sap()
        self.executa_sap_sp02()
        print("Login OK! Processo Iniciado! Por favor aguarde...")
        self.extrair_me3l()
        self.extrair_zmm_contrato()
        self.extrair_acomp_pedido()
        self.extrair_po_contrato()
        self.extrair_zmm_fornecedores()
        self.extrair_ydv1()
        self.extrair_zmm_qualif_forn_rel()
        self.type_keys(["f3", "f3"])
        self.formatar_exportar_arquivos()
        print("Processo Finalizado!")

        pyautogui.alert(
            "PROCESSO CONCLUÍDO! \
                \nFavor aguardar macros de formatação e importação"
        )

    def login_sap(self):
        """
        Executa e realiza login no SAP GUI.
        Assim que inicia é solicitado usuario e senha no prompt de comando.

        Args:
            usuario (str): Usuario SAP
            senha (str): Senha SAP
        """
        usuario = pyautogui.prompt("Informar Usuario SAP ")
        senha = pyautogui.password("Informar Senha SAP ", mask="*")

        self.execute(
            "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
        )

        if not self.find(
            "localiza_b05_pd4", matching=0.97, waiting_time=50000
        ):
            self.not_found("localiza_b05_pd4")
        self.double_click()

        if not self.find(
            "localiza_usuario", matching=0.97, waiting_time=50000
        ):
            self.not_found("localiza_usuario")
        self.click_relative(158, 7)

        self.paste(usuario)
        self.tab(wait=100)
        self.paste(senha)
        self.key_enter()

    def executa_sap_sp02(self):
        self.wait(2000)
        self.type_keys(["S", "P", "0", "2", "return"])
        self.wait(1000)
        self.type_keys(["ctrl", "shift", "f10"])
        self.wait(1000)
        self.type_keys(["tab", "tab", "tab", "delete", "return"])

    def salva_arquivo(
        self, nome_arquivo, caminho_diretorio, wait_time, convert
    ):
        if not self.find(
            "btn_gravar_file_local", matching=0.97, waiting_time=wait_time
        ):
            self.not_found("btn_gravar_file_local")
        self.click()

        if not convert:
            if not self.find(
                "selecionar_texto_com_tabuladores",
                matching=0.97,
                waiting_time=500000,
            ):
                self.not_found("selecionar_texto_com_tabuladores")
            self.click_relative(12, 9)
        else:
            if not self.find("nao_convert", matching=0.97, waiting_time=500000):
                self.not_found("nao_convert")
            self.click_relative(-8, 9)

        if not self.find(
            "btn_confirmar_tipo_arquivo", matching=0.97, waiting_time=300000
        ):
            self.not_found("btn_confirmar_tipo_arquivo")
        self.click()

        if not self.find(
            "localiza_diretorio", matching=0.97, waiting_time=wait_time
        ):
            self.not_found("localiza_diretorio")
        self.click()

        self.tab(wait=1000)
        self.paste(text=caminho_diretorio)
        self.tab(wait=1000)
        self.paste(text=nome_arquivo)
        self.tab(wait=1000)
        self.tab(wait=1000)
        self.space()
        if not self.find(
            "confirma_arq_salvo", matching=0.97, waiting_time=wait_time
        ):
            self.not_found("confirma_arq_salvo")
        self.click_relative(248, -655)

    def extrair_acomp_pedido(self):
        if not self.find(
            "localiza_acomp_pedido", matching=0.97, waiting_time=20000
        ):
            self.not_found("localiza_acomp_pedido")
        self.click_relative(-337, 3)

        self.salva_arquivo(
            convert=False,
            wait_time=1000000,
            nome_arquivo="Acomp_Pedido.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_po_contrato(self):
        if not self.find(
            "localiza_po_contrato", matching=0.97, waiting_time=200000
        ):
            self.not_found("localiza_po_contrato")
        self.click_relative(-336, 2)

        self.salva_arquivo(
            convert=False,
            wait_time=4000000,
            nome_arquivo="POContrato.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_ydv1(self):
        if not self.find("localiza_ydv1", matching=0.97, waiting_time=200000):
            self.not_found("localiza_ydv1")
        self.click_relative(-335, 4)

        self.salva_arquivo(
            convert=False,
            wait_time=100000,
            nome_arquivo="YDV1.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_zmm_contrato(self):
        if not self.find(
            "localiza_zmm_contrato", matching=0.97, waiting_time=200000
        ):
            self.not_found("localiza_zmm_contrato")
        self.click_relative(-334, 6)

        self.salva_arquivo(
            convert=False,
            wait_time=10000000,
            nome_arquivo="ZMM_contratos.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_zmm_fornecedores(self):
        if not self.find(
            "localiza_zmm_fornecedores", matching=0.97, waiting_time=200000
        ):
            self.not_found("localiza_zmm_fornecedores")
        self.click_relative(-333, 4)

        self.salva_arquivo(
            convert=False,
            wait_time=500000,
            nome_arquivo="ZMM_fornecedores.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_zmm_qualif_forn_rel(self):
        if not self.find(
            "localiza_zmm_qualif", matching=0.97, waiting_time=300000
        ):
            self.not_found("localiza_zmm_qualif")
        self.click_relative(-335, 4)

        self.salva_arquivo(
            convert=False,
            wait_time=300000,
            nome_arquivo="ZMM_qualif_forn_rel.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def extrair_me3l(self):
        if not self.find("localiza_me3l", matching=0.97, waiting_time=300000):
            self.not_found("localiza_me3l")
        self.click_relative(-334, 4)

        self.salva_arquivo(
            convert=True,
            wait_time=6000000,
            nome_arquivo="ME3L.txt",
            caminho_diretorio="P:\\Central_Abastecimento\\Automações\\Projeto de Contratos\\Parametros\\",
        )

    def formatar_exportar_arquivos(self):
        app_path = "P:\Central_Abastecimento\Automações\Projeto de Contratos\Documentos\ExportarArquivos - Contratos.xlsm"
        self.execute(app_path)
        if not self.find("localiza_menu", matching=0.97, waiting_time=1000000):
            self.not_found("localiza_menu")
        self.click()
        if not self.find(
            "localiza_formatar_arquivos", matching=0.97, waiting_time=1000000
        ):
            self.not_found("localiza_formatar_arquivos")
        self.click()
        if not self.find(
            "localiza_importar_arquivos", matching=0.97, waiting_time=6000000
        ):
            self.not_found("localiza_importar_arquivos")
        self.click()

    def not_found(self, label):
        print(f"Element not found (Elemento nao encontrado): {label}")


if __name__ == "__main__":
    Bot.main()
