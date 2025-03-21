import customtkinter as ctk
import pandas as pd
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import time
import logging
from datetime import datetime
import os
from PIL import Image, ImageTk
import traceback
import threading
from pywinauto.timings import wait_until

class AutomacaoGUI:
    def __init__(self):
        # Configuração do tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - Folha de Ponto")
        self.window.geometry("700x500")
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)
        
        # Flag para controle de execução
        self.executando = False
        self.thread_automacao = None

        # Configurar ícone
        self.set_window_icon()

        # Criar diretório de logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)
        
        # Configurar logging para arquivos
        self.setup_file_logging()
        
        # Variáveis
        self.arquivo_excel = ctk.StringVar()
        self.linha_inicial = ctk.StringVar(value="1")
        self.status_var = ctk.StringVar(value="Aguardando início...")
        
        # Logger
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []
        
        # Adicionar GUIHandler uma única vez
        class GUIHandler(logging.Handler):
            def __init__(self, gui):
                super().__init__()
                self.gui = gui
                
            def emit(self, record):
                msg = self.format(record)
                self.gui.adicionar_log(msg)
        
        self.gui_handler = GUIHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)
        
        self.criar_interface()

    def setup_file_logging(self):
        """Configura o logging para arquivos"""
        data_atual = datetime.now().strftime("%Y-%m-%d")
        
        # Logger de sucesso
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            success_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'),
                encoding='utf-8'
            )
            success_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.success_logger.addHandler(success_handler)
        
        # Logger de erro
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            error_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'),
                encoding='utf-8'
            )
            error_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.error_logger.addHandler(error_handler)

    def set_window_icon(self):
        """Configura o ícone da janela"""
        try:
            # Caminho para o ícone (ajuste para o caminho correto)
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "bot-folha-de-pagamento.ico")
            
            # Para Windows
            if os.name == 'nt' and os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
                
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")
    
    def selecionar_arquivo(self):
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {filename}")
        
    def criar_interface(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Frame superior para inputs
        input_frame = ctk.CTkFrame(main_frame)
        input_frame.pack(fill="x", padx=10, pady=10)
        
        # Seleção do arquivo Excel
        ctk.CTkLabel(input_frame, text="Arquivo Excel:").pack(anchor="w", padx=5, pady=2)
        
        file_frame = ctk.CTkFrame(input_frame)
        file_frame.pack(fill="x", pady=2)
        
        ctk.CTkEntry(file_frame, textvariable=self.arquivo_excel, width=400).pack(side="left", padx=5)
        ctk.CTkButton(file_frame, text="Procurar", command=self.selecionar_arquivo, width=100).pack(side="left", padx=5)
        
        # Linha inicial
        linha_frame = ctk.CTkFrame(input_frame)
        linha_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(linha_frame, text="Iniciar da linha:").pack(side="left", padx=5)
        ctk.CTkEntry(linha_frame, textvariable=self.linha_inicial, width=100).pack(side="left", padx=5)
        
        # Botões de controle
        buttons_frame = ctk.CTkFrame(input_frame)
        buttons_frame.pack(fill="x", pady=10)
        
        self.btn_iniciar = ctk.CTkButton(
            buttons_frame, 
            text="Iniciar", 
            command=self.iniciar_automacao_thread,
            height=35
        )
        self.btn_iniciar.pack(side="left", expand=True, fill="x", padx=5, pady=5)
        
        self.btn_parar = ctk.CTkButton(
            buttons_frame, 
            text="Parar", 
            command=self.parar_automacao,
            height=35,
            state="disabled"
        )
        self.btn_parar.pack(side="left", expand=True, fill="x", padx=5, pady=5)
        
        # Barra de Progresso
        self.progress_var = ctk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(main_frame)
        self.progress_bar.pack(fill="x", padx=10, pady=10)
        self.progress_bar.set(0)
        
        # Status
        ctk.CTkLabel(
            main_frame, 
            textvariable=self.status_var,
            height=25
        ).pack(pady=5)
        
        # Área de log
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.log_text = ctk.CTkTextbox(log_frame, height=200)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botão para limpar logs
        ctk.CTkButton(
            log_frame, 
            text="Limpar Logs", 
            command=self.limpar_logs,
            height=25
        ).pack(pady=5)

    def limpar_logs(self):
        """Limpa a área de logs"""
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo")
    
    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso"""
        porcentagem = (atual / total)
        self.progress_bar.set(porcentagem)
        self.status_var.set(f"Processando: {atual}/{total} ({porcentagem*100:.1f}%)")
        self.window.update_idletasks()
        
    def adicionar_log(self, mensagem):
        """Adiciona mensagem ao log visual"""
        self.log_text.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {mensagem}\n")
        self.log_text.see("end")
        self.window.update_idletasks()
    
    def iniciar_automacao_thread(self):
        """Inicia a automação em uma thread separada"""
        if self.executando:
            self.adicionar_log("Automação já em execução")
            return
        
        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()
        
        # Atualiza interface
        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
    
    def parar_automacao(self):
        """Para a execução da automação"""
        if self.executando:
            self.executando = False
            self.adicionar_log("Solicitação de parada enviada. Aguardando conclusão...")
            self.status_var.set("Interrompendo...")
        
    def ao_fechar(self):
        """Tratamento do fechamento da janela"""
        if self.executando:
            # Pede confirmação
            if ctk.messagebox.askyesno("Confirmação", 
                                      "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()
        
    def iniciar_automacao(self):
        if not self.arquivo_excel.get():
            self.adicionar_log("Erro: Selecione um arquivo Excel")
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")
            return
        
        try:
            linha_inicial = int(self.linha_inicial.get())
            if linha_inicial < 1:
                raise ValueError
        except ValueError:
            self.adicionar_log("Erro: Linha inicial inválida")
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")
            return
            
        self.adicionar_log("Iniciando automação...")
        self.status_var.set("Em execução...")
        self.executando = True
        
        try:
            # Carregar Excel
            df = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(df) - (linha_inicial - 1)
            self.adicionar_log(f"Arquivo Excel carregado com {total_linhas} linhas para processar")
            
            # Resetar barra de progresso
            self.progress_bar.set(0)
            
            # Iniciar automação
            automacao = DominioAutomation(self.logger, self)
            
            # Conectar ao Domínio
            if not automacao.connect_to_dominio():
                erro_msg = "Erro: Não foi possível conectar ao Domínio"
                self.adicionar_log(erro_msg)
                self.error_logger.error(erro_msg)
                return
                
            # Processar linhas
            for idx, (index, row) in enumerate(df.iloc[linha_inicial-1:].iterrows()):
                # Verificar se deve parar
                if not self.executando:
                    self.adicionar_log("Automação interrompida pelo usuário")
                    break
                
                # Atualizar progresso
                self.atualizar_progresso(idx + 1, total_linhas)
                
                try:
                    log_msg = (f"Linha {index + 1} - Nº {row['Nº']} - "
                            f"EMPRESA: {row.get('EMPRESA', 'N/A')}")
                    
                    success = automacao.processar_linha(row, index)
                    
                    # Log do resultado
                    if success:
                        self.success_logger.info(f"{log_msg} - Enviado com sucesso")
                        self.adicionar_log(f"Linha {index + 1} processada com sucesso")
                    else:
                        self.error_logger.error(f"{log_msg} - Erro no envio")
                        self.adicionar_log(f"Processo interrompido na linha {index + 1}")
                        break
                    
                    time.sleep(2)
                    
                except Exception as e:
                    erro_msg = f"{log_msg} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg)
                    self.adicionar_log(f"Detalhes do erro: {traceback.format_exc()}")
                    break
            
            self.status_var.set("Processamento concluído")
            self.progress_bar.set(1.0)
            
        except Exception as e:
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg)
            self.adicionar_log(f"Detalhes do erro: {traceback.format_exc()}")
            self.status_var.set("Erro no processamento")
        finally:
            self.executando = False
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")
            
    def executar(self):
        self.window.mainloop()

class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui
        
    def log(self, message):
        self.logger.info(message)

    def find_dominio_window(self):
        try:
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                return windows[0]
            self.log("Nenhuma janela do Domínio Folha encontrada.")
            return None
        except Exception as e:
            self.log(f"Erro ao procurar a janela do Domínio Folha: {str(e)}")
            return None

    def connect_to_dominio(self):
        try:
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            # Restaura a janela se estiver minimizada
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            # Traz a janela para o primeiro plano
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)
            return True
        except Exception as e:
            self.log(f"Erro ao conectar ao Domínio Folha: {str(e)}")
            return False

    def wait_for_window(self, titulo, timeout=30):
        """Espera por uma janela com o título especificado"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Tenta encontrar a janela pelo título
                window = self.app.window(title=titulo)
                if window.exists():
                    return window
            except Exception:
                pass
            time.sleep(0.5)
        raise TimeoutError(f"Timeout esperando pela janela: {titulo}")

    def wait_and_check_window_closed(self, window, window_title, timeout=30):
        """Espera até que uma janela seja fechada"""
        start_time = time.time()
        while time.time() - start_time < timeout:

            if not window.exists() or not window.is_visible():
                self.log(f"Janela '{window_title}' fechada com sucesso")
                return True
            else:
                self.log(f"Janela '{window_title}' não foi fechada")
                return False

        time.sleep(0.5)
        self.log(f"Aviso: Tempo máximo de espera atingido para fechamento da janela '{window_title}'")
        return False

    def processar_linha(self, row, index):
        try:
            self.log(f"Processando linha {index + 1}"
            )

            # Conectar à janela principal
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)

            app = Application(backend="uia").connect(handle=handle)
            main_window = app.window(handle=handle)
        
            # Garantir que a janela principal está em foco
            main_window.set_focus()
            time.sleep(0.5)
            
            # Enviar F8
            self.log("Enviando F8 para troca de empresas")
            send_keys('{F8}')
            time.sleep(1.5)

            # Verificar e focar na janela "Troca de empresas"
            try:
                troca_empresas_window = None
                max_attempts = 3
                
                for attempt in range(max_attempts):
                    try:
                        # Tentar encontrar usando critérios mais específicos
                        troca_empresas_window = main_window.child_window(
                            title="Troca de empresas",
                            class_name="FNWND3190"  # Adicione a classe correta aqui
                        )
                        
                        if troca_empresas_window.exists():
                            break
                            
                        # Se não encontrar com critérios específicos, pegar o primeiro elemento
                        troca_empresas_windows = main_window.children(title="Troca de empresas")
                        if troca_empresas_windows:
                            troca_empresas_window = troca_empresas_windows[0]
                            break
                    except Exception:
                        if attempt == max_attempts - 1:
                            self.log("Janela 'Troca de empresas' não encontrada após várias tentativas.")
                            return False
                        time.sleep(1)
                
                if not troca_empresas_window:
                    self.log("Janela 'Troca de empresas' não encontrada.")
                    return False
            except Exception as e:
                self.log(f"Erro ao localizar janela 'Troca de empresas': {str(e)}")
                return False

            self.log("Janela 'Troca de empresas' visível")

            # Enviar o valor pelo teclado
            empresa_num = str(int(row['Nº']))
            self.log(f"Enviando código da empresa: {empresa_num}")
            send_keys(empresa_num)
            time.sleep(0.3)

            # Pressionar Enter
            send_keys('{ENTER}')
            time.sleep(6)  # Reduzido tempo de espera

            # # tentando encontrar janela de aviso bloq honorario
            # honorario_window = main_window.child_window(
            #     auto_id="4888",
            #     class_name="#32770" 
            # )

            # time.sleep(1)

            # if honorario_window.exists():
            #     self.gui.adicionar_log(f"Janela de aviso de bloqueio honorário encontrada na linha {index + 1}")

            #     try:
            #         # Clicar no botão "OK" da janela de aviso
            #         button_ok_honorario = honorario_window.child_window(auto_id="2", class_name="Button")
            #         button_ok_honorario.click_input()
            #         time.sleep(1)

            #         # Registrar no log de erros
            #         self.gui.error_logger.error(f"Linha {index + 1} - Bloqueado por honorário")
                    
            #         # Incrementar contador de linhas puladas
            #         self.gui.linhas_puladas += 1
            #         self.gui.adicionar_log(f"Linhas puladas devido a bloqueio honorário: {self.gui.linhas_puladas}")

            #         # Pular para a próxima iteração do loop
            #         return False

            #     except Exception as e:
            #         self.gui.adicionar_log(f"Erro ao interagir com janela de aviso: {str(e)}")
            
            # Esperar até que a janela de troca feche
            self.wait_and_check_window_closed(troca_empresas_window, "Troca de empresas")
            
            # Fechar janela de avisos de vencimentos se estiver aberta
            try:
                aviso_window = main_window.child_window(
                    title="Avisos de Vencimento",
                    class_name="FNWND3190"
                )
                
                if aviso_window.exists() and aviso_window.is_visible():
                    self.log("Janela 'Avisos de Vencimento' encontrada - executando fechamento")
                    aviso_window.set_focus()
                    send_keys('{ESC}')
                    time.sleep(1)
                    send_keys('{ESC}')
                    self.log("ESCs executados para fechar 'Avisos de Vencimento'")
            except Exception:
                self.log("Nenhuma janela de 'Avisos de Vencimento' encontrada")

            # Voltar para janela principal e enviar comandos
            self.log("Enviando comandos para acessar relatórios")
            main_window.set_focus()
            send_keys('%r')  # ALT+R
            time.sleep(0.5)
            send_keys('i')
            time.sleep(0.5)
            send_keys('i')
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(1)

            # Gerenciador de Relatórios
            try:
                # Encontrar a janela do Gerenciador de Relatórios
                max_attempts = 3
                relatorio_window = None
                
                for attempt in range(max_attempts):
                    try:
                        relatorio_window = main_window.child_window(
                            title="Gerenciador de Relatórios",
                            class_name="FNWND3190"
                        )
                        
                        if relatorio_window.exists():
                            break
                    except Exception:
                        if attempt == max_attempts - 1:
                            self.log("Janela 'Gerenciador de Relatórios' não encontrada após várias tentativas.")
                            return False
                        time.sleep(1)
                
                if not relatorio_window or not relatorio_window.exists():
                    self.log("Janela 'Gerenciador de Relatórios' não encontrada.")
                    return False

                self.log("Gerenciador de Relatórios localizado")
                
                # Usar UIA para encontrar o elemento
                rel_app = Application(backend='uia').connect(handle=relatorio_window.handle)
                tree = rel_app.window(class_name="FNWND3190").child_window(class_name="PBTreeView32_100")
                
                # Encontrar e clicar nos itens
                try:
                    self.log("Localizando itens de menu para Folha de Ponto")
                    folha_ponto = tree.child_window(title="Folha - Ponto")
                    
                    if folha_ponto.exists():
                        folha_ponto.set_focus()
                        folha_ponto.click_input()
                        time.sleep(0.5)
                        
                        # Expandir o nó se necessário
                        folha_ponto.click_input(double=True)
                        time.sleep(1)
                        
                        # Localizar o relatório específico
                        folha_21_20 = tree.child_window(title="Folha de Ponto_21 a 20 - II")
                        if folha_21_20.exists():
                            folha_21_20.click_input(double=True)
                            time.sleep(1)
                        else:
                            self.log("Item 'Folha de Ponto_21 a 20 - II' não encontrado na árvore.")
                            return False
                    else:
                        self.log("Item 'Folha - Ponto' não encontrado na árvore.")
                        return False
                    
                except Exception as e:
                    self.log(f"Erro ao clicar nos itens da árvore de relatórios: {str(e)}")
                    return False

                # Navegação por tab e preenchimento
                self.log("Preenchendo os campos de data")
                send_keys('{TAB}*')
                time.sleep(0.3)
                send_keys('{TAB}' + str(row['data inicio']))
                time.sleep(0.3)
                send_keys('{TAB}' + str(row['data final']))
                time.sleep(0.5)

                # Clica em executar
                self.log("Clicando em executar relatório")
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()
                time.sleep(2)

                # Clica na maleta (ícone publicação)
                self.log("Clicando no ícone de publicação")
                main_window.set_focus()
                button_publicacao = main_window.child_window(auto_id="1005", class_name="FNUDO3190")
                button_publicacao.click_input()
                time.sleep(1)

                # Gerenciar janela de publicação
                try:
                    # Encontrar a janela de publicação de documentos
                    pub_doc_window = main_window.child_window(
                        title="Publicação de Documentos",
                        class_name="FNWNS3190"
                    )
                    
                    if not pub_doc_window.exists():
                        self.log("Janela 'Publicação de Documentos' não encontrada.")
                        return False
                    
                    self.log("Janela de publicação localizada")

                    # Escolher pasta de destino
                    combo_box = pub_doc_window.child_window(auto_id="1007", class_name="ComboBox")
                    combo_box.click_input()
                    time.sleep(0.5)
                    send_keys("Pessoal/Folha de Ponto{ENTER}")
                    time.sleep(0.5)

                    # Campo do nome do pdf
                    nome_pdf = str(row['nome pdf'])
                    self.log(f"Definindo nome do PDF: {nome_pdf}")
                    edit_field = pub_doc_window.child_window(auto_id="1014", class_name="Edit")
                    edit_field.set_text(nome_pdf)
                    time.sleep(0.5)

                    # Botão gravar
                    self.log("Clicando em gravar")
                    button_gravar = pub_doc_window.child_window(auto_id="1016", class_name="Button")
                    button_gravar.click_input()
                    time.sleep(0.5)
                    
                    # Esperar até que a janela não esteja mais visível ou timeout
                    self.wait_and_check_window_closed(pub_doc_window, "Publicação de Documentos")
                    
                    # Gerar PDF
                    self.log("Gerando PDF")
                    button_pdf = main_window.child_window(auto_id="1014", class_name="FNUDO3190")
                    button_pdf.click_input()
                    time.sleep(1)

                    # Salvar PDF
                    try:
                         # Esperar até que a janela de salvamento esteja visível
                        wait_until(timeout=10, retry_interval=0.5, func=lambda: main_window.child_window(
                            title="Salvar em PDF",
                            class_name="#32770"
                        ).exists())

                        # Encontrar a janela de salvamento do pdf
                        save_window = main_window.child_window(
                            title="Salvar em PDF",
                            class_name="#32770"
                        )
                        
                        if save_window.exists():
                            self.log("Janela de salvamento PDF localizada")
                            
                            # Definir nome do arquivo
                            time.sleep(0.5)
                            name_field = save_window.child_window(auto_id="1148", class_name="Edit")
                            name_field.set_text(nome_pdf)
                            time.sleep(0.5)

                            # Botão salvar
                            self.log("Salvando PDF")
                            button_salvar = save_window.child_window(auto_id="1", class_name="Button")
                            button_salvar.click_input()
                            time.sleep(3)
                    except Exception as e:
                        self.log(f"Erro ao interagir com janela de salvamento: {str(e)}")
                        return False

                    # Fechar janelas
                    self.log("Fechando janelas")
                    send_keys('^w')  # Ctrl+W
                    time.sleep(1)
                
                except Exception as e:
                    self.log(f"Erro na publicação: {str(e)}")
                    return False

                # Primeiro ESC
                main_window.set_focus()
                send_keys('{ESC}')
                time.sleep(1)
                
                # Segundo ESC
                main_window.set_focus()
                send_keys('{ESC}')
                time.sleep(1)
                
                # Verificar se ainda existe alguma janela aberta que deveria estar fechada
                try:
                    relatorio_window = main_window.child_window(
                        title="Gerenciador de Relatórios",
                        class_name="FNWND3190"
                    )
                    
                    if relatorio_window.exists() and relatorio_window.is_visible():
                        self.log("Janela de relatórios ainda aberta, enviando ESC adicional")
                        main_window.set_focus()
                        send_keys('{ESC}')
                        time.sleep(1)
                except Exception:
                    pass
                
            except Exception as e:
                self.log(f"Erro ao interagir com o Gerenciador de Relatórios: {str(e)}")
                self.log(f"Detalhes do erro: {traceback.format_exc()}")
                return False
            
            self.log(f"Linha {index + 1} processada com sucesso")
            return True

        except Exception as e:
            self.log(f"Erro ao processar linha {index + 1}: {str(e)}")
            self.log(f"Detalhes do erro: {traceback.format_exc()}")
            return False

def main():
    gui = AutomacaoGUI()
    gui.executar()

if __name__ == "__main__":
    main()