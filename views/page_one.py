import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading
import calendar

from controllers.selenium_automation import SeleniumHandler
from controllers.tratamento_da_planilha import tratar_planilha

class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        self.access_frame = ttk.LabelFrame(self, text=" 1º Passo: Acesso e Download de Relatório ", padding=(10, 10))
        self.access_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.access_frame.grid_columnconfigure(0, weight=0)
        self.access_frame.grid_columnconfigure(1, weight=1)
        self.access_frame.grid_columnconfigure(2, weight=0)
        self.access_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.access_frame, text="Email:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_entry = tk.Entry(self.access_frame, font=("Arial", 10), bd=2, relief="groove")
        self.email_entry.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.access_frame, text="Senha:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.password_entry = tk.Entry(self.access_frame, font=("Arial", 10), show="*", bd=2, relief="groove")
        self.password_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.access_frame, text="Competência (MM/YYYY):", font=("Arial", 10)).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.competencia_mes_entry = tk.Entry(self.access_frame, font=("Arial", 10), width=3, bd=2, relief="groove", justify="center")
        self.competencia_mes_entry.grid(row=2, column=1, sticky="e", padx=(5, 2), pady=5)
        self.competencia_mes_entry.config(validate="key", validatecommand=(self.controller.root.register(lambda P: len(P) <= 2 and (P.isdigit() or P == "")), "%P"))

        tk.Label(self.access_frame, text="/", font=("Arial", 10, "bold")).grid(row=2, column=2, sticky="w", pady=5)

        self.competencia_ano_entry = tk.Entry(self.access_frame, font=("Arial", 10), width=5, bd=2, relief="groove", justify="center")
        self.competencia_ano_entry.grid(row=2, column=3, sticky="w", padx=(2, 5), pady=5)
        self.competencia_ano_entry.config(validate="key", validatecommand=(self.controller.root.register(lambda P: len(P) <= 4 and (P.isdigit() or P == "")), "%P"))

        self.extract_report_button = tk.Button(self.access_frame, text="Extrair Relatório de Análise Comparativa", command=self.start_login_thread,
                                                font=("Arial", 10, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.extract_report_button.grid(row=3, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.treatment_frame = ttk.LabelFrame(self, text=" 2º Passo: Tratamento do Arquivo ", padding=(10, 10))
        self.treatment_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.treatment_frame.grid_columnconfigure(0, weight=0)
        self.treatment_frame.grid_columnconfigure(1, weight=1)
        self.treatment_frame.grid_columnconfigure(2, weight=0)
        self.treatment_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.treatment_frame, text="Competência (MM/YYYY):", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.tratamento_competencia_mes_entry = tk.Entry(self.treatment_frame, font=("Arial", 10), width=3, bd=2, relief="groove", justify="center")
        self.tratamento_competencia_mes_entry.grid(row=0, column=1, sticky="e", padx=(5, 2), pady=5)
        self.tratamento_competencia_mes_entry.config(validate="key", validatecommand=(self.controller.root.register(lambda P: len(P) <= 2 and (P.isdigit() or P == "")), "%P"))
        tk.Label(self.treatment_frame, text="/", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w", pady=5)
        self.tratamento_competencia_ano_entry = tk.Entry(self.treatment_frame, font=("Arial", 10), width=5, bd=2, relief="groove", justify="center")
        self.tratamento_competencia_ano_entry.grid(row=0, column=3, sticky="w", padx=(2, 5), pady=5)
        self.tratamento_competencia_ano_entry.config(validate="key", validatecommand=(self.controller.root.register(lambda P: len(P) <= 4 and (P.isdigit() or P == "")), "%P"))

        tk.Label(self.treatment_frame, text="Anexar Arquivo(s) para Tratamento:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.tratamento_file_path_label = tk.Label(self.treatment_frame, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.tratamento_file_path_label.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tratamento_file_button = tk.Button(self.treatment_frame, text="Selecionar Arquivo(s)", command=self.select_tratamento_file,
                                                 font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.tratamento_file_button.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        self.tratamento_button = tk.Button(self.treatment_frame, text="Tratar Arquivo(s)", command=self.tratar_arquivo,
                                            font=("Arial", 10, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.tratamento_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.navigation_frame = tk.Frame(self)
        self.navigation_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.next_button = tk.Button(self.navigation_frame, text="Próximo >>",
                                      command=lambda: self.controller.show_frame("PageTwo"),
                                      font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", relief="raised", bd=2, cursor="hand2", width=15)
        self.next_button.pack(side="right", padx=10)

    def select_tratamento_file(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if file_paths:
            self.controller.tratamento_file_paths = file_paths
            num_files = len(file_paths)
            self.tratamento_file_path_label.config(text=f"{num_files} arquivo(s) selecionado(s)")
        else:
            self.controller.tratamento_file_paths = []
            self.tratamento_file_path_label.config(text="Nenhum arquivo selecionado")

    def _set_buttons_state(self, state: str):
        self.extract_report_button.config(state=state)
        self.tratamento_button.config(state=state)
        self.tratamento_file_button.config(state=state)
        self.next_button.config(state=state)

    def start_login_thread(self):
        self.controller.set_status_message("")
        self.controller.set_loading_state(True, "Realizando login e download do relatório...")
        self._set_buttons_state(tk.DISABLED)

        login_thread = threading.Thread(target=self._login_in_thread)
        login_thread.start()

    def _login_in_thread(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        mes = self.competencia_mes_entry.get()
        ano = self.competencia_ano_entry.get()

        if not all([email, password, mes, ano]):
            messagebox.showerror("Erro de Validação", "Preencha o Email, Senha e Competência para realizar a extração.")
            self.controller.set_loading_state(False)
            self._set_buttons_state(tk.NORMAL)
            return

        try:
            success = self.controller.selenium_handler._login(email, password, mes, ano)
            self.controller.set_loading_state(False)
            if success:
                self.controller.set_status_message("Download da análise comparativa concluído!", "green")
            else:
                self.controller.set_status_message("Erro durante o login ou navegação. Verifique as credenciais e a competência.", "red")
        except Exception as e:
            self.controller.set_loading_state(False)
            self.controller.set_status_message(f"Ocorreu um erro inesperado: {e}", "red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")
        finally:
            self.controller.selenium_handler.quit_driver()
            self._set_buttons_state(tk.NORMAL)

    def tratar_arquivo(self):
        self.controller.set_status_message("")
        if not hasattr(self.controller, 'tratamento_file_paths') or not self.controller.tratamento_file_paths:
            messagebox.showerror("Erro", "Selecione um ou mais arquivos para tratamento.")
            return

        mes = self.tratamento_competencia_mes_entry.get()
        ano = self.tratamento_competencia_ano_entry.get()

        if not all([mes, ano]):
            messagebox.showerror("Erro de Validação", "Preencha a competência (Mês/Ano) para o tratamento do arquivo.")
            return

        competencia = f"{mes}/{ano}"

        self.controller.set_loading_state(True, "Tratando o(s) arquivo(s)...")
        self._set_buttons_state(tk.DISABLED)

        treatment_thread = threading.Thread(target=self._tratar_arquivo_in_thread, args=(competencia, self.controller.tratamento_file_paths))
        treatment_thread.start()

    def _tratar_arquivo_in_thread(self, competencia: str, file_paths: list):
        erros_encontrados = []
        arquivos_tratados = []
        for file_path in file_paths:
            try:
                tratado_path = tratar_planilha(file_path, competencia)
                if tratado_path.startswith("ERRO:"):
                    erros_encontrados.append(f"Arquivo '{file_path.split('/')[-1]}': {tratado_path}")
                else:
                    arquivos_tratados.append(tratado_path)
            except Exception as e:
                erros_encontrados.append(f"Arquivo '{file_path.split('/')[-1]}': Erro inesperado - {e}")
        
        self.controller.set_loading_state(False)
        self._set_buttons_state(tk.NORMAL)

        if erros_encontrados:
            error_message = "Erros encontrados no tratamento dos seguintes arquivos:\n" + "\n".join(erros_encontrados)
            messagebox.showerror("Erros de Tratamento", error_message)
            self.controller.set_status_message("Erro(s) no tratamento de arquivo(s).", "red")
        
        if arquivos_tratados:
            success_message = "Tratamento de arquivo(s) realizado com sucesso para:\n" + "\n".join([f"- {path.split('/')[-1]}" for path in arquivos_tratados])
            messagebox.showinfo("Tratamento Concluído", success_message)
            self.controller.set_status_message("Arquivo(s) tratado(s) com sucesso!", "green")
            
            if arquivos_tratados:
                self.controller.tratada_path = arquivos_tratados[-1] 
                if "PageTwo" in self.controller.frames and hasattr(self.controller.frames["PageTwo"], 'tratada_path_label'):
                    self.controller.frames["PageTwo"].tratada_path_label.config(text=f"Última planilha tratada: {self.controller.tratada_path.split('/')[-1] if '/' in self.controller.tratada_path else self.controller.tratada_path.split('\\')[-1]}")

        if not erros_encontrados and not arquivos_tratados:
            messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado ou tratado.")
            self.controller.set_status_message("Nenhum arquivo selecionado para tratamento.", "orange")