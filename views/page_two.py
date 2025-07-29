import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading

# Importando módulos da pasta controllers
from controllers.ajuste_contingencial import ajustar_contingencial
from controllers.inclui_dados_na_base import incluir_dados_na_base

class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        """
        Inicializa a Página Dois, contendo os dois últimos passos da aplicação.

        Args:
            parent (ttk.Frame): O frame contêiner da aplicação principal.
            controller (AtenaCommanderApp): Referência à instância principal da aplicação.
        """
        super().__init__(parent)
        self.controller = controller # Adicionado para acessar o controlador principal
        self._root = controller.root # Mantido para compatibilidade com a validação se necessário, mas controller.root é melhor

        # Configura o grid para este frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0) # Contingency Frame
        self.grid_rowconfigure(1, weight=1) # Send Data Frame
        self.grid_rowconfigure(2, weight=0) # Navigation buttons

        # --- 3º Passo: Ajustar Valor Contingencial ---
        self.contingency_frame = ttk.LabelFrame(self, text=" 3º Passo: Ajustar Valor Contingencial ", padding=(10, 10))
        self.contingency_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.contingency_frame.grid_columnconfigure(0, weight=0)
        self.contingency_frame.grid_columnconfigure(1, weight=1)
        self.contingency_frame.grid_columnconfigure(2, weight=1)
        self.contingency_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.contingency_frame, text="Anexar Planilha para Ajuste:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.planilha_path_label = tk.Label(self.contingency_frame, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.planilha_path_label.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.planilha_button = tk.Button(self.contingency_frame, text="Selecionar Arquivo", command=self.select_planilha,
                                         font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.planilha_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.contingency_frame, text="Novo Valor Contingencial:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.contingencia_entry = tk.Entry(self.contingency_frame, font=("Arial", 10), bd=2, relief="groove")
        self.contingencia_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        self.adjust_contingency_button = tk.Button(self.contingency_frame, text="Ajustar Valor Contingencial na Planilha", command=self.ajustar_valor_contingencial,
                                                   font=("Arial", 10, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.adjust_contingency_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        # --- 4º Passo: Enviar Dados para a Base ---
        self.send_data_frame = ttk.LabelFrame(self, text=" 4º Passo: Enviar Dados para a Base ", padding=(10, 10))
        self.send_data_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.send_data_frame.grid_columnconfigure(0, weight=0)
        self.send_data_frame.grid_columnconfigure(1, weight=1)
        self.send_data_frame.grid_columnconfigure(2, weight=1)
        self.send_data_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.send_data_frame, text="Caminho da Base de Dados:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.base_path_display = tk.Label(self.send_data_frame, text="Nenhum caminho selecionado", font=("Arial", 10), fg="gray")
        self.base_path_display.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.base_path_button = tk.Button(self.send_data_frame, text="Procurar Base", command=self.select_base_path,
                                          font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.base_path_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.send_data_frame, text="Planilha Tratada para Envio:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.tratada_path_label = tk.Label(self.send_data_frame, text="Nenhuma planilha tratada selecionada", font=("Arial", 10), fg="gray")
        self.tratada_path_label.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tratada_path_button = tk.Button(self.send_data_frame, text="Selecionar Planilha Tratada", command=self.select_tratada_file,
                                             font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.tratada_path_button.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        self.envio_button = tk.Button(self.send_data_frame, text="Enviar Dados para a Base", command=self.enviar_dados_base,
                                       font=("Arial", 10, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.envio_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        # Botões de Navegação
        self.navigation_frame = tk.Frame(self)
        self.navigation_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.back_button = tk.Button(self.navigation_frame, text="<< Voltar",
                                     command=lambda: self.controller.show_frame("PageOne"),
                                     font=("Arial", 10, "bold"), bg="#FF5722", fg="white", relief="raised", bd=2, cursor="hand2", width=15)
        self.back_button.pack(side="left", padx=10)

        self.finish_button = tk.Button(self.navigation_frame, text="Finalizar",
                                       command=self.finish_application,
                                       font=("Arial", 10, "bold"), bg="#009688", fg="white", relief="raised", bd=2, cursor="hand2", width=15)
        self.finish_button.pack(side="right", padx=10)

        # Ao inicializar a PageTwo, atualize o label com o caminho da planilha tratada, se existir
        # Isso garante que a informação passada da PageOne seja exibida
        if self.controller.tratada_path:
            file_name = self.controller.tratada_path.split('/')[-1] if '/' in self.controller.tratada_path else self.controller.tratada_path.split('\\')[-1]
            self.tratada_path_label.config(text=f"Planilha tratada: {file_name}")
        else:
            self.tratada_path_label.config(text="Nenhuma planilha tratada selecionada")


    def select_planilha(self):
        """Abre uma caixa de diálogo para selecionar a planilha para ajuste de contingência."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.controller.planilha_path = file_path
            self.planilha_path_label.config(text=f"Arquivo selecionado: {file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]}")
        else:
            self.controller.planilha_path = ""
            self.planilha_path_label.config(text="Nenhum arquivo selecionado")

    def select_base_path(self):
        """Abre uma caixa de diálogo para selecionar o arquivo do banco de dados."""
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo da base",
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.controller.base_path = file_path
            nome_arquivo = file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]
            self.base_path_display.config(text=f"Caminho selecionado: {nome_arquivo}")
        else:
            self.controller.base_path = ""
            self.base_path_display.config(text="Nenhum caminho selecionado")

    def select_tratada_file(self):
        """Abre uma caixa de diálogo para selecionar a planilha tratada."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.controller.tratada_path = file_path
            self.tratada_path_label.config(text=f"Planilha tratada: {file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]}")
        else:
            self.controller.tratada_path = ""
            self.tratada_path_label.config(text="Nenhuma planilha tratada selecionada")

    def _set_buttons_state(self, state: str):
        """Função auxiliar para definir o estado de todos os botões nesta página."""
        self.adjust_contingency_button.config(state=state)
        self.planilha_button.config(state=state)
        self.envio_button.config(state=state)
        self.base_path_button.config(state=state)
        self.tratada_path_button.config(state=state)
        self.back_button.config(state=state)
        self.finish_button.config(state=state)

    def ajustar_valor_contingencial(self):
        """Inicia o processo de ajuste do valor contingencial."""
        self.controller.set_status_message("")
        if not self.controller.planilha_path:
            messagebox.showerror("Erro", "Selecione a planilha antes de ajustar o valor contingencial.")
            return
        novo_valor = self.contingencia_entry.get()
        if not novo_valor:
            messagebox.showerror("Erro", "Preencha o valor contingencial.")
            return

        self.controller.set_loading_state(True, "Ajustando valor contingencial...")
        self._set_buttons_state(tk.DISABLED)

        adjust_thread = threading.Thread(target=self._ajustar_valor_contingencial_in_thread, args=(novo_valor,))
        adjust_thread.start()

    def _ajustar_valor_contingencial_in_thread(self, novo_valor: str):
        """Executa o ajuste do valor contingencial em uma thread separada."""
        try:
            resultado = ajustar_contingencial(self.controller.planilha_path, novo_valor)
            self.controller.set_loading_state(False)
            if resultado is True:
                messagebox.showinfo("Contingencial", "Valor contingencial ajustado com sucesso!")
                self.controller.set_status_message("Valor contingencial ajustado!", "green")
            else:
                messagebox.showerror("Erro", f"Erro ao ajustar valor contingencial:\n{resultado}")
                self.controller.set_status_message("Erro ao ajustar valor contingencial.", "red")
        except Exception as e:
            self.controller.set_loading_state(False)
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao ajustar valor contingencial: {e}")
            self.controller.set_status_message(f"Erro inesperado no ajuste: {e}", "red")
        finally:
            self._set_buttons_state(tk.NORMAL)

    def enviar_dados_base(self):
        """Inicia o processo de envio de dados para o banco de dados."""
        self.controller.set_status_message("")
        if not self.controller.base_path:
            messagebox.showerror("Erro", "Selecione o caminho da base antes de enviar os dados.")
            return
        if not self.controller.tratada_path:
            messagebox.showerror("Erro", "Selecione a planilha tratada.")
            return

        self.controller.set_loading_state(True, "Enviando dados para a base...")
        self._set_buttons_state(tk.DISABLED)

        send_thread = threading.Thread(target=self._enviar_dados_base_in_thread)
        send_thread.start()

    def _enviar_dados_base_in_thread(self):
        """Executa a operação de envio de dados em uma thread separada."""
        try:
            resultado = incluir_dados_na_base(self.controller.tratada_path, self.controller.base_path)
            self.controller.set_loading_state(False)
            if resultado is True:
                messagebox.showinfo("Envio", "Dados incluídos na base com sucesso!")
                self.controller.set_status_message("Dados incluídos na base!", "green")
            else:
                messagebox.showerror("Erro", f"Erro ao incluir dados na base:\n{resultado}")
                self.controller.set_status_message("Erro ao incluir dados na base.", "red")
        except Exception as e:
            self.controller.set_loading_state(False)
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao enviar dados para a base: {e}")
            self.controller.set_status_message(f"Erro inesperado no envio: {e}", "red")
        finally:
            self._set_buttons_state(tk.NORMAL)

    def finish_application(self):
        """Fecha a aplicação."""
        if messagebox.askyesno("Finalizar", "Tem certeza que deseja finalizar a aplicação?"):
            self.controller.root.quit()