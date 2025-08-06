import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading

from controllers.consolidar_planilhas import consolidar_planilhas
from controllers.inclui_dados_na_base import incluir_dados_na_base

from .tooltip import Tooltip

class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self._root = controller.root

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        # 3º Passo: Consolidar planilhas
        self.additional_files_frame = ttk.LabelFrame(self, text=" 4º Passo: Consolidar planilhas ", padding=(10, 10))
        self.additional_files_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.additional_files_frame.grid_columnconfigure(0, weight=0)
        self.additional_files_frame.grid_columnconfigure(1, weight=1)
        self.additional_files_frame.grid_columnconfigure(2, weight=0)
        self.additional_files_frame.grid_columnconfigure(3, weight=1)

        # Adicionar o botão de interrogação e o tooltip para o 3º Passo
        info_button_step3 = tk.Label(self.additional_files_frame, text="?", font=("Arial", 8, "bold"), fg="blue", cursor="hand2")
        info_button_step3.grid(row=0, column=4, sticky="w", padx=2, pady=5)
        Tooltip(info_button_step3, "Selecione as 3 planilhas que serão consolidadas em um único arquivo. Segurar Ctrl e selecionar os 3 arquivos tratados.")

        tk.Label(self.additional_files_frame, text="Anexar até 3 Planilhas:", font=("Arial", 8)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.additional_files_path_label = tk.Label(self.additional_files_frame, text="Nenhum arquivo selecionado", font=("Arial", 8), fg="gray")
        self.additional_files_path_label.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.additional_files_button = tk.Button(self.additional_files_frame, text="Selecionar Arquivo(s)", command=self.select_additional_files,
                                                 font=("Arial", 8), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.additional_files_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        self.consolidar_button = tk.Button(self.additional_files_frame, text="Consolidar Dados", command=self.consolidar_dados,
                                           font=("Arial", 8, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.consolidar_button.grid(row=1, column=0, columnspan=4, sticky="", padx=5, pady=10)


        # 4º Passo: Enviar Dados para a Base
        self.send_data_frame = ttk.LabelFrame(self, text=" 5º Passo: Enviar Dados para a Base ", padding=(10, 10))
        self.send_data_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=6, sticky="nsew")
        self.send_data_frame.grid_columnconfigure(0, weight=0)
        self.send_data_frame.grid_columnconfigure(1, weight=1)
        self.send_data_frame.grid_columnconfigure(2, weight=1)
        self.send_data_frame.grid_columnconfigure(3, weight=1)

        # Adicionar o botão de interrogação e o tooltip para o 4º Passo
        info_button_step4 = tk.Label(self.send_data_frame, text="?", font=("Arial", 8, "bold"), fg="blue", cursor="hand2")
        info_button_step4.grid(row=0, column=4, sticky="w", padx=2, pady=5)
        Tooltip(info_button_step4, "Neste passo, você enviará os dados da planilha consolidada e tratada "
                                   "para a base de dados central. Certifique-se de que a base e a planilha "
                                   "estejam corretas antes de prosseguir. Local da base: FS_GCO_CCR - Documentos/2025/03 - Relatórios/3 - Relatório Gerencial da Gerência de Controladoria/Total - 2025/Base Geral - 2025")


        tk.Label(self.send_data_frame, text="Caminho da Base de Dados:", font=("Arial", 8)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.base_path_display = tk.Label(self.send_data_frame, text="Nenhum caminho selecionado", font=("Arial", 8), fg="gray")
        self.base_path_display.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.base_path_button = tk.Button(self.send_data_frame, text="Procurar Base", command=self.select_base_path,
                                          font=("Arial", 8), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.base_path_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.send_data_frame, text="Planilha Tratada para Envio:", font=("Arial", 8)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.tratada_path_label = tk.Label(self.send_data_frame, text="Nenhuma planilha tratada selecionada", font=("Arial", 8), fg="gray")
        self.tratada_path_label.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tratada_path_button = tk.Button(self.send_data_frame, text="Selecionar Planilha Tratada", command=self.select_tratada_file,
                                             font=("Arial", 8), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.tratada_path_button.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        self.envio_button = tk.Button(self.send_data_frame, text="Enviar Dados para a Base", command=self.enviar_dados_base,
                                       font=("Arial", 8, "bold"), bg="#004662", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.envio_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.navigation_frame = tk.Frame(self)
        self.navigation_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.back_button = tk.Button(self.navigation_frame, text="<< Voltar",
                                     command=lambda: self.controller.show_frame("PageOne"),
                                     font=("Arial", 8, "bold"), bg="#FF5722", fg="white", relief="raised", bd=2, cursor="hand2", width=15)
        self.back_button.pack(side="left", padx=10)

        self.finish_button = tk.Button(self.navigation_frame, text="Finalizar",
                                       command=self.finish_application,
                                       font=("Arial", 8, "bold"), bg="#009688", fg="white", relief="raised", bd=2, cursor="hand2", width=15)
        self.finish_button.pack(side="right", padx=10)

        if self.controller.tratada_path:
            file_name = self.controller.tratada_path.split('/')[-1] if '/' in self.controller.tratada_path else self.controller.tratada_path.split('\\')[-1]
            self.tratada_path_label.config(text=f"Planilha tratada: {file_name}")
        else:
            self.tratada_path_label.config(text="Nenhuma planilha tratada selecionada")

    def select_additional_files(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_paths:
            if len(file_paths) > 3:
                messagebox.showwarning("Limite de Arquivos", "Você pode selecionar no máximo 3 planilhas.")
                self.controller.additional_file_paths = file_paths[:3]
                self.additional_files_path_label.config(text=f"3 arquivo(s) selecionado(s) (limite excedido, 3 primeiros)")
            else:
                self.controller.additional_file_paths = file_paths
                num_files = len(file_paths)
                self.additional_files_path_label.config(text=f"{num_files} arquivo(s) selecionado(s)")
        else:
            self.controller.additional_file_paths = []
            self.additional_files_path_label.config(text="Nenhum arquivo selecionado")

    def select_base_path(self):
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
        self.consolidar_button.config(state=state)
        self.additional_files_button.config(state=state)
        self.envio_button.config(state=state)
        self.base_path_button.config(state=state)
        self.tratada_path_button.config(state=state)
        self.back_button.config(state=state)
        self.finish_button.config(state=state)

    def consolidar_dados(self):
        self.controller.set_status_message("")
        if not hasattr(self.controller, 'additional_file_paths') or not self.controller.additional_file_paths:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para consolidação. Por favor, selecione até 3 planilhas no 3º passo.")
            self.controller.set_status_message("Nenhum arquivo selecionado para consolidação.", "orange")
            return

        arquivos_para_consolidar = self.controller.additional_file_paths

        self.controller.set_loading_state(True, "Consolidando as planilhas...")
        self._set_buttons_state(tk.DISABLED)

        consolidation_thread = threading.Thread(target=self._consolidar_dados_in_thread, args=(arquivos_para_consolidar,))
        consolidation_thread.start()

    def _consolidar_dados_in_thread(self, arquivos_para_consolidar: list):
        try:
            output_file_name = "planilha_consolidada.xlsx"
            consolidated_path = consolidar_planilhas(arquivos_para_consolidar, output_file_name)

            self.controller.set_loading_state(False)
            self._set_buttons_state(tk.NORMAL)

            if consolidated_path.startswith("ERRO:"):
                messagebox.showerror("Erro de Consolidação", consolidated_path)
                self.controller.set_status_message("Erro na consolidação das planilhas.", "red")
            else:
                messagebox.showinfo("Consolidação Concluída", f"Planilhas consolidadas com sucesso em: {consolidated_path}")
                self.controller.set_status_message(f"Planilhas consolidadas em: {consolidated_path}", "green")
        except Exception as e:
            self.controller.set_loading_state(False)
            self._set_buttons_state(tk.NORMAL)
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado durante a consolidação: {e}")
            self.controller.set_status_message(f"Erro inesperado na consolidação: {e}", "red")

    def enviar_dados_base(self):
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
        if messagebox.askyesno("Finalizar", "Tem certeza que deseja finalizar a aplicação?"):
            self.controller.root.quit()