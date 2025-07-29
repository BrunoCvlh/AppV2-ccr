import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

# Importando módulos da pasta controllers
from controllers.selenium_automation import SeleniumHandler
from controllers.tratamento_da_planilha import tratar_planilha
from controllers.ajuste_contingencial import ajustar_contingencial
from controllers.inclui_dados_na_base import incluir_dados_na_base

# Importando as classes de página da pasta views
from views.page_one import PageOne
from views.page_two import PageTwo

class AtenaCommanderApp:
    def __init__(self, root: tk.Tk):
        """
        Inicializa a janela principal da aplicação e configura a navegação entre as páginas.

        Args:
            root (tk.Tk): A janela raiz do Tkinter.
        """
        self.root = root
        self.root.title("Coordenação de Controladoria - GCO")
        self.root.geometry("700x550")

        # Configura o grid para a janela principal
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=0) # Para o título principal
        self.root.grid_rowconfigure(1, weight=1) # Para o contêiner das páginas
        self.root.grid_rowconfigure(2, weight=0) # Para os rótulos de carregamento/status

        # Inicializa as variáveis de estado compartilhadas da aplicação
        self.selenium_handler = SeleniumHandler()
        self.tratamento_file_path = ""
        self.planilha_path = ""
        self.base_path = ""
        self.tratada_path = "" # Isso armazenará o caminho do arquivo tratado da PageOne

        # Rótulo do título principal
        self.main_title_label = tk.Label(root, text="Automação de Relatórios e Bases",
                                         font=("Arial", 14, "bold"), fg="#1A2F4B")
        self.main_title_label.grid(row=0, column=0, pady=20, sticky="nsew")

        # Frame contêiner para as páginas
        self.container = ttk.Frame(root)
        self.container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # Rótulos de carregamento e status (compartilhados entre as páginas)
        self.loading_label = tk.Label(root, text="Aguarde enquanto a operação é processada...\nNão feche o programa até a conclusão.",
                                      font=("Arial", 10, "italic"), fg="blue", wraplength=700, justify="center")
        self.loading_label.grid(row=2, column=0, sticky="nsew", padx=20, pady=5)
        self.loading_label.grid_forget() # Ocultar inicialmente

        self.status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.status_label.grid(row=3, column=0, sticky="nsew", padx=20, pady=5)

        # Dicionário para armazenar as diferentes páginas
        self.frames = {}
        for F in (PageOne, PageTwo):
            page_name = F.__name__
            # Passa o contêiner e uma referência a esta instância principal do aplicativo (controlador)
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            # Coloca todos os frames um em cima do outro
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("PageOne") # Mostra a primeira página inicialmente

    def show_frame(self, page_name: str):
        """
        Mostra o frame da página especificada e oculta os outros.

        Args:
            page_name (str): O nome da página a ser exibida (ex: "PageOne", "PageTwo").
        """
        frame = self.frames[page_name]
        frame.tkraise()
        # Garante que os rótulos de carregamento e status estejam ocultos ao trocar de página
        self.loading_label.grid_forget()
        self.status_label.config(text="")

    def set_loading_state(self, is_loading: bool, message: str = ""):
        """
        Gerencia a visibilidade e o texto do rótulo de carregamento.

        Args:
            is_loading (bool): True para mostrar carregamento, False para ocultar.
            message (str): Mensagem opcional para exibir no rótulo de carregamento.
        """
        if is_loading:
            self.loading_label.config(text=message or "Aguarde enquanto a operação é processada...\nNão feche o programa até a conclusão.")
            self.loading_label.grid(row=2, column=0, sticky="nsew", padx=20, pady=5)
            self.status_label.config(text="") # Limpa o status ao carregar
        else:
            self.loading_label.grid_forget()

    def set_status_message(self, message: str, color: str = "green"):
        """
        Define o texto e a cor do rótulo de status.

        Args:
            message (str): A mensagem de status a ser exibida.
            color (str): A cor da mensagem de status (ex: "green", "red").
        """
        self.status_label.config(text=message, fg=color)

if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()