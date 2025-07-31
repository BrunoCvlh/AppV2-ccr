import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

from controllers.selenium_automation import SeleniumHandler
from controllers.tratamento_da_planilha import tratar_planilha
from controllers.ajuste_contingencial import ajustar_contingencial
from controllers.inclui_dados_na_base import incluir_dados_na_base

from views.page_one import PageOne
from views.page_two import PageTwo

class AtenaCommanderApp:
    def __init__(self, root: tk.Tk):

        self.root = root
        self.root.title("Coordenação de Controladoria - GCO")
        self.root.geometry("700x670")

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)

        self.selenium_handler = SeleniumHandler()
        self.tratamento_file_path = ""
        self.planilha_path = ""
        self.base_path = ""
        self.tratada_path = ""

        self.main_title_label = tk.Label(root, text="Automação de Relatórios e Bases",
                                         font=("Arial", 14, "bold"), fg="#1A2F4B")
        self.main_title_label.grid(row=0, column=0, pady=20, sticky="nsew")

        self.container = ttk.Frame(root)
        self.container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.loading_label = tk.Label(root, text="Aguarde enquanto a operação é processada...\nNão feche o programa até a conclusão.",
                                      font=("Arial", 10, "italic"), fg="blue", wraplength=700, justify="center")
        self.loading_label.grid(row=2, column=0, sticky="nsew", padx=20, pady=5)
        self.loading_label.grid_forget()

        self.status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.status_label.grid(row=3, column=0, sticky="nsew", padx=20, pady=5)

        self.frames = {}
        for F in (PageOne, PageTwo):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("PageOne")

    def show_frame(self, page_name: str):

        frame = self.frames[page_name]
        frame.tkraise()
        self.loading_label.grid_forget()
        self.status_label.config(text="")

    def set_loading_state(self, is_loading: bool, message: str = ""):

        if is_loading:
            self.loading_label.config(text=message or "Aguarde enquanto a operação é processada...\nNão feche o programa até a conclusão.")
            self.loading_label.grid(row=2, column=0, sticky="nsew", padx=20, pady=5)
            self.status_label.config(text="")
        else:
            self.loading_label.grid_forget()

    def set_status_message(self, message: str, color: str = "green"):
        self.status_label.config(text=message, fg=color)

if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()