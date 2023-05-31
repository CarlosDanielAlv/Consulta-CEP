import os
import pandas as pd
import requests
import tkinter as tk
from tkinter import messagebox, filedialog
from threading import Thread


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Preenchimento de Endereço")
        self.geometry("400x150")
        self.configure(bg="#F2F2F2")  # Cor de fundo elegante

        # Centralizar o formulário na tela
        self.center_window()

        self.label = tk.Label(self, text="Validar Planilha:", font=("Helvetica", 14), bg="#F2F2F2")
        self.label.pack(pady=10)

        self.button_frame = tk.Frame(self, bg="#F2F2F2")  # Widget de preenchimento para centralizar os botões
        self.button_frame.pack()

        self.button_validate = tk.Button(self.button_frame, text="Validar", font=("Helvetica", 12), command=self.validate_file, bg="#008CBA", fg="white")
        self.button_validate.pack(side=tk.LEFT, padx=10, pady=10)
        self.button_validate.configure(width=12)  # Definir largura do botão

        self.button_open = tk.Button(self.button_frame, text="Abrir Planilha", font=("Helvetica", 12), command=self.open_file, bg="#008CBA", fg="white")
        self.button_open.pack(side=tk.LEFT, padx=10, pady=10)
        self.button_open.configure(width=12)  # Definir largura do botão

        self.progress_label = tk.Label(self, text="", font=("Helvetica", 12), bg="#F2F2F2")
        self.progress_label.pack()

        # Centralizar elementos
        self.label.pack(anchor=tk.CENTER)
        self.button_frame.pack(anchor=tk.CENTER)
        self.progress_label.pack(anchor=tk.CENTER)

    def center_window(self):
        # Obter as dimensões da tela
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calcular as coordenadas x e y para centralizar o formulário
        x = int((screen_width - self.winfo_reqwidth()) / 2)
        y = int((screen_height - self.winfo_reqheight()) / 2)

        # Definir a posição do formulário
        self.geometry(f"+{x}+{y}")

    def validate_file(self):
        file_path = os.path.join(os.getcwd(), "CEPS.xlsx")
        if os.path.exists(file_path):
            try:
                planilha = pd.read_excel(file_path, dtype={'CEP': str})
                if self.validate_columns(planilha):
                    if self.validate_ceps(planilha):
                        cep_count = len(planilha)
                        message = f"A planilha contém {cep_count} CEP(s). Deseja iniciar a pesquisa?"
                        if messagebox.askyesno("Validação Concluída", message):

                            self.button_validate.configure(state="disabled")
                            self.update_progress()
                            self.preencher_dados_api_na_planilha(planilha)

                    else:
                        messagebox.showerror("Erro de Validação", "A coluna CEP está vazia.")
                else:
                    messagebox.showerror("Erro de Validação", "A planilha não possui todas as colunas necessárias.")
            except PermissionError:
                messagebox.showerror("Erro", "A planilha está aberta. Por favor, feche-a e tente novamente.")
        else:
            messagebox.showerror("Erro", "Arquivo da planilha não encontrado.")

    def validate_columns(self, planilha):
        required_columns = {"CEP", "Endereço", "Bairro", "Cidade", "Estado"}
        return set(planilha.columns) == required_columns

    def validate_ceps(self, planilha):
        return "CEP" in planilha.columns and not planilha["CEP"].isnull().all()

    def update_progress(self):
        self.progress_label.configure(text="Preenchendo dados...")
        self.button_validate.configure(state="disabled")

    def consultar_cep(self, cep):
        try:
            url = f'https://viacep.com.br/ws/{cep}/json/'
            response = requests.get(url)
            if response.status_code == 200:
                return response.json()
            else:
                return None
        except requests.exceptions.RequestException as e:
            print(f"Ocorreu um erro na consulta do CEP {cep}: {e}")
            return None

    def preencher_dados_api_na_planilha(self, planilha):
        try:
            for index, row in planilha.iterrows():
                cep = row['CEP']

                if len(str(cep)) == 8 and str(cep).isdigit():
                    resultado = self.consultar_cep(cep)

                    if resultado is not None:
                        planilha.at[index, 'Endereço'] = resultado.get('logradouro', '')
                        planilha.at[index, 'Bairro'] = resultado.get('bairro', '')
                        planilha.at[index, 'Cidade'] = resultado.get('localidade', '')
                        planilha.at[index, 'Estado'] = resultado.get('uf', '')
                        self.progress_label.configure(text=f"Preenchimento do CEP {cep}")
                else:
                    planilha.at[index, 'Endereço'] = ''
                    planilha.at[index, 'Bairro'] = ''
                    planilha.at[index, 'Cidade'] = ''
                    planilha.at[index, 'Estado'] = ''

            planilha.to_excel("CEPS.xlsx", index=False)
            messagebox.showinfo("Concluído", "Preenchimento de dados concluído com sucesso!")
            self.button_validate.configure(state="normal")
            self.progress_label.configure(text="Preenchimento de dados concluído.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante o preenchimento dos dados: {e}")

    def open_file(self):
        file_path = os.path.join(os.getcwd(), "CEPS.xlsx")
        if os.path.exists(file_path):
            try:
                os.startfile(file_path)
            except Exception:
                messagebox.showerror("Erro", "Não foi possível abrir a planilha.")
        else:
            messagebox.showerror("Erro", "Arquivo da planilha não encontrado.")


if __name__ == "__main__":
    app = Application()
    app.mainloop()
