#importando bibliotecas
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font, simpledialog
import pandas as pd
import os
import sys

class ConciliaFacil:
    def __init__(self, root):
        self.root = root
        self.root.title("Concilia Fácil - Suporta Qualquer Formato")
        self.root.geometry("500x300")
        self.root.configure(bg="#f0f0f0")

        # Verificar dependências
        self.verificar_dependencias()

        # Configurar fonte grande
        self.fonte_grande = font.Font(size=12)

        # Variáveis
        self.arquivo1 = ""
        self.arquivo2 = ""
        self.df1 = None
        self.df2 = None
        self.colunas_selecionadas = []

        # Criar interface
        self.criar_interface()

    def verificar_dependencias(self):
        try:
            import pandas
            import openpyxl
        except ImportError as e:
            messagebox.showerror(
                "Erro de Dependência",
                f"Faltam bibliotecas necessárias. Instale-as com:\n\n"
                f"pip install pandas openpyxl\n\n"
                f"Erro: {str(e)}"
            )
            sys.exit(1)

    def criar_interface(self):
        # Título
        titulo = ttk.Label(
            self.root,
            text=" Concilia Fácil - Suporta Qualquer Formato",
            font=("Arial", 16, "bold"),
            background="#f0f0f0"
        )
        titulo.pack(pady=20)

        # Passo 1: Carregar Planilha 1
        frame_passo1 = ttk.Frame(self.root)
        frame_passo1.pack(pady=10)
        ttk.Label(frame_passo1, text="1. Clique para escolher a PRIMEIRA planilha:", font=self.fonte_grande).pack(side="left")
        btn_plan1 = ttk.Button(frame_passo1, text="Abrir Planilha 1", command=lambda: self.carregar_arquivo(1))
        btn_plan1.pack(side="left", padx=10)

        # Passo 2: Carregar Planilha 2
        frame_passo2 = ttk.Frame(self.root)
        frame_passo2.pack(pady=10)
        ttk.Label(frame_passo2, text="2. Clique para escolher a SEGUNDA planilha:", font=self.fonte_grande).pack(side="left")
        btn_plan2 = ttk.Button(frame_passo2, text="Abrir Planilha 2", command=lambda: self.carregar_arquivo(2))
        btn_plan2.pack(side="left", padx=10)

        # Passo 3: Selecionar Colunas
        frame_passo3 = ttk.Frame(self.root)
        frame_passo3.pack(pady=10)
        ttk.Label(frame_passo3, text="3. Selecione as colunas para comparar:", font=self.fonte_grande).pack(side="left")
        btn_selecionar_colunas = ttk.Button(frame_passo3, text="Selecionar Colunas", command=self.selecionar_colunas)
        btn_selecionar_colunas.pack(side="left", padx=10)

        # Passo 4: Botão Mágico
        btn_conciliar = ttk.Button(
            self.root,
            text=" CLIQUE AQUI PARA CONCILIAR! ",
            command=self.conciliar,
            style="Big.TButton"
        )
        btn_conciliar.pack(pady=30)
        self.root.style = ttk.Style()
        self.root.style.configure("Big.TButton", font=("Arial", 14, "bold"), foreground="green")

        # Status
        self.status = ttk.Label(self.root, text="", background="#f0f0f0")
        self.status.pack()

    def carregar_arquivo(self, num):
        arquivo = filedialog.askopenfilename(
            title=f"Selecione a Planilha {num}",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            if num == 1:
                self.arquivo1 = arquivo
                self.df1 = self.ler_arquivo(arquivo)
            else:
                self.arquivo2 = arquivo
                self.df2 = self.ler_arquivo(arquivo)
            self.status.config(text=f" Planilha {num} carregada!")

    def ler_arquivo(self, caminho):
        # Pergunta ao usuário qual é a linha do cabeçalho
        linha_cabecalho = simpledialog.askinteger(
            "Linha do Cabeçalho",
            "Qual é a linha do cabeçalho? (1 para a primeira linha)",
            parent=self.root
        )
        if not linha_cabecalho:
            linha_cabecalho = 1  # Padrão: primeira linha

        # Verifica a extensão do arquivo
        extensao = os.path.splitext(caminho)[1].lower()

        try:
            if extensao in [".xlsx", ".xls"]:
                return pd.read_excel(caminho, header=linha_cabecalho - 1)
            elif extensao == ".csv":
                delimitador = simpledialog.askstring(
                    "Delimitador",
                    "Qual é o delimitador do arquivo CSV? (Ex: , ou ;)",
                    parent=self.root
                )
                if not delimitador:
                    delimitador = ","  # Padrão: vírgula
                return pd.read_csv(caminho, delimiter=delimitador, header=linha_cabecalho - 1)
            elif extensao == ".txt":
                delimitador = simpledialog.askstring(
                    "Delimitador",
                    "Qual é o delimitador do arquivo TXT? (Ex: tabulação ou ,)",
                    parent=self.root
                )
                if not delimitador:
                    delimitador = "\t"  # Padrão: tabulação
                return pd.read_csv(caminho, delimiter=delimitador, header=linha_cabecalho - 1)
            elif extensao == ".ods":
                return pd.read_excel(caminho, engine="odf", header=linha_cabecalho - 1)
            elif extensao == ".json":
                return pd.read_json(caminho)
            else:
                messagebox.showerror(
                    "Formato não suportado",
                    f"O formato '{extensao}' não é suportado. Use .xlsx, .xls, .csv, .txt, .ods ou .json."
                )
                return None
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler o arquivo",
                f"Não foi possível ler o arquivo. Verifique o formato e o conteúdo.\nErro: {str(e)}"
            )
            return None

    def selecionar_colunas(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showerror("Erro", "Carregue as duas planilhas primeiro!")
            return

        colunas_comuns = list(set(self.df1.columns) & set(self.df2.columns))
        if not colunas_comuns:
            messagebox.showerror("Erro", "Não há colunas iguais nas planilhas!")
            return

        self.colunas_selecionadas = simpledialog.askstring("Selecionar Colunas", f"Digite as colunas a serem comparadas, separadas por vírgula:\n{', '.join(colunas_comuns)}", parent=self.root)
        if self.colunas_selecionadas:
            self.colunas_selecionadas = [col.strip() for col in self.colunas_selecionadas.split(",")]

    def conciliar(self):
        if not self.arquivo1 or not self.arquivo2:
            messagebox.showerror("Ops!", "Carregue as duas planilhas primeiro!")
            return

        if not self.colunas_selecionadas:
            messagebox.showerror("Erro", "Selecione as colunas para comparar!")
            return

        try:
            # Merge das planilhas
            merged = pd.merge(self.df1, self.df2, on=self.colunas_selecionadas, how="outer", suffixes=("_Plan1", "_Plan2"), indicator=True)

            # Imprimir colunas do DataFrame mesclado
            print("Colunas do DataFrame mesclado:", merged.columns)

            # Identificar divergências
            divergencias = merged[merged["_merge"] != "both"]

            # Comparar valores nas colunas selecionadas
            diferencas_valores = pd.DataFrame()
            for coluna in self.colunas_selecionadas:
                diferencas = merged[~merged.apply(lambda row: abs(row[f"{coluna}_Plan1"] - row[f"{coluna}_Plan2"]) <= 0.01, axis=1)]
                diferencas_valores = pd.concat([diferencas_valores, diferencas])

            # Criar relatório
            relatorio_path = os.path.join(os.path.expanduser("~"), "Desktop", "RelatorioConciliação.xlsx")
            with pd.ExcelWriter(relatorio_path) as writer:
                if not divergencias.empty:
                    divergencias.to_excel(writer, sheet_name="Divergências", index=False)
                if not diferencas_valores.empty:
                    diferencas_valores.to_excel(writer, sheet_name="Valores Diferentes", index=False)

            # Abrir relatório automaticamente
            os.startfile(relatorio_path)
            self.status.config(text=" Pronto! O relatório foi salvo na ÁREA DE TRABALHO!")

        except Exception as e:
            messagebox.showerror("Erro", f"Algo deu errado. Verifique o formato do arquivo!\nErro: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConciliaFacil(root)
    root.mainloop()