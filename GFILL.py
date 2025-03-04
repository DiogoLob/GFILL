import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from fpdf import FPDF  # Biblioteca para gerar PDFs

class ProductionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Produção Industrial - v2.0")
        self.root.geometry("1200x800")
        self.setup_styles()
        self.load_data()
        self.create_widgets()
        self.setup_bindings()
        self.update_ui()
        
    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('Helvetica', 10), padding=6)
        self.style.configure('Header.TLabel', font=('Helvetica', 14, 'bold'))
        self.style.configure('Stats.TLabel', font=('Helvetica', 12, 'bold'))
        self.style.map('TButton', 
                      foreground=[('active', '!disabled', 'white'), ('!active', 'black')],
                      background=[('active', '#0052cc'), ('!active', '#4a90e2')])

    def load_data(self):
        self.file_name = "producao.xlsx"
        try:
            self.df_producao = pd.read_excel(self.file_name, sheet_name="Producao")
            self.df_manutencao = pd.read_excel(self.file_name, sheet_name="Manutencao")
        except FileNotFoundError:
            self.df_producao = pd.DataFrame(columns=["NS", "Data", "Hora"])
            self.df_manutencao = pd.DataFrame(columns=["NS", "Status", "Data", "Hora"])

    def save_data(self):
        with pd.ExcelWriter(self.file_name, engine="xlsxwriter") as writer:
            self.df_producao.to_excel(writer, sheet_name="Producao", index=False)
            self.df_manutencao.to_excel(writer, sheet_name="Manutencao", index=False)

    def create_widgets(self):
        self.create_notebook()
        self.create_production_tab()
        self.create_maintenance_tab()
        self.create_status_bar()

    def create_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both')

    def create_production_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Controle de Produção")

        # Painel esquerdo
        left_panel = ttk.Frame(tab)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

        self.create_input_section(left_panel)
        self.create_filter_section(left_panel)
        self.create_quick_stats(left_panel)

        # Painel direito
        right_panel = ttk.Frame(tab)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.create_production_table(right_panel)
        self.create_chart_section(right_panel)

    def create_input_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Registro de Produção", padding=10)
        frame.pack(fill=tk.X, pady=5)

        ttk.Label(frame, text="Número de Série:").grid(row=0, column=0, sticky=tk.W)
        self.entry_ns_producao = ttk.Entry(frame, width=25)
        self.entry_ns_producao.grid(row=0, column=1, padx=5)

        btn_produzir = ttk.Button(frame, text="Registrar Produção", command=self.produzir)
        btn_produzir.grid(row=0, column=2, padx=5)

        btn_personalizado = ttk.Button(frame, text="Inserir Personalizado", command=self.abrir_janela_personalizada)
        btn_personalizado.grid(row=1, column=0, columnspan=3, pady=5)

        btn_relatorio = ttk.Button(frame, text="Gerar Relatório", command=self.gerar_relatorio)
        btn_relatorio.grid(row=2, column=0, columnspan=3, pady=5)

        self.lbl_counter = ttk.Label(frame, text="Máquinas Produzidas: 0", style='Stats.TLabel')
        self.lbl_counter.grid(row=3, column=0, columnspan=3, pady=5)

    def gerar_relatorio(self):
        # Configurações gerais do PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_margins(15, 15, 15)  # Margens equilibradas
        pdf.set_auto_page_break(True, margin=15)
        
        # Borda estilizada
        pdf.set_line_width(0.5)
        pdf.rect(5, 5, 200, 287)  # Moldura fina

        # ---- CABEÇALHO ---- #
        # Logo
        logo_path = "logoall.jpg"
        pdf.image(logo_path, x=20, y=12, w=25)
        
        # Título principal
        pdf.set_font("Arial", 'B', 18)
        pdf.set_xy(0, 15)
        pdf.cell(0, 10, "Report Diário", 0, 1, 'C')
        
        # Data formatada
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(100, 100, 100)  # Cinza profissional
        hoje = datetime.now().strftime("%d/%m/%Y")
        pdf.cell(0, 5, f"Emitido em: {hoje}", 0, 1, 'C')
        pdf.ln(8)

        # ---- SEÇÃO DE RESUMO ---- #
        pdf.set_line_width(0.1)  # Linhas ultra finas
        
        # Tabela de Resumo
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Resumo Operacional", 0, 1)
        
        dados_resumo = [
            ("Máquinas Hoje", len(self.df_producao[self.df_producao["Data"] == datetime.now().strftime("%Y-%m-%d")])),
            ("Total Produção", len(self.df_producao)),
            ("Manutenções Hoje", len(self.df_manutencao[self.df_manutencao['Data'] == datetime.now().strftime("%Y-%m-%d")])),
            ("Total Manutenções", len(self.df_manutencao))
        ]
        
        # Cabeçalho da tabela
        pdf.set_fill_color(240, 240, 240)  # Fundo cinza claro
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(80, 8, "Indicador", 1, 0, 'C', 1)
        pdf.cell(40, 8, "Valor", 1, 1, 'C', 1)
        
        # Dados da tabela
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(50, 50, 50)  # Cinza escuro
        for indicador, valor in dados_resumo:
            pdf.cell(80, 8, indicador, 1, 0, 'L')
            pdf.cell(40, 8, str(valor), 1, 1, 'C')
        
        pdf.ln(10)

        # ---- IMAGEM DA MÁQUINA ---- #
        imagem_maquina_path = "maquina.png"
        pdf.image(imagem_maquina_path, x=150, y=46, w=41)  # Posicionamento preciso

        # ---- TABELA DE PRODUÇÃO ---- #
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Produção do Dia", 0, 1)
        pdf.set_font("Arial", '', 9)  # Fonte menor
        
        # Cabeçalho da tabela
        colunas = ["NS", "Data", "Hora"]
        larguras = [60, 50, 40]
        
        pdf.set_fill_color(240, 240, 240)
        for i, col in enumerate(colunas):
            pdf.cell(larguras[i], 7, col, 1, 0, 'C', 1)
        pdf.ln()
        
        # Dados da tabela
        df_producao_hoje = self.df_producao[self.df_producao["Data"] == datetime.now().strftime("%Y-%m-%d")]
        for _, row in df_producao_hoje.iterrows():
            pdf.cell(larguras[0], 7, str(row["NS"]), 1, 0, 'L')
            pdf.cell(larguras[1], 7, str(row["Data"]), 1, 0, 'C')
            pdf.cell(larguras[2], 7, str(row["Hora"]), 1, 1, 'C')
        
        pdf.ln(12)

        # ---- GRÁFICO DE PRODUÇÃO ---- #
        if not self.df_producao.empty:
            # Geração do gráfico
            plt.figure(figsize=(8, 3))  # Formato mais alongado
            self.df_producao['Data'] = pd.to_datetime(self.df_producao['Data'])
            daily_production = self.df_producao.groupby('Data').size()
            
            plt.bar(daily_production.index, daily_production.values, color='#4A90E2', width=0.8)
            plt.xlabel('')
            plt.ylabel('Unidades')
            plt.gca().spines['top'].set_visible(False)
            plt.gca().spines['right'].set_visible(False)
            plt.grid(axis='y', linestyle='--', alpha=0.7)
            
            grafico_path = "grafico_producao.png"
            plt.savefig(grafico_path, bbox_inches='tight', dpi=150)
            plt.close()

            # Inserção no PDF
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 8, "Desempenho de Produção", 0, 1)
            pdf.image(grafico_path, x=20, y=pdf.get_y(), w=170)  # Gráfico alinhado

        # ---- RODAPÉ ---- #
        pdf.set_y(265)
        pdf.set_font("Arial", 'I', 8)
        pdf.set_text_color(150, 150, 150)
        pdf.cell(0, 5, "Relatório gerado automaticamente pelo Sistema de Gestão de Produção", 0, 0, 'C')

        # Salvamento final
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            pdf.output(file_path)
            messagebox.showinfo("Sucesso", f"Relatório salvo em: {file_path}")

    def abrir_janela_personalizada(self):
        top = tk.Toplevel(self.root)
        top.title("Inserção Personalizada")
        top.geometry("300x200")

        ttk.Label(top, text="Número de Série:").grid(row=0, column=0, padx=5, pady=5)
        entry_ns = ttk.Entry(top, width=20)
        entry_ns.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(top, text="Data:").grid(row=1, column=0, padx=5, pady=5)
        entry_data = DateEntry(top, width=12, date_pattern='yyyy-mm-dd')
        entry_data.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(top, text="Hora:").grid(row=2, column=0, padx=5, pady=5)
        entry_hora = ttk.Entry(top, width=20)
        entry_hora.grid(row=2, column=1, padx=5, pady=5)

        def salvar_personalizado():
            ns = entry_ns.get().strip()
            data = entry_data.get()
            hora = entry_hora.get().strip()

            if not ns or not data or not hora:
                messagebox.showwarning("Aviso", "Preencha todos os campos!")
                return

            if ns in self.df_producao["NS"].values:
                messagebox.showwarning("Aviso", "Número de série já registrado!")
                return

            novo_registro = pd.DataFrame([{
                "NS": ns,
                "Data": data,
                "Hora": hora
            }])

            self.df_producao = pd.concat([self.df_producao, novo_registro], ignore_index=True)
            self.save_data()
            self.update_ui()
            top.destroy()
            messagebox.showinfo("Sucesso", "Produção personalizada registrada com sucesso!")

        btn_salvar = ttk.Button(top, text="Salvar", command=salvar_personalizado)
        btn_salvar.grid(row=3, column=0, columnspan=2, pady=10)

    def create_quick_stats(self, parent):
        frame = ttk.LabelFrame(parent, text="Estatísticas Rápidas", padding=10)
        frame.pack(fill=tk.BOTH, pady=5)

        self.stats_labels = []
        stats = [
            ("Média Diária", "0"),
            ("Último NS", "-"),
            ("Máquinas em Manutenção", "0"),
            ("Manutenções Hoje", "0")
        ]

        for i, (label, value) in enumerate(stats):
            row = ttk.Frame(frame)
            row.pack(fill=tk.X, pady=2)
            
            ttk.Label(row, text=label+":", width=20, anchor=tk.W).pack(side=tk.LEFT)
            lbl_value = ttk.Label(row, text=value, style='Stats.TLabel')
            lbl_value.pack(side=tk.LEFT)
            self.stats_labels.append(lbl_value)

    def filtrar_producao(self):
        try:
            data_inicio = self.entry_data_inicio.get()
            data_fim = self.entry_data_fim.get()

            if data_inicio and data_fim:
                df_filtrado = self.df_producao[
                    (self.df_producao["Data"] >= data_inicio) & 
                    (self.df_producao["Data"] <= data_fim)
                ]
                self.update_table(self.tree_producao, df_filtrado)
                self.lbl_counter.config(text=f"Máquinas Filtradas: {len(df_filtrado)}")
            else:
                messagebox.showwarning("Aviso", "Selecione ambas as datas para filtrar!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao filtrar: {str(e)}")

    def create_filter_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Filtros Avançados", padding=10)
        frame.pack(fill=tk.X, pady=5)

        ttk.Label(frame, text="Data Início:").grid(row=0, column=0)
        self.entry_data_inicio = DateEntry(frame, width=12, date_pattern='yyyy-mm-dd')
        self.entry_data_inicio.grid(row=0, column=1, padx=5)

        ttk.Label(frame, text="Data Fim:").grid(row=0, column=2)
        self.entry_data_fim = DateEntry(frame, width=12, date_pattern='yyyy-mm-dd')
        self.entry_data_fim.grid(row=0, column=3, padx=5)

        btn_filtrar = ttk.Button(frame, text="Aplicar Filtro", command=self.filtrar_producao)
        btn_filtrar.grid(row=0, column=4, padx=5)

        btn_export = ttk.Button(frame, text="Exportar CSV", command=self.export_csv)
        btn_export.grid(row=0, column=5, padx=5)

    def create_production_table(self, parent):
        frame = ttk.LabelFrame(parent, text="Histórico de Produção", padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        columns = ("NS", "Data", "Hora")
        self.tree_producao = ttk.Treeview(frame, columns=columns, show="headings", selectmode='browse')
        
        for col in columns:
            self.tree_producao.heading(col, text=col)
            self.tree_producao.column(col, width=100)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree_producao.yview)
        self.tree_producao.configure(yscroll=scrollbar.set)
        
        self.tree_producao.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def create_chart_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Desempenho de Produção", padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        self.fig = plt.Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_maintenance_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Controle de Manutenção")

        # Painel esquerdo
        left_panel = ttk.Frame(tab)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

        # Seção de registro de manutenção
        frame = ttk.LabelFrame(left_panel, text="Registro de Manutenção", padding=10)
        frame.pack(fill=tk.X, pady=5)

        ttk.Label(frame, text="Número de Série:").grid(row=0, column=0)
        self.entry_ns_manutencao = ttk.Entry(frame, width=25)
        self.entry_ns_manutencao.grid(row=0, column=1, padx=5)

        self.var_status = tk.StringVar(value="Estoque")
        ttk.Radiobutton(frame, text="Estoque", variable=self.var_status, value="Estoque").grid(row=1, column=0)
        ttk.Radiobutton(frame, text="Produção", variable=self.var_status, value="Produção").grid(row=1, column=1)

        btn_registrar = ttk.Button(frame, text="Registrar Manutenção", command=self.registrar_manutencao)
        btn_registrar.grid(row=2, column=0, columnspan=2, pady=5)

        # Painel direito
        right_panel = ttk.Frame(tab)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tabela de manutenção
        frame = ttk.LabelFrame(right_panel, text="Histórico de Manutenção", padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        columns = ("NS", "Status", "Data", "Hora")
        self.tree_manutencao = ttk.Treeview(frame, columns=columns, show="headings", selectmode='browse')
        
        for col in columns:
            self.tree_manutencao.heading(col, text=col)
            self.tree_manutencao.column(col, width=100)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree_manutencao.yview)
        self.tree_manutencao.configure(yscroll=scrollbar.set)
        
        self.tree_manutencao.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Botão de liberação
        btn_liberar = ttk.Button(frame, text="Liberar Manutenção", command=self.liberar_manutencao)
        btn_liberar.pack(pady=5)

    def liberar_manutencao(self):
        selected = self.tree_manutencao.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma máquina para liberar!")
            return

        for item in selected:
            ns = self.tree_manutencao.item(item, "values")[0]
            
            # Remove da manutenção
            self.df_manutencao = self.df_manutencao[self.df_manutencao["NS"] != ns]
            
            # Adiciona de volta à produção
            data_atual = datetime.now().strftime("%Y-%m-%d")
            hora_atual = datetime.now().strftime("%H:%M:%S")
            novo_registro = pd.DataFrame([{
                "NS": ns,
                "Data": data_atual,
                "Hora": hora_atual
            }])
            self.df_producao = pd.concat([self.df_producao, novo_registro], ignore_index=True)

        self.save_data()
        self.update_ui()
        messagebox.showinfo("Sucesso", "Máquina(s) liberada(s) com sucesso!")

    def create_status_bar(self):
        self.status_bar = ttk.Label(self.root, text="Pronto", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_bindings(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.tree_producao.bind("<Double-1>", self.on_item_double_click)

    def update_ui(self):
        self.update_table(self.tree_producao, self.df_producao)
        self.update_table(self.tree_manutencao, self.df_manutencao)
        self.update_chart()
        self.update_stats()
        self.lbl_counter.config(text=f"Máquinas Produzidas: {len(self.df_producao)}")

    def update_table(self, tree, df):
        tree.delete(*tree.get_children())
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def update_stats(self):
        if not self.df_producao.empty:
            # Calcula estatísticas de produção
            daily_avg = len(self.df_producao) / self.df_producao['Data'].nunique()
            last_ns = self.df_producao['NS'].iloc[-1]
            maintenance_count = len(self.df_manutencao)
            
            # Calcula manutenções do dia
            hoje = datetime.now().strftime("%Y-%m-%d")
            manutencoes_hoje = len(self.df_manutencao[self.df_manutencao['Data'] == hoje])
            
            # Atualiza labels
            self.stats_labels[0].config(text=f"{daily_avg:.1f}")
            self.stats_labels[1].config(text=last_ns)
            self.stats_labels[2].config(text=str(maintenance_count))
            self.stats_labels[3].config(text=str(manutencoes_hoje))

    def update_chart(self):
        if not self.df_producao.empty:
            self.ax.clear()
            self.df_producao['Data'] = pd.to_datetime(self.df_producao['Data'])
            daily_production = self.df_producao.groupby('Data').size()
            
            self.ax.bar(daily_production.index, daily_production.values)
            self.ax.set_xlabel('Data')
            self.ax.set_ylabel('Quantidade Produzida')
            self.ax.set_title('Produção Diária')
            self.ax.grid(True)
            
            self.canvas.draw()

    def produzir(self):
        ns = self.entry_ns_producao.get().strip()
        if not ns:
            messagebox.showwarning("Aviso", "Digite o número de série!")
            return

        if ns in self.df_producao["NS"].values:
            messagebox.showwarning("Aviso", "Número de série já registrado!")
            return

        data_atual = datetime.now().strftime("%Y-%m-%d")
        hora_atual = datetime.now().strftime("%H:%M:%S")

        novo_registro = pd.DataFrame([{
            "NS": ns,
            "Data": data_atual,
            "Hora": hora_atual
        }])

        self.df_producao = pd.concat([self.df_producao, novo_registro], ignore_index=True)
        self.save_data()
        self.update_ui()
        self.entry_ns_producao.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Produção registrada com sucesso!")

    def registrar_manutencao(self):
        ns = self.entry_ns_manutencao.get().strip()
        if not ns:
            messagebox.showwarning("Aviso", "Digite o número de série!")
            return

        if ns not in self.df_producao["NS"].values:
            messagebox.showwarning("Aviso", "Número de série não encontrado na produção!")
            return

        # Remove da produção
        self.df_producao = self.df_producao[self.df_producao["NS"] != ns]

        data_atual = datetime.now().strftime("%Y-%m-%d")
        hora_atual = datetime.now().strftime("%H:%M:%S")
        status = self.var_status.get()

        novo_registro = pd.DataFrame([{
            "NS": ns,
            "Status": status,
            "Data": data_atual,
            "Hora": hora_atual
        }])

        self.df_manutencao = pd.concat([self.df_manutencao, novo_registro], ignore_index=True)
        self.save_data()
        self.update_ui()
        self.entry_ns_manutencao.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Manutenção registrada com sucesso!")

    def on_item_double_click(self, event):
        item = self.tree_producao.selection()[0]
        values = self.tree_producao.item(item, 'values')
        
        top = tk.Toplevel(self.root)
        top.title("Editar Registro")
        
        ttk.Label(top, text="Número de Série:").grid(row=0, column=0)
        entry_ns = ttk.Entry(top)
        entry_ns.grid(row=0, column=1)
        entry_ns.insert(0, values[0])
        
        ttk.Label(top, text="Data:").grid(row=1, column=0)
        entry_data = DateEntry(top, width=12, date_pattern='yyyy-mm-dd')
        entry_data.grid(row=1, column=1)
        entry_data.set_date(values[1])
        
        ttk.Label(top, text="Hora:").grid(row=2, column=0)
        entry_hora = ttk.Entry(top)
        entry_hora.grid(row=2, column=1)
        entry_hora.insert(0, values[2])
        
        def salvar_edicao():
            self.df_producao.at[self.df_producao["NS"] == values[0], "NS"] = entry_ns.get()
            self.df_producao.at[self.df_producao["NS"] == values[0], "Data"] = entry_data.get()
            self.df_producao.at[self.df_producao["NS"] == values[0], "Hora"] = entry_hora.get()
            self.save_data()
            self.update_ui()
            top.destroy()
            messagebox.showinfo("Sucesso", "Registro atualizado!")
        
        ttk.Button(top, text="Salvar", command=salvar_edicao).grid(row=3, column=0, columnspan=2)

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            self.df_producao.to_csv(file_path, index=False)
            self.status_bar.config(text=f"Arquivo exportado: {file_path}")

    def on_close(self):
        if messagebox.askokcancel("Sair", "Deseja realmente sair?"):
            self.save_data()
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductionApp(root)
    root.mainloop()