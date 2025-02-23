import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
from tkcalendar import DateEntry  # Biblioteca para calendário

# Criando ou carregando os dados
file_name = "producao.xlsx"

try:
    df_producao = pd.read_excel(file_name, sheet_name="Producao")
    df_manutencao = pd.read_excel(file_name, sheet_name="Manutencao")
except FileNotFoundError:
    df_producao = pd.DataFrame(columns=["NS", "Data", "Hora"])
    df_manutencao = pd.DataFrame(columns=["NS", "Status", "Data", "Hora"])

def salvar_dados():
    with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
        df_producao.to_excel(writer, sheet_name="Producao", index=False)
        df_manutencao.to_excel(writer, sheet_name="Manutencao", index=False)

def atualizar_contador():
    lbl_produzidas.config(text=f"{len(df_producao)}")

def produzir():
    global df_producao  # Declare a variável global no início da função

    ns = entry_ns_producao.get().strip()  # Remove espaços extras
    if not ns:
        messagebox.showwarning("Aviso", "Digite o número de NS!")
        return

    # Verifica se o NS já foi produzido
    if ns in df_producao["NS"].values:
        messagebox.showerror("Erro", f"A máquina com NS {ns} já foi produzida! Escolha outro número.")
        return

    data = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")

    # Adiciona a nova máquina à tabela de produção
    novo_registro = pd.DataFrame([{"NS": ns, "Data": data, "Hora": hora}])
    df_producao = pd.concat([df_producao, novo_registro], ignore_index=True)

    salvar_dados()
    atualizar_tabelas()
    atualizar_contador()
    calcular_media_tempo()

    messagebox.showinfo("Sucesso", "Máquina produzida com sucesso!")

def registrar_manutencao():
    global df_producao  # Declara a variável global antes de usá-la
    global df_manutencao  # Também declara df_manutencao como global

    ns = entry_ns_manutencao.get().strip()  # Remove espaços extras
    status = "Produção" if var_producao.get() else "Estoque"

    if not ns:
        messagebox.showwarning("Aviso", "Digite o número de NS!")
        return

    # Verifica se o NS já foi produzido
    if ns in df_producao["NS"].values:
        if not var_producao.get():  # Se a caixa de confirmação não estiver marcada
            messagebox.showerror(
                "Erro",
                f"A máquina com NS {ns} já foi produzida!\n"
                "Para registrar a manutenção, marque a caixa de confirmação."
            )
            return  # Impede o registro até que a caixa seja marcada

    # Registra na tabela de manutenção
    data = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")

    df_manutencao = pd.concat(
        [df_manutencao, pd.DataFrame([{"NS": ns, "Status": status, "Data": data, "Hora": hora}])],
        ignore_index=True
    )

    # Se for produção, remove da tabela de produção
    if var_producao.get():
        df_producao = df_producao[df_producao.NS != ns]

    salvar_dados()
    atualizar_tabelas()

    messagebox.showinfo("Sucesso", "Máquina registrada na manutenção!")

def atualizar_tabelas():
    for row in tree_producao.get_children():
        tree_producao.delete(row)
    for _, row in df_producao.iterrows():
        tree_producao.insert("", "end", values=list(row))
    for row in tree_manutencao.get_children():
        tree_manutencao.delete(row)
    for _, row in df_manutencao.iterrows():
        tree_manutencao.insert("", "end", values=list(row))
    atualizar_contador()
    calcular_media_tempo()

def liberar_manutencao():
    selected = tree_manutencao.selection()
    if selected:
        for item in selected:
            ns = tree_manutencao.item(item, "values")[0]  # Obtém o NS corretamente
            global df_manutencao, df_producao
            df_manutencao = df_manutencao[df_manutencao["NS"] != ns]  # Remove da manutenção
            df_producao = pd.concat([df_producao, pd.DataFrame([{"NS": ns, "Data": datetime.now().strftime("%Y-%m-%d"),
                                                                     "Hora": datetime.now().strftime("%H:%M:%S")}])],
                                     ignore_index=True)  # Adiciona na produção

        salvar_dados()
        atualizar_tabelas()
        messagebox.showinfo("Sucesso", "Máquina liberada e contabilizada na produção!")
    else:
        messagebox.showwarning("Aviso", "Selecione uma máquina para liberar!")

def filtrar_producao():
    try:
        data_inicio = entry_data_inicio.get()
        data_fim = entry_data_fim.get()

        # Filtra pelo intervalo de datas
        df_filtrado = df_producao[
            (df_producao["Data"] >= data_inicio) & (df_producao["Data"] <= data_fim)
        ]

        # Atualiza a tabela de produção apenas com os dados filtrados
        for row in tree_producao.get_children():
            tree_producao.delete(row)

        for _, row in df_filtrado.iterrows():
            tree_producao.insert("", "end", values=list(row))

        # Atualiza o contador de máquinas filtradas
        lbl_resultado_filtro.config(text=f"Máquinas no período: {len(df_filtrado)}")
        calcular_media_tempo()
        atualizar_contador()

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao filtrar: {e}")

# New: Calculate average production time
def calcular_media_tempo():
    if not df_producao.empty:
        df_producao['Data Hora'] = pd.to_datetime(df_producao['Data'] + ' ' + df_producao['Hora'])
        tempo_producao = (df_producao['Data Hora'].max() - df_producao['Data Hora'].min()).total_seconds() / 60  # in minutes
        num_maquinas = len(df_producao)
        media_tempo = tempo_producao / num_maquinas if num_maquinas else 0
        lbl_media_tempo.config(text=f"{media_tempo:.2f} min")
    else:
        lbl_media_tempo.config(text="0 min")


# GUI setup
root = tk.Tk()
root.title("Controle de Produção")
root.geometry("950x700") # Aumentando a largura para melhor visualização

# Notebook (tabs)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

frame_producao = ttk.Frame(notebook)
frame_manutencao = ttk.Frame(notebook)
notebook.add(frame_producao, text="Produção")
notebook.add(frame_manutencao, text="Manutenção")

# Frame principal para layout da produção
frame_producao_Geral = ttk.Frame(frame_producao) # Frame master para organizar tudo na aba Produção
frame_producao_Geral.pack(expand=True, fill="both", padx=10, pady=10)

# Div 1: Título
div_titulo = ttk.Frame(frame_producao_Geral)
div_titulo.pack(pady=10, fill="x")
ttk.Label(div_titulo, text="CONTROLE DE PRODUÇÃO", font=("Arial", 16, "bold")).pack(pady=10)
ttk.Separator(div_titulo).pack(fill='x', pady=5)

# Div 2: Entrada de NS e Botão Produzir (Horizontal)
div_entrada_produzir = ttk.Frame(frame_producao_Geral)
div_entrada_produzir.pack(pady=10, fill="x")

    # Div 2.1: Entrada NS (Esquerda)
div_entrada_ns = ttk.Frame(div_entrada_produzir)
div_entrada_ns.pack(side=tk.LEFT, padx=10, pady=10, anchor='w') # Alinhado à esquerda
ttk.Label(div_entrada_ns, text="Entry NS:").pack(anchor='w')
entry_ns_producao = ttk.Entry(div_entrada_ns, width=20)
entry_ns_producao.pack(anchor='w')

    # Div 2.2: Botão Produzir (Centro)
div_botao_produzir = ttk.Frame(div_entrada_produzir)
div_botao_produzir.pack(side=tk.LEFT, padx=20, pady=10) # Espaço entre os divs
btn_produzir = ttk.Button(div_botao_produzir, text="Produzir", width=15, command=produzir)
btn_produzir.pack()

# Div 3: Filtro por Data (Horizontal)
div_filtro_data = ttk.Frame(frame_producao_Geral)
div_filtro_data.pack(pady=10, fill="x")

    # Div 3.1: Data Início
div_data_inicio = ttk.Frame(div_filtro_data)
div_data_inicio.pack(side=tk.LEFT, padx=10, pady=10, anchor='w') # Alinhado à esquerda
ttk.Label(div_data_inicio, text="Data Início:").pack(side=tk.LEFT)
entry_data_inicio = DateEntry(div_data_inicio, width=12, date_pattern='yyyy-mm-dd')
entry_data_inicio.pack(side=tk.LEFT)

    # Div 3.2: Data Fim
div_data_fim = ttk.Frame(div_filtro_data)
div_data_fim.pack(side=tk.LEFT, padx=10, pady=10, anchor='w') # Alinhado à esquerda
ttk.Label(div_data_fim, text="Data Fim:").pack(side=tk.LEFT)
entry_data_fim = DateEntry(div_data_fim, width=12, date_pattern='yyyy-mm-dd')
entry_data_fim.pack(side=tk.LEFT)

    # Div 3.3: Botão Filtrar
div_botao_filtrar = ttk.Frame(div_filtro_data)
div_botao_filtrar.pack(side=tk.LEFT, padx=20, pady=10) # Espaço entre os divs
btn_filtrar = ttk.Button(div_botao_filtrar, text="Filtrar", command=lambda: filtrar_producao())
btn_filtrar.pack()

    # Div 3.4: Label Resultado Filtro
div_label_filtro = ttk.Frame(div_filtro_data)
div_label_filtro.pack(side=tk.LEFT, padx=10, pady=10, anchor='w') # Alinhado à esquerda
lbl_resultado_filtro = ttk.Label(div_label_filtro, text="", font=("Arial", 12))
lbl_resultado_filtro.pack(side=tk.LEFT)

# Div 4: Contador de Produção
div_contador = ttk.Frame(frame_producao_Geral)
div_contador.pack(pady=10, fill="x")
lbl_contador = ttk.Label(div_contador, text="", font=("Arial", 14))
lbl_contador.pack(side=tk.LEFT)

# Div 5: Tabela de Produção
div_tabela_producao = ttk.Frame(frame_producao_Geral)
div_tabela_producao.pack(pady=10, fill="both", expand=True)
ttk.Label(div_tabela_producao, text="Tabela com as máquinas registradas", font=("Arial", 12)).pack()
tree_producao = ttk.Treeview(div_tabela_producao, columns=("NS", "Data", "Hora"), show="headings")
tree_producao.heading("NS", text="NS")
tree_producao.heading("Data", text="Data")
tree_producao.heading("Hora", text="Hora")
tree_producao.pack(fill=tk.BOTH, expand=True)

# Div 6: Imagem e Labels de Estatísticas (Horizontal)
div_imagem_estatisticas = ttk.Frame(frame_producao_Geral)
div_imagem_estatisticas.pack(pady=10, fill="x")

    # Div 6.1: Imagem (Esquerda)
div_imagem = ttk.Frame(div_imagem_estatisticas)
div_imagem.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.X, expand=True)
lbl_imagem_maquina = ttk.Label(div_imagem, text="[Imagem da Máquina]", relief="solid", width=20)
lbl_imagem_maquina.pack(pady=5, fill=tk.X, expand=True)

    # Div 6.2: Labels de Produzidas e Média de Tempo (Direita, Vertical)
div_estatisticas_labels = ttk.Frame(div_imagem_estatisticas)
div_estatisticas_labels.pack(side=tk.LEFT, padx=20, pady=10)

        # Div 6.2.1: Label Produzidas
div_label_produzidas = ttk.Frame(div_estatisticas_labels)
div_label_produzidas.pack(pady=5)
ttk.Label(div_label_produzidas, text="PRODUZIDAS", font=("Arial", 12)).pack()
lbl_produzidas = ttk.Label(div_label_produzidas, text=f"{len(df_producao)}", font=("Arial", 30))  # Replace with actual count
lbl_produzidas.pack()

        # Div 6.2.2: Label Média de Tempo
div_label_media_tempo = ttk.Frame(div_estatisticas_labels)
div_label_media_tempo.pack(pady=5)
ttk.Label(div_label_media_tempo, text="Média de tempo por máquina:", font=("Arial", 12)).pack()
lbl_media_tempo = ttk.Label(div_label_media_tempo, text="0 min", font=("Arial", 16))
lbl_media_tempo.pack()


# Frame Manutenção (Aba Manutenção) - Layout similar, com "divs" para organização
frame_manutencao_Geral = ttk.Frame(frame_manutencao)
frame_manutencao_Geral.pack(expand=True, fill="both", padx=10, pady=10)

# Div Manutenção 1: Entrada NS Manutenção e Botões
div_manutencao_entrada = ttk.Frame(frame_manutencao_Geral)
div_manutencao_entrada.pack(pady=10, fill="x")

    # Div Manutenção 1.1: Entrada NS Manutenção
div_manutencao_entry_ns = ttk.Frame(div_manutencao_entrada)
div_manutencao_entry_ns.pack(side=tk.LEFT, padx=10, pady=10, anchor='w')
entry_ns_manutencao = ttk.Entry(div_manutencao_entry_ns, width=20)
entry_ns_manutencao.pack()

    # Div Manutenção 1.2: Checkbox Produção
div_manutencao_checkbox = ttk.Frame(div_manutencao_entrada)
div_manutencao_checkbox.pack(side=tk.LEFT, padx=10, pady=10)
var_producao = tk.BooleanVar()
chk_producao = ttk.Checkbutton(div_manutencao_checkbox, text="Produção", variable=var_producao)
chk_producao.pack()

    # Div Manutenção 1.3: Botão Registrar Manutenção
div_manutencao_botao_registrar = ttk.Frame(div_manutencao_entrada)
div_manutencao_botao_registrar.pack(side=tk.LEFT, padx=20, pady=10)
btn_manutencao = ttk.Button(div_manutencao_botao_registrar, text="Registrar Manutenção", command=registrar_manutencao)
btn_manutencao.pack()

# Div Manutenção 2: Tabela de Manutenção
div_manutencao_tabela = ttk.Frame(frame_manutencao_Geral)
div_manutencao_tabela.pack(pady=10, fill="both", expand=True)
tree_manutencao = ttk.Treeview(div_manutencao_tabela, columns=("NS", "Status", "Data", "Hora"), show="headings")
tree_manutencao.heading("NS", text="NS")
tree_manutencao.heading("Status", text="Status")
tree_manutencao.heading("Data", text="Data")
tree_manutencao.heading("Hora", text="Hora")
tree_manutencao.pack(fill=tk.BOTH, expand=True)

# Div Manutenção 3: Botão Liberar Manutenção
div_manutencao_liberar = ttk.Frame(frame_manutencao_Geral)
div_manutencao_liberar.pack(pady=10, fill="x")
btn_liberar = ttk.Button(div_manutencao_liberar, text="Liberar", command=liberar_manutencao)
btn_liberar.pack()


atualizar_tabelas()
calcular_media_tempo()

root.mainloop()