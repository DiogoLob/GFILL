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
except:
    df_producao = pd.DataFrame(columns=["NS", "Data", "Hora"])
    df_manutencao = pd.DataFrame(columns=["NS", "Status", "Data", "Hora"])

def salvar_dados():
    with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
        df_producao.to_excel(writer, sheet_name="Producao", index=False)
        df_manutencao.to_excel(writer, sheet_name="Manutencao", index=False)

def atualizar_contador():
    lbl_contador.config(text=f"Máquinas Produzidas: {len(df_producao)}")

def produzir():
    ns = entry_ns_producao.get()
    if ns:
        data = datetime.now().strftime("%Y-%m-%d")
        hora = datetime.now().strftime("%H:%M:%S")
        global df_producao
        df_producao = pd.concat([df_producao, pd.DataFrame([{"NS": ns, "Data": data, "Hora": hora}])], ignore_index=True)
        salvar_dados()
        atualizar_tabelas()
        atualizar_contador()
        messagebox.showinfo("Sucesso", "Máquina produzida com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Digite o número de NS!")

def registrar_manutencao():
    ns = entry_ns_manutencao.get()
    status = "Produção" if var_producao.get() else "Estoque"
    if ns:
        data = datetime.now().strftime("%Y-%m-%d")
        hora = datetime.now().strftime("%H:%M:%S")
        global df_manutencao
        df_manutencao = pd.concat([df_manutencao, pd.DataFrame([{"NS": ns, "Status": status, "Data": data, "Hora": hora}])], ignore_index=True)
        if var_producao.get():
            global df_producao
            df_producao = df_producao[df_producao.NS != ns]
        salvar_dados()
        atualizar_tabelas()
        messagebox.showinfo("Sucesso", "Máquina registrada na manutenção!")
    else:
        messagebox.showwarning("Aviso", "Digite o número de NS!")

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

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao filtrar: {e}")


root = tk.Tk()
root.title("Controle de Produção")
root.geometry("900x600")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

frame_producao = ttk.Frame(notebook)
frame_manutencao = ttk.Frame(notebook)
notebook.add(frame_producao, text="Produção")
notebook.add(frame_manutencao, text="Manutenção")

frame_top = ttk.Frame(frame_producao)
frame_top.pack(pady=10)
entry_ns_producao = ttk.Entry(frame_top, width=20)
entry_ns_producao.pack(side=tk.LEFT, padx=5)
btn_produzir = ttk.Button(frame_top, text="Produzir", command=produzir)
btn_produzir.pack(side=tk.LEFT, padx=5)

# Campos para selecionar período de filtragem
frame_filtro = ttk.Frame(frame_producao)
frame_filtro.pack(pady=10)

ttk.Label(frame_filtro, text="Data Início:").pack(side=tk.LEFT, padx=5)
entry_data_inicio = DateEntry(frame_filtro, width=12, date_pattern='yyyy-mm-dd')
entry_data_inicio.pack(side=tk.LEFT, padx=5)

ttk.Label(frame_filtro, text="Data Fim:").pack(side=tk.LEFT, padx=5)
entry_data_fim = DateEntry(frame_filtro, width=12, date_pattern='yyyy-mm-dd')
entry_data_fim.pack(side=tk.LEFT, padx=5)

btn_filtrar = ttk.Button(frame_filtro, text="Filtrar", command=lambda: filtrar_producao())
btn_filtrar.pack(side=tk.LEFT, padx=10)

lbl_resultado_filtro = ttk.Label(frame_filtro, text="", font=("Arial", 12))
lbl_resultado_filtro.pack(side=tk.LEFT, padx=10)

frame_contador = ttk.Frame(frame_producao)
frame_contador.pack(pady=10)
lbl_contador = ttk.Label(frame_contador, text=f"Máquinas Produzidas: {len(df_producao)}", font=("Arial", 14))
lbl_contador.pack(side=tk.LEFT)

frame_tree_producao = ttk.Frame(frame_producao)
frame_tree_producao.pack(fill=tk.BOTH, expand=True, pady=10)
tree_producao = ttk.Treeview(frame_tree_producao, columns=("NS", "Data", "Hora"), show="headings")
tree_producao.heading("NS", text="NS")
tree_producao.heading("Data", text="Data")
tree_producao.heading("Hora", text="Hora")
tree_producao.pack(fill=tk.BOTH, expand=True)

frame_manutencao_top = ttk.Frame(frame_manutencao)
frame_manutencao_top.pack(pady=10)
entry_ns_manutencao = ttk.Entry(frame_manutencao_top, width=20)
entry_ns_manutencao.pack(side=tk.LEFT, padx=5)
var_producao = tk.BooleanVar()
chk_producao = ttk.Checkbutton(frame_manutencao_top, text="Produção", variable=var_producao)
chk_producao.pack(side=tk.LEFT, padx=5)
btn_manutencao = ttk.Button(frame_manutencao_top, text="Registrar Manutenção", command=registrar_manutencao)
btn_manutencao.pack(side=tk.LEFT, padx=5)

frame_tree_manutencao = ttk.Frame(frame_manutencao)
frame_tree_manutencao.pack(fill=tk.BOTH, expand=True, pady=10)
tree_manutencao = ttk.Treeview(frame_tree_manutencao, columns=("NS", "Status", "Data", "Hora"), show="headings")
tree_manutencao.heading("NS", text="NS")
tree_manutencao.heading("Status", text="Status")
tree_manutencao.heading("Data", text="Data")
tree_manutencao.heading("Hora", text="Hora")
tree_manutencao.pack(fill=tk.BOTH, expand=True)

btn_liberar = ttk.Button(frame_manutencao, text="Liberar", command=liberar_manutencao)
btn_liberar.pack(pady=10)

atualizar_tabelas()
root.mainloop()
