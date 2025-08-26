import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
from fpdf import FPDF
import glob
from PIL import Image, ImageTk, ImageSequence
import sys

def resource_path(relative_path):
    """Funciona em .exe e no modo normal"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Configurações iniciais
df = pd.DataFrame(columns=["Data", "Fornecedor", "Descrição", "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"])
cores_barras = ['#4CAF50', '#FF9800', '#2196F3', '#9C27B0', '#F44336', '#00BCD4']

def salvar_df():
    df.to_csv("orcamentos.csv", index=False)

# Funções

def adicionar_item():
    try:
        fornecedor = entry_fornecedor.get()
        descricao = entry_descricao.get()
        preco_unitario = float(entry_preco.get())
        quantidade = int(entry_quantidade.get())
        ipi = float(entry_ipi.get())
        desconto = float(entry_desconto.get())
        data = datetime.now().strftime("%Y-%m-%d")

        preco_com_ipi = preco_unitario * (1 + ipi / 100)
        preco_com_desconto = preco_com_ipi * (1 - desconto / 100)
        total = preco_com_desconto * quantidade

        novo_dado = {
            "Data": data,
            "Fornecedor": fornecedor,
            "Descrição": descricao,
            "Preço Unitário": preco_unitario,
            "Quantidade": quantidade,
            "IPI": ipi,
            "Desconto": desconto,
            "Total Final": total
        }

        global df
        df = pd.concat([df, pd.DataFrame([novo_dado])], ignore_index=True)
        salvar_df()
        atualizar_tabela()

        for entry in [entry_fornecedor, entry_descricao, entry_preco, entry_quantidade, entry_ipi, entry_desconto]:
            entry.delete(0, tk.END)

    except ValueError:
        messagebox.showerror("Erro", "Insira valores válidos")

def remover_selecionado():
    item_selecionado = tabela.selection()
    if not item_selecionado:
        messagebox.showinfo("Info", "Selecione um item para remover.")
        return
    valores = tabela.item(item_selecionado[0])['values']
    fornecedor = valores[0]
    descricao = valores[1]
    resposta = messagebox.askyesno(
        "Confirmação",
        f"Deseja mesmo excluir\nFornecedor: {fornecedor}\nDescrição: {descricao}?"
    )
    if resposta:
        global df
        idx = df[(df['Fornecedor'] == fornecedor) & (df['Descrição'] == descricao)].index
        if not idx.empty:
            df = df.drop(idx[0])
            salvar_df()
            atualizar_tabela()


def atualizar_tabela(dataframe=None):
    for row in tabela.get_children():
        tabela.delete(row)
    dados = dataframe if dataframe is not None else df
    for index, row in dados.iterrows():
        preco_unitario = row["Preço Unitário"] if pd.notna(row["Preço Unitário"]) else 0
        quantidade = row["Quantidade"] if pd.notna(row["Quantidade"]) else 0
        ipi = row["IPI"] if pd.notna(row["IPI"]) else 0
        desconto = row["Desconto"] if pd.notna(row["Desconto"]) else 0
        total_final = row["Total Final"] if pd.notna(row["Total Final"]) else 0
        tabela.insert("", "end", values=[
            row["Fornecedor"], row["Descrição"], f"R$ {preco_unitario:.2f}", quantidade,
            f"{ipi:.1f}%", f"{desconto:.1f}%", f"R$ {total_final:.2f}"
        ])

def validar_numero(P):
    # Permite apenas números positivos e ponto ou vírgula para decimais
    if P == "":
        return True
    try:
        # Troca vírgula por ponto para aceitar ambos
        float(P.replace(",", "."))
        return float(P.replace(",", ".")) >= 0
    except ValueError:
        return False

def gerar_grafico():
    if df.empty:
        messagebox.showinfo("Info", "Nenhum dado para exibir no gráfico")
        return
    agrupado = df.groupby("Fornecedor")["Total Final"].sum().sort_values()
    plt.figure(figsize=(10, 5))
    barras = plt.bar(agrupado.index, agrupado.values, color=cores_barras[:len(agrupado)])
    plt.title("Comparação de Orçamentos por Fornecedor")
    plt.ylabel("Valor Total (R$)")
    plt.xticks(rotation=45)
    for barra in barras:
        yval = barra.get_height()
        plt.text(barra.get_x() + barra.get_width()/2.0, yval, f"R$ {yval:.2f}", ha='center', va='bottom')
    plt.tight_layout()
    plt.show()

def gerar_pdf():
    dados = tabela.get_children()
    if not dados:
        messagebox.showinfo("Info", "Nenhum dado disponível para exportar")
        return

    # Caminho da área de trabalho do usuário
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_pdf = os.path.join(desktop, "Orçamentos_NM_Napoleão")
    if not os.path.exists(pasta_pdf):
        os.makedirs(pasta_pdf)

    # Sugere nome padrão com data/hora
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    nome_padrao = f"Orçamento_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.pdf"
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Salvar PDF como",
        initialdir=pasta_pdf,
        initialfile=nome_padrao
    )
    if not arquivo:
        return

    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, txt=f"Relatório de Orçamentos {data_hoje}", ln=True, align='C')
    pdf.ln(10)

    # Cabeçalho
    col_widths = [40, 80, 30, 20, 20, 20, 30]
    cabecalho = ["Fornecedor", "Descrição", "Preço", "Qtd", "IPI", "Desc", "Total"]
    for i, col in enumerate(cabecalho):
        pdf.cell(col_widths[i], 10, col, border=1, align='C')
    pdf.ln()

    # Linhas
    for item in dados:
        valores = tabela.item(item)['values']
        for i, valor in enumerate(valores):
            pdf.cell(col_widths[i], 10, str(valor), border=1, align='C')
        pdf.ln()

    pdf.output(arquivo)
    messagebox.showinfo("Sucesso", f"PDF gerado com sucesso como '{os.path.basename(arquivo)}'")

def aplicar_filtros():
    filtrado = df.copy()
    fornecedor = filtro_fornecedor.get()
    descricao = filtro_descricao.get()

    if fornecedor:
        filtrado = filtrado[filtrado['Fornecedor'].str.contains(fornecedor, case=False, na=False)]
    if descricao:
        filtrado = filtrado[filtrado['Descrição'].str.contains(descricao, case=False, na=False)]

    if var_ordem.get() == 'A-Z':
        filtrado = filtrado.sort_values(by='Fornecedor')
    elif var_ordem.get() == 'Maior Desconto':
        filtrado = filtrado.sort_values(by='Desconto', ascending=False)
    elif var_ordem.get() == 'Menor IPI':
        filtrado = filtrado.sort_values(by='IPI')
    elif var_ordem.get() == 'Maior IPI':
        filtrado = filtrado.sort_values(by='IPI', ascending=False)
    elif var_ordem.get() == 'Menor Preço':
        filtrado = filtrado.sort_values(by='Total Final')
    elif var_ordem.get() == 'Maior Valor':
        filtrado = filtrado.sort_values(by='Total Final', ascending=False)

    atualizar_tabela(filtrado)

def limpar_filtros():
    filtro_fornecedor.delete(0, tk.END)
    filtro_descricao.delete(0, tk.END)
    var_ordem.set("Ordenar por...")
    atualizar_tabela()

def confirmar_saida():
    if messagebox.askyesno("Fechar", "Deseja mesmo fechar o aplicativo?"):
        salvar_df()
        tela.destroy()


def mostrar_orcamentos():
    # Get correct desktop path and folder
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_pdf = os.path.join(desktop, "Orçamentos_NM_Napoleão")
    
    # Create folder if it doesn't exist
    if not os.path.exists(pasta_pdf):
        os.makedirs(pasta_pdf)
        messagebox.showinfo("Orçamentos Salvos", "Nenhum orçamento PDF encontrado na pasta.")
        return

    # Search for PDFs in the correct folder
    arquivos_pdf = glob.glob(os.path.join(pasta_pdf, "*.pdf"))
    if not arquivos_pdf:
        messagebox.showinfo("Orçamentos Salvos", "Nenhum orçamento PDF encontrado na pasta.")
        return

    # Cria uma nova janela para mostrar a lista
    janela = tk.Toplevel(tela)
    janela.title("Orçamentos Salvos")
    janela.geometry("550x350")

    tk.Label(janela, text="Orçamentos PDF Salvos", font=("Arial", 12, "bold")).pack(pady=10)

    frame_lista = tk.Frame(janela)
    frame_lista.pack(padx=10, pady=10, fill="both", expand=True)

    lista = tk.Listbox(frame_lista, font=("Arial", 10), width=60)
    lista.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_lista, orient="vertical", command=lista.yview)
    scrollbar.pack(side="left", fill="y")
    lista.config(yscrollcommand=scrollbar.set)

    for arquivo in arquivos_pdf:
        data_criacao = datetime.fromtimestamp(os.path.getctime(arquivo)).strftime("%d/%m/%Y %H:%M")
        lista.insert(tk.END, f"{os.path.basename(arquivo)}  |  Criado em: {data_criacao}")

    def abrir_pdf():
        selecionado = lista.curselection()
        if selecionado:
            nome_arquivo = arquivos_pdf[selecionado[0]]
            os.startfile(nome_arquivo)  # Abre o PDF com o visualizador padrão do Windows

    def apagar_pdf():
        selecionado = lista.curselection()
        if not selecionado:
            messagebox.showinfo("Info", "Selecione um orçamento para apagar.")
            return
        nome_arquivo = arquivos_pdf[selecionado[0]]
        nome_exibicao = os.path.basename(nome_arquivo)
        confirmar = messagebox.askyesno("Confirmação", f"Deseja realmente apagar o orçamento '{nome_exibicao}'?")
        if confirmar:
            try:
                os.remove(nome_arquivo)
                lista.delete(selecionado[0])
                del arquivos_pdf[selecionado[0]]
                messagebox.showinfo("Sucesso", f"Orçamento '{nome_exibicao}' apagado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível apagar o arquivo.\n{e}")

    btns_frame = tk.Frame(janela)
    btns_frame.pack(pady=5)

    btn_abrir = tk.Button(btns_frame, text="Abrir PDF Selecionado", command=abrir_pdf, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_abrir.pack(side="left", padx=5)

    btn_apagar = tk.Button(btns_frame, text="Apagar Orçamento", command=apagar_pdf, bg="#F44336", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_apagar.pack(side="left", padx=5)

def novo_orcamento():
    global df
    # Verifica se há dados para salvar
    if not df.empty:
        if messagebox.askyesno("Novo Orçamento", 
            "Deseja salvar o orçamento atual antes de começar um novo?"):
            
            # Pasta principal
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            pasta_pdf = os.path.join(desktop, "Orçamentos_NM_Napoleão")
            
            # Cria pasta de backups dos CSVs
            pasta_backup = os.path.join(pasta_pdf, "Orçamentos_Antigos")
            if not os.path.exists(pasta_backup):
                os.makedirs(pasta_backup)
            
            # Salva backup do CSV atual
            timestamp = datetime.now().strftime('%d-%m-%Y_%H-%M-%S')
            csv_backup = os.path.join(pasta_backup, f"orcamento_{timestamp}.csv")
            df.to_csv(csv_backup, index=False)
            
            # Gera nome automático para o PDF
            nome_pdf = os.path.join(pasta_pdf, f"Orçamento_{timestamp}.pdf")
            
            # Gera o PDF automaticamente
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 10, txt=f"Relatório de Orçamentos {datetime.now().strftime('%d/%m/%Y')}", 
                    ln=True, align='C')
            pdf.ln(10)

            # Cabeçalho
            col_widths = [40, 80, 30, 20, 20, 20, 30]
            cabecalho = ["Fornecedor", "Descrição", "Preço", "Qtd", "IPI", "Desc", "Total"]
            for i, col in enumerate(cabecalho):
                pdf.cell(col_widths[i], 10, col, border=1, align='C')
            pdf.ln()

            # Dados
            for _, row in df.iterrows():
                valores = [
                    row["Fornecedor"],
                    row["Descrição"],
                    f"R$ {row['Preço Unitário']:.2f}",
                    str(row["Quantidade"]),
                    f"{row['IPI']:.1f}%",
                    f"{row['Desconto']:.1f}%",
                    f"R$ {row['Total Final']:.2f}"
                ]
                for i, valor in enumerate(valores):
                    pdf.cell(col_widths[i], 10, str(valor), border=1, align='C')
                pdf.ln()

            pdf.output(nome_pdf)
            messagebox.showinfo("Sucesso", 
                f"Orçamento anterior salvo como:\nPDF: {os.path.basename(nome_pdf)}\nCSV: {os.path.basename(csv_backup)}")

    # Limpa o DataFrame

    df = pd.DataFrame(columns=["Data", "Fornecedor", "Descrição", "Preço Unitário", 
                              "Quantidade", "IPI", "Desconto", "Total Final"])
    
    # Salva CSV vazio
    salvar_df()
    
    # Limpa a tabela
    atualizar_tabela()
    
    # Limpa os campos de entrada
    for entry in [entry_fornecedor, entry_descricao, entry_preco, 
                 entry_quantidade, entry_ipi, entry_desconto]:
        entry.delete(0, tk.END)
    
    messagebox.showinfo("Novo Orçamento", "Pronto para um novo orçamento!")

def mostrar_csvs_antigos():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_backup = os.path.join(desktop, "Orçamentos_NM_Napoleão", "Orçamentos_Antigos")
    
    if not os.path.exists(pasta_backup):
        messagebox.showinfo("Info", "Nenhum orçamento antigo encontrado.")
        return

    arquivos_csv = glob.glob(os.path.join(pasta_backup, "*.csv"))
    if not arquivos_csv:
        messagebox.showinfo("Info", "Nenhum orçamento antigo encontrado.")
        return

    janela = tk.Toplevel(tela)
    janela.title("Orçamentos Antigos (CSV)")
    janela.geometry("550x350")

    tk.Label(janela, text="Orçamentos CSV Salvos", font=("Arial", 12, "bold")).pack(pady=10)

    frame_lista = tk.Frame(janela)
    frame_lista.pack(padx=10, pady=10, fill="both", expand=True)

    lista = tk.Listbox(frame_lista, font=("Arial", 10), width=60)
    lista.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_lista, orient="vertical", command=lista.yview)
    scrollbar.pack(side="left", fill="y")
    lista.config(yscrollcommand=scrollbar.set)

    for arquivo in arquivos_csv:
        data_criacao = datetime.fromtimestamp(os.path.getctime(arquivo)).strftime("%d/%m/%Y %H:%M")
        lista.insert(tk.END, f"{os.path.basename(arquivo)}  |  Criado em: {data_criacao}")

    def carregar_csv():
        selecionado = lista.curselection()
        if selecionado:
            global df
            nome_arquivo = arquivos_csv[selecionado[0]]
            if messagebox.askyesno("Carregar", "Deseja carregar este orçamento? (O atual será substituído)"):
                df = pd.read_csv(nome_arquivo)
                atualizar_tabela()
                janela.destroy()
                messagebox.showinfo("Sucesso", "Orçamento carregado com sucesso!")

    btns_frame = tk.Frame(janela)
    btns_frame.pack(pady=5)

    btn_carregar = tk.Button(btns_frame, text="Carregar Orçamento", command=carregar_csv, 
                           bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_carregar.pack(side="left", padx=5)

# Interface
tela = tk.Tk()
tela.title("App de Orçamento com Fornecedores")
tela.geometry("1300x800")

# Adiciona frame para logo
frame_logo = tk.Frame(tela)
frame_logo.pack(pady=10)

# Carrega e redimensiona a logo
try:
    # Usa resource_path para localizar a imagem
    logo_path = resource_path("n_multifibra_logo-removebg-preview.png")
    imagem = Image.open(logo_path)
    basewidth = 180
    wpercent = (basewidth/float(imagem.size[0]))
    hsize = int((float(imagem.size[1])*float(wpercent)))
    imagem = imagem.resize((basewidth, hsize), Image.Resampling.LANCZOS)
    logo = ImageTk.PhotoImage(imagem)
    label_logo = tk.Label(frame_logo, image=logo)
    label_logo.image = logo
    label_logo.pack()
except Exception as e:
    print(f"Erro ao carregar logo: {e}")

# Create menu first
menu_bar = tk.Menu(tela)
menu_opcoes = tk.Menu(menu_bar, tearoff=0)
tela.config(menu=menu_bar)

# commands to menu
menu_opcoes.add_command(label="Mostrar Orçamentos", command=mostrar_orcamentos)
menu_opcoes.add_separator()
menu_opcoes.add_command(label="Novo Orçamento", command=novo_orcamento)
menu_opcoes.add_command(label="Orçamentos Antigos", command=mostrar_csvs_antigos)
menu_bar.add_cascade(label="Opções", menu=menu_opcoes)

fonte_label = ("Arial", 11, "bold")
padrao_entry = {"font": ("Arial", 10), "width": 30}

frame_entrada = tk.Frame(tela)
frame_entrada.pack(pady=10)

# Labels e Entradas (padronizadas)
labels_entrada = ["Fornecedor", "Descrição", "Preço Unitário (R$)", "Quantidade", "IPI (%)", "Desconto (%)"]
entries_entrada = []

vcmd = (tela.register(validar_numero), "%P")

for i, label in enumerate(labels_entrada):
    tk.Label(frame_entrada, text=label, font=fonte_label).grid(row=i, column=0, sticky="e", padx=5, pady=2)
    if label in ["Preço Unitário (R$)", "Quantidade", "IPI (%)", "Desconto (%)"]:
        entry = tk.Entry(frame_entrada, **padrao_entry, validate="key", validatecommand=vcmd)
    else:
        entry = tk.Entry(frame_entrada, **padrao_entry)
    entry.grid(row=i, column=1, padx=5, pady=2)
    entries_entrada.append(entry)

entry_fornecedor, entry_descricao, entry_preco, entry_quantidade, entry_ipi, entry_desconto = entries_entrada

# Botões principais
frame_botoes = tk.Frame(tela)
frame_botoes.pack(pady=10)

btn_add = tk.Button(frame_botoes, text="Adicionar Item", command=adicionar_item, bg="#4CAF50", fg="white", font=fonte_label, width=15)
btn_add.grid(row=0, column=0, padx=5)

btn_grafico = tk.Button(frame_botoes, text="Gerar Gráfico", command=gerar_grafico, bg="#9C27B0", fg="white", font=fonte_label, width=15)
btn_grafico.grid(row=0, column=1, padx=5)

btn_pdf = tk.Button(frame_botoes, text="Gerar PDF", command=gerar_pdf, bg="#607D8B", fg="white", font=fonte_label, width=15)
btn_pdf.grid(row=0, column=2, padx=5)

btn_remover = tk.Button(frame_botoes, text="Remover Orçamento", command=remover_selecionado, bg="#F00707", fg="white", font=fonte_label, width=20)
btn_remover.grid(row=0, column=3, padx=5)


# Filtros
filtros_frame = tk.LabelFrame(tela, text="🔍 Filtros Avançados", padx=10, pady=10, font=("Arial", 10, "bold"))
filtros_frame.pack(pady=10, fill="x")

tk.Label(filtros_frame, text="Fornecedor:").grid(row=0, column=0)
filtro_fornecedor = tk.Entry(filtros_frame, font=("Arial", 10), width=25)
filtro_fornecedor.grid(row=0, column=1, padx=5)

tk.Label(filtros_frame, text="Descrição:").grid(row=0, column=2)
filtro_descricao = tk.Entry(filtros_frame, font=("Arial", 10), width=25)
filtro_descricao.grid(row=0, column=3, padx=5)

var_ordem = tk.StringVar()
ordem_menu = ttk.Combobox(
    filtros_frame,
    textvariable=var_ordem,
    values=["", "A-Z", "Maior Desconto", "Menor IPI", "Maior IPI", "Menor Preço", "Maior Valor"],
    state="readonly",
    width=18
)
ordem_menu.grid(row=0, column=4, padx=5)
ordem_menu.set("Ordenar por...")

btn_filtrar = tk.Button(filtros_frame, text="Aplicar Filtros", command=aplicar_filtros, bg="#2196F3", fg="white", font=fonte_label, width=15)
btn_filtrar.grid(row=0, column=5, padx=5)

btn_limpar = tk.Button(filtros_frame, text="Limpar Filtros", command=limpar_filtros, bg="#F44336", fg="white", font=fonte_label, width=15)
btn_limpar.grid(row=0, column=6, padx=5)

# Tabela
frame_tabela = tk.Frame(tela)
frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

colunas = ["Fornecedor", "Descrição", "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"]
tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")
tabela.pack(fill="both", expand=True)

for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, width=150, anchor="center")

# Inicializa tabela com dados
try:
    df = pd.read_csv("orcamentos.csv")
    df["Preço Unitário"] = pd.to_numeric(df["Preço Unitário"], errors="coerce")
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce")
    df["IPI"] = pd.to_numeric(df["IPI"], errors="coerce")
    df["Desconto"] = pd.to_numeric(df["Desconto"], errors="coerce")
    df["Total Final"] = pd.to_numeric(df["Total Final"], errors="coerce")
except FileNotFoundError:
    df = pd.DataFrame(columns=colunas)

atualizar_tabela()

tela.protocol("WM_DELETE_WINDOW", confirmar_saida)
tela.mainloop()