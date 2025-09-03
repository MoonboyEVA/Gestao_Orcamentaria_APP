import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, EXTENDED, simpledialog
import matplotlib.pyplot as plt
from datetime import datetime
import os
from fpdf import FPDF
import glob
from PIL import Image, ImageTk, ImageSequence
import sys
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import shutil  # Importa o módulo shutil

# Função para obter o caminho correto dos recursos
def resource_path(relative_path):
    """Funciona em .exe e no modo normal"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)



# Configurações iniciais
df = pd.DataFrame(columns=["Data", "Fornecedor", "Produto", "Descrição", "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"])
cores_barras = ['#4CAF50', '#FF9800', '#2196F3', '#9C27B0', '#F44336', '#00BCD4']

def salvar_df():
    try:
        # Salva o arquivo principal
        df.to_csv("orcamentos.csv", index=False)
        
        # Se existe um arquivo de backup aberto, salva nele também
        if hasattr(tela, 'arquivo_backup_atual'):
            df.to_csv(tela.arquivo_backup_atual, index=False)
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar dados:\n{e}")

# Funções

def adicionar_item():
    try:
        # Validação de campos vazios
        if not all([entry_fornecedor.get(), entry_produto.get()]):
            messagebox.showerror("Erro", "Fornecedor e Produto são campos obrigatórios!")
            return
            
        # Validação de números negativos
        if float(entry_preco.get()) < 0 or int(entry_quantidade.get()) < 0:
            messagebox.showerror("Erro", "Preço e Quantidade não podem ser negativos!")
            return

        fornecedor = entry_fornecedor.get()
        produto = entry_produto.get()
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
            "Produto": produto,
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

        for entry in [entry_fornecedor, entry_produto, entry_descricao, entry_preco, entry_quantidade, entry_ipi, entry_desconto]:
            entry.delete(0, tk.END)

    except ValueError:
        messagebox.showerror("Erro", "Insira valores válidos")

def remover_selecionado():
    itens_selecionados = tabela.selection()
    if not itens_selecionados:
        messagebox.showinfo("Info", "Selecione um ou mais itens para remover.")
        return

    qtd_selecionados = len(itens_selecionados)
    mensagem = f"Você selecionou {qtd_selecionados} item(ns) para remover.\n\nPrimeiros itens selecionados:"
    
    # Mostra até 3 itens como exemplo
    for i, item in enumerate(itens_selecionados[:3]):
        valores = tabela.item(item)['values']
        mensagem += f"\n- {valores[0]} | {valores[1]} | R$ {valores[7]}"
    
    if qtd_selecionados > 3:
        mensagem += f"\n\nE mais {qtd_selecionados - 3} outro(s) item(ns)..."
    
    mensagem += "\n\nDeseja realmente remover estes itens?"

    if messagebox.askyesno("Confirmação de Remoção", mensagem):
        global df
        for item in itens_selecionados:
            valores = tabela.item(item)['values']
            idx = df[(df['Fornecedor'] == valores[0]) & 
                    (df['Produto'] == valores[1]) & 
                    (df['Descrição'] == valores[2])].index
            if not idx.empty:
                df = df.drop(idx[0])
        
        salvar_df()
        atualizar_tabela()
        messagebox.showinfo("Sucesso", f"{qtd_selecionados} item(ns) removido(s) com sucesso!")




def atualizar_tabela(dataframe=None):
    tabela.delete(*tabela.get_children())  # Forma mais eficiente
    dados = dataframe if dataframe is not None else df
    for _, row in dados.iterrows():
        tabela.insert("", "end", values=[
            row["Fornecedor"],
            row["Produto"],
            row["Descrição"],
            f"R$ {float(row['Preço Unitário']):.2f}",
            int(row["Quantidade"]),
            f"{float(row['IPI']):.1f}%",
            f"{float(row['Desconto']):.1f}%",
            f"R$ {float(row['Total Final']):.2f}"
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
    plt.close('all')  # Fecha gráficos anteriores
    if df.empty:
        messagebox.showinfo("Info", "Nenhum dado para exibir no gráfico")
        return

    # Criar janela para os gráficos
    janela_grafico = tk.Toplevel(tela)
    janela_grafico.title("Análise Gráfica dos Orçamentos")
    janela_grafico.geometry("800x600")
    janela_grafico.transient(tela)
    janela_grafico.focus_force()

    # Criar notebook (sistema de abas)
    notebook = ttk.Notebook(janela_grafico)
    notebook.pack(fill='both', expand=True, padx=10, pady=5)

    # Função auxiliar para criar gráficos
    def criar_aba_grafico(titulo):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=titulo)
        return frame

    # Aba 1: Valor Total por Fornecedor
    frame_valor = criar_aba_grafico("Valor Total")
    fig_valor = plt.Figure(figsize=(8, 5))
    ax_valor = fig_valor.add_subplot(111)
    
    valor_total = df.groupby("Fornecedor")["Total Final"].sum().sort_values(ascending=True)
    bars_valor = ax_valor.barh(valor_total.index, valor_total.values, color='#2196F3')
    
    ax_valor.set_title("Valor Total por Fornecedor")
    ax_valor.set_xlabel("Valor Total (R$)")
    
    # Adicionar valores nas barras
    for bar in bars_valor:
        width = bar.get_width()
        ax_valor.text(width, bar.get_y() + bar.get_height()/2, 
                     f'R$ {width:,.2f}', ha='left', va='center', fontsize=8)
    
    canvas_valor = FigureCanvasTkAgg(fig_valor, frame_valor)
    canvas_valor.draw()
    canvas_valor.get_tk_widget().pack(fill='both', expand=True)

    # Aba 2: IPI Médio por Fornecedor
    frame_ipi = criar_aba_grafico("IPI Médio")
    fig_ipi = plt.Figure(figsize=(8, 5))
    ax_ipi = fig_ipi.add_subplot(111)
    
    ipi_medio = df.groupby("Fornecedor")["IPI"].mean().sort_values(ascending=True)
    bars_ipi = ax_ipi.barh(ipi_medio.index, ipi_medio.values, color='#FF9800')
    
    ax_ipi.set_title("IPI Médio por Fornecedor")
    ax_ipi.set_xlabel("IPI (%)")
    
    for bar in bars_ipi:
        width = bar.get_width()
        ax_ipi.text(width, bar.get_y() + bar.get_height()/2, 
                   f'{width:.1f}%', ha='left', va='center', fontsize=8)
    
    canvas_ipi = FigureCanvasTkAgg(fig_ipi, frame_ipi)
    canvas_ipi.draw()
    canvas_ipi.get_tk_widget().pack(fill='both', expand=True)

    # Aba 3: Desconto Médio por Fornecedor
    frame_desconto = criar_aba_grafico("Desconto Médio")
    fig_desconto = plt.Figure(figsize=(8, 5))
    ax_desconto = fig_desconto.add_subplot(111)
    
    desconto_medio = df.groupby("Fornecedor")["Desconto"].mean().sort_values(ascending=True)
    bars_desconto = ax_desconto.barh(desconto_medio.index, desconto_medio.values, color='#4CAF50')
    
    ax_desconto.set_title("Desconto Médio por Fornecedor")
    ax_desconto.set_xlabel("Desconto (%)")
    
    for bar in bars_desconto:
        width = bar.get_width()
        ax_desconto.text(width, bar.get_y() + bar.get_height()/2, 
                        f'{width:.1f}%', ha='left', va='center', fontsize=8)
    
    canvas_desconto = FigureCanvasTkAgg(fig_desconto, frame_desconto)
    canvas_desconto.draw()
    canvas_desconto.get_tk_widget().pack(fill='both', expand=True)

    # Aba 4: Quantidade Total por Fornecedor
    frame_qtd = criar_aba_grafico("Quantidade Total")
    fig_qtd = plt.Figure(figsize=(8, 5))
    ax_qtd = fig_qtd.add_subplot(111)
    
    qtd_total = df.groupby("Fornecedor")["Quantidade"].sum().sort_values(ascending=True)
    bars_qtd = ax_qtd.barh(qtd_total.index, qtd_total.values, color='#9C27B0')
    
    ax_qtd.set_title("Quantidade Total por Fornecedor")
    ax_qtd.set_xlabel("Quantidade")
    
    for bar in bars_qtd:
        width = bar.get_width()
        ax_qtd.text(width, bar.get_y() + bar.get_height()/2, 
                   f'{int(width)}', ha='left', va='center', fontsize=8)
    
    canvas_qtd = FigureCanvasTkAgg(fig_qtd, frame_qtd)
    canvas_qtd.draw()
    canvas_qtd.get_tk_widget().pack(fill='both', expand=True)

    # Botão para exportar todos os gráficos
    def exportar_graficos():
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            pasta_graficos = os.path.join(desktop, "Orçamentos_NM_Napoleão", "Gráficos")
            if not os.path.exists(pasta_graficos):
                os.makedirs(pasta_graficos)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Salvar cada gráfico
            fig_valor.savefig(os.path.join(pasta_graficos, f'valor_total_{timestamp}.png'))
            fig_ipi.savefig(os.path.join(pasta_graficos, f'ipi_medio_{timestamp}.png'))
            fig_desconto.savefig(os.path.join(pasta_graficos, f'desconto_medio_{timestamp}.png'))
            fig_qtd.savefig(os.path.join(pasta_graficos, f'quantidade_total_{timestamp}.png'))
            
            messagebox.showinfo("Sucesso", f"Gráficos exportados com sucesso para:\n{pasta_graficos}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar gráficos:\n{e}")

    btn_exportar = tk.Button(
        janela_grafico,
        text="Exportar Gráficos",
        command=exportar_graficos,
        bg="#2196F3",
        fg="white",
        font=("Arial", 10, "bold")
    )
    btn_exportar.pack(pady=5)

def gerar_pdf():
    dados = tabela.get_children()
    if not dados:
        messagebox.showinfo("Info", "Nenhum dado disponível para exportar")
        return

    # Organiza os dados por produto e fornecedor
    produtos_dados = {}
    fornecedores_set = set()
    
    for item in dados:
        valores = tabela.item(item)['values']
        fornecedor = valores[0]
        produto = valores[1]
        
        if produto not in produtos_dados:
            produtos_dados[produto] = {}
        
        if fornecedor not in produtos_dados[produto]:
            produtos_dados[produto][fornecedor] = []
        
        produtos_dados[produto][fornecedor].append({
            'Preço': valores[3],
            'Qtd': valores[4],
            'IPI': valores[5],
            'Desc': valores[6],
            'Total': valores[7]
        })
        
        fornecedores_set.add(fornecedor)

    # Configuração do PDF
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_pdf = os.path.join(desktop, "Orçamentos_NM_Napoleão")
    if not os.path.exists(pasta_pdf):
        os.makedirs(pasta_pdf)

    data_hoje = datetime.now().strftime('%d/%m/%Y')
    nome_padrao = f"Orçamento_Comparativo_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.pdf"
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Salvar PDF como",
        initialdir=pasta_pdf,
        initialfile=nome_padrao
    )
    if not arquivo:
        return

    # Criar PDF
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # Título
    pdf.set_font("Arial", "B", size=12)
    pdf.cell(0, 10, txt=f"Comparativo de Orçamentos - {data_hoje}", ln=True, align='C')
    pdf.ln(5)

    # Lista de fornecedores ordenada
    fornecedores = sorted(list(fornecedores_set))
    
    # Definir larguras
    margem = 10
    largura_disponivel = pdf.w - (2 * margem)
    num_fornecedores = len(fornecedores)
    largura_coluna = largura_disponivel / num_fornecedores

    # Para cada produto
    for produto in produtos_dados:
        # Adiciona nova página se necessário
        if pdf.get_y() > pdf.h - 60:
            pdf.add_page()
            pdf.set_font("Arial", "B", size=12)
            pdf.cell(0, 10, txt=f"Comparativo de Orçamentos - {data_hoje}", ln=True, align='C')
            pdf.ln(5)

        # Nome do produto
        pdf.set_font("Arial", "B", size=10)
        pdf.cell(0, 7, f"Produto: {produto}", ln=True, align='L')
        pdf.ln(2)
        
        # Cabeçalho dos fornecedores
        y_start = pdf.get_y()
        x_start = margem
        
        # Desenha células do cabeçalho
        for fornecedor in fornecedores:
            pdf.set_xy(x_start, y_start)
            pdf.cell(largura_coluna, 7, fornecedor, border=1, align='C')
            x_start += largura_coluna
        
        pdf.ln(7)  # Espaço após cabeçalho
        
        # Reset posições para conteúdo
        y_content = pdf.get_y()
        altura_maxima_secao = 0
        
        # Primeira passagem para calcular altura máxima
        for fornecedor in fornecedores:
            if fornecedor in produtos_dados[produto]:
                itens = produtos_dados[produto][fornecedor]
                altura_total = 0
                for item in itens:
                    altura_total += 20  # Altura estimada por item
                altura_maxima_secao = max(altura_maxima_secao, altura_total)
        
        # Segunda passagem para desenhar conteúdo
        x_start = margem
        for fornecedor in fornecedores:
            pdf.set_xy(x_start, y_content)
            pdf.set_font("Arial", size=8)
            
            if fornecedor in produtos_dados[produto]:
                y_item = y_content
                for item in produtos_dados[produto][fornecedor]:
                    pdf.set_xy(x_start, y_item)
                    texto = (f"Preço: {item['Preço']}\n"
                            f"Qtd: {item['Qtd']}\n"
                            f"IPI: {item['IPI']}\n"
                            f"Desc: {item['Desc']}\n"
                            f"Total: {item['Total']}")
                    pdf.multi_cell(largura_coluna, 4, texto, border=1)
                    y_item = pdf.get_y() + 2
            else:
                pdf.cell(largura_coluna, altura_maxima_secao, "N/A", border=1, align='C')
            
            x_start += largura_coluna
        
        # Move para o próximo produto
        pdf.set_y(y_content + altura_maxima_secao + 10)

    pdf.output(arquivo)
    messagebox.showinfo("Sucesso", f"PDF comparativo gerado com sucesso como '{os.path.basename(arquivo)}'")

def aplicar_filtros():
    filtrado = df.copy()
    fornecedor = filtro_fornecedor.get()
    produto = filtro_produto.get()
    descricao = filtro_descricao.get()

    if fornecedor:
        filtrado = filtrado[filtrado['Fornecedor'].str.contains(fornecedor, case=False, na=False)]
    if produto:
        filtrado = filtrado[filtrado['Produto'].str.contains(produto, case=False, na=False)]
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
    filtro_produto.delete(0, tk.END)
    filtro_descricao.delete(0, tk.END)
    var_ordem.set("Ordenar por...")
    atualizar_tabela()

def confirmar_saida():
    global df
    try:
        # Se o DataFrame estiver vazio, fecha direto
        if df.empty:
            tela.destroy()
            return
            
        # Verifica se existem alterações para salvar
        if os.path.exists("orcamentos.csv"):
            df_original = pd.read_csv("orcamentos.csv")
            
            # Se houver diferenças, salva automaticamente no arquivo atual
            if not df.equals(df_original):
                # Salva no arquivo principal
                df.to_csv("orcamentos.csv", index=False)
                
                # Se existe um arquivo de backup aberto, salva nele também
                if hasattr(tela, 'arquivo_backup_atual'):
                    df.to_csv(tela.arquivo_backup_atual, index=False)
            
            tela.destroy()
        else:
            # Se não existe arquivo original mas há dados
            if not df.empty:
                # Salva apenas no arquivo principal
                df.to_csv("orcamentos.csv", index=False)
            
            tela.destroy()
                
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar alterações:\n{e}")
        tela.destroy()

def mostrar_orcamentos():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_pdf = os.path.join(desktop, "Orçamentos_NM_Napoleão")
    
    if not os.path.exists(pasta_pdf):
        os.makedirs(pasta_pdf)
        messagebox.showinfo("Orçamentos Salvos", "Nenhum orçamento PDF encontrado na pasta.")
        return

    arquivos_pdf = glob.glob(os.path.join(pasta_pdf, "*.pdf"))
    if not arquivos_pdf:
        messagebox.showinfo("Orçamentos Salvos", "Nenhum orçamento PDF encontrado na pasta.")
        return

    # Cria uma nova janela para mostrar a lista
    janela = tk.Toplevel(tela)
    janela.title("Orçamentos Salvos")
    janela.geometry("550x350")
    janela.transient(tela)  # Faz a janela ser dependente da principal
    janela.focus_force()    # Força o foco para esta janela

    tk.Label(janela, text="Orçamentos PDF Salvos", font=("Arial", 12, "bold")).pack(pady=10)

    frame_lista = tk.Frame(janela)
    frame_lista.pack(padx=10, pady=10, fill="both", expand=True)

    lista = tk.Listbox(frame_lista, font=("Arial", 10), width=60)
    lista.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_lista, orient="vertical", command=lista.yview)
    scrollbar.pack(side="right", fill="y")
    lista.config(yscrollcommand=scrollbar.set)

    for arquivo in arquivos_pdf:
        data_criacao = datetime.fromtimestamp(os.path.getctime(arquivo)).strftime("%d/%m/%Y %H:%M")
        lista.insert(tk.END, f"{os.path.basename(arquivo)}  |  Criado em: {data_criacao}")

    def abrir_pdf():
        selecionado = lista.curselection()
        if selecionado:
            nome_arquivo = arquivos_pdf[selecionado[0]]
            os.startfile(nome_arquivo)

    def apagar_pdf():
        selecionado = lista.curselection()
        if not selecionado:
            messagebox.showinfo("Info", "Selecione um orçamento para apagar.")
            return
        nome_arquivo = arquivos_pdf[selecionado[0]]
        nome_exibicao = os.path.basename(nome_arquivo)
        if messagebox.askyesno("Confirmação", f"Deseja realmente apagar o orçamento '{nome_exibicao}'?"):
            try:
                os.remove(nome_arquivo)
                lista.delete(selecionado[0])
                arquivos_pdf.pop(selecionado[0])
                messagebox.showinfo("Sucesso", f"Orçamento '{nome_exibicao}' apagado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível apagar o arquivo.\n{e}")

    btns_frame = tk.Frame(janela)
    btns_frame.pack(pady=5)

    btn_abrir = tk.Button(btns_frame, text="Abrir PDF Selecionado", command=abrir_pdf,
                         bg="#2196F3", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_abrir.pack(side="left", padx=5)

    btn_apagar = tk.Button(btns_frame, text="Apagar Orçamento", command=apagar_pdf,
                          bg="#F44336", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_apagar.pack(side="left", padx=5)

def mostrar_csvs_antigos():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pasta_backup = os.path.join(desktop, "Orçamentos_NM_Napoleão", "Orçamentos_Antigos")
    
    if not os.path.exists(pasta_backup):
        os.makedirs(pasta_backup)
        messagebox.showinfo("Info", "Nenhum orçamento antigo encontrado.")
        return

    arquivos_csv = glob.glob(os.path.join(pasta_backup, "*.csv"))
    if not arquivos_csv:
        messagebox.showinfo("Info", "Nenhum orçamento antigo encontrado.")
        return

    janela = tk.Toplevel(tela)
    janela.title("Orçamentos Antigos (CSV)")
    janela.geometry("550x350")
    janela.transient(tela)
    janela.focus_force()

    tk.Label(janela, text="Orçamentos CSV Salvos", font=("Arial", 12, "bold")).pack(pady=10)

    frame_lista = tk.Frame(janela)
    frame_lista.pack(padx=10, pady=10, fill="both", expand=True)

    lista = tk.Listbox(frame_lista, font=("Arial", 10), width=60)
    lista.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_lista, orient="vertical", command=lista.yview)
    scrollbar.pack(side="right", fill="y")
    lista.config(yscrollcommand=scrollbar.set)

    for arquivo in arquivos_csv:
        data_criacao = datetime.fromtimestamp(os.path.getctime(arquivo)).strftime("%d/%m/%Y %H:%M")
        lista.insert(tk.END, f"{os.path.basename(arquivo)}  |  Criado em: {data_criacao}")

    def carregar_csv():
        selecionado = lista.curselection()
        if selecionado:
            nome_arquivo = arquivos_csv[selecionado[0]]
            if messagebox.askyesno("Carregar", "Deseja carregar este orçamento? (O atual será substituído)"):
                try:
                    global df
                    df = pd.read_csv(nome_arquivo)
                    # Armazena o caminho do arquivo aberto
                    tela.arquivo_backup_atual = nome_arquivo
                    atualizar_tabela()
                    janela.destroy()
                    messagebox.showinfo("Sucesso", "Orçamento carregado com sucesso!")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")

    btns_frame = tk.Frame(janela)
    btns_frame.pack(pady=5)

    btn_carregar = tk.Button(btns_frame, text="Carregar Orçamento", command=carregar_csv,
                           bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=20)
    btn_carregar.pack(pady=5)

def novo_orcamento():
    """Cria um novo orçamento, salvando o atual como backup"""
    global df
    
    if not df.empty:
        # Verifica se existem dados para salvar
        resposta = messagebox.askyesno(
            "Novo Orçamento",
            "Deseja salvar o orçamento atual antes de criar um novo?"
        )
        
        if resposta:
            # Cria pasta de backup se não existir
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            pasta_backup = os.path.join(desktop, "Orçamentos_NM_Napoleão", "Orçamentos_Antigos")
            if not os.path.exists(pasta_backup):
                os.makedirs(pasta_backup)
            
            # Salva o arquivo atual com timestamp
            nome_arquivo = f"orcamento_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            caminho_backup = os.path.join(pasta_backup, nome_arquivo)
            df.to_csv(caminho_backup, index=False)
            messagebox.showinfo("Backup", f"Orçamento atual salvo em:\n{caminho_backup}")
    
    # Cria novo DataFrame vazio
    df = pd.DataFrame(columns=[
        "Data", "Fornecedor", "Produto", "Descrição", 
        "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"
    ])
    
    # Limpa a tabela
    atualizar_tabela()
    
    # Limpa os campos de entrada
    for entry in [entry_fornecedor, entry_produto, entry_descricao, 
                 entry_preco, entry_quantidade, entry_ipi, entry_desconto]:
        entry.delete(0, tk.END)
    
    # Limpa os filtros
    limpar_filtros()
    
    # Remove o arquivo atual ao criar novo orçamento
    if hasattr(tela, 'arquivo_atual'):
        delattr(tela, 'arquivo_atual')
    
    messagebox.showinfo("Sucesso", "Novo orçamento criado com sucesso!")

def get_sugestoes():
    """Retorna dicionário com sugestões para cada campo"""
    return {
        'fornecedor': sorted(list(df['Fornecedor'].unique())),
        'produto': sorted(list(df['Produto'].unique())),
        'descricao': sorted(list(df['Descrição'].unique()))
    }

def salvar_alteracoes():
    """Salva explicitamente as alterações no arquivo CSV"""
    try:
        salvar_df()
        messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar alterações:\n{e}")

# Interface
tela = tk.Tk()
tela.title("App de Orçamento com Fornecedores")
tela.geometry("1300x800")

# Adiciona frame para 
frame_logo = tk.Frame(tela)
frame_logo.pack(pady=10)

# Carrega e redimensiona a logo
try:
    # Usa resource_path para localizar a imagem
    logo_path = resource_path("logo.png")
    imagem = Image.open(logo_path)
    basewidth = 200
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
menu_atalhos = tk.Menu(menu_bar, tearoff=0)  # Novo menu para atalhos
tela.config(menu=menu_bar)

# commands to menu
menu_opcoes.add_command(label="Mostrar Orçamentos", command=mostrar_orcamentos)
menu_opcoes.add_separator()
menu_opcoes.add_command(label="Novo Orçamento", command=novo_orcamento)
menu_opcoes.add_command(label="Orçamentos Antigos", command=mostrar_csvs_antigos)

# Adiciona os atalhos no submenu
menu_atalhos.add_command(label="Salvar (Ctrl+S)")
menu_atalhos.add_command(label="Novo Orçamento (Ctrl+N)")
menu_atalhos.add_command(label="Abrir Orçamentos Antigos (Ctrl+O)")
menu_atalhos.add_command(label="Gerar PDF (Ctrl+P)")
menu_atalhos.add_command(label="Gerar Gráfico (Ctrl+G)")
menu_atalhos.add_command(label="Focar Filtro (Ctrl+F)")
menu_atalhos.add_command(label="Atualizar Tabela (F5)")
menu_atalhos.add_command(label="Remover Selecionado (Delete)")

# Adiciona os menus na barra
menu_bar.add_cascade(label="Opções", menu=menu_opcoes)
menu_bar.add_cascade(label="Atalhos", menu=menu_atalhos)

fonte_label = ("Arial", 11, "bold")
padrao_entry = {"font": ("Arial", 10), "width": 30}

frame_entrada = tk.Frame(tela)
frame_entrada.pack(pady=10)

# Labels e Entradas (padronizadas)
labels_entrada = ["Fornecedor", "Produto", "Descrição", "Preço Unitário (R$)", "Quantidade", "IPI (%)", "Desconto (%)"]
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

entry_fornecedor, entry_produto, entry_descricao, entry_preco, entry_quantidade, entry_ipi, entry_desconto = entries_entrada

# Botões principais
frame_botoes = tk.Frame(tela)
frame_botoes.pack(pady=10)

btn_add = tk.Button(frame_botoes, text="Adicionar Item", command=adicionar_item, 
                   bg="#4CAF50", fg="white", font=fonte_label, width=15)
btn_add.grid(row=0, column=0, padx=5)

btn_grafico = tk.Button(frame_botoes, text="Gerar Gráfico", command=gerar_grafico, 
                       bg="#9C27B0", fg="white", font=fonte_label, width=15)
btn_grafico.grid(row=0, column=1, padx=5)

btn_pdf = tk.Button(frame_botoes, text="Gerar PDF", command=gerar_pdf, 
                   bg="#607D8B", fg="white", font=fonte_label, width=15)
btn_pdf.grid(row=0, column=2, padx=5)

btn_salvar = tk.Button(frame_botoes, text="Salvar Alterações", command=salvar_alteracoes, 
                      bg="#FF9800", fg="white", font=fonte_label, width=15)
btn_salvar.grid(row=0, column=3, padx=5)

btn_remover = tk.Button(frame_botoes, text="Remover Orçamento", command=remover_selecionado, 
                       bg="#F00707", fg="white", font=fonte_label, width=20)
btn_remover.grid(row=0, column=4, padx=5)

# Filtros
filtros_frame = tk.LabelFrame(tela, text="🔍 Filtros Avançados", padx=10, pady=10, font=("Arial", 10, "bold"))
filtros_frame.pack(pady=10, fill="x")

tk.Label(filtros_frame, text="Fornecedor:").grid(row=0, column=0)
filtro_fornecedor = tk.Entry(filtros_frame, font=("Arial", 10), width=25)
filtro_fornecedor.grid(row=0, column=1, padx=5)

tk.Label(filtros_frame, text="Produto:").grid(row=0, column=2)
filtro_produto = tk.Entry(filtros_frame, font=("Arial", 10), width=25)
filtro_produto.grid(row=0, column=3, padx=5)

tk.Label(filtros_frame, text="Descrição:").grid(row=0, column=4)
filtro_descricao = tk.Entry(filtros_frame, font=("Arial", 10), width=25)
filtro_descricao.grid(row=0, column=5, padx=5)

var_ordem = tk.StringVar()
ordem_menu = ttk.Combobox(
    filtros_frame,
    textvariable=var_ordem,
    values=["", "A-Z", "Maior Desconto", "Menor IPI", "Maior IPI", "Menor Preço", "Maior Valor"],
    state="readonly",
    width=18
)
ordem_menu.grid(row=0, column=6, padx=5)
ordem_menu.set("Ordenar por...")

btn_filtrar = tk.Button(filtros_frame, text="Aplicar Filtros", command=aplicar_filtros, bg="#2196F3", fg="white", font=fonte_label, width=15)
btn_filtrar.grid(row=0, column=7, padx=5)

btn_limpar = tk.Button(filtros_frame, text="Limpar Filtros", command=limpar_filtros, bg="#F44336", fg="white", font=fonte_label, width=15)
btn_limpar.grid(row=0, column=8, padx=5)

# Tabela
frame_tabela = tk.Frame(tela)
frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

# Criar scrollbars
scrollbar_y = ttk.Scrollbar(frame_tabela)
scrollbar_y.pack(side="right", fill="y")

scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal")
scrollbar_x.pack(side="bottom", fill="x")

# Criar tabela com suporte a seleção múltipla e scrollbars
colunas = ["Fornecedor", "Produto", "Descrição", "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"]
tabela = ttk.Treeview(
    frame_tabela, 
    columns=colunas, 
    show="headings", 
    selectmode="extended",  # Permite seleção múltipla
    yscrollcommand=scrollbar_y.set,
    xscrollcommand=scrollbar_x.set
)

# Configurar scrollbars
scrollbar_y.config(command=tabela.yview)
scrollbar_x.config(command=tabela.xview)

# Configurar colunas
for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, width=150, anchor="center")

tabela.pack(fill="both", expand=True)

# Adicione este estilo para melhorar a visualização da seleção
style = ttk.Style()
style.map('Treeview',
    foreground=[('selected', 'white')],
    background=[('selected', '#0078D7')]  # Cor azul do Windows para seleção
)

# Adicione estas funções para manipular a seleção
def on_select(event):
    """Atualiza a contagem de itens selecionados"""
    selecionados = len(tabela.selection())
    if selecionados > 0:
        btn_remover.config(text=f"Remover ({selecionados})")
    else:
        btn_remover.config(text="Remover Orçamento")

# Vincule o evento de seleção à tabela
tabela.bind('<<TreeviewSelect>>', on_select)

# Inicializa tabela com dados
# Configuração inicial do DataFrame vazio
df = pd.DataFrame(columns=[
    "Data", "Fornecedor", "Produto", "Descrição", 
    "Preço Unitário", "Quantidade", "IPI", "Desconto", "Total Final"
])

# Atualiza a tabela vazia
atualizar_tabela()

tela.protocol("WM_DELETE_WINDOW", confirmar_saida)

# Adicione esta função para lidar com os atalhos de teclado
def setup_hotkeys():
    # Adiciona verificação de plataforma
    mod = "Command" if sys.platform == "darwin" else "Control"
    tela.bind(f"<{mod}-s>", lambda e: salvar_alteracoes())
    tela.bind(f"<{mod}-n>", lambda e: novo_orcamento())
    tela.bind(f"<{mod}-o>", lambda e: mostrar_csvs_antigos())
    tela.bind(f"<{mod}-p>", lambda e: gerar_pdf())
    tela.bind(f"<{mod}-g>", lambda e: gerar_grafico())
    tela.bind("<Delete>", lambda e: remover_selecionado())
    tela.bind("<F5>", lambda e: atualizar_tabela())

def focar_filtro(event=None):
    """Função para focar no campo de filtro de fornecedor"""
    filtro_fornecedor.focus_set()

setup_hotkeys()

tela.mainloop()