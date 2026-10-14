import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def limpar_numero(x):
    """Converte os valores em formato string do Brasil para float."""
    if pd.isna(x):
        return 0.0
    s = str(x).replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

def extrair_nome_produto(nome_completo):
    """Remove os prefixos 'C.G.O :' ou 'FILIAL :' para deixar apenas o nome do produto."""
    if pd.isna(nome_completo):
        return "Desconhecido"
    partes = str(nome_completo).split(' : ', 1)
    return partes[1] if len(partes) > 1 else str(nome_completo)

def selecionar_arquivo_cgo():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
    if caminho:
        lbl_cgo.config(text=caminho)
        janela.caminho_cgo = caminho

def selecionar_arquivo_filial():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
    if caminho:
        lbl_filial.config(text=caminho)
        janela.caminho_filial = caminho

def processar_e_exportar():
    if not hasattr(janela, 'caminho_cgo') or not hasattr(janela, 'caminho_filial'):
        messagebox.showwarning("Atenção", "Por favor, selecione os dois arquivos antes de processar.")
        return

    caminho_salvar = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        title="Salvar Relatório Como"
    )
    
    if not caminho_salvar:
        return

    try:
        # Lê os arquivos (tenta utf-8, se falhar, usa latin1)
        try:
            df_cgo = pd.read_csv(janela.caminho_cgo, sep=';', encoding='utf-8')
        except UnicodeDecodeError:
            df_cgo = pd.read_csv(janela.caminho_cgo, sep=';', encoding='latin1')
            
        try:
            df_filial = pd.read_csv(janela.caminho_filial, sep=';', encoding='utf-8')
        except UnicodeDecodeError:
            df_filial = pd.read_csv(janela.caminho_filial, sep=';', encoding='latin1')

        # Limpeza e formatação de colunas numéricas
        for df in [df_cgo, df_filial]:
            df['Qtd Saída Outras Num'] = df['Qtd Saída Outras'].apply(limpar_numero)
            df['Vlr Saída Outras Num'] = df['Vlr Saída Outras'].apply(limpar_numero)

        # 1. Base CGO (Detalhada por CGO para manter o Nome do Produto fiel ao CGO)
        cgos_detalhado = df_cgo.groupby(['Código Produto Base', 'Código C.G.O.']).agg({
            'C.G.O. : Produto Base': 'first'
        }).reset_index()
        cgos_detalhado['PRODUTO_CGO'] = cgos_detalhado['C.G.O. : Produto Base'].apply(extrair_nome_produto)
        cgos_detalhado.rename(columns={'Código C.G.O.': 'CGO'}, inplace=True)
        cgos_detalhado['CGO'] = cgos_detalhado['CGO'].apply(lambda x: str(int(x)) if pd.notna(x) else '-')

        # Conta quantos CGOs existem para cada produto
        contagem_cgos = cgos_detalhado.groupby('Código Produto Base')['CGO'].nunique().reset_index()
        contagem_cgos.rename(columns={'CGO': 'Qtd_CGOs'}, inplace=True)

        # 2. Base FILIAL
        filiais_detalhado = df_filial.groupby(['Código Produto Base', 'Código Empresa']).agg({
            'Qtd Saída Outras Num': 'sum',
            'Vlr Saída Outras Num': 'sum'
        }).reset_index()
        filiais_detalhado.rename(columns={'Código Empresa': 'FILIAL', 'Qtd Saída Outras Num': 'Qtd Saída Outras', 'Vlr Saída Outras Num': 'Vlr Saída Outras'}, inplace=True)
        filiais_detalhado['FILIAL'] = filiais_detalhado['FILIAL'].apply(lambda x: str(int(x)) if pd.notna(x) else 'Sem Filial')

        # Filtro: Remove as linhas sem movimentação na filial
        filiais_detalhado = filiais_detalhado[(filiais_detalhado['Qtd Saída Outras'] != 0) & (filiais_detalhado['Qtd Saída Outras'].notna())]

        # Junta a contagem de CGOs na base da filial para saber o que separar
        filiais_c_contagem = pd.merge(filiais_detalhado, contagem_cgos, on='Código Produto Base', how='inner')

        # 3. SEPARAÇÃO DOS DADOS
        
        # A) Produtos com apenas 1 CGO (Vão para as abas de Filiais)
        df_unico = filiais_c_contagem[filiais_c_contagem['Qtd_CGOs'] == 1]
        df_unico_final = pd.merge(df_unico, cgos_detalhado[['Código Produto Base', 'CGO', 'PRODUTO_CGO']], on='Código Produto Base', how='left')
        df_unico_final.rename(columns={'PRODUTO_CGO': 'PRODUTO'}, inplace=True)
        
        colunas_padrao = ['PRODUTO', 'CGO', 'FILIAL', 'Qtd Saída Outras', 'Vlr Saída Outras']
        df_unico_final = df_unico_final[colunas_padrao]

        # B) Produtos com MÚLTIPLOS CGOs (Vão para a aba de auditoria)
        df_multi = filiais_c_contagem[filiais_c_contagem['Qtd_CGOs'] > 1]
        # Aqui o 'inner join' vai multiplicar as linhas (uma linha para cada CGO), exatamente como você quer visualizar
        df_multi_final = pd.merge(df_multi, cgos_detalhado[['Código Produto Base', 'CGO', 'PRODUTO_CGO']], on='Código Produto Base', how='left')
        df_multi_final.rename(columns={'PRODUTO_CGO': 'PRODUTO'}, inplace=True)
        
        # Organiza a aba multi na ordem que você pediu: Filial, CGO, Produto, Valor...
        colunas_multi = ['FILIAL', 'CGO', 'PRODUTO', 'Qtd Saída Outras', 'Vlr Saída Outras']
        df_multi_final = df_multi_final[colunas_multi].sort_values(by=['FILIAL', 'PRODUTO', 'CGO'])

        # 4. EXPORTAÇÃO
        with pd.ExcelWriter(caminho_salvar, engine='openpyxl') as writer:
            
            # Exporta as Filiais (apenas os itens 100% corretos/com 1 CGO)
            if not df_unico_final.empty:
                filiais_unicas = df_unico_final['FILIAL'].unique()
                filiais_num = sorted([f for f in filiais_unicas if f.isdigit()], key=int)
                filiais_outras = [f for f in filiais_unicas if not f.isdigit()]
                
                for filial in filiais_num + filiais_outras:
                    df_aba = df_unico_final[df_unico_final['FILIAL'] == filial]
                    nome_aba = f"Filial {filial}" if filial != 'Sem Filial' else "Sem Filial"
                    df_aba.to_excel(writer, sheet_name=nome_aba, index=False)
            
            # Exporta os Multi CGOs (detalhado linha a linha)
            if not df_multi_final.empty:
                df_multi_final.to_excel(writer, sheet_name="PRODUTOS CGOS DIFERENTES", index=False)

        messagebox.showinfo("Sucesso", "Relatório gerado! Itens com múltiplos CGOs foram separados e detalhados na aba final.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

# --- Interface Gráfica ---
janela = tk.Tk()
janela.title("Conciliador Rigoroso (CGO x Filial)")
janela.geometry("500x250")
janela.eval('tk::PlaceWindow . center')

tk.Label(janela, text="Selecione os arquivos para cruzar:", font=("Arial", 10, "bold")).pack(pady=10)

btn_cgo = tk.Button(janela, text="1. Selecionar CGO SINT", command=selecionar_arquivo_cgo, width=30)
btn_cgo.pack(pady=5)
lbl_cgo = tk.Label(janela, text="Nenhum selecionado", fg="gray")
lbl_cgo.pack()

btn_filial = tk.Button(janela, text="2. Selecionar FILIAL SINT", command=selecionar_arquivo_filial, width=30)
btn_filial.pack(pady=5)
lbl_filial = tk.Label(janela, text="Nenhum selecionado", fg="gray")
lbl_filial.pack()

btn_exportar = tk.Button(janela, text="Gerar Excel", command=processar_e_exportar, width=30, bg="green", fg="white", font=("Arial", 10, "bold"))
btn_exportar.pack(pady=20)

janela.mainloop()