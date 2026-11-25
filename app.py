# ==============================================================================
# APLICATIVO DESKTOP - PREENCHIMENTO IAUTOMÁTICO
# Objetivo: Fornecer uma interface visual para que os engenheiros selecionem
# planilhas, submetam os dados à IA e injetem as decisões de volta no Excel.
# ==============================================================================

# Importação das bibliotecas essenciais
import tkinter as tk                             # Para criar a interface gráfica (janela, botões)
from tkinter import filedialog, messagebox       # Para caixas de seleção de arquivos e alertas
import pandas as pd                              # Para leitura e manipulação da planilha em memória
import joblib                                    # Para carregar o cérebro da IA (modelo .pkl)
import os                                        # Para manipulação de caminhos de arquivos
import sys                                       # Para identificar se está rodando como .exe ou script
import sklearn                                   # Biblioteca base de Machine Learning
from sklearn.ensemble import RandomForestClassifier # O algoritmo utilizado
import win32com.client as win32                  # Para controlar o aplicativo do Excel silenciosamente

# ==============================================================================
# 1. LOCALIZANDO O CÉREBRO DA IA (CARREGAMENTO DOS ARQUIVOS .PKL)
# Lógica adaptativa que acha os arquivos de IA independente de como o app foi aberto.
# ==============================================================================
try:
    # Se 'frozen' for True, significa que o código foi compilado em um .exe pelo PyInstaller
    if getattr(sys, 'frozen', False):
        diretorio_actual = os.path.dirname(sys.executable)
    # Caso contrário, está rodando direto do VSCode/Python
    else:
        diretorio_actual = os.path.dirname(os.path.abspath(__file__))
        
    # Carrega a rede de decisões matemáticas da IA e a ordem exata das colunas esperadas
    modelo_rf = joblib.load(os.path.join(diretorio_actual, 'cerebro_ia.pkl'))
    colunas_treinadas = joblib.load(os.path.join(diretorio_actual, 'colunas_ia.pkl'))
except Exception as e:
    messagebox.showerror("Erro de Leitura", f"Ocorreu o seguinte erro ao carregar a IA:\n\n{e}")
    sys.exit() # Encerra o programa caso os arquivos não sejam encontrados

# ==============================================================================
# 2. FUNÇÃO DE RADAR (CÃO FAREJADOR)
# Escaneia o Excel do usuário para descobrir onde as colunas foram parar.
# ==============================================================================
def mapear_colunas(df_raw):
    mapa = {}
    
    # Passo A: Busca os cabeçalhos fixos (Métricas e Decisão) varrendo o topo da planilha
    for col in range(df_raw.shape[1]):
        texto_coluna = " ".join([str(x).lower() for x in df_raw.iloc[0:15, col].tolist()])
        
        if 'decision' in texto_coluna and 'final' in texto_coluna:
            mapa['Decisao'] = col
        if 'qo [w]' in texto_coluna:
            mapa['Qo_A'] = col; mapa['Qo_B'] = col + 1
        if 'wc [w]' in texto_coluna:
            mapa['Wc_A'] = col; mapa['Wc_B'] = col + 1
        if 'cop [-]' in texto_coluna:
            mapa['COP_A'] = col; mapa['COP_B'] = col + 1
        if 'speed [rpm]' in texto_coluna and 'deviation' not in texto_coluna:
            mapa['RPM_A'] = col; mapa['RPM_B'] = col + 1
        if 'deviation to avg' in texto_coluna:
            mapa['Var_Qo'] = col; mapa['Var_Wc'] = col + 1; mapa['Var_COP'] = col + 2; mapa['Var_RPM'] = col + 3

    # Passo B: Farejador de OK e Condição (Bypass de cabeçalhos mesclados)
    for col in range(30, min(60, df_raw.shape[1])):
        # Ampliamos a amostra (linha 10 a 30) para detectar variações no tamanho do cabeçalho
        amostra_dados = df_raw.iloc[9:30, col].astype(str).str.strip().str.upper().tolist()
        
        if 'OK' in amostra_dados:
            mapa['Flag_OK'] = col
            mapa['Condicao'] = col + 1  # A Condição é sempre a vizinha da direita!
            
            # Descobre a linha exata onde os dados começam achando a posição do primeiro 'OK'
            # (Somamos 9 pois a nossa amostra_dados começou no índice 9 do Pandas)
            mapa['Indice_Inicio_Dados'] = 9 + amostra_dados.index('OK')
            break
            
    return mapa

# ==============================================================================
# 3. FUNÇÃO PRINCIPAL (ANÁLISE E PREENCHIMENTO NATIVO EM BLOCO)
# ==============================================================================
def processar_planilha():
    # Abre a caixa de diálogo do Windows para o usuário escolher o arquivo
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a Planilha do Teste",
        filetypes=[("Excel", "*.xlsm;*.xlsx")]
    )
    
    # Se o usuário clicar em "Cancelar", a função encerra silenciosamente
    if not caminho_arquivo:
        return 
        
    caminho_arquivo = os.path.normpath(caminho_arquivo)
    
    try:
        label_status.config(text="Escaneando estrutura da planilha...", fg="blue")
        janela.update() # Força a tela a atualizar a mensagem imediatamente
        
        # Lê a planilha bruta e mapeia as coordenadas com o Radar
        df_raw = pd.read_excel(caminho_arquivo, sheet_name='CAL', header=None)
        mapa = mapear_colunas(df_raw)
        
        # Trava de segurança: impede que a IA tente chutar dados se o template for muito diferente
        if 'Condicao' not in mapa or 'Decisao' not in mapa:
            raise ValueError("O Radar não conseguiu localizar as colunas vitais ('OK', Condição ou Decisão Final) nesta planilha.")
            
        # Constrói a tabela estruturada usando as coordenadas do mapa
        df_novo = pd.DataFrame()
        df_novo['Flag_AW'] = df_raw.iloc[:, mapa['Flag_OK']]
        df_novo['Condicao_Ensaio'] = df_raw.iloc[:, mapa['Condicao']]
        df_novo['Qo_A'] = df_raw.iloc[:, mapa['Qo_A']]; df_novo['Qo_B'] = df_raw.iloc[:, mapa['Qo_B']]
        df_novo['Wc_A'] = df_raw.iloc[:, mapa['Wc_A']]; df_novo['Wc_B'] = df_raw.iloc[:, mapa['Wc_B']]
        df_novo['COP_A'] = df_raw.iloc[:, mapa['COP_A']]; df_novo['COP_B'] = df_raw.iloc[:, mapa['COP_B']]
        df_novo['RPM_A'] = df_raw.iloc[:, mapa['RPM_A']]; df_novo['RPM_B'] = df_raw.iloc[:, mapa['RPM_B']]
        df_novo['Var_Qo'] = df_raw.iloc[:, mapa['Var_Qo']]; df_novo['Var_Wc'] = df_raw.iloc[:, mapa['Var_Wc']]
        df_novo['Var_COP'] = df_raw.iloc[:, mapa['Var_COP']]; df_novo['Var_RPM'] = df_raw.iloc[:, mapa['Var_RPM']]
        
        # Aplica o filtro de linhas válidas ("OK")
        linha_inicial = mapa['Indice_Inicio_Dados']
        
        # TRAVA DE SEGURANÇA: Se o radar marcou linha 12 (índice 12), mantemos. 
        # Se marcou linha 13 (índice 12 no Pandas), forçamos começar na 11 (linha 12 no Excel) 
        # para garantir que nenhuma linha útil seja esquecida no topo.
        if linha_inicial > 11:
            linha_inicial = 11
            
        df_novo = df_novo.iloc[linha_inicial:].reset_index(drop=True)
        
        # Mantém apenas as linhas onde o engenheiro atestou como "OK"
        df_novo = df_novo[df_novo['Flag_AW'].astype(str).str.strip().str.upper() == 'OK']
        df_novo = df_novo.drop(columns=['Flag_AW']) # Descarta a flag pois ela não vai para a IA
        
        if df_novo.empty:
            raise ValueError("Não foi encontrada nenhuma linha marcada com 'OK' nesta planilha.")
            
        # Separa as condições de ambiente (Texto -> Colunas)
        cond_sep = df_novo['Condicao_Ensaio'].astype(str).str.split(expand=True)
        df_novo['Tev'] = cond_sep[0]; df_novo['Tcd'] = cond_sep[1]
        df_novo['Tsuc'] = cond_sep[2] if 2 in cond_sep.columns else None
        df_novo['speed'] = cond_sep[3] if 3 in cond_sep.columns else None
        
        # Transforma o texto (categorias) em colunas binárias, assim como foi feito no treino
        colunas_existentes = [col for col in ['Tev', 'Tcd', 'Tsuc', 'speed'] if col in df_novo.columns]
        df_novo = pd.get_dummies(df_novo, columns=colunas_existentes, dummy_na=False)
        df_novo = df_novo.drop(columns=['Condicao_Ensaio'])
        
        # Reordena e preenche colunas ausentes para garantir que a IA receba o formato exato que ela conhece
        X_novo = df_novo.reindex(columns=colunas_treinadas, fill_value=0)
        
        # O momento mágico: A IA avalia as linhas e toma as decisões
        previsoes_ia = modelo_rf.predict(X_novo)
        
        # ==============================================================================
        # INJEÇÃO CIRÚRGICA TURBO (BULK INSERT NO EXCEL)
        # Ao invés de gravar célula por célula, gravamos um bloco gigantesco de uma vez.
        # ==============================================================================
        label_status.config(text="Injetando as decisões no Excel...", fg="orange")
        janela.update()
        
        # Acorda o aplicativo real do Excel no Windows (mas deixa ele invisível)
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False 
        excel.DisplayAlerts = False # Desativa pop-ups irritantes do Excel
        
        try:
            # Abre o arquivo fisicamente
            wb = excel.Workbooks.Open(caminho_arquivo)
            ws = wb.Worksheets("CAL")
            
            # Descobre em qual coluna (Ex: BW = 75) a IA deve colocar as respostas
            coluna_decisao_excel = mapa['Decisao'] + 1 
            
            # O Excel começa a contar linhas a partir do 1 (o Pandas a partir do 0)
            # Aplica a mesma trava de segurança para a escrita no Excel
            linha_inicial_pandas = mapa['Indice_Inicio_Dados']
            if linha_inicial_pandas > 11:
                linha_inicial_pandas = 11
                
            linha_inicial_excel = linha_inicial_pandas + 1 
            total_linhas = len(previsoes_ia)
            
            # Formata a resposta da IA em Lista de Listas (o padrão que o Excel exige)
            valores_coluna = [[decisao] for decisao in previsoes_ia]
            
            # Demarca a área exata onde as respostas vão cair
            celula_inicio = ws.Cells(linha_inicial_excel, coluna_decisao_excel)
            celula_fim = ws.Cells(linha_inicial_excel + total_linhas - 1, coluna_decisao_excel)
            
            # Despeja as milhares de previsões no Excel em menos de 1 segundo!
            ws.Range(celula_inicio, celula_fim).Value = valores_coluna
            
            wb.Save()
            wb.Close()
        finally:
            excel.Quit() # É crucial encerrar o processo para não travar o Windows do usuário
        
        # Feedback visual de sucesso
        label_status.config(text="Sucesso Total!", fg="green")
        messagebox.showinfo("Sucesso", f"Análise instantânea concluída!\n{total_linhas} linhas válidas foram preenchidas na coluna {coluna_decisao_excel}.")
        
    except Exception as e:
        label_status.config(text="Erro.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar a planilha:\n{e}")

# ==============================================================================
# 4. CRIAÇÃO DA INTERFACE VISUAL (TKINTER)
# Desenho da janelinha amigável do software.
# ==============================================================================
janela = tk.Tk()
janela.title("Validador IA - Preenchimento IAutomático")
janela.geometry("450x200") # Tamanho fixo da janela (Largura x Altura)

# Rótulo de Título Principal
tk.Label(janela, text="Auditoria Automática de Compressores", font=("Arial", 12, "bold")).pack(pady=15)

# Botão de Ação que aciona toda a mágica
tk.Button(janela, text="Analisar e Preencher Planilha", command=processar_planilha, bg="#007ACC", fg="white", font=("Arial", 10, "bold"), width=30).pack(pady=10)

# Barra de status no rodapé para acalmar a ansiedade do usuário
label_status = tk.Label(janela, text="Aguardando arquivo...", fg="gray")
label_status.pack(pady=10)

# Inicia o loop eterno do aplicativo (mantém ele aberto até o usuário clicar no X)
janela.mainloop()