import streamlit as st
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt # Matplotlib não está sendo usado nos gráficos
# import seaborn as sns # Seaborn não está sendo usado
import io
import base64
import re
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import traceback # Para exibir detalhes do erro
import math # Para truncamento (se necessário, mas vamos focar na formatação)

# Configuração da página
st.set_page_config(
    page_title="Gerador de Curva ABC - SINAPI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado (mantido como estava)
st.markdown("""
<style>
    .reportview-container {
        background-color: #f0f2f6;
    }
    .sidebar .sidebar-content {
        background-color: #f0f2f6;
    }
    h1, h2, h3 {
        color: #1e3c72;
    }
    .stButton>button {
        background-color: #1e3c72;
        color: white;
        border-radius: 5px; /* Adicionado arredondamento */
        border: none; /* Removido borda padrão */
        padding: 10px 15px; /* Ajustado padding */
    }
    .stButton>button:hover {
        background-color: #2a5298; /* Cor um pouco mais clara no hover */
    }
    .highlight {
        background-color: #ffffff; /* Fundo branco para destaque */
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1e3c72;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1); /* Sombra suave */
        margin-bottom: 20px; /* Espaçamento inferior */
    }
    .footer {
        margin-top: 40px; /* Mais espaço antes do rodapé */
        padding-top: 10px;
        border-top: 1px solid #ddd;
        text-align: center;
        font-size: 0.8em;
        color: #666;
    }
    /* Melhorar aparência dos expanders */
    .streamlit-expanderHeader {
        background-color: #e8eaf6;
        color: #1e3c72;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.title("📊 Gerador de Curva ABC - SINAPI")
st.markdown("### Automatize a geração da Curva ABC a partir de planilhas sintéticas do SINAPI")

# --- Funções Auxiliares ---

def detectar_delimitador(sample_content):
    """Detecta o delimitador mais provável em uma amostra de conteúdo CSV."""
    delimiters = [';', ',', '\t', '|']
    counts = {d: sample_content.count(d) for d in delimiters}
    # Prioriza ';' se a contagem for similar a ',', comum no Brasil
    if counts.get(';', 0) > 0 and counts[';'] >= counts.get(',', 0) * 0.8:
         return ';'
    # Remove delimitadores com contagem zero antes de encontrar o máximo
    counts = {k: v for k, v in counts.items() if v > 0}
    if not counts:
        return ',' # Retorna vírgula como padrão se nada for encontrado
    return max(counts, key=counts.get)

def encontrar_linha_cabecalho(df_preview):
    """Encontra a linha que provavelmente contém os cabeçalhos."""
    # Expandido para incluir mais termos comuns e os novos campos
    cabecalhos_possiveis = [
        'CÓDIGO', 'ITEM', 'DESCRIÇÃO', 'CUSTO', 'VALOR', 'TOTAL', 'PREÇO', 'SERVIÇO',
        'UNID', 'UNIDADE', 'UM', 'QUANT', 'QTD', 'QUANTIDADE', 'UNITÁRIO', 'UNITARIO'
    ]
    max_matches = 0
    header_row_index = 0

    for i in range(min(20, len(df_preview))):
        try:
            # Pega valores não nulos da linha, converte para string e maiúsculas
            row_values = df_preview.iloc[i].dropna().astype(str).str.upper().tolist()
            # Conta quantas células na linha contêm alguma das palavras-chave
            current_matches = sum(any(keyword in cell for keyword in cabecalhos_possiveis) for cell in row_values)

            # Considera a linha com mais palavras-chave como cabeçalho
            # Adiciona um bônus se encontrar 'DESCRIÇÃO' ou 'CODIGO', que são mais prováveis
            if 'DESCRIÇÃO' in row_values or 'CODIGO' in row_values or 'CÓDIGO' in row_values:
                 current_matches += 2

            if current_matches > max_matches:
                max_matches = current_matches
                header_row_index = i
        except Exception:
            continue # Ignora linhas problemáticas

    # Lógica de fallback (mantida)
    if max_matches < 2 and df_preview.iloc[0].isnull().all():
         if len(df_preview) > 1 and not df_preview.iloc[1].isnull().all(): return 1
         else: return 0
    elif max_matches == 0:
         return 0

    return header_row_index

def sanitizar_dataframe(df):
    """Sanitiza o DataFrame para garantir compatibilidade com Streamlit/PyArrow."""
    if df is None: return None
    df_clean = df.copy()
    new_columns = []
    seen_columns = {}
    for i, col in enumerate(df_clean.columns):
        col_str = str(col).strip() if pd.notna(col) else f"coluna_{i}"
        if not col_str: col_str = f"coluna_{i}" # Garante que não seja vazio
        while col_str in seen_columns: # Renomeia duplicatas
            seen_columns[col_str] = seen_columns.get(col_str, 0) + 1
            col_str = f"{col_str.rsplit('_', 1)[0]}_{seen_columns[col_str]}" if '_' in col_str and col_str.rsplit('_', 1)[-1].isdigit() else f"{col_str}_1"
        seen_columns[col_str] = 0
        new_columns.append(col_str)
    df_clean.columns = new_columns

    for col in df_clean.columns:
        try:
            # Tenta converter para numérico primeiro (ignora erros, mantém object se falhar)
            df_clean[col] = pd.to_numeric(df_clean[col], errors='ignore')
            # Se não for numérico, tenta datetime (ignora erros)
            if df_clean[col].dtype == 'object':
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='ignore')
                except Exception: # Captura exceções mais amplas de datetime
                     pass
            # Se ainda for 'object' ou se for tipo misto, converte para string
            if df_clean[col].dtype == 'object' or df_clean[col].apply(type).nunique() > 1:
                 # Verifica se a conversão para string é segura
                 if df_clean[col].isnull().all() or df_clean[col].apply(lambda x: isinstance(x, (str, int, float, bool))).all():
                      df_clean[col] = df_clean[col].astype(str).replace('nan', '', regex=False).replace('NaT', '', regex=False)
                 # else: st.warning(f"Coluna {col} com tipos complexos não convertida para string.")

            # Remove caracteres nulos se for string
            if isinstance(df_clean[col].dtype, pd.StringDtype) or df_clean[col].dtype == 'object':
                 if df_clean[col].apply(lambda x: isinstance(x, str)).any():
                      df_clean[col] = df_clean[col].str.replace('\x00', '', regex=False)

        except Exception as e:
            # st.warning(f"Erro ao sanitizar coluna '{col}': {e}. Tentando converter para string.")
            try: df_clean[col] = df_clean[col].astype(str)
            except Exception: st.error(f"Falha crítica ao converter coluna '{col}' para string.")

    df_clean = df_clean.dropna(how='all').dropna(axis=1, how='all')
    return df_clean


# --- Função Principal de Processamento ---

def processar_arquivo(uploaded_file):
    """Carrega e processa o arquivo CSV ou Excel, identificando o cabeçalho."""
    df = None
    delimitador = None
    linha_cabecalho = 0 # Default
    encodings_to_try = ['utf-8', 'latin1', 'cp1252']

    try:
        file_name = uploaded_file.name.lower()
        file_content = uploaded_file.getvalue()

        # --- Processamento Excel ---
        if file_name.endswith(('.xlsx', '.xls')):
            try:
                # Tenta ler com openpyxl primeiro (mais novo)
                df_preview = pd.read_excel(io.BytesIO(file_content), engine='openpyxl', nrows=25, header=None)
            except Exception:
                # Se falhar, tenta com xlrd (mais antigo, pode precisar instalar)
                try:
                    st.warning("Falha ao ler com 'openpyxl', tentando com 'xlrd' (pode ser necessário instalar: pip install xlrd)")
                    df_preview = pd.read_excel(io.BytesIO(file_content), engine='xlrd', nrows=25, header=None)
                except Exception as e_xlrd:
                    st.error(f"Erro ao ler preview do Excel com ambos os engines: {e_xlrd}")
                    return None, None

            linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            # Ler arquivo completo com o cabeçalho correto e engine que funcionou
            engine_to_use = 'openpyxl' if 'df_preview' in locals() and isinstance(df_preview, pd.DataFrame) else 'xlrd'
            df = pd.read_excel(io.BytesIO(file_content), engine=engine_to_use, header=linha_cabecalho)


        # --- Processamento CSV ---
        elif file_name.endswith('.csv'):
            detected_encoding = None
            decoded_content = None
            for enc in encodings_to_try:
                try:
                    decoded_content = file_content.decode(enc)
                    detected_encoding = enc
                    break
                except UnicodeDecodeError: continue
            if decoded_content is None:
                st.error("Erro de decodificação CSV. Verifique o encoding.")
                return None, None
            if not decoded_content.strip():
                 st.error("Arquivo CSV vazio.")
                 return None, None

            sample_for_delimiter = decoded_content[:5000]
            delimitador = detectar_delimitador(sample_for_delimiter)

            try:
                df_preview = pd.read_csv(io.StringIO(decoded_content), delimiter=delimitador, nrows=25, header=None, skipinitialspace=True, low_memory=False)
                linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            except Exception as e:
                st.warning(f"Erro ao ler preview do CSV: {e}. Assumindo linha 0.")
                linha_cabecalho = 0

            df = pd.read_csv(
                io.StringIO(decoded_content), delimiter=delimitador, header=linha_cabecalho,
                encoding=detected_encoding, on_bad_lines='warn', skipinitialspace=True, low_memory=False
            )

        else:
            st.error("Formato de arquivo não suportado.")
            return None, None

        # --- Pós-processamento Comum ---
        if df is not None:
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty:
                 st.error("Arquivo vazio após remover linhas/colunas nulas.")
                 return None, delimitador
            df = sanitizar_dataframe(df)
            if df is None or df.empty:
                 st.error("Falha na sanitização do DataFrame.")
                 return None, delimitador
            return df, delimitador
        else:
             return None, delimitador

    except Exception as e:
        st.error(f"Erro fatal ao processar o arquivo: {str(e)}")
        with st.expander("Detalhes técnicos do erro"): st.text(traceback.format_exc())
        return None, None

# --- Funções da Curva ABC (limpeza, identificação, geração) ---

def limpar_valor(valor):
    """Limpa e converte valores monetários para float, tratando diversos formatos."""
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float, np.number)): return float(valor)
    valor_str = str(valor).strip()
    if not valor_str: return 0.0
    valor_str = re.sub(r'[R$€£¥\s]', '', valor_str)
    has_comma = ',' in valor_str
    has_dot = '.' in valor_str
    if has_comma and has_dot:
        last_comma_pos = valor_str.rfind(',')
        last_dot_pos = valor_str.rfind('.')
        if last_comma_pos > last_dot_pos: valor_str = valor_str.replace('.', '').replace(',', '.')
        else: valor_str = valor_str.replace(',', '')
    elif has_comma: valor_str = valor_str.replace(',', '.')
    valor_str = re.sub(r'[^\d.]', '', valor_str)
    try: return float(valor_str) if valor_str else 0.0
    except ValueError: return 0.0

# Função para limpar quantidade (similar a valor, mas sem R$)
def limpar_quantidade(qtd):
    """Limpa e converte valores de quantidade para float."""
    if pd.isna(qtd): return 0.0
    if isinstance(qtd, (int, float, np.number)): return float(qtd)
    qtd_str = str(qtd).strip()
    if not qtd_str: return 0.0
    # Remove espaços, mas mantém . e ,
    qtd_str = re.sub(r'[\s]', '', qtd_str)
    # Lógica de separador decimal (igual a limpar_valor)
    has_comma = ',' in qtd_str
    has_dot = '.' in qtd_str
    if has_comma and has_dot:
        last_comma_pos = qtd_str.rfind(',')
        last_dot_pos = qtd_str.rfind('.')
        if last_comma_pos > last_dot_pos: qtd_str = qtd_str.replace('.', '').replace(',', '.')
        else: qtd_str = qtd_str.replace(',', '')
    elif has_comma: qtd_str = qtd_str.replace(',', '.')
    qtd_str = re.sub(r'[^\d.]', '', qtd_str) # Remove não numéricos exceto ponto
    try: return float(qtd_str) if qtd_str else 0.0
    except ValueError: return 0.0


def identificar_colunas(df):
    """Identifica heuristicamente as colunas de código, descrição, valor e as novas colunas."""
    coluna_codigo = None
    coluna_descricao = None
    coluna_valor = None
    coluna_unidade = None # Novo
    coluna_quantidade = None # Novo
    coluna_custo_unitario = None # Novo

    # Prioridade para nomes exatos comuns (adicionando novos)
    exact_matches = {
        'codigo': ['código', 'codigo', 'cod.', 'item', 'ref', 'referencia', 'referência'],
        'descricao': ['descrição', 'descricao', 'desc', 'especificação', 'especificacao', 'serviço', 'servico'],
        'valor': ['valor total', 'custo total', 'preço total', 'total', 'valor', 'custo', 'preço'],
        'unidade': ['unid', 'unidade', 'und', 'um'],
        'quantidade': ['quantidade', 'quant', 'qtd', 'qtde'],
        'custo_unitario': ['custo unitário', 'custo unitario', 'preço unitário', 'preco unitario', 'valor unitário', 'valor unitario', 'unitário', 'unitario']
    }

    cols_lower = {str(col).lower().strip(): col for col in df.columns}
    identified_cols = {} # Guarda as colunas já identificadas para não reatribuir

    # 1. Busca por nomes exatos
    for target, patterns in exact_matches.items():
        current_col_var = locals().get(f'coluna_{target}') # Pega a variável correspondente (coluna_codigo, etc.)
        if current_col_var is not None: continue # Pula se já identificou

        for pattern in patterns:
            if pattern in cols_lower:
                col_original = cols_lower[pattern]
                if col_original not in identified_cols.values(): # Verifica se essa coluna original já foi usada
                    # Verificação adicional para valor/custo_unitario/quantidade (devem parecer numéricos)
                    is_numeric_like = False
                    if target in ['valor', 'custo_unitario', 'quantidade']:
                         try: is_numeric_like = pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].dropna().astype(str).str.contains(r'[\d,.]').any()
                         except Exception: pass # Ignora erro na verificação
                    else: # Para código, descrição, unidade, não precisa ser numérico
                         is_numeric_like = True

                    if is_numeric_like:
                        identified_cols[target] = col_original
                        # Atualiza a variável local dinamicamente (cuidado com escopo, mas ok aqui)
                        globals()[f'coluna_{target}'] = col_original
                        break # Vai para o próximo target

    # 2. Busca por padrões parciais se não encontrou por nome exato (simplificado)
    # (Opcional: pode adicionar lógica parcial aqui se necessário, mas nomes exatos costumam bastar)

    # 3. Heurísticas baseadas em conteúdo (mantidas para descrição e valor como fallback)
    available_cols = [c for c in df.columns if c not in identified_cols.values()]

    if 'descricao' not in identified_cols and available_cols:
         try:
              mean_lengths = {col: df[col].astype(str).str.len().mean() for col in available_cols}
              if mean_lengths: identified_cols['descricao'] = max(mean_lengths, key=mean_lengths.get)
         except Exception: pass

    if 'valor' not in identified_cols and available_cols:
         max_sum = -1
         best_val_col = None
         potential_val_cols = [c for c in available_cols if c != identified_cols.get('descricao')]
         for col in potential_val_cols:
              try:
                   numeric_vals = df[col].apply(limpar_valor)
                   current_sum = numeric_vals.sum()
                   if current_sum > max_sum and numeric_vals.count() > len(df)*0.5:
                        max_sum = current_sum
                        best_val_col = col
              except Exception: continue
         if best_val_col: identified_cols['valor'] = best_val_col

    # Retorna as colunas encontradas (pode ser None se não achou)
    return (identified_cols.get('codigo'), identified_cols.get('descricao'), identified_cols.get('valor'),
            identified_cols.get('unidade'), identified_cols.get('quantidade'), identified_cols.get('custo_unitario'))


def gerar_curva_abc(df, coluna_codigo, coluna_descricao, coluna_valor,
                    coluna_unidade=None, coluna_quantidade=None, coluna_custo_unitario=None, # Novas colunas opcionais
                    limite_a=80, limite_b=95):
    """Gera a curva ABC a partir do DataFrame processado, incluindo novas colunas."""

    # Validação das colunas essenciais
    if not all([coluna_codigo, coluna_descricao, coluna_valor]):
        st.error("Erro interno: Colunas essenciais (Código, Descrição, Valor) não foram fornecidas.")
        return None, 0
    essential_cols = [coluna_codigo, coluna_descricao, coluna_valor]
    if not all(col in df.columns for col in essential_cols):
         missing = [col for col in essential_cols if col not in df.columns]
         st.error(f"Erro interno: Colunas essenciais não encontradas no DataFrame: {missing}")
         return None, 0

    # Lista de colunas a serem usadas (inclui opcionais se existirem no DF)
    cols_to_use = essential_cols
    optional_cols_map = { # Mapeia nome do parâmetro para nome da coluna no DF
        'unidade': coluna_unidade,
        'quantidade': coluna_quantidade,
        'custo_unitario': coluna_custo_unitario
    }
    valid_optional_cols = {}
    for key, col_name in optional_cols_map.items():
        if col_name and col_name in df.columns:
            cols_to_use.append(col_name)
            valid_optional_cols[key] = col_name # Guarda as opcionais válidas

    try:
        # Seleciona e copia as colunas relevantes
        df_work = df[list(set(cols_to_use))].copy() # Usa set para evitar duplicatas

        # --- Limpeza e Conversão ---
        df_work['valor_numerico'] = df_work[coluna_valor].apply(limpar_valor)
        df_work['codigo_str'] = df_work[coluna_codigo].astype(str).str.strip()
        df_work['descricao_str'] = df_work[coluna_descricao].astype(str).str.strip()
        # Limpa opcionais se existirem
        if 'unidade' in valid_optional_cols:
            df_work['unidade_str'] = df_work[valid_optional_cols['unidade']].astype(str).str.strip()
        if 'quantidade' in valid_optional_cols:
            df_work['quantidade_num'] = df_work[valid_optional_cols['quantidade']].apply(limpar_quantidade)
        if 'custo_unitario' in valid_optional_cols:
            df_work['custo_unitario_num'] = df_work[valid_optional_cols['custo_unitario']].apply(limpar_valor)

        # Filtra itens com valor zero ou negativo e código inválido
        df_work = df_work[(df_work['valor_numerico'] > 0) & (df_work['codigo_str'] != '')]
        if df_work.empty:
            st.error("Nenhum item com valor positivo e código válido encontrado após limpeza.")
            return None, 0

        # --- Agrupamento ---
        agg_dict = {
            'descricao': ('descricao_str', 'first'),
            'valor': ('valor_numerico', 'sum') # Valor total é sempre somado
        }
        # Adiciona agregação para colunas opcionais (usando 'first')
        if 'unidade' in valid_optional_cols: agg_dict['unidade'] = ('unidade_str', 'first')
        # Para quantidade e custo unitário, 'first' é apropriado para planilha sintética
        if 'quantidade' in valid_optional_cols: agg_dict['quantidade'] = ('quantidade_num', 'first')
        if 'custo_unitario' in valid_optional_cols: agg_dict['custo_unitario'] = ('custo_unitario_num', 'first')

        df_agrupado = df_work.groupby('codigo_str').agg(**agg_dict).reset_index()
        df_agrupado = df_agrupado.rename(columns={'codigo_str': 'codigo'})

        # --- Cálculo da Curva ABC ---
        valor_total_geral = df_agrupado['valor'].sum()
        if valor_total_geral == 0:
            st.error("Valor total dos itens é zero.")
            return None, 0

        df_curva = df_agrupado.sort_values('valor', ascending=False).reset_index(drop=True)

        # Adicionar colunas de percentual e classificação
        df_curva['percentual'] = (df_curva['valor'] / valor_total_geral * 100)
        df_curva['percentual_acumulado'] = df_curva['percentual'].cumsum()
        df_curva['custo_total_acumulado'] = df_curva['valor'].cumsum() # Custo acumulado

        def classificar(perc_acum):
            # Adiciona pequena tolerância para evitar problemas de ponto flutuante no limite
            if perc_acum <= limite_a + 1e-9: return 'A'
            elif perc_acum <= limite_b + 1e-9: return 'B'
            else: return 'C'
        df_curva['classificacao'] = df_curva['percentual_acumulado'].apply(classificar)

        # Adicionar posição
        df_curva.insert(0, 'posicao', range(1, len(df_curva) + 1))

        # --- Montar DataFrame Final ---
        # Define a ordem desejada das colunas
        col_order = ['posicao', 'codigo', 'descricao']
        if 'unidade' in valid_optional_cols: col_order.append('unidade')
        if 'quantidade' in valid_optional_cols: col_order.append('quantidade')
        if 'custo_unitario' in valid_optional_cols: col_order.append('custo_unitario')
        col_order.extend(['valor', 'custo_total_acumulado', 'percentual_acumulado', 'classificacao'])
        # Adiciona 'percentual' se quiser
        # col_order.append('percentual')

        # Seleciona e reordena, preenchendo com NaN se alguma opcional não existia
        df_final = df_curva.reindex(columns=col_order)

        return df_final, valor_total_geral

    except Exception as e:
        st.error(f"Erro ao gerar a curva ABC: {str(e)}")
        with st.expander("Detalhes técnicos do erro"): st.text(traceback.format_exc())
        return None, 0

# --- Funções de Visualização e Download ---

def criar_graficos_plotly(df_curva, valor_total, limite_a, limite_b): # Passa limites
    """Cria gráficos interativos usando Plotly."""
    # (Função mantida como na versão anterior, pois os gráficos principais não mudam)
    # (Poderia adicionar gráficos baseados nas novas colunas se desejado)
    if df_curva is None or df_curva.empty: return None
    try:
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=("Diagrama de Pareto (Itens x Valor Acumulado)", "Distribuição por Valor (%)",
                            "Distribuição por Quantidade de Itens (%)", "Top 10 Itens por Valor"),
            specs=[[{"secondary_y": True}, {"type": "pie"}], [{"type": "pie"}, {"type": "bar"}]],
            vertical_spacing=0.15, horizontal_spacing=0.1
        )
        colors = {'A': '#2ca02c', 'B': '#ff7f0e', 'C': '#d62728'}
        # Pareto
        fig.add_trace(go.Bar(x=df_curva['posicao'], y=df_curva['valor'], name='Valor Item',
                           marker_color=df_curva['classificacao'].map(colors), text=df_curva['codigo'],
                           hoverinfo='x+y+text+name'), secondary_y=False, row=1, col=1)
        fig.add_trace(go.Scatter(x=df_curva['posicao'], y=df_curva['percentual_acumulado'], name='Perc. Acumulado',
                               mode='lines+markers', line=dict(color='#1f77b4', width=2), marker=dict(size=4)),
                      secondary_y=True, row=1, col=1)
        fig.add_hline(y=limite_a, line_dash="dash", line_color="grey", annotation_text=f"Classe A ({limite_a}%)", secondary_y=True, row=1, col=1)
        fig.add_hline(y=limite_b, line_dash="dash", line_color="grey", annotation_text=f"Classe B ({limite_b}%)", secondary_y=True, row=1, col=1)
        # Pizza Valor
        valor_por_classe = df_curva.groupby('classificacao')['valor'].sum().reindex(['A', 'B', 'C']).fillna(0)
        fig.add_trace(go.Pie(labels=valor_por_classe.index, values=valor_por_classe.values, name='Valor',
                           marker_colors=[colors.get(c, '#888') for c in valor_por_classe.index], hole=0.4, pull=[0.05 if c == 'A' else 0 for c in valor_por_classe.index],
                           textinfo='percent+label', hoverinfo='label+percent+value+name'), row=1, col=2)
        # Pizza Quantidade
        qtd_por_classe = df_curva['classificacao'].value_counts().reindex(['A', 'B', 'C']).fillna(0)
        fig.add_trace(go.Pie(labels=qtd_por_classe.index, values=qtd_por_classe.values, name='Quantidade',
                           marker_colors=[colors.get(c, '#888') for c in qtd_por_classe.index], hole=0.4, pull=[0.05 if c == 'A' else 0 for c in qtd_por_classe.index],
                           textinfo='percent+label', hoverinfo='label+percent+value+name'), row=2, col=1)
        # Top 10
        top10 = df_curva.head(10).sort_values('valor', ascending=True)
        fig.add_trace(go.Bar(y=top10['codigo'] + ' (' + top10['descricao'].str[:30] + '...)', x=top10['valor'], name='Top 10 Valor',
                           orientation='h', marker_color=top10['classificacao'].map(colors), text=top10['valor'].map('R$ {:,.2f}'.format),
                           textposition='outside', hoverinfo='y+x+name'), row=2, col=2)
        # Layout
        fig.update_layout(height=850, showlegend=False, title_text="Análise Gráfica da Curva ABC", title_x=0.5, title_font_size=22,
                          margin=dict(l=20, r=20, t=80, b=20), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        fig.update_yaxes(title_text="Valor do Item (R$)", secondary_y=False, row=1, col=1)
        fig.update_yaxes(title_text="Percentual Acumulado (%)", secondary_y=True, row=1, col=1, range=[0, 101])
        fig.update_xaxes(title_text="Posição do Item", row=1, col=1)
        fig.update_xaxes(title_text="Valor (R$)", row=2, col=2)
        fig.update_yaxes(title_text="Item", autorange="reversed", row=2, col=2, tickfont_size=10)
        fig.layout.annotations[0].update(text="<b>Diagrama de Pareto</b>")
        fig.layout.annotations[1].update(text="<b>Distribuição Valor (%)</b>")
        fig.layout.annotations[2].update(text="<b>Distribuição Quantidade (%)</b>")
        fig.layout.annotations[3].update(text="<b>Top 10 Itens (Valor)</b>")
        return fig
    except Exception as e:
        st.error(f"Erro ao criar gráficos: {str(e)}")
        with st.expander("Detalhes técnicos do erro"): st.text(traceback.format_exc())
        return None

def get_download_link(df, filename, text, file_format='csv'):
    """Gera um botão de download para o DataFrame como CSV ou Excel."""
    try:
        # Prepara o DataFrame para download (nomes amigáveis e formatação se Excel)
        df_download = df.copy()
        # Renomeia colunas para o arquivo baixado
        rename_map = {
            'posicao': 'Posição', 'codigo': 'Código', 'descricao': 'Descrição',
            'unidade': 'Unidade', 'quantidade': 'Quantidade', 'custo_unitario': 'Custo Unitário (R$)',
            'valor': 'Custo Total (R$)', 'custo_total_acumulado': 'Custo Total Acumulado (R$)',
            'percentual': 'Percentual (%)', 'percentual_acumulado': 'Percentual Acumulado (%)',
            'classificacao': 'Classificação'
        }
        df_download.rename(columns=rename_map, inplace=True)
        # Seleciona apenas colunas que realmente existem após renomear
        df_download = df_download[[col for col in rename_map.values() if col in df_download.columns]]

        if file_format == 'csv':
            # Usa utf-8-sig para BOM, ajudando Excel a reconhecer UTF-8
            data = df_download.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig')
            mime = 'text/csv'
        elif file_format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_download.to_excel(writer, sheet_name='Curva ABC', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Curva ABC']
                # Formatos (simplificado para evitar erros complexos)
                header_format = workbook.add_format({'bold': True, 'bg_color': '#1e3c72', 'color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                currency_format = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                percent_format = workbook.add_format({'num_format': '0.00"%"', 'border': 1}) # Formato percentual
                number_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1}) # Formato número geral
                center_format = workbook.add_format({'align': 'center', 'border': 1})

                # Aplica formato de cabeçalho
                for col_num, value in enumerate(df_download.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                # Aplica formatos às colunas baseado no nome *final*
                col_map = {name: i for i, name in enumerate(df_download.columns)}
                currency_cols = ['Custo Unitário (R$)', 'Custo Total (R$)', 'Custo Total Acumulado (R$)']
                percent_cols = ['Percentual (%)', 'Percentual Acumulado (%)']
                number_cols = ['Quantidade']
                center_cols = ['Posição', 'Classificação']

                for col_name in df_download.columns:
                     col_idx = col_map[col_name]
                     if col_name in currency_cols: fmt = currency_format
                     elif col_name in percent_cols: fmt = percent_format
                     elif col_name in number_cols: fmt = number_format
                     elif col_name in center_cols: fmt = center_format
                     else: fmt = None # Formato padrão

                     # Ajusta largura da coluna
                     try: width = max(df_download[col_name].astype(str).map(len).max(), len(col_name)) + 2
                     except: width = len(col_name) + 5 # Fallback
                     worksheet.set_column(col_idx, col_idx, min(width, 50), fmt) # Aplica formato e largura

            output.seek(0)
            data = output.read()
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            st.error("Formato de download inválido.")
            return

        # Cria o botão de download
        st.download_button(
            label=text,
            data=data,
            file_name=filename,
            mime=mime,
            key=f"download_{file_format}" # Chave única para o botão
        )

    except Exception as e:
        st.error(f"Erro ao gerar link de download ({file_format}): {e}")
        # Não retorna HTML aqui, apenas loga o erro

# --- Interface Streamlit ---

# Sidebar (mantida)
with st.sidebar:
    # st.image("logo.png", use_column_width=True) # Comente ou substitua pelo seu logo
    st.header("⚙️ Configurações")
    st.subheader("Parâmetros da Curva ABC")
    if 'limite_a' not in st.session_state: st.session_state.limite_a = 80
    if 'limite_b' not in st.session_state: st.session_state.limite_b = 95
    limite_a = st.slider("Limite Classe A (%)", 50, 95, st.session_state.limite_a, 1, key='limite_a_slider')
    limite_b_min = limite_a + 1
    limite_b = st.slider("Limite Classe B (%)", limite_b_min, 99, max(st.session_state.limite_b, limite_b_min), 1, key='limite_b_slider')
    st.session_state.limite_a = limite_a
    st.session_state.limite_b = limite_b
    st.markdown("---")
    st.subheader("ℹ️ Sobre")
    st.info("""
    Gera Curvas ABC de planilhas sintéticas.
    **Funcionalidades:** Upload CSV/Excel, Detecção auto, Limpeza, Visualização, Exportação.
    """)
    st.markdown("---")
    st.caption(f"Versão 1.2 - {datetime.now().year}")

# Conteúdo Principal
st.markdown('<div class="highlight">', unsafe_allow_html=True)
st.markdown("""
#### Como usar:
1.  **Upload** da planilha (CSV/Excel).
2.  **Confirme/Selecione** as colunas (Código, Descrição, Valor Total são obrigatórias).
3.  Ajuste os **limites A/B** (opcional).
4.  Clique em **"Gerar Curva ABC"**.
5.  **Analise** e **Baixe** os resultados.
""")
st.markdown('</div>', unsafe_allow_html=True)

# Upload
uploaded_file = st.file_uploader("📂 Selecione a planilha sintética (CSV, XLSX, XLS)", type=["csv", "xlsx", "xls"], key="file_uploader")

# Inicialização do Estado da Sessão
default_state = {
    'df_processed': None, 'col_codigo': None, 'col_descricao': None, 'col_valor': None,
    'col_unidade': None, 'col_quantidade': None, 'col_custo_unitario': None, # Novas colunas
    'curva_gerada': False, 'curva_abc': None, 'valor_total': 0,
    'last_uploaded_filename': None
}
for key, value in default_state.items():
    if key not in st.session_state:
        st.session_state[key] = value

# Processamento do Arquivo
if uploaded_file is not None:
    if st.session_state.last_uploaded_filename != uploaded_file.name:
        st.session_state.last_uploaded_filename = uploaded_file.name
        # Reseta o estado ao carregar novo arquivo
        for key in default_state: st.session_state[key] = default_state[key]

        with st.spinner('Processando arquivo...'):
            df_processed, delimitador = processar_arquivo(uploaded_file)
            st.session_state.df_processed = df_processed

            if df_processed is not None:
                st.success(f"Arquivo '{uploaded_file.name}' carregado!")
                with st.spinner('Identificando colunas...'):
                     # Atualiza o estado com as colunas identificadas
                     cols = identificar_colunas(df_processed)
                     st.session_state.col_codigo, st.session_state.col_descricao, st.session_state.col_valor, \
                     st.session_state.col_unidade, st.session_state.col_quantidade, st.session_state.col_custo_unitario = cols
            else:
                st.error("Falha ao processar o arquivo.")
                st.session_state.last_uploaded_filename = None # Permite tentar carregar de novo

# Exibição e Controles (se arquivo processado)
if st.session_state.df_processed is not None:
    df = st.session_state.df_processed

    with st.expander("🔍 Visualizar amostra dos dados carregados", expanded=False):
        try: st.dataframe(df.head(10))
        except Exception as e: st.warning(f"Erro ao exibir amostra: {e}")

    st.subheader("Confirme as Colunas para a Curva ABC")
    available_columns = [''] + list(df.columns) # Adiciona opção vazia

    def get_safe_index(col_name, options):
        try: return options.index(col_name) if col_name and col_name in options else 0
        except ValueError: return 0

    # Layout em 2 linhas para melhor organização
    row1_cols = st.columns(3)
    row2_cols = st.columns(3)

    with row1_cols[0]:
        sel_codigo = st.selectbox("Código*", available_columns, index=get_safe_index(st.session_state.col_codigo, available_columns), key='sel_cod')
    with row1_cols[1]:
        sel_descricao = st.selectbox("Descrição*", available_columns, index=get_safe_index(st.session_state.col_descricao, available_columns), key='sel_desc')
    with row1_cols[2]:
        sel_valor = st.selectbox("Valor Total*", available_columns, index=get_safe_index(st.session_state.col_valor, available_columns), key='sel_val')

    with row2_cols[0]:
        sel_unidade = st.selectbox("Unidade (Opcional)", available_columns, index=get_safe_index(st.session_state.col_unidade, available_columns), key='sel_un')
    with row2_cols[1]:
        sel_quantidade = st.selectbox("Quantidade (Opcional)", available_columns, index=get_safe_index(st.session_state.col_quantidade, available_columns), key='sel_qtd')
    with row2_cols[2]:
        sel_custo_unitario = st.selectbox("Custo Unitário (Opcional)", available_columns, index=get_safe_index(st.session_state.col_custo_unitario, available_columns), key='sel_cu')

    # Atualiza estado com seleções
    st.session_state.col_codigo = sel_codigo if sel_codigo else None
    st.session_state.col_descricao = sel_descricao if sel_descricao else None
    st.session_state.col_valor = sel_valor if sel_valor else None
    st.session_state.col_unidade = sel_unidade if sel_unidade else None
    st.session_state.col_quantidade = sel_quantidade if sel_quantidade else None
    st.session_state.col_custo_unitario = sel_custo_unitario if sel_custo_unitario else None

    # Validação e Botão Gerar
    cols_ok = st.session_state.col_codigo and st.session_state.col_descricao and st.session_state.col_valor
    if not cols_ok: st.warning("Selecione as colunas obrigatórias (*) para continuar.")

    if st.button("🚀 Gerar Curva ABC", key="gerar_btn", disabled=not cols_ok):
        with st.spinner('Gerando Curva ABC...'):
            resultado, valor_total = gerar_curva_abc(
                df, st.session_state.col_codigo, st.session_state.col_descricao, st.session_state.col_valor,
                st.session_state.col_unidade, st.session_state.col_quantidade, st.session_state.col_custo_unitario, # Passa novas colunas
                st.session_state.limite_a, st.session_state.limite_b
            )
            if resultado is not None:
                st.session_state.curva_abc = resultado
                st.session_state.valor_total = valor_total
                st.session_state.curva_gerada = True
            else:
                st.error("Falha ao gerar a Curva ABC.")
                st.session_state.curva_gerada = False

# Exibir Resultados
if st.session_state.curva_gerada and st.session_state.curva_abc is not None:
    st.markdown("---")
    st.header("✅ Resultados da Curva ABC")

    resultado_final = st.session_state.curva_abc # DataFrame com dados numéricos
    valor_total_final = st.session_state.valor_total

    # --- Resumo Estatístico ---
    st.subheader("📊 Resumo Estatístico")
    stats_cols = st.columns(4)
    classes_count = resultado_final['classificacao'].value_counts().to_dict()
    valor_por_classe = resultado_final.groupby('classificacao')['valor'].sum().to_dict()
    with stats_cols[0]: st.metric("Total Itens Agrupados", f"{len(resultado_final):,}")
    with stats_cols[1]: st.metric("Valor Total (R$)", f"{valor_total_final:,.2f}")
    count_a = classes_count.get('A', 0); perc_count_a = (count_a / len(resultado_final) * 100) if len(resultado_final) > 0 else 0
    with stats_cols[2]: st.metric("Itens Classe A", f"{count_a} ({perc_count_a:.1f}%)")
    value_a = valor_por_classe.get('A', 0); perc_value_a = (value_a / valor_total_final * 100) if valor_total_final > 0 else 0
    with stats_cols[3]: st.metric("Valor Classe A", f"{perc_value_a:.1f}%")

    # --- Gráficos ---
    st.subheader("📈 Análise Gráfica")
    with st.spinner("Gerando gráficos..."):
        fig = criar_graficos_plotly(resultado_final, valor_total_final, st.session_state.limite_a, st.session_state.limite_b)
        if fig: st.plotly_chart(fig, use_container_width=True)
        else: st.warning("Não foi possível gerar os gráficos.")

    # --- Tabela Detalhada ---
    st.subheader("📋 Tabela de Resultados Detalhada")
    filter_cols = st.columns([0.3, 0.7])
    with filter_cols[0]:
        classe_filtro = st.multiselect("Filtrar Classe", options=sorted(resultado_final['classificacao'].unique()), default=sorted(resultado_final['classificacao'].unique()), key='filter_classe')
    with filter_cols[1]:
        busca = st.text_input("Buscar Código/Descrição", key='search_term')

    # Filtra o DataFrame original (numérico)
    df_filtrado_orig = resultado_final[resultado_final['classificacao'].isin(classe_filtro)]
    if busca:
        df_filtrado_orig = df_filtrado_orig[
            df_filtrado_orig['codigo'].astype(str).str.contains(busca, case=False, na=False) |
            df_filtrado_orig['descricao'].astype(str).str.contains(busca, case=False, na=False)
        ]

    # Cria cópia para exibição e aplica formatação
    df_exibicao = df_filtrado_orig.copy()
    # Renomeia colunas para exibição
    rename_map_display = {
        'posicao': 'Pos', 'codigo': 'Código', 'descricao': 'Descrição', 'unidade': 'Und',
        'quantidade': 'Qtd', 'custo_unitario': 'Custo Unit.', 'valor': 'Custo Total',
        'custo_total_acumulado': 'Custo Acum.', 'percentual_acumulado': '% Acum.',
        'classificacao': 'Classe'
        # 'percentual': '%' # Se quiser mostrar
    }
    df_exibicao.rename(columns=rename_map_display, inplace=True)
    # Formata colunas numéricas como string para exibição
    format_map = {
        'Custo Unit.': 'R$ {:,.2f}', 'Custo Total': 'R$ {:,.2f}', 'Custo Acum.': 'R$ {:,.2f}',
        '% Acum.': '{:.2f}%', #'%' : '{:.2f}%',
        'Qtd': '{:,.2f}' # Formata quantidade com 2 decimais
    }
    for col, fmt in format_map.items():
         if col in df_exibicao.columns:
              # Aplica formatação, tratando possíveis erros
              try: df_exibicao[col] = df_exibicao[col].map(lambda x: fmt.format(x) if pd.notna(x) else '')
              except (TypeError, ValueError): pass # Ignora erro de formatação

    # Seleciona apenas colunas que existem após renomear
    cols_to_display = [col for col in rename_map_display.values() if col in df_exibicao.columns]
    st.dataframe(df_exibicao[cols_to_display], height=400, use_container_width=True)

    # --- Downloads ---
    st.subheader("📥 Downloads")
    dl_cols = st.columns(2)
    data_atual = datetime.now().strftime("%Y%m%d_%H%M")
    csv_filename = f"curva_abc_{data_atual}.csv"
    excel_filename = f"curva_abc_{data_atual}.xlsx"

    with dl_cols[0]:
         # Passa o DataFrame original (resultado_final) para a função de download
         get_download_link(resultado_final, csv_filename, "Baixar como CSV", file_format='csv')
    with dl_cols[1]:
         get_download_link(resultado_final, excel_filename, "Baixar como Excel", file_format='excel')

# Rodapé
st.markdown('<div class="footer">', unsafe_allow_html=True)
st.markdown(f"© {datetime.now().year} - Gerador de Curva ABC | Adaptado para Planilhas SINAPI")
st.markdown('</div>', unsafe_allow_html=True)
