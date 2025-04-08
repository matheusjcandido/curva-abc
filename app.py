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
import math # Importa math para truncamento
import xlsxwriter # Importado explicitamente para referência

# Configuração da página
st.set_page_config(
    page_title="Gerador de Curva ABC",
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
    /* Ajustar tamanho da fonte na tabela de resultados */
    .stDataFrame div[data-testid="stTable"] {
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.title("📊 Gerador de Curva ABC")
st.markdown("### Automatize a geração da Curva ABC a partir de planilhas sintéticas")

# --- Funções Auxiliares ---

def detectar_delimitador(sample_content):
    """Detecta o delimitador mais provável em uma amostra de conteúdo CSV."""
    delimiters = [';', ',', '\t', '|']
    counts = {d: sample_content.count(d) for d in delimiters}
    if counts.get(';', 0) > 0 and counts[';'] >= counts.get(',', 0) * 0.8: return ';'
    counts = {k: v for k, v in counts.items() if v > 0}
    if not counts: return ','
    return max(counts, key=counts.get)

def encontrar_linha_cabecalho(df_preview):
    """Encontra a linha que provavelmente contém os cabeçalhos."""
    cabecalhos_possiveis = [
        'CÓDIGO', 'ITEM', 'DESCRIÇÃO', 'CUSTO', 'VALOR', 'TOTAL', 'PREÇO', 'SERVIÇO',
        'UNID', 'UNIDADE', 'UM', 'QUANT', 'QTD', 'QUANTIDADE', 'UNITÁRIO', 'UNITARIO'
    ]
    max_matches = 0
    header_row_index = 0
    for i in range(min(20, len(df_preview))):
        try:
            row_values = df_preview.iloc[i].dropna().astype(str).str.upper().tolist()
            current_matches = sum(any(keyword in cell for keyword in cabecalhos_possiveis) for cell in row_values)
            if 'DESCRIÇÃO' in row_values or 'CODIGO' in row_values or 'CÓDIGO' in row_values: current_matches += 2
            if current_matches > max_matches:
                max_matches = current_matches
                header_row_index = i
        except Exception: continue
    if max_matches < 2 and df_preview.iloc[0].isnull().all():
         if len(df_preview) > 1 and not df_preview.iloc[1].isnull().all(): return 1
         else: return 0
    elif max_matches == 0: return 0
    return header_row_index

def sanitizar_dataframe(df):
    """Sanitiza o DataFrame para garantir compatibilidade com Streamlit/PyArrow."""
    if df is None: return None
    df_clean = df.copy()
    new_columns = []
    seen_columns = {}
    for i, col in enumerate(df_clean.columns):
        col_str = str(col).strip() if pd.notna(col) else f"coluna_{i}"
        if not col_str: col_str = f"coluna_{i}"
        col_base = col_str.rsplit('_', 1)[0] if '_' in col_str and col_str.rsplit('_', 1)[-1].isdigit() else col_str
        count = seen_columns.get(col_base, 0)
        while col_str in seen_columns:
             count += 1
             col_str = f"{col_base}_{count}"
        seen_columns[col_str] = 0 # Marca como visto
        new_columns.append(col_str)
    df_clean.columns = new_columns

    for col in df_clean.columns:
        try:
            col_dtype = df_clean[col].dtype
            if pd.api.types.is_numeric_dtype(col_dtype): continue
            converted_col = pd.to_numeric(df_clean[col], errors='coerce')
            if not converted_col.isnull().all() and converted_col.notnull().sum() / len(df_clean[col]) > 0.5:
                 df_clean[col] = converted_col
                 continue
            if df_clean[col].dtype == 'object':
                 try:
                      converted_dt = pd.to_datetime(df_clean[col], errors='coerce')
                      if not converted_dt.isnull().all() and converted_dt.notnull().sum() / len(df_clean[col]) > 0.5:
                           df_clean[col] = converted_dt
                           continue
                 except Exception: pass
            if df_clean[col].dtype == 'object' or df_clean[col].apply(type).nunique() > 1:
                 df_clean[col] = df_clean[col].astype(str).replace('nan', '', regex=False).replace('NaT', '', regex=False)
            if isinstance(df_clean[col].dtype, pd.StringDtype) or df_clean[col].dtype == 'object':
                 if df_clean[col].apply(lambda x: isinstance(x, str)).any():
                      df_clean[col] = df_clean[col].str.replace('\x00', '', regex=False)
        except Exception:
            try: df_clean[col] = df_clean[col].astype(str)
            except Exception: st.error(f"Falha crítica ao converter coluna '{col}' para string.")

    df_clean = df_clean.dropna(how='all').dropna(axis=1, how='all')
    return df_clean

# --- Função Principal de Processamento ---

def processar_arquivo(uploaded_file):
    """Carrega e processa o arquivo CSV ou Excel, identificando o cabeçalho."""
    df = None; delimitador = None; linha_cabecalho = 0
    encodings_to_try = ['utf-8', 'latin1', 'cp1252']; engine_to_use = 'openpyxl'
    try:
        file_name = uploaded_file.name.lower(); file_content = uploaded_file.getvalue()
        if file_name.endswith(('.xlsx', '.xls')):
            try: df_preview = pd.read_excel(io.BytesIO(file_content), engine='openpyxl', nrows=25, header=None)
            except Exception:
                try: df_preview = pd.read_excel(io.BytesIO(file_content), engine='xlrd', nrows=25, header=None); engine_to_use = 'xlrd'
                except Exception as e: st.error(f"Erro preview Excel: {e}"); return None, None
            linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            df = pd.read_excel(io.BytesIO(file_content), engine=engine_to_use, header=linha_cabecalho)
        elif file_name.endswith('.csv'):
            decoded_content = None; detected_encoding = None
            for enc in encodings_to_try:
                try: decoded_content = file_content.decode(enc); detected_encoding = enc; break
                except UnicodeDecodeError: continue
            if decoded_content is None: st.error("Erro decodificação CSV."); return None, None
            if not decoded_content.strip(): st.error("Arquivo CSV vazio."); return None, None
            delimitador = detectar_delimitador(decoded_content[:5000])
            try: df_preview = pd.read_csv(io.StringIO(decoded_content), delimiter=delimitador, nrows=25, header=None, skipinitialspace=True, low_memory=False); linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            except Exception: linha_cabecalho = 0
            df = pd.read_csv(io.StringIO(decoded_content), delimiter=delimitador, header=linha_cabecalho, encoding=detected_encoding, on_bad_lines='warn', skipinitialspace=True, low_memory=False)
        else: st.error("Formato não suportado."); return None, None
        if df is not None:
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty: st.error("Arquivo vazio após limpeza."); return None, delimitador
            df = sanitizar_dataframe(df)
            if df is None or df.empty: st.error("Falha sanitização."); return None, delimitador
            return df, delimitador
        else: return None, delimitador
    except Exception as e: st.error(f"Erro fatal processar: {e}"); # with st.expander("Detalhes"): st.text(traceback.format_exc())
    return None, None

# --- Funções da Curva ABC (limpeza, identificação, geração) ---

def limpar_valor(valor):
    """Limpa e converte valores monetários para float."""
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float, np.number)): return float(valor)
    s = str(valor).strip(); s = re.sub(r'[R$€£¥\s]', '', s)
    if ',' in s and '.' in s: s = s.replace('.', '', s.count('.') - 1).replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    s = re.sub(r'[^\d.]', '', s)
    try: return float(s) if s else 0.0
    except ValueError: return 0.0

def limpar_quantidade(qtd):
    """Limpa e converte valores de quantidade para float."""
    if pd.isna(qtd): return 0.0
    if isinstance(qtd, (int, float, np.number)): return float(qtd)
    s = str(qtd).strip();
    if not s: return 0.0
    s = re.sub(r'[\s]', '', s)
    if ',' in s and '.' in s: s = s.replace('.', '', s.count('.') - 1).replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    s = re.sub(r'[^\d.]', '', s)
    try: return float(s) if s else 0.0
    except ValueError: return 0.0

# Função para truncar um número em n casas decimais
def truncate(number, decimals=0):
    """Trunca um número para um número específico de casas decimais."""
    if not isinstance(number, (int, float, np.number)): return number
    if math.isnan(number) or math.isinf(number): return number
    factor = 10.0 ** decimals
    return math.trunc(number * factor) / factor

def identificar_colunas(df):
    """Identifica heuristicamente as colunas necessárias."""
    identified_cols = {}; cols_lower_map = {str(col).lower().strip(): col for col in df.columns}; df_cols_list = list(df.columns)
    exact_matches = {'codigo': ['código do serviço', 'código', 'codigo', 'cod.', 'item', 'ref'],'descricao': ['descrição do serviço', 'descrição', 'descricao', 'desc', 'especificação', 'serviço'],'valor': ['custo total', 'valor total', 'preço total', 'total', 'valor'],'unidade': ['unidade de medida', 'unid', 'unidade', 'und', 'um'],'quantidade': ['quantidade', 'quantidade total', 'quant', 'qtd', 'qtde'],'custo_unitario': ['custo unitário', 'preço unitário', 'valor unitário', 'custo unitario', 'preco unitario', 'valor unitario', 'unitário']}
    for target, patterns in exact_matches.items():
        if target in identified_cols: continue
        for pattern in patterns:
            if pattern in cols_lower_map:
                col_original = cols_lower_map[pattern]
                if col_original not in identified_cols.values():
                    is_num_ok = True
                    if target in ['valor', 'custo_unitario', 'quantidade']:
                         is_num_ok = False
                         try: is_num_ok = pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].dropna().astype(str).str.contains(r'[\d,.]').any()
                         except Exception: pass
                    if is_num_ok: identified_cols[target] = col_original; break
    if 'quantidade' not in identified_cols and 'unidade' in identified_cols:
        col_unidade_nome = identified_cols['unidade']
        try:
            unidade_idx = df_cols_list.index(col_unidade_nome)
            if unidade_idx + 1 < len(df_cols_list):
                col_seguinte = df_cols_list[unidade_idx + 1]
                if col_seguinte not in identified_cols.values():
                    is_num_ok = False
                    try: is_num_ok = pd.api.types.is_numeric_dtype(df[col_seguinte]) or df[col_seguinte].dropna().astype(str).str.contains(r'[\d,.]').any()
                    except Exception: pass
                    if is_num_ok: identified_cols['quantidade'] = col_seguinte
        except ValueError: pass
    available = [c for c in df.columns if c not in identified_cols.values()]
    if 'descricao' not in identified_cols and available:
         try: identified_cols['descricao'] = max(available, key=lambda c: df[c].astype(str).str.len().mean())
         except Exception: pass
    available = [c for c in df.columns if c not in identified_cols.values()]
    if 'valor' not in identified_cols and available:
        col_sums = {}
        for col in available:
            try:
                vals = df[col].apply(limpar_valor); s = vals.sum(); c = vals.count(); l = len(df)
                if c > l * 0.1: col_sums[col] = s
            except Exception: continue
        if col_sums:
            best_val_col = max(col_sums, key=col_sums.get)
            if col_sums[best_val_col] > 0: identified_cols['valor'] = best_val_col
    return (identified_cols.get('codigo'), identified_cols.get('descricao'), identified_cols.get('valor'), identified_cols.get('unidade'), identified_cols.get('quantidade'), identified_cols.get('custo_unitario'))


def gerar_curva_abc(df, col_cod, col_desc, col_val, col_un=None, col_qtd=None, col_cu=None, lim_a=80, lim_b=95):
    """Gera a curva ABC com cálculo de custo total ajustado para porcentagem e valor do CU."""
    essential = {'Código': col_cod, 'Descrição': col_desc, 'Valor': col_val}
    if not all(essential.values()) or not all(c in df.columns for c in essential.values()): st.error("Colunas essenciais inválidas."); return None, 0, None
    optional = {'unidade': col_un, 'quantidade': col_qtd, 'custo_unitario': col_cu}
    cols_to_use = list(essential.values()) + [c for c in optional.values() if c and c in df.columns]
    valid_optional = {k: v for k, v in optional.items() if v and v in df.columns}
    can_calculate_total = 'quantidade' in valid_optional and 'custo_unitario' in valid_optional
    unit_col_present = 'unidade' in valid_optional

    try:
        df_work = df[list(set(cols_to_use))].copy()
        df_work['valor_original_num'] = df_work[col_val].apply(limpar_valor)
        df_work['codigo_str'] = df_work[col_cod].astype(str).str.strip()
        df_work['descricao_str'] = df_work[col_desc].astype(str).str.strip()
        if unit_col_present: df_work['unidade_str'] = df_work[valid_optional['unidade']].astype(str).str.strip()
        if 'quantidade' in valid_optional: df_work['quantidade_num'] = df_work[valid_optional['quantidade']].apply(limpar_quantidade)
        if 'custo_unitario' in valid_optional: df_work['custo_unitario_num'] = df_work[valid_optional['custo_unitario']].apply(limpar_valor)

        df_work = df_work[(df_work['valor_original_num'] > 0) & (df_work['codigo_str'] != '')]
        if df_work.empty: st.error("Nenhum item válido encontrado (valor > 0)."); return None, 0, None

        agg_config = {'descricao': ('descricao_str', 'first'), 'valor_original_sum': ('valor_original_num', 'sum')}
        if unit_col_present: agg_config['unidade'] = ('unidade_str', 'first')
        if 'quantidade' in valid_optional: agg_config['quantidade'] = ('quantidade_num', 'sum')
        if 'custo_unitario' in valid_optional: agg_config['custo_unitario'] = ('custo_unitario_num', 'first')

        df_agg = df_work.groupby('codigo_str').agg(**agg_config).reset_index().rename(columns={'codigo_str': 'codigo'})

        # Calcula Custo Total (Qtd * CU) se possível, ajustando para unidade '%' e valor do CU
        if can_calculate_total:
            # *** AJUSTE CÁLCULO % BASEADO NO VALOR DO CU ***
            def calculate_row_total(row):
                qtd = row.get('quantidade', 0)
                cu = row.get('custo_unitario', 0)
                unit = str(row.get('unidade', '')).strip() if 'unidade' in df_agg.columns else ''

                if not (isinstance(qtd, (int, float, np.number)) and pd.notna(qtd)): qtd = 0.0
                if not (isinstance(cu, (int, float, np.number)) and pd.notna(cu)): cu = 0.0

                if unit == '%':
                    # Verifica o valor do Custo Unitário
                    if cu < 500:
                        # CU < 500: Assume CU é por ponto percentual. Não divide Qtd.
                        return float(qtd) * float(cu)
                    else:
                        # CU >= 500: Assume CU é valor total para 100%. Divide Qtd por 100.
                        return (float(qtd) / 100.0) * float(cu)
                else:
                    # Unidade não é '%', cálculo padrão
                    return float(qtd) * float(cu)

            df_agg['valor_calc'] = df_agg.apply(calculate_row_total, axis=1)
            df_agg['valor'] = df_agg['valor_calc'].apply(lambda x: truncate(x, 2))
            df_agg = df_agg[df_agg['valor'] > 0]
            if df_agg.empty: st.error("Nenhum item com Custo Total calculado > 0."); return None, 0, None
            df_agg = df_agg.drop(columns=['valor_calc'])
        else:
            st.warning("Usando o 'Valor Total' original para a curva ABC.")
            df_agg['valor'] = df_agg['valor_original_sum']
            df_agg = df_agg[df_agg['valor'] > 0]
            if df_agg.empty: st.error("Nenhum item com Valor Total original > 0."); return None, 0, None
        if 'valor_original_sum' in df_agg.columns: df_agg = df_agg.drop(columns=['valor_original_sum'])

        v_total = df_agg['valor'].sum()
        if v_total <= 0: st.error("Valor total final é zero ou negativo."); return None, 0, None

        df_curve = df_agg.sort_values('valor', ascending=False).reset_index(drop=True)
        df_curve['percentual'] = (df_curve['valor'] / v_total * 100)
        df_curve['percentual_acumulado'] = df_curve['percentual'].cumsum()
        df_curve['custo_total_acumulado'] = df_curve['valor'].cumsum()
        df_curve['classificacao'] = df_curve['percentual_acumulado'].apply(lambda p: 'A' if p <= lim_a + 1e-9 else ('B' if p <= lim_b + 1e-9 else 'C'))
        df_curve['posicao'] = range(1, len(df_curve) + 1)

        final_cols_internal = ['codigo', 'descricao']
        if 'unidade' in df_curve.columns: final_cols_internal.append('unidade')
        if 'quantidade' in df_curve.columns: final_cols_internal.append('quantidade')
        if 'custo_unitario' in df_curve.columns: final_cols_internal.append('custo_unitario')
        final_cols_internal.extend(['valor', 'custo_total_acumulado', 'percentual', 'percentual_acumulado', 'classificacao'])
        df_final = df_curve.reindex(columns=final_cols_internal)
        df_curve_with_pos = df_curve

        return df_final, v_total, df_curve_with_pos
    except KeyError as e: st.error(f"Erro: Coluna essencial não encontrada: {e}."); return None, 0, None
    except Exception as e: st.error(f"Erro inesperado gerar curva: {str(e)}"); return None, 0, None


# --- Funções de Visualização e Download ---

def criar_graficos_plotly(df_curva_with_pos, valor_total, limite_a, limite_b):
    """Cria gráficos interativos usando Plotly."""
    if df_curva_with_pos is None or df_curva_with_pos.empty: return None
    try:
        df_curva = df_curva_with_pos
        fig = make_subplots(rows=2, cols=2, subplot_titles=("Diagrama de Pareto", "Distribuição Valor (%)", "Distribuição Quantidade (%)", "Top 10 Itens (Valor)"), specs=[[{"secondary_y": True}, {"type": "pie"}], [{"type": "pie"}, {"type": "bar"}]], vertical_spacing=0.15, horizontal_spacing=0.1)
        colors = {'A': '#2ca02c', 'B': '#ff7f0e', 'C': '#d62728'}
        fig.add_trace(go.Bar(x=df_curva['posicao'], y=df_curva['valor'], name='Valor', marker_color=df_curva['classificacao'].map(colors), text=df_curva['codigo'], hoverinfo='x+y+text+name'), secondary_y=False, row=1, col=1)
        fig.add_trace(go.Scatter(x=df_curva['posicao'], y=df_curva['percentual_acumulado'], name='% Acum.', mode='lines+markers', line=dict(color='#1f77b4', width=2), marker=dict(size=4)), secondary_y=True, row=1, col=1)
        fig.add_hline(y=limite_a, line_dash="dash", line_color="grey", annotation_text=f"A ({limite_a}%)", secondary_y=True, row=1, col=1); fig.add_hline(y=limite_b, line_dash="dash", line_color="grey", annotation_text=f"B ({limite_b}%)", secondary_y=True, row=1, col=1)
        valor_classe = df_curva.groupby('classificacao')['valor'].sum().reindex(['A', 'B', 'C']).fillna(0)
        pull_valor = [0.05 if label == 'A' else 0 for label in valor_classe.index]
        fig.add_trace(go.Pie(labels=valor_classe.index, values=valor_classe.values, name='Valor', marker_colors=[colors.get(c, '#888') for c in valor_classe.index], hole=0.4, pull=pull_valor, textinfo='percent+label', hoverinfo='label+percent+value+name'), row=1, col=2)
        qtd_classe = df_curva['classificacao'].value_counts().reindex(['A', 'B', 'C']).fillna(0)
        pull_qtd = [0.05 if label == 'A' else 0 for label in qtd_classe.index]
        fig.add_trace(go.Pie(labels=qtd_classe.index, values=qtd_classe.values, name='Qtd', marker_colors=[colors.get(c, '#888') for c in qtd_classe.index], hole=0.4, pull=pull_qtd, textinfo='percent+label', hoverinfo='label+percent+value+name'), row=2, col=1)
        top10 = df_curva.head(10).sort_values('valor', ascending=True)
        fig.add_trace(go.Bar(y=top10['codigo'] + ' (' + top10['descricao'].str[:30] + '...)', x=top10['valor'], name='Top 10', orientation='h', marker_color=top10['classificacao'].map(colors), text=top10['valor'].map('R$ {:,.2f}'.format), textposition='outside', hoverinfo='y+x+name'), row=2, col=2)
        fig.update_layout(height=850, showlegend=False, title_text="Análise Gráfica da Curva ABC", title_x=0.5, title_font_size=22, margin=dict(l=20, r=20, t=80, b=20), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        fig.update_yaxes(title_text="Valor (R$)", secondary_y=False, row=1, col=1); fig.update_yaxes(title_text="% Acumulado", secondary_y=True, row=1, col=1, range=[0, 101])
        fig.update_xaxes(title_text="Posição", row=1, col=1); fig.update_xaxes(title_text="Valor (R$)", row=2, col=2); fig.update_yaxes(title_text="Item", autorange="reversed", row=2, col=2, tickfont_size=10)
        for i, title in enumerate(["<b>Pareto</b>", "<b>Valor (%)</b>", "<b>Quantidade (%)</b>", "<b>Top 10 Itens</b>"]): fig.layout.annotations[i].update(text=title)
        return fig
    except Exception as e: st.error(f"Erro gráficos: {e}");
    return None

def get_download_link(df_orig, filename, text, file_format='csv'):
    """Gera botão de download para CSV ou Excel com colunas e nomes corretos."""
    try:
        df_download = df_orig.copy()
        rename_map_download = {'codigo': 'CÓDIGO DO SERVIÇO', 'descricao': 'DESCRIÇÃO DO SERVIÇO', 'unidade': 'UNIDADE DE MEDIDA', 'quantidade': 'QUANTIDADE TOTAL', 'custo_unitario': 'CUSTO UNITÁRIO', 'valor': 'CUSTO TOTAL', 'custo_total_acumulado': 'CUSTO TOTAL ACUMULADO', 'percentual': '% DO ITEM', 'percentual_acumulado': '% ACUMULADO', 'classificacao': 'FAIXA'}
        df_download.rename(columns=rename_map_download, inplace=True)
        download_col_order = ['CÓDIGO DO SERVIÇO', 'DESCRIÇÃO DO SERVIÇO', 'UNIDADE DE MEDIDA', 'QUANTIDADE TOTAL', 'CUSTO UNITÁRIO', 'CUSTO TOTAL', 'CUSTO TOTAL ACUMULADO', '% DO ITEM', '% ACUMULADO', 'FAIXA']
        df_download = df_download[[col for col in download_col_order if col in df_download.columns]]

        if file_format == 'csv':
            data = df_download.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig')
            mime = 'text/csv'
        elif file_format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_download.to_excel(writer, sheet_name='Curva ABC', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Curva ABC']
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1e3c72', 'color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                # Aplica formato cabeçalho e largura básica
                for col_num, value in enumerate(df_download.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                    try: width = min(max(df_download[value].astype(str).map(len).max(), len(value)) + 2, 60)
                    except: width = len(value) + 5
                    worksheet.set_column(col_num, col_num, width) # Simplificado: Sem formato de célula
            output.seek(0)
            data = output.read()
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            st.error("Formato inválido.")
            return

        st.download_button(label=text, data=data, file_name=filename, mime=mime, key=f"dl_{file_format}")
    except Exception as e:
        print(f"DEBUG: Erro download ({file_format}): {e}") # Opcional: log no console

# --- Interface Streamlit ---

# Sidebar
with st.sidebar:
    st.header("⚙️ Configurações"); st.subheader("Parâmetros ABC")
    if 'limite_a' not in st.session_state: st.session_state.limite_a = 50 # Padrão A=50
    if 'limite_b' not in st.session_state: st.session_state.limite_b = 80 # Padrão B=80
    lim_a = st.slider("Limite A (%)", 50, 95, st.session_state.limite_a, 1, key='lim_a_sld')
    lim_b_min = lim_a + 1; lim_b = st.slider("Limite B (%)", lim_b_min, 99, max(st.session_state.limite_b, lim_b_min), 1, key='lim_b_sld')
    st.session_state.limite_a, st.session_state.limite_b = lim_a, lim_b
    st.markdown("---"); st.subheader("ℹ️ Sobre"); st.info("Gera Curvas ABC. v1.11"); st.markdown("---"); st.caption(f"© {datetime.now().year}")

# Conteúdo Principal
st.markdown('<div class="highlight">', unsafe_allow_html=True); st.markdown("#### Como usar:\n1. Faça o upload da aba de planilha sintética.\n2. **Confirme Colunas**.\n3. Ajuste **Limites**.\n4. Clique **Gerar**.\n5. **Analise/Baixe**."); st.markdown('</div>', unsafe_allow_html=True)

# Upload
uploaded_file = st.file_uploader("📂 Selecione a planilha", type=["csv", "xlsx", "xls"], key="file_uploader")

# Estado da Sessão
default_state = {'df_processed': None,'col_codigo': None,'col_descricao': None,'col_valor': None,'col_unidade': None,'col_quantidade': None,'col_custo_unitario': None,'curva_gerada': False,'curva_abc': None,'valor_total': 0,'last_uploaded_filename': None, 'df_curve_with_pos': None}
for k, v in default_state.items():
    if k not in st.session_state: st.session_state[k] = v

# Processamento Arquivo
if uploaded_file:
    if st.session_state.last_uploaded_filename != uploaded_file.name:
        st.session_state.last_uploaded_filename = uploaded_file.name
        for k in default_state: st.session_state[k] = default_state[k] # Reset
        with st.spinner('Processando...'):
            df_proc, delim = processar_arquivo(uploaded_file)
            st.session_state.df_processed = df_proc
            if df_proc is not None:
                st.success(f"Arquivo '{uploaded_file.name}' carregado!")
                with st.spinner('Identificando colunas...'):
                     cols = identificar_colunas(df_proc)
                     st.session_state.col_codigo, st.session_state.col_descricao, st.session_state.col_valor, \
                     st.session_state.col_unidade, st.session_state.col_quantidade, st.session_state.col_custo_unitario = cols
            else: st.error("Falha processamento."); st.session_state.last_uploaded_filename = None

# Controles e Geração
if st.session_state.df_processed is not None:
    df = st.session_state.df_processed
    with st.expander("🔍 Amostra Dados", expanded=False):
        try: st.dataframe(df.head(10))
        except Exception as e: st.warning(f"Erro amostra: {e}")

    st.subheader("Confirme as Colunas")
    cols = list(df.columns); available_cols = [''] + cols
    def get_idx(col_name): return cols.index(col_name) + 1 if col_name and col_name in cols else 0

    r1c1, r1c2, r1c3 = st.columns(3); r2c1, r2c2, r2c3 = st.columns(3)
    with r1c1: st.selectbox("Código*", available_cols, index=get_idx(st.session_state.col_codigo), key='sel_cod')
    with r1c2: st.selectbox("Descrição*", available_cols, index=get_idx(st.session_state.col_descricao), key='sel_desc')
    with r1c3: st.selectbox("Valor Total*", available_cols, index=get_idx(st.session_state.col_valor), key='sel_val') # Ainda pede o Valor Total original (usado como fallback)
    with r2c1: st.selectbox("Unidade", available_cols, index=get_idx(st.session_state.col_unidade), key='sel_un')
    with r2c2: st.selectbox("Quantidade", available_cols, index=get_idx(st.session_state.col_quantidade), key='sel_qtd') # Usará heurística se identificada
    with r2c3: st.selectbox("Custo Unitário", available_cols, index=get_idx(st.session_state.col_custo_unitario), key='sel_cu')

    cols_ok = st.session_state.sel_cod and st.session_state.sel_desc and st.session_state.sel_val # Validação ainda usa valor original
    if not cols_ok: st.warning("Selecione colunas obrigatórias (*).")
    if cols_ok and not (st.session_state.sel_qtd and st.session_state.sel_cu):
         st.info("Se as colunas 'Quantidade' e 'Custo Unitário' não forem selecionadas, o 'Valor Total' original será usado para a curva ABC.")


    if st.button("🚀 Gerar Curva ABC", key="gen_btn", disabled=not cols_ok):
        with st.spinner('Gerando...'):
            res, v_tot, df_curve_with_pos = gerar_curva_abc(
                df,
                st.session_state.sel_cod, st.session_state.sel_desc, st.session_state.sel_val, # Passa o valor original aqui
                st.session_state.sel_un, st.session_state.sel_qtd, st.session_state.sel_cu, # Passa Qtd e CU para cálculo interno
                st.session_state.limite_a, st.session_state.limite_b
            )
            if res is not None:
                st.session_state.update({
                    'curva_abc': res, 'valor_total': v_tot, 'curva_gerada': True,
                    'df_curve_with_pos': df_curve_with_pos
                })
            else: st.error("Falha ao gerar."); st.session_state.curva_gerada = False

# Exibir Resultados
if st.session_state.curva_gerada and st.session_state.curva_abc is not None:
    st.markdown("---"); st.header("✅ Resultados da Curva ABC")
    resultado_final = st.session_state.curva_abc
    valor_total_final = st.session_state.valor_total
    df_grafico = st.session_state.df_curve_with_pos

    # Resumo
    st.subheader("📊 Resumo"); stats_cols = st.columns(4)
    classes_count = resultado_final['classificacao'].value_counts().to_dict(); val_classe = resultado_final.groupby('classificacao')['valor'].sum().to_dict() # 'valor' aqui é o calculado/truncado
    with stats_cols[0]: st.metric("Itens", f"{len(resultado_final):,}")
    with stats_cols[1]: st.metric("Valor Total", f"R$ {valor_total_final:,.2f}") # v_total agora é baseado no valor calculado/truncado
    count_a = classes_count.get('A', 0); perc_ca = (count_a/len(resultado_final)*100) if len(resultado_final) else 0
    with stats_cols[2]: st.metric("Itens A", f"{count_a} ({perc_ca:.1f}%)")
    val_a = val_classe.get('A', 0); perc_va = (val_a/valor_total_final*100) if valor_total_final else 0
    with stats_cols[3]: st.metric("Valor A", f"{perc_va:.1f}%")

    # Gráficos
    st.subheader("📈 Gráficos");
    fig = criar_graficos_plotly(df_grafico, valor_total_final, st.session_state.limite_a, st.session_state.limite_b)
    if fig: st.plotly_chart(fig, use_container_width=True)
    else: st.warning("Erro gráficos.")

    # Tabela Detalhada
    st.subheader("📋 Tabela Detalhada"); f_c1, f_c2 = st.columns([0.3, 0.7])
    with f_c1: classe_f = st.multiselect("Filtrar Classe", sorted(resultado_final['classificacao'].unique()), default=sorted(resultado_final['classificacao'].unique()), key='f_cls')
    with f_c2: busca = st.text_input("Buscar Código/Descrição", key='f_busca')

    df_filt_orig = resultado_final[resultado_final['classificacao'].isin(classe_f)]
    if busca: df_filt_orig = df_filt_orig[df_filt_orig['codigo'].astype(str).str.contains(busca, na=False, case=False) | df_filt_orig['descricao'].astype(str).str.contains(busca, na=False, case=False)]

    # Preparar para exibição (renomear, formatar, ordenar)
    df_exib = df_filt_orig.copy()
    rename_map_disp = {'codigo': 'CÓDIGO DO SERVIÇO', 'descricao': 'DESCRIÇÃO DO SERVIÇO', 'unidade': 'UNIDADE', 'quantidade': 'QTD TOTAL', 'custo_unitario': 'CUSTO UNIT.', 'valor': 'CUSTO TOTAL', 'custo_total_acumulado': 'CUSTO ACUM.', 'percentual': '% ITEM', 'percentual_acumulado': '% ACUM.', 'classificacao': 'FAIXA'}
    df_exib.rename(columns=rename_map_disp, inplace=True)
    display_col_order = ['CÓDIGO DO SERVIÇO', 'DESCRIÇÃO DO SERVIÇO', 'UNIDADE', 'QTD TOTAL', 'CUSTO UNIT.', 'CUSTO TOTAL', 'CUSTO ACUM.', '% ITEM', '% ACUM.', 'FAIXA']
    df_exib = df_exib[[col for col in display_col_order if col in df_exib.columns]]

    # Formatação para exibição (aplicada ao df_exib)
    format_disp = {'CUSTO UNIT.': 'R$ {:,.2f}', 'CUSTO TOTAL': 'R$ {:,.2f}', 'CUSTO ACUM.': 'R$ {:,.2f}', '% ITEM': '{:.2f}%', '% ACUM.': '{:.2f}%', 'QTD TOTAL': '{:,.2f}'}
    reverse_rename_map = {v: k for k, v in rename_map_disp.items()}
    for col_display, fmt in format_disp.items():
        if col_display in df_exib.columns:
            original_col_name = reverse_rename_map.get(col_display)
            if original_col_name and original_col_name in df_filt_orig.columns:
                 try: df_exib[col_display] = df_filt_orig[original_col_name].apply(lambda x: fmt.format(x) if pd.notna(x) and isinstance(x, (int, float, np.number)) else (x if pd.notna(x) else ''))
                 except (TypeError, ValueError, KeyError):
                      try: df_exib[col_display] = df_exib[col_display].apply(lambda x: fmt.format(float(x)) if pd.notna(x) else '')
                      except: pass

    st.dataframe(df_exib, height=400, use_container_width=True)

    # Downloads
    st.subheader("📥 Downloads"); dl_c1, dl_c2 = st.columns(2)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    with dl_c1: get_download_link(resultado_final, f"curva_abc_{ts}.csv", "Baixar CSV", 'csv')
    with dl_c2: get_download_link(resultado_final, f"curva_abc_{ts}.xlsx", "Baixar Excel", 'excel')

# Rodapé
st.markdown('<div class="footer">', unsafe_allow_html=True); st.markdown(f"© {datetime.now().year} - Gerador Curva ABC"); st.markdown('</div>', unsafe_allow_html=True)
