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
    if counts[';'] > 0 and counts[';'] >= counts[','] * 0.8:
         return ';'
    # Remove delimitadores com contagem zero antes de encontrar o máximo
    counts = {k: v for k, v in counts.items() if v > 0}
    if not counts:
        return ',' # Retorna vírgula como padrão se nada for encontrado
    return max(counts, key=counts.get)

def encontrar_linha_cabecalho(df_preview):
    """Encontra a linha que provavelmente contém os cabeçalhos."""
    cabecalhos_possiveis = ['CÓDIGO', 'ITEM', 'DESCRIÇÃO', 'CUSTO', 'VALOR', 'TOTAL', 'PREÇO', 'SERVIÇO']
    max_matches = 0
    header_row_index = 0

    # Verifica as primeiras 20 linhas (ou menos se o df for menor)
    for i in range(min(20, len(df_preview))):
        try:
            row_values = df_preview.iloc[i].astype(str).str.upper().tolist()
            current_matches = sum(any(keyword in str(cell).upper() for keyword in cabecalhos_possiveis) for cell in row_values if pd.notna(cell))

            # Considera a linha com mais palavras-chave como cabeçalho
            if current_matches > max_matches:
                max_matches = current_matches
                header_row_index = i
        except Exception:
            # Ignora linhas que causem erro na verificação
            continue

    # Se nenhuma correspondência significativa for encontrada, assume 0
    # mas só se a linha 0 tiver algum conteúdo útil
    if max_matches < 2 and df_preview.iloc[0].isnull().all(): # Se poucas correspondências e linha 0 vazia
         if len(df_preview) > 1 and not df_preview.iloc[1].isnull().all():
              return 1 # Tenta a linha 1 se a 0 for vazia
         else:
              return 0 # Mantém 0 como último recurso
    elif max_matches == 0:
         return 0 # Assume 0 se nenhuma correspondência

    return header_row_index

def sanitizar_dataframe(df):
    """Sanitiza o DataFrame para garantir compatibilidade com Streamlit/PyArrow."""
    if df is None:
        return None

    df_clean = df.copy()

    # Garante que nomes das colunas sejam strings únicas e limpas
    new_columns = []
    seen_columns = {}
    for i, col in enumerate(df_clean.columns):
        col_str = str(col).strip() if pd.notna(col) else f"coluna_{i}" # Nome padrão se for nulo
        if col_str in seen_columns:
            seen_columns[col_str] += 1
            new_columns.append(f"{col_str}_{seen_columns[col_str]}")
        else:
            seen_columns[col_str] = 0
            new_columns.append(col_str)
    df_clean.columns = new_columns

    # Converte colunas com tipos mistos ou 'object' para string, tratando erros
    for col in df_clean.columns:
        try:
            # Tenta detectar tipos mistos (exceto None/NaN)
            non_na_types = df_clean[col].dropna().apply(type).unique()
            if len(non_na_types) > 1:
                 # Se misto, tenta converter para numérico se possível, senão string
                 df_clean[col] = pd.to_numeric(df_clean[col], errors='ignore')
                 if df_clean[col].dtype == 'object': # Se ainda for object após to_numeric
                      df_clean[col] = df_clean[col].astype(str)

            elif df_clean[col].dtype == 'object':
                 # Tenta converter para numérico ou datetime se parecer apropriado
                 df_clean[col] = pd.to_numeric(df_clean[col], errors='ignore')
                 # Se ainda for object, tenta datetime
                 if df_clean[col].dtype == 'object':
                      try:
                           # Tenta converter para datetime, mas volta se falhar muito
                           original_col = df_clean[col].copy()
                           df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
                           # Se mais de 50% falhou, reverte para string
                           if df_clean[col].isna().sum() / len(df_clean[col]) > 0.5:
                                df_clean[col] = original_col.astype(str)
                      except Exception:
                           df_clean[col] = df_clean[col].astype(str) # Último recurso: string
                 # Se ainda for object após tentativas, converte para string
                 if df_clean[col].dtype == 'object':
                      df_clean[col] = df_clean[col].astype(str)

            # Remove caracteres nulos de colunas string
            if df_clean[col].dtype == 'string' or df_clean[col].dtype == 'object':
                 # Verifica se a coluna realmente contém strings antes de usar .str
                 if df_clean[col].apply(lambda x: isinstance(x, str)).any():
                      df_clean[col] = df_clean[col].astype(str).str.replace('\0', '', regex=False)

        except Exception as e:
            # Em caso de erro inesperado na sanitização da coluna, converte para string
            # st.warning(f"Erro ao sanitizar coluna '{col}': {e}. Convertendo para string.")
            try:
                df_clean[col] = df_clean[col].astype(str)
            except Exception:
                 st.error(f"Falha crítica ao converter coluna '{col}' para string.")
                 # Poderia remover a coluna ou parar, mas vamos deixar seguir por enquanto

    # Remove linhas que são completamente nulas
    df_clean = df_clean.dropna(how='all')
    return df_clean

# --- Função Principal de Processamento ---

def processar_arquivo(uploaded_file):
    """Carrega e processa o arquivo CSV ou Excel, identificando o cabeçalho."""
    df = None
    delimitador = None
    linha_cabecalho = 0 # Default
    encodings_to_try = ['utf-8', 'latin1', 'cp1252'] # ISO-8859-1 é similar a latin1

    try:
        file_name = uploaded_file.name.lower()
        file_content = uploaded_file.getvalue() # Ler bytes uma vez

        # --- Processamento Excel ---
        if file_name.endswith(('.xlsx', '.xls')):
            # 1. Ler preview para encontrar cabeçalho
            try:
                df_preview = pd.read_excel(io.BytesIO(file_content), engine='openpyxl', nrows=25, header=None)
                linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            except Exception as e:
                st.warning(f"Não foi possível ler o preview do Excel para achar cabeçalho: {e}. Assumindo linha 0.")
                linha_cabecalho = 0

            # 2. Ler arquivo completo com o cabeçalho correto
            df = pd.read_excel(io.BytesIO(file_content), engine='openpyxl', header=linha_cabecalho)

        # --- Processamento CSV ---
        elif file_name.endswith('.csv'):
            detected_encoding = None
            decoded_content = None

            # 1. Tentar decodificar com diferentes encodings
            for enc in encodings_to_try:
                try:
                    decoded_content = file_content.decode(enc)
                    detected_encoding = enc
                    # st.info(f"Arquivo CSV decodificado com {enc}")
                    break # Sai do loop se decodificar com sucesso
                except UnicodeDecodeError:
                    continue # Tenta o próximo encoding
            if decoded_content is None:
                st.error("Erro de decodificação: Não foi possível ler o arquivo CSV com os encodings testados (UTF-8, Latin-1, CP1252). Verifique o encoding do arquivo.")
                return None, None

            # Verificar se o conteúdo decodificado está vazio
            if not decoded_content.strip():
                 st.error("O arquivo CSV enviado está vazio.")
                 return None, None

            # 2. Detectar delimitador a partir de uma amostra
            sample_for_delimiter = decoded_content[:5000] # Usa os primeiros 5000 chars
            delimitador = detectar_delimitador(sample_for_delimiter)
            # st.info(f"Delimitador detectado: '{delimitador}'")

            # 3. Ler preview para encontrar cabeçalho
            try:
                df_preview = pd.read_csv(io.StringIO(decoded_content), delimiter=delimitador, nrows=25, header=None, skipinitialspace=True)
                linha_cabecalho = encontrar_linha_cabecalho(df_preview)
                # st.info(f"Linha de cabeçalho detectada: {linha_cabecalho}")
            except Exception as e:
                st.warning(f"Não foi possível ler o preview do CSV para achar cabeçalho: {e}. Assumindo linha 0.")
                linha_cabecalho = 0

            # 4. Ler arquivo completo com cabeçalho e delimitador corretos
            df = pd.read_csv(
                io.StringIO(decoded_content),
                delimiter=delimitador,
                header=linha_cabecalho,
                encoding=detected_encoding,
                on_bad_lines='warn', # Avisa sobre linhas ruins mas tenta continuar
                skipinitialspace=True, # Ignora espaços após o delimitador
                low_memory=False # Ajuda com tipos mistos, mas usa mais memória
            )

        else:
            st.error("Formato de arquivo não suportado. Use CSV, XLSX ou XLS.")
            return None, None

        # --- Pós-processamento Comum ---
        if df is not None:
            # Remover linhas e colunas totalmente vazias
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty:
                 st.error("O arquivo parece vazio após remover linhas/colunas nulas.")
                 return None, delimitador

            # Sanitizar o DataFrame para compatibilidade
            df = sanitizar_dataframe(df)
            if df is None or df.empty:
                 st.error("Falha na sanitização do DataFrame.")
                 return None, delimitador

            return df, delimitador

        else:
             # Caso df não tenha sido atribuído (erro anterior)
             return None, delimitador

    except Exception as e:
        st.error(f"Erro fatal ao processar o arquivo: {str(e)}")
        with st.expander("Detalhes técnicos do erro"):
            st.text(traceback.format_exc())
        return None, None

# --- Funções da Curva ABC (limpeza, identificação, geração) ---

def limpar_valor(valor):
    """Limpa e converte valores monetários para float, tratando diversos formatos."""
    if pd.isna(valor):
        return 0.0

    # Se já for numérico, retorna como float
    if isinstance(valor, (int, float, np.number)):
        return float(valor)

    valor_str = str(valor).strip()
    if not valor_str:
        return 0.0

    # Remove símbolos de moeda comuns e espaços extras
    valor_str = re.sub(r'[R$€£¥\s]', '', valor_str)

    # Verifica se temos '.' e ',' para determinar o formato
    has_comma = ',' in valor_str
    has_dot = '.' in valor_str

    if has_comma and has_dot:
        # Descobre qual é o último separador
        last_comma_pos = valor_str.rfind(',')
        last_dot_pos = valor_str.rfind('.')
        # Se vírgula vem depois do ponto, assume formato BR/EU (1.234,56)
        if last_comma_pos > last_dot_pos:
            valor_str = valor_str.replace('.', '').replace(',', '.')
        # Se ponto vem depois da vírgula, assume formato US (1,234.56)
        else:
            valor_str = valor_str.replace(',', '')
    elif has_comma:
        # Apenas vírgula, assume como separador decimal (1234,56)
        valor_str = valor_str.replace(',', '.')
    # Se tem apenas ponto ou nenhum separador, já deve estar no formato correto (1234.56 ou 1234)

    # Remove qualquer caractere não numérico restante (exceto o ponto decimal)
    valor_str = re.sub(r'[^\d.]', '', valor_str)

    try:
        return float(valor_str) if valor_str else 0.0
    except ValueError:
        # st.warning(f"Não foi possível converter '{valor}' para número.")
        return 0.0

def identificar_colunas(df):
    """Identifica heuristicamente as colunas de código, descrição e valor."""
    coluna_codigo = None
    coluna_descricao = None
    coluna_valor = None

    # Prioridade para nomes exatos comuns
    exact_matches = {
        'codigo': ['código', 'codigo', 'cod.', 'item', 'ref', 'referencia', 'referência'],
        'descricao': ['descrição', 'descricao', 'desc', 'especificação', 'especificacao', 'serviço', 'servico'],
        'valor': ['valor total', 'custo total', 'preço total', 'total', 'valor', 'custo', 'preço']
    }

    cols_lower = {str(col).lower().strip(): col for col in df.columns}

    # 1. Busca por nomes exatos
    for target, patterns in exact_matches.items():
        for pattern in patterns:
            if pattern in cols_lower:
                col_original = cols_lower[pattern]
                if target == 'codigo' and coluna_codigo is None:
                    coluna_codigo = col_original
                elif target == 'descricao' and coluna_descricao is None:
                    coluna_descricao = col_original
                elif target == 'valor' and coluna_valor is None:
                     # Verifica se a coluna de valor parece numérica
                     if pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].astype(str).str.contains(r'[\d,.]').any():
                          coluna_valor = col_original
                # Remove a coluna encontrada para evitar re-identificação
                # del cols_lower[pattern] # Cuidado ao modificar dict durante iteração, melhor só checar se já achou
                break # Vai para o próximo target (codigo, descricao, valor)

    # 2. Busca por padrões parciais se não encontrou por nome exato
    partial_patterns = {
        'codigo': ['cod', 'item', 'ref'],
        'descricao': ['desc', 'especif', 'serv'],
        'valor': ['total', 'valor', 'custo', 'preço']
    }
    remaining_cols = {k: v for k, v in cols_lower.items() if v not in [coluna_codigo, coluna_descricao, coluna_valor]}

    for target, patterns in partial_patterns.items():
        if (target == 'codigo' and coluna_codigo) or \
           (target == 'descricao' and coluna_descricao) or \
           (target == 'valor' and coluna_valor):
            continue # Já encontrou essa coluna

        for col_lower, col_original in remaining_cols.items():
             if any(p in col_lower for p in patterns):
                  if target == 'codigo':
                       coluna_codigo = col_original
                       break
                  elif target == 'descricao':
                       coluna_descricao = col_original
                       break
                  elif target == 'valor':
                       if pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].astype(str).str.contains(r'[\d,.]').any():
                            coluna_valor = col_original
                            break # Para de procurar valor se achar um candidato

    # 3. Heurísticas baseadas em conteúdo (último recurso)
    available_cols = list(df.columns)
    potential_desc_cols = [c for c in available_cols if c not in [coluna_codigo, coluna_valor]]
    potential_val_cols = [c for c in available_cols if c not in [coluna_codigo, coluna_descricao]]

    if not coluna_descricao and potential_desc_cols:
         # Coluna com strings mais longas é provavelmente descrição
         try:
              mean_lengths = {col: df[col].astype(str).str.len().mean() for col in potential_desc_cols}
              if mean_lengths:
                   coluna_descricao = max(mean_lengths, key=mean_lengths.get)
         except Exception: pass # Ignora erro se cálculo de média falhar

    if not coluna_valor and potential_val_cols:
         # Coluna com maior soma de valores numéricos é provavelmente valor
         max_sum = -1
         best_val_col = None
         for col in potential_val_cols:
              try:
                   numeric_vals = df[col].apply(limpar_valor)
                   current_sum = numeric_vals.sum()
                   # Verifica se a coluna tem valores significativos e é predominantemente numérica
                   if current_sum > max_sum and numeric_vals.count() > len(df)*0.5: # Pelo menos 50% de valores numéricos
                        max_sum = current_sum
                        best_val_col = col
              except Exception: continue
         if best_val_col:
              coluna_valor = best_val_col

    # Garante que as colunas sejam diferentes
    if coluna_codigo == coluna_descricao or coluna_codigo == coluna_valor: coluna_codigo = None
    if coluna_descricao == coluna_valor: coluna_descricao = None
    # Tenta re-identificar se houver conflito (simplificado: apenas anula)

    return coluna_codigo, coluna_descricao, coluna_valor

def gerar_curva_abc(df, coluna_codigo, coluna_descricao, coluna_valor, limite_a=80, limite_b=95):
    """Gera a curva ABC a partir do DataFrame processado."""
    if not all([coluna_codigo, coluna_descricao, coluna_valor]):
        st.error("Erro interno: Uma ou mais colunas essenciais não foram identificadas.")
        return None, 0
    if coluna_codigo not in df.columns or coluna_descricao not in df.columns or coluna_valor not in df.columns:
         st.error(f"Erro interno: Colunas selecionadas ({coluna_codigo}, {coluna_descricao}, {coluna_valor}) não encontradas no DataFrame processado.")
         return None, 0

    try:
        # Seleciona e copia as colunas relevantes
        df_work = df[[coluna_codigo, coluna_descricao, coluna_valor]].copy()

        # Limpa e converte a coluna de valor
        df_work['valor_numerico'] = df_work[coluna_valor].apply(limpar_valor)

        # Converte código para string e remove espaços
        df_work['codigo_str'] = df_work[coluna_codigo].astype(str).str.strip()
        df_work['descricao_str'] = df_work[coluna_descricao].astype(str).str.strip()


        # Filtra itens com valor zero ou negativo e código inválido (opcional, mas bom)
        df_work = df_work[(df_work['valor_numerico'] > 0) & (df_work['codigo_str'] != '')]

        if df_work.empty:
            st.error("Nenhum item com valor positivo e código válido encontrado após limpeza.")
            return None, 0

        # Agrupa por código, somando valores e pegando a primeira descrição
        df_agrupado = df_work.groupby('codigo_str').agg(
            descricao=('descricao_str', 'first'),
            valor=('valor_numerico', 'sum')
        ).reset_index() # Converte o índice (codigo_str) de volta para coluna

        # Renomeia a coluna de código agrupado
        df_agrupado = df_agrupado.rename(columns={'codigo_str': 'codigo'})

        # Calcula valor total
        valor_total = df_agrupado['valor'].sum()
        if valor_total == 0:
            st.error("Valor total dos itens é zero. Não é possível gerar a curva.")
            return None, 0

        # Ordenar por valor
        df_curva = df_agrupado.sort_values('valor', ascending=False).reset_index(drop=True)

        # Adicionar colunas de percentual e classificação
        df_curva['percentual'] = (df_curva['valor'] / valor_total * 100)
        df_curva['percentual_acumulado'] = df_curva['percentual'].cumsum()

        def classificar(perc_acum):
            if perc_acum <= limite_a: return 'A'
            elif perc_acum <= limite_b: return 'B'
            else: return 'C'
        df_curva['classificacao'] = df_curva['percentual_acumulado'].apply(classificar)

        # Adicionar posição
        df_curva.insert(0, 'posicao', range(1, len(df_curva) + 1))

        # Selecionar e reordenar colunas finais
        df_final = df_curva[['posicao', 'codigo', 'descricao', 'valor', 'percentual', 'percentual_acumulado', 'classificacao']]

        return df_final, valor_total

    except Exception as e:
        st.error(f"Erro ao gerar a curva ABC: {str(e)}")
        with st.expander("Detalhes técnicos do erro"):
            st.text(traceback.format_exc())
        return None, 0

# --- Funções de Visualização e Download ---

def criar_graficos_plotly(df_curva, valor_total):
    """Cria gráficos interativos usando Plotly."""
    if df_curva is None or df_curva.empty:
        return None

    try:
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=("Diagrama de Pareto (Itens x Valor Acumulado)",
                            "Distribuição por Valor (%)",
                            "Distribuição por Quantidade de Itens (%)",
                            "Top 10 Itens por Valor"),
            specs=[
                [{"secondary_y": True}, {"type": "pie"}], # Pareto com eixo Y secundário
                [{"type": "pie"}, {"type": "bar"}]
            ],
            vertical_spacing=0.15, # Aumenta espaço vertical
            horizontal_spacing=0.1
        )

        # Cores padrão para as classes
        colors = {'A': '#2ca02c', 'B': '#ff7f0e', 'C': '#d62728'} # Verde, Laranja, Vermelho

        # --- Gráfico 1: Pareto ---
        fig.add_trace(
            go.Bar(
                x=df_curva['posicao'],
                y=df_curva['valor'], # Usar valor absoluto na barra
                name='Valor do Item',
                marker_color=df_curva['classificacao'].map(colors),
                text=df_curva['codigo'], # Mostrar código no hover da barra
                hoverinfo='x+y+text+name'
            ),
            secondary_y=False, row=1, col=1
        )
        fig.add_trace(
            go.Scatter(
                x=df_curva['posicao'],
                y=df_curva['percentual_acumulado'],
                name='Percentual Acumulado',
                mode='lines+markers',
                line=dict(color='#1f77b4', width=2), # Azul para linha acumulada
                marker=dict(size=4)
            ),
            secondary_y=True, row=1, col=1
        )
        # Linhas de referência A e B
        fig.add_hline(y=limite_a, line_dash="dash", line_color="grey", annotation_text=f"Classe A ({limite_a}%)", secondary_y=True, row=1, col=1)
        fig.add_hline(y=limite_b, line_dash="dash", line_color="grey", annotation_text=f"Classe B ({limite_b}%)", secondary_y=True, row=1, col=1)

        # --- Gráfico 2: Pizza por Valor ---
        valor_por_classe = df_curva.groupby('classificacao')['valor'].sum().reindex(['A', 'B', 'C']).fillna(0)
        fig.add_trace(
            go.Pie(
                labels=valor_por_classe.index,
                values=valor_por_classe.values,
                name='Valor',
                marker_colors=[colors[c] for c in valor_por_classe.index],
                hole=0.4,
                pull=[0.05 if c == 'A' else 0 for c in valor_por_classe.index], # Destaca A
                textinfo='percent+label',
                hoverinfo='label+percent+value+name'
            ),
            row=1, col=2
        )

        # --- Gráfico 3: Pizza por Quantidade ---
        qtd_por_classe = df_curva['classificacao'].value_counts().reindex(['A', 'B', 'C']).fillna(0)
        fig.add_trace(
            go.Pie(
                labels=qtd_por_classe.index,
                values=qtd_por_classe.values,
                name='Quantidade',
                marker_colors=[colors[c] for c in qtd_por_classe.index],
                hole=0.4,
                pull=[0.05 if c == 'A' else 0 for c in qtd_por_classe.index], # Destaca A
                textinfo='percent+label',
                hoverinfo='label+percent+value+name'
            ),
            row=2, col=1
        )

        # --- Gráfico 4: Top 10 Itens ---
        top10 = df_curva.head(10).sort_values('valor', ascending=True) # Ordena para gráfico hbar
        fig.add_trace(
            go.Bar(
                y=top10['codigo'] + ' (' + top10['descricao'].str[:30] + '...)', # Combina código e descrição no eixo Y
                x=top10['valor'],
                name='Top 10 Valor',
                orientation='h',
                marker_color=top10['classificacao'].map(colors),
                text=top10['valor'].map('R$ {:,.2f}'.format), # Formata valor como texto
                textposition='outside', # Coloca texto fora da barra
                hoverinfo='y+x+name'
            ),
            row=2, col=2
        )

        # --- Layout Geral ---
        fig.update_layout(
            height=850, # Aumenta altura
            showlegend=False,
            title_text="Análise Gráfica da Curva ABC",
            title_x=0.5,
            title_font_size=22,
            margin=dict(l=20, r=20, t=80, b=20), # Ajusta margens
            paper_bgcolor='rgba(0,0,0,0)', # Fundo transparente
            plot_bgcolor='rgba(0,0,0,0)'  # Fundo transparente
        )

        # Layout Eixos Pareto
        fig.update_yaxes(title_text="Valor do Item (R$)", secondary_y=False, row=1, col=1)
        fig.update_yaxes(title_text="Percentual Acumulado (%)", secondary_y=True, row=1, col=1, range=[0, 101])
        fig.update_xaxes(title_text="Posição do Item (Ordenado por Valor)", row=1, col=1)

        # Layout Eixo Top 10
        fig.update_xaxes(title_text="Valor (R$)", row=2, col=2)
        fig.update_yaxes(title_text="Item (Código + Descrição)", autorange="reversed", row=2, col=2, tickfont_size=10)

        # Atualiza títulos dos subplots para clareza
        fig.layout.annotations[0].update(text="<b>Diagrama de Pareto</b> (Valor x Acumulado %)")
        fig.layout.annotations[1].update(text="<b>Distribuição do Valor Total (%)</b> por Classe")
        fig.layout.annotations[2].update(text="<b>Distribuição da Quantidade de Itens (%)</b> por Classe")
        fig.layout.annotations[3].update(text="<b>Top 10 Itens</b> por Valor")

        return fig

    except Exception as e:
        st.error(f"Erro ao criar gráficos: {str(e)}")
        with st.expander("Detalhes técnicos do erro"):
            st.text(traceback.format_exc())
        return None

def get_download_link(df, filename, text, file_format='csv'):
    """Gera um link para download do DataFrame como CSV ou Excel."""
    try:
        if file_format == 'csv':
            data = df.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig') # utf-8-sig para Excel ler BOM
            mime = 'text/csv'
            href_data = f'data:{mime};base64,{base64.b64encode(data.encode("utf-8-sig")).decode()}'
        elif file_format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Curva ABC', index=False)
                # Formatação (opcional, pode ser removida se causar lentidão/erros)
                workbook = writer.book
                worksheet = writer.sheets['Curva ABC']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#1e3c72', 'color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                currency_format = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
                center_format = workbook.add_format({'align': 'center', 'border': 1})

                # Aplicar formato de cabeçalho
                for col_num, value in enumerate(df.columns.values):
                     worksheet.write(0, col_num, value, header_format)

                # Aplicar formatos às colunas (ajuste os nomes das colunas conforme necessário)
                col_map = {name: i for i, name in enumerate(df.columns)}
                if 'Valor (R$)' in col_map: worksheet.set_column(col_map['Valor (R$)'], col_map['Valor (R$)'], 15, currency_format)
                if 'Percentual (%)' in col_map: worksheet.set_column(col_map['Percentual (%)'], col_map['Percentual (%)'], 12, percent_format)
                if 'Percentual Acumulado (%)' in col_map: worksheet.set_column(col_map['Percentual Acumulado (%)'], col_map['Percentual Acumulado (%)'], 15, percent_format)
                if 'Classificação' in col_map: worksheet.set_column(col_map['Classificação'], col_map['Classificação'], 12, center_format)
                if 'Posição' in col_map: worksheet.set_column(col_map['Posição'], col_map['Posição'], 8, center_format)

                # Ajustar largura das outras colunas
                for col_name in ['Código', 'Descrição']:
                     if col_name in col_map:
                          col_idx = col_map[col_name]
                          width = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
                          worksheet.set_column(col_idx, col_idx, min(width, 60)) # Limita largura máxima

            output.seek(0)
            data = output.read()
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            href_data = f'data:{mime};base64,{base64.b64encode(data).decode()}'
        else:
            return "Formato inválido"

        # Usa st.download_button para melhor compatibilidade
        st.download_button(
            label=text,
            data=data,
            file_name=filename,
            mime=mime,
        )
        return "" # Retorna string vazia pois o botão é criado diretamente

    except Exception as e:
        st.error(f"Erro ao gerar link de download ({file_format}): {e}")
        return f"<i>Erro ao gerar link ({file_format})</i>"


# --- Interface Streamlit ---

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/150x50.png?text=Logo+Empresa", use_column_width=True) # Placeholder para logo
    st.header("⚙️ Configurações")
    st.subheader("Parâmetros da Curva ABC")
    # Usar session_state para manter os valores dos sliders entre reruns
    if 'limite_a' not in st.session_state: st.session_state.limite_a = 80
    if 'limite_b' not in st.session_state: st.session_state.limite_b = 95

    limite_a = st.slider("Limite Classe A (%)", 50, 95, st.session_state.limite_a, 1, key='limite_a_slider')
    # Garante que limite B seja maior que A
    limite_b_min = limite_a + 1
    limite_b = st.slider("Limite Classe B (%)", limite_b_min, 99, max(st.session_state.limite_b, limite_b_min), 1, key='limite_b_slider')

    # Atualiza session_state se os sliders mudarem
    st.session_state.limite_a = limite_a
    st.session_state.limite_b = limite_b

    st.markdown("---")
    st.subheader("ℹ️ Sobre")
    st.info("""
    Aplicativo para gerar Curvas ABC a partir de planilhas sintéticas (SINAPI ou outras).
    **Funcionalidades:**
    - Upload de CSV/Excel.
    - Detecção automática de cabeçalho e colunas.
    - Limpeza de dados e agrupamento.
    - Visualizações interativas.
    - Exportação em CSV e Excel.
    """)
    st.markdown("---")
    st.caption(f"Versão 1.1 - {datetime.now().year}")

# Conteúdo Principal
st.markdown('<div class="highlight">', unsafe_allow_html=True)
st.markdown("""
#### Como usar:
1.  **Faça o upload** da planilha (CSV ou Excel).
2.  **Confirme as colunas** de Código, Descrição e Valor Total detectadas (ou selecione manualmente).
3.  Ajuste os **limites das classes A e B** na barra lateral (opcional).
4.  Clique em **"Gerar Curva ABC"**.
5.  **Analise** os gráficos e a tabela.
6.  **Baixe** os resultados se desejar.
""")
st.markdown('</div>', unsafe_allow_html=True)

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Selecione a planilha sintética (CSV, XLSX, XLS)", type=["csv", "xlsx", "xls"], key="file_uploader")

# Inicializa session_state se necessário
if 'df_processed' not in st.session_state: st.session_state.df_processed = None
if 'col_codigo' not in st.session_state: st.session_state.col_codigo = None
if 'col_descricao' not in st.session_state: st.session_state.col_descricao = None
if 'col_valor' not in st.session_state: st.session_state.col_valor = None
if 'curva_gerada' not in st.session_state: st.session_state.curva_gerada = False
if 'curva_abc' not in st.session_state: st.session_state.curva_abc = None
if 'valor_total' not in st.session_state: st.session_state.valor_total = 0

# Processa o arquivo se um novo for carregado
if uploaded_file is not None:
    # Verifica se é um arquivo diferente do anterior para reprocessar
    if 'last_uploaded_filename' not in st.session_state or st.session_state.last_uploaded_filename != uploaded_file.name:
        st.session_state.last_uploaded_filename = uploaded_file.name
        st.session_state.df_processed = None # Reseta dados processados
        st.session_state.curva_gerada = False # Reseta flag de curva gerada
        st.session_state.col_codigo = None
        st.session_state.col_descricao = None
        st.session_state.col_valor = None

        with st.spinner('Processando arquivo...'):
            df_processed, delimitador = processar_arquivo(uploaded_file)
            st.session_state.df_processed = df_processed # Armazena no estado da sessão

            if df_processed is not None:
                st.success(f"Arquivo '{uploaded_file.name}' carregado e processado!")
                # Identifica colunas após processar
                with st.spinner('Identificando colunas...'):
                     st.session_state.col_codigo, st.session_state.col_descricao, st.session_state.col_valor = identificar_colunas(df_processed)
            else:
                st.error("Falha ao processar o arquivo.")
                # Limpa o estado se o processamento falhar
                st.session_state.df_processed = None
                st.session_state.last_uploaded_filename = None


# Exibe controles e botão de gerar se o arquivo foi processado
if st.session_state.df_processed is not None:
    df = st.session_state.df_processed # Pega o df do estado

    with st.expander("🔍 Visualizar amostra dos dados carregados", expanded=False):
        try:
            st.dataframe(df.head(10))
        except Exception as e:
            st.warning(f"Não foi possível exibir a amostra em tabela: {e}. Mostrando como texto.")
            st.text(df.head(10).to_string())

    st.subheader("Confirme as Colunas para a Curva ABC")
    available_columns = [''] + list(df.columns) # Adiciona opção vazia

    # Função auxiliar para obter índice seguro
    def get_safe_index(col_name, options):
        try:
            return options.index(col_name) if col_name and col_name in options else 0
        except ValueError:
            return 0

    col1, col2, col3 = st.columns(3)
    with col1:
        col_codigo_selecionada = st.selectbox(
            "Coluna de Código*", options=available_columns,
            index=get_safe_index(st.session_state.col_codigo, available_columns),
            key='select_codigo'
        )
    with col2:
        col_descricao_selecionada = st.selectbox(
            "Coluna de Descrição*", options=available_columns,
            index=get_safe_index(st.session_state.col_descricao, available_columns),
            key='select_descricao'
        )
    with col3:
        col_valor_selecionada = st.selectbox(
            "Coluna de Valor Total*", options=available_columns,
            index=get_safe_index(st.session_state.col_valor, available_columns),
            key='select_valor'
        )

    # Atualiza o estado da sessão com as colunas selecionadas
    st.session_state.col_codigo = col_codigo_selecionada if col_codigo_selecionada else None
    st.session_state.col_descricao = col_descricao_selecionada if col_descricao_selecionada else None
    st.session_state.col_valor = col_valor_selecionada if col_valor_selecionada else None

    # Verifica se todas as colunas foram selecionadas
    cols_ok = st.session_state.col_codigo and st.session_state.col_descricao and st.session_state.col_valor
    if not cols_ok:
         st.warning("Selecione todas as colunas marcadas com * para continuar.")

    # Botão para gerar a curva ABC (habilitado apenas se colunas OK)
    if st.button("🚀 Gerar Curva ABC", key="gerar_btn", disabled=not cols_ok):
        with st.spinner('Gerando Curva ABC...'):
            resultado, valor_total = gerar_curva_abc(
                df,
                st.session_state.col_codigo,
                st.session_state.col_descricao,
                st.session_state.col_valor,
                st.session_state.limite_a,
                st.session_state.limite_b
            )

            if resultado is not None:
                st.session_state.curva_abc = resultado
                st.session_state.valor_total = valor_total
                st.session_state.curva_gerada = True
                # Não precisa de rerun aqui, a exibição abaixo cuidará disso
            else:
                st.error("Falha ao gerar a Curva ABC. Verifique os dados e as colunas selecionadas.")
                st.session_state.curva_gerada = False

# Exibir resultados se a curva ABC foi gerada com sucesso
if st.session_state.curva_gerada and st.session_state.curva_abc is not None:
    st.markdown("---")
    st.header("✅ Resultados da Curva ABC")

    resultado = st.session_state.curva_abc
    valor_total = st.session_state.valor_total

    # Renomear colunas para exibição amigável
    df_exibicao = resultado.copy()
    df_exibicao.columns = [
        'Posição', 'Código', 'Descrição', 'Valor (R$)',
        'Percentual (%)', 'Percentual Acumulado (%)', 'Classificação'
    ]
    # Formatar colunas numéricas para exibição
    df_exibicao['Valor (R$)'] = df_exibicao['Valor (R$)'].map('{:,.2f}'.format)
    df_exibicao['Percentual (%)'] = df_exibicao['Percentual (%)'].map('{:.2f}%'.format)
    df_exibicao['Percentual Acumulado (%)'] = df_exibicao['Percentual Acumulado (%)'].map('{:.2f}%'.format)


    # --- Estatísticas Resumo ---
    st.subheader("📊 Resumo Estatístico")
    stats_cols = st.columns(4)
    classes_count = resultado['classificacao'].value_counts().to_dict()
    valor_por_classe = resultado.groupby('classificacao')['valor'].sum().to_dict()

    with stats_cols[0]:
        st.metric("Total de Itens", f"{len(resultado):,}")
    with stats_cols[1]:
        st.metric("Valor Total (R$)", f"{valor_total:,.2f}")
    with stats_cols[2]:
        count_a = classes_count.get('A', 0)
        perc_count_a = (count_a / len(resultado) * 100) if len(resultado) > 0 else 0
        st.metric("Itens Classe A", f"{count_a} ({perc_count_a:.1f}%)")
    with stats_cols[3]:
        value_a = valor_por_classe.get('A', 0)
        perc_value_a = (value_a / valor_total * 100) if valor_total > 0 else 0
        st.metric("Valor Classe A", f"{perc_value_a:.1f}%")
    # Adicionar mais métricas se desejar (Classe B, C etc.)

    # --- Gráficos ---
    st.subheader("📈 Análise Gráfica")
    with st.spinner("Gerando gráficos..."):
        fig = criar_graficos_plotly(resultado, valor_total) # Passa o DF original com números
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Não foi possível gerar os gráficos.")

    # --- Tabela de Resultados com Filtros ---
    st.subheader("📋 Tabela de Resultados Detalhada")
    filter_cols = st.columns([0.3, 0.7]) # Colunas para filtros
    with filter_cols[0]:
        classe_filtro = st.multiselect(
            "Filtrar por Classe",
            options=sorted(resultado['classificacao'].unique()),
            default=sorted(resultado['classificacao'].unique()),
            key='filter_classe'
        )
    with filter_cols[1]:
        busca = st.text_input("Buscar por Código ou Descrição", key='search_term')

    # Aplicar filtros ao DataFrame *original* antes da formatação para exibição
    df_filtrado_orig = resultado[resultado['classificacao'].isin(classe_filtro)]
    if busca:
        df_filtrado_orig = df_filtrado_orig[
            df_filtrado_orig['codigo'].astype(str).str.contains(busca, case=False, na=False) |
            df_filtrado_orig['descricao'].astype(str).str.contains(busca, case=False, na=False)
        ]

    # Formatar o DataFrame filtrado para exibição
    df_filtrado_exibicao = df_filtrado_orig.copy()
    df_filtrado_exibicao.columns = [
        'Posição', 'Código', 'Descrição', 'Valor (R$)',
        'Percentual (%)', 'Percentual Acumulado (%)', 'Classificação'
    ]
    # Aplicar formatação apenas para exibição na tabela
    df_filtrado_exibicao['Valor (R$)'] = df_filtrado_orig['valor'].map('R$ {:,.2f}'.format)
    df_filtrado_exibicao['Percentual (%)'] = df_filtrado_orig['percentual'].map('{:.2f}%'.format)
    df_filtrado_exibicao['Percentual Acumulado (%)'] = df_filtrado_orig['percentual_acumulado'].map('{:.2f}%'.format)

    st.dataframe(df_filtrado_exibicao, height=400, use_container_width=True)

    # --- Downloads ---
    st.subheader("📥 Downloads")
    dl_cols = st.columns(2)
    data_atual = datetime.now().strftime("%Y%m%d_%H%M")
    csv_filename = f"curva_abc_{data_atual}.csv"
    excel_filename = f"curva_abc_{data_atual}.xlsx"

    # Usar o DataFrame com nomes de coluna amigáveis para download
    df_download = resultado.copy()
    df_download.columns = [
        'Posição', 'Código', 'Descrição', 'Valor_R$',
        'Percentual_%', 'Percentual_Acumulado_%', 'Classificação'
    ]


    with dl_cols[0]:
         get_download_link(df_download, csv_filename, "Baixar como CSV", file_format='csv')
    with dl_cols[1]:
         get_download_link(df_download, excel_filename, "Baixar como Excel", file_format='excel')


# Rodapé
st.markdown('<div class="footer">', unsafe_allow_html=True)
st.markdown(f"© {datetime.now().year} - Gerador de Curva ABC | Adaptado para Planilhas SINAPI")
st.markdown('</div>', unsafe_allow_html=True)

