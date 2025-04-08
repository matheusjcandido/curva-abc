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

# Estilo CSS personalizado
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
    cabecalhos_possiveis = ['CÓDIGO', 'ITEM', 'DESCRIÇÃO', 'CUSTO', 'VALOR', 'TOTAL', 'PREÇO', 'SERVIÇO']
    max_matches = 0
    header_row_index = 0

    # Verifica as primeiras 20 linhas (ou menos se o df for menor)
    for i in range(min(20, len(df_preview))):
        try:
            # Tenta acessar a linha e verificar valores
            row_values = df_preview.iloc[i]
            # Pula linhas totalmente vazias rapidamente
            if row_values.isnull().all():
                continue
            # Converte para string para busca (ignora erros individuais de célula)
            row_values_str = row_values.astype(str).str.upper().tolist()
            current_matches = sum(any(keyword in str(cell).upper() for keyword in cabecalhos_possiveis) for cell in row_values_str if pd.notna(cell))

            # Considera a linha com mais palavras-chave como cabeçalho
            # Requer pelo menos 2 correspondências para ser considerado um cabeçalho forte
            if current_matches > max_matches and current_matches >= 2:
                max_matches = current_matches
                header_row_index = i
        except Exception:
            # Ignora linhas que causem erro na verificação
            continue

    # Se encontrou um cabeçalho forte, retorna o índice
    if max_matches >= 2:
        return header_row_index

    # Se não encontrou um cabeçalho forte, retorna 0 como padrão
    # (a leitura com header=0 tentará usar a primeira linha)
    return 0


# REMOVIDA: Função sanitizar_dataframe() foi removida pois estava causando o erro.

# --- Função Principal de Processamento ---

def processar_arquivo(uploaded_file):
    """Carrega e processa o arquivo CSV ou Excel, identificando o cabeçalho."""
    df = None
    delimitador = None
    linha_cabecalho = 0 # Default
    encodings_to_try = ['utf-8', 'latin1', 'cp1252'] # ISO-8859-1 é similar a latin1

    try:
        file_name = uploaded_file.name.lower()
        # Usar BytesIO para ler o arquivo em memória uma vez
        file_content_io = io.BytesIO(uploaded_file.getvalue())

        # --- Processamento Excel ---
        if file_name.endswith(('.xlsx', '.xls')):
            # 1. Ler preview para encontrar cabeçalho
            try:
                # Ler sem cabeçalho para inspeção
                df_preview = pd.read_excel(file_content_io, engine='openpyxl', nrows=25, header=None)
                linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            except Exception as e:
                st.warning(f"Não foi possível ler o preview do Excel para achar cabeçalho: {e}. Assumindo linha 0.")
                linha_cabecalho = 0

            # 2. Ler arquivo completo com o cabeçalho correto
            # Resetar o ponteiro do BytesIO antes de ler novamente
            file_content_io.seek(0)
            df = pd.read_excel(file_content_io, engine='openpyxl', header=linha_cabecalho)

        # --- Processamento CSV ---
        elif file_name.endswith('.csv'):
            detected_encoding = None
            decoded_content = None
            file_bytes = file_content_io.getvalue() # Pega os bytes do BytesIO

            # 1. Tentar decodificar com diferentes encodings
            for enc in encodings_to_try:
                try:
                    decoded_content = file_bytes.decode(enc)
                    detected_encoding = enc
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

            # 3. Ler preview para encontrar cabeçalho
            try:
                # Usar StringIO para ler o conteúdo decodificado
                df_preview = pd.read_csv(io.StringIO(decoded_content), delimiter=delimitador, nrows=25, header=None, skipinitialspace=True, low_memory=False)
                linha_cabecalho = encontrar_linha_cabecalho(df_preview)
            except Exception as e:
                st.warning(f"Não foi possível ler o preview do CSV para achar cabeçalho: {e}. Assumindo linha 0.")
                linha_cabecalho = 0

            # 4. Ler arquivo completo com cabeçalho e delimitador corretos
            df = pd.read_csv(
                io.StringIO(decoded_content), # Usa o conteúdo já decodificado
                delimiter=delimitador,
                header=linha_cabecalho,
                # encoding=detected_encoding, # Não precisa mais, já decodificado
                on_bad_lines='warn', # Avisa sobre linhas ruins mas tenta continuar
                skipinitialspace=True, # Ignora espaços após o delimitador
                low_memory=False # Ajuda com tipos mistos, mas usa mais memória
            )

        else:
            st.error("Formato de arquivo não suportado. Use CSV, XLSX ou XLS.")
            return None, None

        # --- Pós-processamento Comum ---
        if df is not None:
            # Limpar nomes das colunas (converter para string, remover espaços extras)
            df.columns = [str(col).strip() if pd.notna(col) else f'coluna_sem_nome_{i}' for i, col in enumerate(df.columns)]
            # Renomear colunas duplicadas para garantir unicidade
            cols = pd.Series(df.columns)
            for dup in cols[cols.duplicated()].unique():
                cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
            df.columns = cols

            # Remover linhas e colunas totalmente vazias
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty:
                 st.error("O arquivo parece vazio após remover linhas/colunas nulas.")
                 return None, delimitador

            # REMOVIDO: Chamada para sanitizar_dataframe()
            # df = sanitizar_dataframe(df)
            # if df is None or df.empty:
            #      st.error("Falha na sanitização do DataFrame.")
            #      return None, delimitador

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
    # Adicionado espaço negativo e outros caracteres comuns em planilhas
    valor_str = re.sub(r'[R$€£¥\s\u00A0\(\)-]', '', valor_str)

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
    # Permite um único ponto decimal
    valor_str_limpo = ""
    decimal_found = False
    for char in valor_str:
        if char.isdigit():
            valor_str_limpo += char
        elif char == '.' and not decimal_found:
            valor_str_limpo += char
            decimal_found = True

    try:
        return float(valor_str_limpo) if valor_str_limpo else 0.0
    except ValueError:
        # st.warning(f"Não foi possível converter '{valor}' para número após limpeza -> '{valor_str_limpo}'.")
        return 0.0


def identificar_colunas(df):
    """Identifica heuristicamente as colunas de código, descrição e valor."""
    coluna_codigo = None
    coluna_descricao = None
    coluna_valor = None

    # Prioridade para nomes exatos comuns
    exact_matches = {
        'codigo': ['código', 'codigo', 'cod.', 'item', 'ref', 'referencia', 'referência'],
        'descricao': ['descrição', 'descricao', 'desc', 'especificação', 'especificacao', 'serviço', 'servico', 'designação', 'designacao'],
        'valor': ['valor total', 'custo total', 'preço total', 'total', 'valor', 'custo', 'preço']
    }

    cols_lower = {str(col).lower().strip(): col for col in df.columns}
    original_cols = list(df.columns) # Mantem a ordem original

    # 1. Busca por nomes exatos, respeitando a ordem original das colunas
    found_cols = set()
    for col_original in original_cols:
        col_lower = str(col_original).lower().strip()
        if col_original in found_cols: continue # Pula se já foi usada

        if coluna_codigo is None and any(pattern == col_lower for pattern in exact_matches['codigo']):
            coluna_codigo = col_original
            found_cols.add(col_original)
            continue

        if coluna_descricao is None and any(pattern == col_lower for pattern in exact_matches['descricao']):
            coluna_descricao = col_original
            found_cols.add(col_original)
            continue

        if coluna_valor is None and any(pattern == col_lower for pattern in exact_matches['valor']):
             # Verifica se a coluna parece numérica
             try:
                 if pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].astype(str).str.contains(r'[\d,.]').any():
                      coluna_valor = col_original
                      found_cols.add(col_original)
                      continue
             except Exception: pass # Ignora erro na verificação

    # 2. Busca por padrões parciais se não encontrou por nome exato
    partial_patterns = {
        'codigo': ['cod', 'item', 'ref'],
        'descricao': ['desc', 'especif', 'serv', 'design'],
        'valor': ['total', 'valor', 'custo', 'preço', 'vlr', 'vlr.', 'custo']
    }

    for col_original in original_cols:
        if col_original in found_cols: continue
        col_lower = str(col_original).lower().strip()

        if coluna_codigo is None and any(p in col_lower for p in partial_patterns['codigo']):
            coluna_codigo = col_original
            found_cols.add(col_original)
            continue

        if coluna_descricao is None and any(p in col_lower for p in partial_patterns['descricao']):
            coluna_descricao = col_original
            found_cols.add(col_original)
            continue

        if coluna_valor is None and any(p in col_lower for p in partial_patterns['valor']):
             try:
                 if pd.api.types.is_numeric_dtype(df[col_original]) or df[col_original].astype(str).str.contains(r'[\d,.]').any():
                      coluna_valor = col_original
                      found_cols.add(col_original)
                      continue
             except Exception: pass

    # 3. Heurísticas baseadas em conteúdo (último recurso)
    remaining_cols = [c for c in original_cols if c not in found_cols]

    if not coluna_descricao and remaining_cols:
         # Coluna com strings mais longas é provavelmente descrição
         mean_lengths = {}
         for col in remaining_cols:
              try:
                   # Calcula o comprimento médio apenas para valores não nulos convertidos para string
                   mean_lengths[col] = df[col].dropna().astype(str).str.len().mean()
              except Exception: continue
         if mean_lengths:
              # Pega a coluna com maior comprimento médio
              potential_desc = max(mean_lengths, key=mean_lengths.get)
              # Verifica se o comprimento médio é razoavelmente longo (evita colunas de códigos curtos)
              if mean_lengths[potential_desc] > 10:
                   coluna_descricao = potential_desc
                   found_cols.add(potential_desc)
                   remaining_cols = [c for c in original_cols if c not in found_cols] # Atualiza restantes

    if not coluna_valor and remaining_cols:
         # Coluna com maior soma de valores numéricos é provavelmente valor
         max_sum = -float('inf') # Inicia com infinito negativo
         best_val_col = None
         for col in remaining_cols:
              try:
                   # Tenta limpar e converter a coluna inteira
                   numeric_vals = df[col].apply(limpar_valor)
                   # Calcula a soma e a contagem de valores não-zero
                   current_sum = numeric_vals[numeric_vals > 0].sum()
                   non_zero_count = (numeric_vals > 0).sum()
                   total_count = len(numeric_vals.dropna())

                   # Condições: Soma significativa E a maioria dos valores são numéricos > 0
                   if current_sum > max_sum and total_count > 0 and (non_zero_count / total_count) > 0.5:
                        max_sum = current_sum
                        best_val_col = col
              except Exception: continue
         if best_val_col:
              coluna_valor = best_val_col
              found_cols.add(best_val_col)
              # remaining_cols = [c for c in original_cols if c not in found_cols] # Atualiza restantes

    # Se ainda falta o código, pega a primeira coluna restante (chute comum)
    if not coluna_codigo and remaining_cols:
        coluna_codigo = remaining_cols[0]
        found_cols.add(remaining_cols[0])

    # Se ainda falta descrição, pega a próxima restante
    remaining_cols = [c for c in original_cols if c not in found_cols]
    if not coluna_descricao and remaining_cols:
         coluna_descricao = remaining_cols[0]
         found_cols.add(remaining_cols[0])

    # Se ainda falta valor, pega a próxima restante
    remaining_cols = [c for c in original_cols if c not in found_cols]
    if not coluna_valor and remaining_cols:
         coluna_valor = remaining_cols[0]
         found_cols.add(remaining_cols[0])

    # Garante que as colunas sejam diferentes (se a heurística falhou)
    final_cols = [coluna_codigo, coluna_descricao, coluna_valor]
    if len(set(filter(None, final_cols))) != len(list(filter(None, final_cols))):
         # Houve duplicação, anula os valores para forçar seleção manual
         st.warning("Detecção automática resultou em colunas duplicadas. Revise a seleção.")
         # Não retorna None, deixa o usuário corrigir na interface
         pass # Mantém os valores duplicados para o usuário ver e corrigir

    return coluna_codigo, coluna_descricao, coluna_valor


def gerar_curva_abc(df, coluna_codigo, coluna_descricao, coluna_valor, limite_a=80, limite_b=95):
    """Gera a curva ABC a partir do DataFrame processado."""
    # Validação inicial
    if not all([coluna_codigo, coluna_descricao, coluna_valor]):
        st.error("Erro: Selecione as colunas de Código, Descrição e Valor.")
        return None, 0
    if coluna_codigo not in df.columns or coluna_descricao not in df.columns or coluna_valor not in df.columns:
         st.error(f"Erro interno: Colunas selecionadas não encontradas no DataFrame processado. Colunas disponíveis: {list(df.columns)}")
         return None, 0
    if len(set([coluna_codigo, coluna_descricao, coluna_valor])) < 3:
         st.error("Erro: As colunas de Código, Descrição e Valor devem ser diferentes.")
         return None, 0

    try:
        # Seleciona e copia as colunas relevantes para evitar modificar o df original
        df_work = df[[coluna_codigo, coluna_descricao, coluna_valor]].copy()

        # Limpa e converte a coluna de valor (aplicar ANTES de agrupar)
        df_work['valor_numerico'] = df_work[coluna_valor].apply(limpar_valor)

        # Converte código e descrição para string e remove espaços extras
        # Usar .loc para evitar SettingWithCopyWarning
        df_work.loc[:, 'codigo_str'] = df_work[coluna_codigo].astype(str).str.strip()
        df_work.loc[:, 'descricao_str'] = df_work[coluna_descricao].astype(str).str.strip()

        # Filtra itens com valor zero ou negativo e código vazio ANTES de agrupar
        df_filtered = df_work[(df_work['valor_numerico'] > 0) & (df_work['codigo_str'] != '')].copy()

        if df_filtered.empty:
            st.error("Nenhum item com valor positivo e código válido encontrado após limpeza inicial.")
            # Mostrar amostra dos dados que foram filtrados para depuração
            with st.expander("Dados filtrados (valor <= 0 ou código vazio)"):
                 st.dataframe(df_work[~((df_work['valor_numerico'] > 0) & (df_work['codigo_str'] != ''))].head(20))
            return None, 0

        # Agrupa por código, somando valores e pegando a primeira descrição não vazia
        # Usar lambda para pegar a primeira descrição válida
        first_valid_desc = lambda x: x.dropna().astype(str).iloc[0] if not x.dropna().empty else ''

        df_agrupado = df_filtered.groupby('codigo_str').agg(
            # Pega a primeira descrição encontrada para aquele código
            descricao=('descricao_str', first_valid_desc),
            valor=('valor_numerico', 'sum')
        ).reset_index() # Converte o índice (codigo_str) de volta para coluna

        # Renomeia a coluna de código agrupado
        df_agrupado = df_agrupado.rename(columns={'codigo_str': 'codigo'})

        # Verifica novamente se o DataFrame agrupado está vazio
        if df_agrupado.empty:
             st.error("O DataFrame ficou vazio após agrupar os itens por código.")
             return None, 0

        # Calcula valor total
        valor_total = df_agrupado['valor'].sum()
        if valor_total == 0:
            st.error("Valor total dos itens agrupados é zero. Não é possível gerar a curva.")
            return None, 0

        # Ordenar por valor
        df_curva = df_agrupado.sort_values('valor', ascending=False).reset_index(drop=True)

        # Adicionar colunas de percentual e classificação
        df_curva['percentual'] = (df_curva['valor'] / valor_total * 100)
        # Evitar percentual acumulado > 100 devido a arredondamentos
        df_curva['percentual_acumulado'] = df_curva['percentual'].cumsum().clip(upper=100.0)

        def classificar(perc_acum):
            # Usar limites ligeiramente ajustados para evitar problemas de ponto flutuante
            if perc_acum <= limite_a + 1e-9: return 'A'
            elif perc_acum <= limite_b + 1e-9: return 'B'
            else: return 'C'
        df_curva['classificacao'] = df_curva['percentual_acumulado'].apply(classificar)

        # Adicionar posição
        df_curva.insert(0, 'posicao', range(1, len(df_curva) + 1))

        # Selecionar e reordenar colunas finais
        df_final = df_curva[['posicao', 'codigo', 'descricao', 'valor', 'percentual', 'percentual_acumulado', 'classificacao']]

        return df_final, valor_total

    except Exception as e:
        st.error(f"Erro inesperado ao gerar a curva ABC: {str(e)}")
        with st.expander("Detalhes técnicos do erro"):
            st.text(traceback.format_exc())
        return None, 0

# --- Funções de Visualização e Download ---

def criar_graficos_plotly(df_curva, valor_total, limite_a, limite_b): # Adicionado limites como parâmetros
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
                hoverinfo='x+y+text+name',
                hovertemplate='<b>Pos:</b> %{x}<br><b>Código:</b> %{text}<br><b>Valor:</b> R$ %{y:,.2f}<extra></extra>' # Template de hover
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
                marker=dict(size=4),
                yaxis='y2', # Especifica o eixo Y secundário
                hovertemplate='<b>Pos:</b> %{x}<br><b>Acumulado:</b> %{y:.2f}%<extra></extra>' # Template de hover
            ),
            # secondary_y=True, # Não precisa aqui, já definido no yaxis='y2'
             row=1, col=1
        )
        # Linhas de referência A e B
        fig.add_hline(y=limite_a, line_dash="dash", line_color="grey", annotation_text=f"Classe A ({limite_a}%)",
                      annotation_position="bottom right", secondary_y=True, row=1, col=1)
        fig.add_hline(y=limite_b, line_dash="dash", line_color="grey", annotation_text=f"Classe B ({limite_b}%)",
                      annotation_position="bottom right", secondary_y=True, row=1, col=1)

        # --- Gráfico 2: Pizza por Valor ---
        valor_por_classe = df_curva.groupby('classificacao')['valor'].sum().reindex(['A', 'B', 'C']).fillna(0)
        fig.add_trace(
            go.Pie(
                labels=valor_por_classe.index,
                values=valor_por_classe.values,
                name='Valor',
                marker_colors=[colors.get(c, '#cccccc') for c in valor_por_classe.index], # Cor cinza se classe não existir
                hole=0.4,
                pull=[0.05 if c == 'A' else 0 for c in valor_por_classe.index], # Destaca A
                textinfo='percent+label',
                hoverinfo='label+percent+value+name',
                hovertemplate='<b>Classe:</b> %{label}<br><b>Valor:</b> R$ %{value:,.2f}<br><b>Percentual:</b> %{percent:.1%}<extra></extra>'
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
                marker_colors=[colors.get(c, '#cccccc') for c in qtd_por_classe.index],
                hole=0.4,
                pull=[0.05 if c == 'A' else 0 for c in qtd_por_classe.index], # Destaca A
                textinfo='percent+label',
                hoverinfo='label+percent+value+name',
                hovertemplate='<b>Classe:</b> %{label}<br><b>Itens:</b> %{value:,d}<br><b>Percentual:</b> %{percent:.1%}<extra></extra>'
            ),
            row=2, col=1
        )

        # --- Gráfico 4: Top 10 Itens ---
        top10 = df_curva.head(10).sort_values('valor', ascending=True) # Ordena para gráfico hbar
        # Criar texto combinado para eixo Y
        top10_labels = top10.apply(lambda row: f"{row['codigo']} ({str(row['descricao'])[:30]}...)", axis=1)
        fig.add_trace(
            go.Bar(
                y=top10_labels, # Usa os labels combinados
                x=top10['valor'],
                name='Top 10 Valor',
                orientation='h',
                marker_color=top10['classificacao'].map(colors),
                text=top10['valor'].map('R$ {:,.2f}'.format), # Formata valor como texto
                textposition='outside', # Coloca texto fora da barra
                hoverinfo='y+x+name',
                hovertemplate='<b>Item:</b> %{y}<br><b>Valor:</b> R$ %{x:,.2f}<extra></extra>'
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
            margin=dict(l=40, r=40, t=80, b=40), # Ajusta margens
            paper_bgcolor='rgba(0,0,0,0)', # Fundo transparente
            plot_bgcolor='rgba(0,0,0,0)',  # Fundo transparente
            hovermode='closest' # Melhora interação do hover
        )

        # Layout Eixos Pareto
        fig.update_yaxes(title_text="Valor do Item (R$)", secondary_y=False, row=1, col=1, zeroline=True, zerolinewidth=1, zerolinecolor='lightgrey')
        fig.update_yaxes(title_text="Percentual Acumulado (%)", secondary_y=True, row=1, col=1, range=[0, 101], tickformat=".0f", gridcolor='lightgrey')
        fig.update_xaxes(title_text="Posição do Item (Ordenado por Valor)", row=1, col=1, gridcolor='lightgrey')

        # Layout Eixo Top 10
        fig.update_xaxes(title_text="Valor (R$)", row=2, col=2, zeroline=True, zerolinewidth=1, zerolinecolor='lightgrey')
        fig.update_yaxes(title_text="Item (Código + Descrição)", autorange="reversed", row=2, col=2, tickfont_size=10)

        # Atualiza títulos dos subplots para clareza
        fig.layout.annotations[0].update(text="<b>Diagrama de Pareto</b> (Valor x Acumulado %)")
        fig.layout.annotations[1].update(text="<b>Distribuição do Valor Total (%)</b> por Classe")
        fig.layout.annotations[2].update(text="<b>Distribuição da Quantidade de Itens (%)</b> por Classe")
        fig.layout.annotations[3].update(text="<b>Top 10 Itens</b> por Valor")

        # Adiciona anotação com valor total no gráfico de Pareto
        fig.add_annotation(
            xref='paper', yref='paper', x=0.01, y=1.05, # Posição no canto superior esquerdo
            text=f"<b>Valor Total: R$ {valor_total:,.2f}</b>",
            showarrow=False, font=dict(size=12, color="#1e3c72")
        )

        return fig

    except Exception as e:
        st.error(f"Erro ao criar gráficos: {str(e)}")
        with st.expander("Detalhes técnicos do erro"):
            st.text(traceback.format_exc())
        return None

def get_download_link(df, filename, text, file_format='csv'):
    """Gera um botão de download para o DataFrame como CSV ou Excel."""
    try:
        if file_format == 'csv':
            # Usar utf-8-sig para garantir que o Excel abra corretamente com acentos
            data = df.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig')
            mime = 'text/csv'
        elif file_format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Curva ABC', index=False)
                # --- Formatação Opcional do Excel ---
                try:
                    workbook = writer.book
                    worksheet = writer.sheets['Curva ABC']
                    # Formatos (definidos apenas se a formatação for aplicada)
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#1e3c72', 'color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                    currency_format = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                    percent_format = workbook.add_format({'num_format': '0.00"%"', 'border': 1}) # Formato percentual correto
                    center_format = workbook.add_format({'align': 'center', 'border': 1})
                    default_format = workbook.add_format({'border': 1}) # Formato padrão com borda

                    # Mapeia nomes das colunas para índices
                    col_map = {name: i for i, name in enumerate(df.columns)}

                    # Aplica formato de cabeçalho e ajusta largura inicial
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        # Define largura inicial baseada no cabeçalho
                        worksheet.set_column(col_num, col_num, len(str(value)) + 2)

                    # Aplica formatos às colunas de dados (da linha 1 em diante)
                    num_rows = len(df)
                    if 'Valor_R$' in col_map: worksheet.set_column(col_map['Valor_R$'], col_map['Valor_R$'], 15, currency_format)
                    if 'Percentual_%' in col_map: worksheet.set_column(col_map['Percentual_%'], col_map['Percentual_%'], 12, percent_format)
                    if 'Percentual_Acumulado_%' in col_map: worksheet.set_column(col_map['Percentual_Acumulado_%'], col_map['Percentual_Acumulado_%'], 15, percent_format)
                    if 'Classificação' in col_map: worksheet.set_column(col_map['Classificação'], col_map['Classificação'], 12, center_format)
                    if 'Posição' in col_map: worksheet.set_column(col_map['Posição'], col_map['Posição'], 8, center_format)
                    if 'Código' in col_map: worksheet.set_column(col_map['Código'], col_map['Código'], 15, default_format) # Código com formato padrão
                    if 'Descrição' in col_map: worksheet.set_column(col_map['Descrição'], col_map['Descrição'], 50, default_format) # Descrição mais larga

                    # Congelar painel superior (cabeçalho)
                    worksheet.freeze_panes(1, 0)
                except Exception as format_e:
                    st.warning(f"Erro durante a formatação do Excel (continuando sem formatação avançada): {format_e}")
            # --- Fim da Formatação Opcional ---
            output.seek(0)
            data = output.getvalue() # Use getvalue() para bytes
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            st.error("Formato de download inválido.")
            return

        # Usa st.download_button
        st.download_button(
            label=text,
            data=data, # Passa bytes diretamente para CSV e Excel
            file_name=filename,
            mime=mime,
            key=f"download_{file_format}" # Chave única para o botão
        )

    except Exception as e:
        st.error(f"Erro ao gerar dados para download ({file_format}): {e}")
        # Não retorna nada, o botão não será criado


# --- Interface Streamlit ---

# Sidebar
with st.sidebar:
    # st.image("https://via.placeholder.com/150x50.png?text=Logo+Empresa", use_column_width=True) # Placeholder para logo
    st.header("⚙️ Configurações")
    st.subheader("Parâmetros da Curva ABC")
    # Usar session_state para manter os valores dos sliders entre reruns
    if 'limite_a' not in st.session_state: st.session_state.limite_a = 80
    if 'limite_b' not in st.session_state: st.session_state.limite_b = 95

    # Define os limites dos sliders
    limite_a_value = st.slider("Limite Classe A (%)", min_value=50, max_value=95, value=st.session_state.limite_a, step=1, key='limite_a_slider')
    # Garante que limite B seja maior que A
    limite_b_min = limite_a_value + 1
    limite_b_value = st.slider("Limite Classe B (%)", min_value=limite_b_min, max_value=99, value=max(st.session_state.limite_b, limite_b_min), step=1, key='limite_b_slider')

    # Atualiza session_state APENAS se os valores mudarem
    if limite_a_value != st.session_state.limite_a:
        st.session_state.limite_a = limite_a_value
        st.session_state.curva_gerada = False # Força recalcular se limite mudar
    if limite_b_value != st.session_state.limite_b:
        st.session_state.limite_b = limite_b_value
        st.session_state.curva_gerada = False # Força recalcular se limite mudar


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
    st.caption(f"Versão 1.2 - {datetime.now().y
