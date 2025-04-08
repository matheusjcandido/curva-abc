import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io
import base64
import re
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

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
    }
    .highlight {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1e3c72;
    }
    .footer {
        margin-top: 20px;
        padding-top: 10px;
        border-top: 1px solid #ddd;
        text-align: center;
        font-size: 0.8em;
        color: #666;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.title("Gerador de Curva ABC - SINAPI")
st.markdown("### Automatize a geração da Curva ABC a partir de planilhas sintéticas do SINAPI")

# Função para detectar o delimitador do CSV
def detectar_delimitador(content):
    """Detecta o delimitador usado no arquivo CSV."""
    primeiras_linhas = content[:5000]  # Analisa apenas os primeiros 5000 caracteres
    
    contagem = {
        ';': primeiras_linhas.count(';'),
        ',': primeiras_linhas.count(','),
        '\t': primeiras_linhas.count('\t')
    }
    
    return max(contagem, key=contagem.get)

# Função para encontrar a linha de cabeçalho
def encontrar_linha_cabecalho(df):
    """Encontra a linha que contém os cabeçalhos da tabela."""
    cabecalhos_possiveis = ['CÓDIGO', 'ITEM', 'DESCRIÇÃO', 'CUSTO TOTAL', 'VALOR TOTAL']
    
    for i in range(min(15, len(df))):  # Verifica as primeiras 15 linhas
        row = df.iloc[i]
        row_text = ' '.join([str(x).upper() for x in row.values])
        
        if any(cabecalho in row_text for cabecalho in cabecalhos_possiveis):
            return i
    
    return 0  # Se não encontrar, assume a primeira linha

# Função para carregar e processar o arquivo CSV
def processar_arquivo(uploaded_file, encoding='utf-8'):
    """Carrega e processa o arquivo CSV."""
    try:
        # Para arquivos Excel
        if uploaded_file.name.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            linha_cabecalho = encontrar_linha_cabecalho(df)
            
            # Define a linha de cabeçalho e reindexa o DataFrame
            cabecalhos = df.iloc[linha_cabecalho].values
            df = pd.DataFrame(df.values[linha_cabecalho+1:], columns=cabecalhos)
            
            # Limpeza e sanitização do DataFrame
            df = sanitizar_dataframe(df)
            return df, None
        
        # Para arquivos CSV
        # Ler o conteúdo do arquivo
        content = uploaded_file.getvalue().decode(encoding)
        
        # Verificar se o arquivo está vazio
        if content.strip() == '':
            st.error("O arquivo enviado está vazio. Por favor, verifique o arquivo e tente novamente.")
            return None, None
        
        # Detectar o delimitador
        delimitador = detectar_delimitador(content)
        
        # Carregar o CSV sem cabeçalhos inicialmente
        df = pd.read_csv(
            io.StringIO(content),
            delimiter=delimitador,
            encoding=encoding,
            on_bad_lines='warn',  # Parâmetro atualizado
            header=None,  # Sem cabeçalhos inicialmente
            low_memory=False
        )
        
        # Encontrar a linha de cabeçalho
        linha_cabecalho = encontrar_linha_cabecalho(df)
        
        # Define a linha de cabeçalho e reindexa o DataFrame
        cabecalhos = df.iloc[linha_cabecalho].values
        df = pd.DataFrame(df.values[linha_cabecalho+1:], columns=cabecalhos)
        
        # Limpar dados
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Sanitizar o DataFrame para evitar problemas com PyArrow/Streamlit
        df = sanitizar_dataframe(df)
        
        return df, delimitador
    except UnicodeDecodeError:
        # Tentar com encoding alternativo se utf-8 falhar
        if encoding == 'utf-8':
            return processar_arquivo(uploaded_file, 'latin1')
        elif encoding == 'latin1':
            return processar_arquivo(uploaded_file, 'cp1252')
        else:
            st.error(f"Erro ao decodificar o arquivo com encoding {encoding}.")
            return None, None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
        import traceback
        with st.expander("Detalhes do erro"):
            st.text(traceback.format_exc())
        return None, None

# Função para sanitizar o DataFrame e garantir compatibilidade com PyArrow/Streamlit
def sanitizar_dataframe(df):
    """Sanitiza o DataFrame para garantir compatibilidade com PyArrow/Streamlit."""
    if df is None:
        return None
    
    # Cria uma cópia para evitar modificar o original
    df_clean = df.copy()
    
    # Converte tipos problemáticos para string para evitar erros PyArrow
    for col in df_clean.columns:
        try:
            # Verifica se a coluna tem tipos mistos
            if df_clean[col].apply(type).nunique() > 1:
                df_clean[col] = df_clean[col].astype(str)
            
            # Converte coluna de objetos para string
            if df_clean[col].dtype == 'object':
                df_clean[col] = df_clean[col].astype(str)
        except:
            # Em caso de erro, converte para string
            df_clean[col] = df_clean[col].astype(str)
    
    # Garante que os nomes das colunas sejam strings
    df_clean.columns = df_clean.columns.astype(str)
    
    # Remove caracteres nulos que podem causar problemas
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object' or df_clean[col].dtype == 'string':
            df_clean[col] = df_clean[col].str.replace('\0', '', regex=False)
    
    return df_clean

# Função para limpar e converter valores numéricos
def limpar_valor(valor_str):
    """Limpa e converte valores monetários para float."""
    if pd.isna(valor_str):
        return 0.0
    
    if isinstance(valor_str, (int, float)):
        return float(valor_str)
    
    # Remove qualquer caractere não numérico exceto . e ,
    valor_str = re.sub(r'[^\d.,]', '', str(valor_str))
    
    if not valor_str:
        return 0.0
    
    # Verifica se temos . e , para determinar o separador decimal
    if ',' in valor_str and '.' in valor_str:
        # Se o último separador for uma vírgula, provavelmente é o separador decimal (formato brasileiro)
        if valor_str.rindex(',') > valor_str.rindex('.'):
            # Formato brasileiro: 1.234,56
            valor_str = valor_str.replace('.', '').replace(',', '.')
        else:
            # Formato americano: 1,234.56
            valor_str = valor_str.replace(',', '')
    elif ',' in valor_str:
        # Apenas vírgulas presentes, assume que é o separador decimal
        valor_str = valor_str.replace(',', '.')
    
    try:
        return float(valor_str)
    except ValueError:
        return 0.0

# Função para identificar colunas relevantes
def identificar_colunas(df):
    """Identifica as colunas de código, descrição e valor."""
    # Padrões para busca (expandidos)
    padroes_codigo = ['cod', 'código', 'codigo', 'referência', 'referencia', 'ref', 'code', 'item code', 'serviço', 'servico']
    padroes_descricao = ['desc', 'descrição', 'descricao', 'serviço', 'servico', 'item', 'especificação', 'especificacao']
    padroes_valor = ['valor total', 'total', 'preço total', 'preco total', 'valor', 'custo', 'custo total', 'preço', 'preco']
    
    # Mapeia as colunas para nomes mais fáceis de processar
    cols_map = {}
    for col in df.columns:
        cols_map[str(col).lower().strip()] = col
    
    coluna_codigo = None
    coluna_descricao = None
    coluna_valor = None
    
    # Buscar exatamente pelos nomes comuns em planilhas SINAPI
    exact_codigo_matches = ['código do serviço', 'código', 'codigo']
    exact_descricao_matches = ['descrição do serviço', 'descrição', 'descricao']
    exact_valor_matches = ['custo total', 'valor total', 'total']
    
    for col_lower in cols_map:
        # Verificar correspondências exatas primeiro
        if not coluna_codigo:
            if col_lower in exact_codigo_matches:
                coluna_codigo = cols_map[col_lower]
                continue
                
        if not coluna_descricao:
            if col_lower in exact_descricao_matches:
                coluna_descricao = cols_map[col_lower]
                continue
                
        if not coluna_valor:
            if col_lower in exact_valor_matches:
                coluna_valor = cols_map[col_lower]
                continue
        
        # Verificar padrões parciais
        if not coluna_codigo and any(p in col_lower for p in padroes_codigo):
            coluna_codigo = cols_map[col_lower]
        
        if not coluna_descricao and any(p in col_lower for p in padroes_descricao):
            coluna_descricao = cols_map[col_lower]
        
        if not coluna_valor and any(p in col_lower for p in padroes_valor):
            coluna_valor = cols_map[col_lower]
    
    # Se não encontrou pelo nome, verificar conteúdo
    if not coluna_codigo:
        for col in df.columns:
            valores = df[col].astype(str)
            # Busca por padrões de códigos SINAPI ou códigos compostos
            if valores.str.match(r'^\d{4,}$').any() or valores.str.match(r'^COMP\s*\d+').any() or valores.str.contains('COMP', case=False).any():
                coluna_codigo = col
                break
    
    if not coluna_descricao:
        for col in df.columns:
            if col == coluna_codigo or col == coluna_valor:
                continue
            
            valores = df[col].astype(str)
            # Textos longos são provavelmente descrições
            if valores.str.len().mean() > 15:
                coluna_descricao = col
                break
    
    if not coluna_valor:
        max_valor = 0
        for col in df.columns:
            if col == coluna_codigo or col == coluna_descricao:
                continue
            
            try:
                # Limpar e converter valores
                valores = df[col].apply(limpar_valor)
                
                if valores.max() > max_valor and not valores.isna().all():
                    max_valor = valores.max()
                    coluna_valor = col
            except:
                continue
    
    # Exibir informações de depuração sobre as colunas encontradas
    with st.expander("Informações de depuração - Colunas detectadas"):
        st.write(f"Coluna de Código: {coluna_codigo}")
        st.write(f"Coluna de Descrição: {coluna_descricao}")
        st.write(f"Coluna de Valor: {coluna_valor}")
        
        # Mostrar amostra
        if coluna_codigo and coluna_descricao and coluna_valor:
            st.write("Amostra de dados:")
            amostra = pd.DataFrame({
                'Código': df[coluna_codigo].head(3),
                'Descrição': df[coluna_descricao].head(3),
                'Valor': df[coluna_valor].head(3)
            })
            st.dataframe(amostra)
    
    return coluna_codigo, coluna_descricao, coluna_valor

# Função para gerar a curva ABC
def gerar_curva_abc(df, coluna_codigo, coluna_descricao, coluna_valor, limite_a=80, limite_b=95):
    """Gera a curva ABC a partir do DataFrame."""
    try:
        # Criar dicionário para agrupar itens com o mesmo código
        itens_agrupados = {}
        
        # Processar cada linha
        for _, row in df.iterrows():
            codigo = str(row[coluna_codigo]).strip()
            descricao = str(row[coluna_descricao]).strip()
            
            # Obter o valor usando a função de limpeza
            valor = limpar_valor(row[coluna_valor])
            
            # Verificar se temos um código válido
            if codigo and (re.match(r'^\d+$', codigo) or re.match(r'^COMP\s*\d+', codigo, re.IGNORECASE) or re.match(r'^COMP\d+', codigo, re.IGNORECASE)):
                if valor > 0:
                    if codigo in itens_agrupados:
                        itens_agrupados[codigo]['valor'] += valor
                    else:
                        itens_agrupados[codigo] = {
                            'codigo': codigo,
                            'descricao': descricao,
                            'valor': valor
                        }
        
        # Verificar se encontramos itens
        if not itens_agrupados:
            st.error("Não foi possível encontrar itens válidos. Verifique se as colunas selecionadas estão corretas.")
            with st.expander("Detalhes do erro"):
                st.write("Não foram encontrados itens com códigos válidos e valores positivos.")
                st.write("Verifique se os dados estão no formato esperado:")
                st.write("- Códigos numéricos ou iniciados com 'COMP'")
                st.write("- Valores numéricos positivos")
                st.write("- Colunas corretamente identificadas")
            return None, 0
        
        # Converter para DataFrame
        df_curva = pd.DataFrame(list(itens_agrupados.values()))
        
        # Calcular valor total
        valor_total = df_curva['valor'].sum()
        
        # Ordenar por valor
        df_curva = df_curva.sort_values('valor', ascending=False).reset_index(drop=True)
        
        # Adicionar colunas
        df_curva['percentual'] = (df_curva['valor'] / valor_total * 100).round(2)
        df_curva['percentual_acumulado'] = df_curva['percentual'].cumsum().round(2)
        
        # Classificar
        def classificar(perc_acum):
            if perc_acum <= limite_a:
                return 'A'
            elif perc_acum <= limite_b:
                return 'B'
            else:
                return 'C'
        
        df_curva['classificacao'] = df_curva['percentual_acumulado'].apply(classificar)
        
        # Adicionar posição
        df_curva.insert(0, 'posicao', range(1, len(df_curva) + 1))
        
        return df_curva, valor_total
    
    except Exception as e:
        st.error(f"Erro ao gerar a curva ABC: {str(e)}")
        import traceback
        with st.expander("Detalhes do erro"):
            st.write(traceback.format_exc())
        return None, 0

# Função para criar gráficos interativos com Plotly
def criar_graficos_plotly(df_curva, valor_total):
    """Cria gráficos interativos usando Plotly."""
    # Criar figura com subplots
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=("Diagrama de Pareto", "Distribuição por Valor", 
                       "Distribuição por Quantidade de Itens", "Top 10 Itens"),
        specs=[
            [{"type": "scatter"}, {"type": "pie"}],
            [{"type": "pie"}, {"type": "bar"}]
        ],
        vertical_spacing=0.1,
        horizontal_spacing=0.1,
    )
    
    # 1. Gráfico de Pareto
    fig.add_trace(
        go.Bar(
            x=df_curva['posicao'], 
            y=df_curva['percentual'],
            name='Percentual',
            marker_color='royalblue'
        ),
        row=1, col=1
    )
    
    fig.add_trace(
        go.Scatter(
            x=df_curva['posicao'], 
            y=df_curva['percentual_acumulado'],
            name='Percentual Acumulado',
            line=dict(color='firebrick', width=3),
            yaxis='y2'
        ),
        row=1, col=1
    )
    
    # Adicionar linhas horizontais
    fig.add_shape(
        type="line", line=dict(dash='dash', color='green', width=2),
        y0=80, y1=80, x0=0, x1=len(df_curva),
        xref='x', yref='y2', row=1, col=1
    )
    
    fig.add_shape(
        type="line", line=dict(dash='dash', color='orange', width=2),
        y0=95, y1=95, x0=0, x1=len(df_curva),
        xref='x', yref='y2', row=1, col=1
    )
    
    # 2. Gráfico de pizza para distribuição por valor
    valor_por_classe = df_curva.groupby('classificacao')['valor'].sum()
    
    fig.add_trace(
        go.Pie(
            labels=valor_por_classe.index,
            values=valor_por_classe.values,
            textinfo='percent+label',
            hole=0.3,
            marker=dict(colors=['green', 'gold', 'darkorange'])
        ),
        row=1, col=2
    )
    
    # 3. Gráfico de pizza para distribuição por quantidade
    qtd_por_classe = df_curva['classificacao'].value_counts()
    
    fig.add_trace(
        go.Pie(
            labels=qtd_por_classe.index,
            values=qtd_por_classe.values,
            textinfo='percent+label',
            hole=0.3,
            marker=dict(colors=['green', 'gold', 'darkorange'])
        ),
        row=2, col=1
    )
    
    # 4. Gráfico de barras para top 10 itens
    top10 = df_curva.head(10)
    
    fig.add_trace(
        go.Bar(
            y=top10['codigo'],
            x=top10['percentual'],
            orientation='h',
            text=top10['percentual'].map('{:.2f}%'.format),
            textposition='auto',
            marker=dict(
                color=top10['classificacao'].map({'A': 'green', 'B': 'gold', 'C': 'darkorange'})
            )
        ),
        row=2, col=2
    )
    
    # Configurações de layout
    fig.update_layout(
        height=800,
        showlegend=False,
        title_text="Análise da Curva ABC",
        title_x=0.5,
        title_font=dict(size=20),
    )
    
    # Configurações específicas para o gráfico de Pareto (duplo eixo y)
    fig.update_layout(
        yaxis=dict(title='Percentual (%)'),
        yaxis2=dict(
            title='Percentual Acumulado (%)',
            overlaying='y',
            side='right',
            range=[0, 100]
        ),
        xaxis=dict(title='Itens'),
    )
    
    # Configurações para o gráfico de barras top 10
    fig.update_layout(
        xaxis4=dict(title='Percentual (%)'),
        yaxis4=dict(title='Código SINAPI')
    )
    
    return fig

# Função para criar link de download
def get_download_link(df, filename, text):
    """Gera um link para download do DataFrame como CSV."""
    csv = df.to_csv(index=False, sep=';', decimal=',')
    b64 = base64.b64encode(csv.encode('utf-8')).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">{text}</a>'
    return href

# Função para criar link de download Excel
def get_excel_download_link(df, filename, text):
    """Gera um link para download do DataFrame como Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Curva ABC', index=False)
        
        # Formatação do Excel
        workbook = writer.book
        worksheet = writer.sheets['Curva ABC']
        
        # Formatos
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#1e3c72',
            'color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Formatar cabeçalho
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajustar largura das colunas
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_width)
    
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">{text}</a>'
    return href

# Criar sidebar
with st.sidebar:
    st.header("Configurações")
    
    st.subheader("Parâmetros da Curva ABC")
    limite_a = st.slider("Limite para Classe A (%)", 60, 90, 80, 5)
    limite_b = st.slider("Limite para Classe B (%)", limite_a + 5, 99, 95, 1)
    
    st.markdown("---")
    
    st.subheader("Sobre")
    st.info("""
    Este aplicativo foi desenvolvido para automatizar a geração de curvas ABC a partir de planilhas sintéticas do SINAPI.
    
    Funcionalidades:
    - Detecção automática de colunas
    - Agrupamento de itens com mesmo código
    - Visualizações interativas
    - Exportação em CSV e Excel
    """)

# Conteúdo principal
st.markdown('<div class="highlight">', unsafe_allow_html=True)
st.markdown("""
### Como usar:
1. Faça upload da planilha sintética do SINAPI (formato CSV ou Excel)
2. Confirme as colunas detectadas automaticamente
3. Clique em "Gerar Curva ABC"
4. Visualize os resultados e baixe o arquivo
""")
st.markdown('</div>', unsafe_allow_html=True)

# Upload do arquivo
uploaded_file = st.file_uploader("Selecione a planilha sintética do SINAPI", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    # Exibir informações do arquivo
    st.write(f"Arquivo: **{uploaded_file.name}**")
    
    with st.spinner('Processando o arquivo...'):
        # Processar o arquivo
        df, delimitador = processar_arquivo(uploaded_file)
        
        if df is not None:
            if uploaded_file.name.lower().endswith(('.xlsx', '.xls')):
                st.success(f"Arquivo Excel carregado com sucesso!")
            else:
                st.success(f"Arquivo CSV carregado com sucesso! Delimitador detectado: '{delimitador}'")
            
            # Amostra dos dados
            with st.expander("Visualizar amostra dos dados"):
                try:
                    st.dataframe(df.head(10))
                except Exception as e:
                    st.warning("Não foi possível exibir os dados em formato tabular devido a um erro de compatibilidade.")
                    st.text("Visualização alternativa dos dados:")
                    # Exibe como texto alternativo
                    for i, row in df.head(10).iterrows():
                        st.text(f"Linha {i+1}: {dict(row)}")
            
            # Identificar colunas
            col_codigo, col_descricao, col_valor = identificar_colunas(df)
            
            # Verificar se as colunas foram encontradas
            if not col_codigo or not col_descricao or not col_valor:
                st.warning("Algumas colunas não foram detectadas automaticamente. Por favor, selecione-as manualmente.")
            
            # Interface para seleção das colunas
            st.subheader("Confirme as colunas")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                col_codigo_selecionada = st.selectbox(
                    "Coluna de Código SINAPI",
                    options=df.columns,
                    index=list(df.columns).index(col_codigo) if col_codigo in df.columns else 0
                )
            
            with col2:
                col_descricao_selecionada = st.selectbox(
                    "Coluna de Descrição",
                    options=df.columns,
                    index=list(df.columns).index(col_descricao) if col_descricao in df.columns else 0
                )
            
            with col3:
                col_valor_selecionada = st.selectbox(
                    "Coluna de Valor Total",
                    options=df.columns,
                    index=list(df.columns).index(col_valor) if col_valor in df.columns else 0
                )
            
            # Exibir informações de depuração
            with st.expander("Informações de depuração - Valores selecionados"):
                st.write("**Colunas selecionadas:**")
                st.write(f"Coluna de Código: {col_codigo_selecionada}")
                st.write(f"Coluna de Descrição: {col_descricao_selecionada}")
                st.write(f"Coluna de Valor: {col_valor_selecionada}")
                
                # Mostrar amostra de valores
                if df is not None:
                    st.write("**Amostra de valores das colunas selecionadas:**")
                    try:
                        # Cria um dicionário de valores para exibir
                        amostra = {
                            'Código': df[col_codigo_selecionada].head(5).tolist(),
                            'Descrição': df[col_descricao_selecionada].head(5).tolist(),
                            'Valor': df[col_valor_selecionada].head(5).tolist()
                        }
                        
                        # Exibe como tabela Markdown em vez de dataframe
                        st.markdown("| Código | Descrição | Valor |")
                        st.markdown("|--------|-----------|-------|")
                        for i in range(min(5, len(amostra['Código']))):
                            st.markdown(f"| {amostra['Código'][i]} | {amostra['Descrição'][i]} | {amostra['Valor'][i]} |")
                    except Exception as e:
                        st.error(f"Erro ao mostrar amostra: {str(e)}")
                        import traceback
                        st.text(traceback.format_exc())
            
            # Botão para gerar a curva ABC
            if st.button("Gerar Curva ABC", key="gerar_btn"):
                with st.spinner('Gerando Curva ABC...'):
                    # Gerar a curva ABC
                    resultado, valor_total = gerar_curva_abc(
                        df, 
                        col_codigo_selecionada, 
                        col_descricao_selecionada, 
                        col_valor_selecionada,
                        limite_a,
                        limite_b
                    )
                    
                    if resultado is not None:
                        st.session_state['curva_abc'] = resultado
                        st.session_state['valor_total'] = valor_total
                        st.session_state['curva_gerada'] = True
                        
                        # Redirecionar para atualizar a página e mostrar os resultados
                        st.rerun()  # Atualizado de st.experimental_rerun()
                    else:
                        st.error("Não foi possível gerar a curva ABC. Verifique as colunas selecionadas.")

# Exibir resultados se a curva ABC foi gerada
if 'curva_gerada' in st.session_state and st.session_state['curva_gerada']:
    st.header("Resultados da Curva ABC")
    
    resultado = st.session_state['curva_abc']
    valor_total = st.session_state['valor_total']
    
    # Renomear colunas para exibição
    df_exibicao = resultado.copy()
    df_exibicao.columns = [
        'Posição', 'Código', 'Descrição', 'Valor (R$)', 
        'Percentual (%)', 'Percentual Acumulado (%)', 'Classificação'
    ]
    
    # Estatísticas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Itens", f"{len(resultado)}")
    
    with col2:
        st.metric("Valor Total", f"R$ {valor_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    
    with col3:
        classes = resultado['classificacao'].value_counts().to_dict()
        st.metric("Itens Classe A", f"{classes.get('A', 0)} ({classes.get('A', 0)/len(resultado)*100:.1f}%)")
    
    with col4:
        valor_classe_a = resultado[resultado['classificacao'] == 'A']['valor'].sum()
        st.metric("Valor Classe A", f"{valor_classe_a/valor_total*100:.1f}%")
    
    # Gráficos
    st.subheader("Análise Gráfica")
    fig = criar_graficos_plotly(resultado, valor_total)
    st.plotly_chart(fig, use_container_width=True)
    
    # Tabela de resultados
    st.subheader("Tabela de Resultados")
    
    # Adicionar filtros
    col1, col2 = st.columns(2)
    
    with col1:
        classe_filtro = st.multiselect(
            "Filtrar por Classificação",
            options=['A', 'B', 'C'],
            default=['A', 'B', 'C']
        )
    
    with col2:
        busca = st.text_input("Buscar por Código ou Descrição")
    
    # Aplicar filtros
    df_filtrado = df_exibicao[df_exibicao['Classificação'].isin(classe_filtro)]
    
    if busca:
        df_filtrado = df_filtrado[
            df_filtrado['Código'].astype(str).str.contains(busca, case=False) |
            df_filtrado['Descrição'].astype(str).str.contains(busca, case=False)
        ]
    
    # Exibir tabela
    try:
        st.dataframe(df_filtrado, height=400)
    except Exception as e:
        st.warning("Não foi possível exibir a tabela completa devido a um erro de compatibilidade.")
        st.write("Resumo dos resultados:")
        
        # Exibe estatísticas em vez da tabela completa
        st.write(f"Total de itens: {len(df_filtrado)}")
        
        # Exibe os primeiros 10 itens em formato de texto
        st.write("Primeiros 10 itens:")
        for i, row in df_filtrado.head(10).iterrows():
            st.write(f"**{row['Código']}** - {row['Descrição'][:50]}... - R$ {row['Valor (R$)']}")
    
    # Links para download
    st.subheader("Downloads")
    
    col1, col2 = st.columns(2)
    
    with col1:
        data_atual = datetime.now().strftime("%Y%m%d")
        csv_filename = f"curva_abc_{data_atual}.csv"
        st.markdown(get_download_link(df_exibicao, csv_filename, "📥 Baixar como CSV"), unsafe_allow_html=True)
    
    with col2:
        excel_filename = f"curva_abc_{data_atual}.xlsx"
        st.markdown(get_excel_download_link(df_exibicao, excel_filename, "📊 Baixar como Excel"), unsafe_allow_html=True)

# Rodapé
st.markdown('<div class="footer">', unsafe_allow_html=True)
st.markdown("© 2025 - Gerador de Curva ABC para planilhas SINAPI")
st.markdown('</div>', unsafe_allow_html=True)
