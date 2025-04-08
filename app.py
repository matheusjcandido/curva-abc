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

# Função para carregar e processar o arquivo CSV
def processar_arquivo(uploaded_file, encoding='utf-8'):
    """Carrega e processa o arquivo CSV."""
    try:
        # Ler o conteúdo do arquivo
        content = uploaded_file.getvalue().decode(encoding)
        
        # Detectar o delimitador
        delimitador = detectar_delimitador(content)
        
        # Carregar o CSV
        df = pd.read_csv(
            io.StringIO(content),
            delimiter=delimitador,
            encoding=encoding,
            error_bad_lines=False,
            warn_bad_lines=True,
            low_memory=False
        )
        
        # Limpar dados
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        return df, delimitador
    except UnicodeDecodeError:
        # Tentar com encoding alternativo se utf-8 falhar
        if encoding == 'utf-8':
            return processar_arquivo(uploaded_file, 'latin1')
        else:
            st.error(f"Erro ao decodificar o arquivo com encoding {encoding}.")
            return None, None
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
        return None, None

# Função para identificar colunas relevantes
def identificar_colunas(df):
    """Identifica as colunas de código, descrição e valor."""
    # Padrões para busca
    padroes_codigo = ['cod', 'código', 'codigo', 'referência', 'referencia', 'ref']
    padroes_descricao = ['desc', 'descrição', 'descricao', 'serviço', 'servico', 'item']
    padroes_valor = ['valor total', 'total', 'preço total', 'preco total', 'valor', 'custo']
    
    cols = df.columns
    coluna_codigo = None
    coluna_descricao = None
    coluna_valor = None
    
    # Buscar por nomes de colunas
    for col in cols:
        col_lower = str(col).lower()
        
        if not coluna_codigo and any(p in col_lower for p in padroes_codigo):
            coluna_codigo = col
        
        if not coluna_descricao and any(p in col_lower for p in padroes_descricao):
            coluna_descricao = col
        
        if not coluna_valor and any(p in col_lower for p in padroes_valor):
            coluna_valor = col
    
    # Se não encontrou pelo nome, verificar conteúdo
    if not coluna_codigo:
        for col in cols:
            valores = df[col].astype(str)
            if valores.str.match(r'^\d{5,}$').any() or valores.str.match(r'^COMP\s*\d+').any():
                coluna_codigo = col
                break
    
    if not coluna_descricao:
        for col in cols:
            if col == coluna_codigo or col == coluna_valor:
                continue
            
            valores = df[col].astype(str)
            if valores.str.len().mean() > 15:  # textos longos são provavelmente descrições
                coluna_descricao = col
                break
    
    if not coluna_valor:
        max_valor = 0
        for col in cols:
            if col == coluna_codigo or col == coluna_descricao:
                continue
            
            try:
                valores = pd.to_numeric(
                    df[col].astype(str).str.replace(r'[^\d.,]', '', regex=True).str.replace(',', '.'), 
                    errors='coerce'
                )
                if valores.max() > max_valor and not valores.isna().all():
                    max_valor = valores.max()
                    coluna_valor = col
            except:
                continue
    
    return coluna_codigo, coluna_descricao, coluna_valor

# Função para gerar a curva ABC
def gerar_curva_abc(df, coluna_codigo, coluna_descricao, coluna_valor, limite_a=80, limite_b=95):
    """Gera a curva ABC a partir do DataFrame."""
    try:
        # Criar dicionário para agrupar itens com o mesmo código
        itens_agrupados = {}
        
        # Verificar se a coluna de valor precisa ser convertida
        if df[coluna_valor].dtype == object:
            df[coluna_valor] = df[coluna_valor].astype(str).str.replace(r'[^\d.,]', '', regex=True)
            df[coluna_valor] = df[coluna_valor].str.replace(',', '.').astype(float)
        
        # Processar cada linha
        for _, row in df.iterrows():
            codigo = str(row[coluna_codigo]).strip()
            descricao = str(row[coluna_descricao]).strip()
            
            # Tentar obter o valor
            try:
                valor = float(row[coluna_valor])
            except:
                valor = 0
            
            # Verificar se temos um código válido
            if codigo and (re.match(r'^\d+$', codigo) or re.match(r'^COMP\s*\d+', codigo) or re.match(r'^COMP\d+', codigo)):
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
            return None
        
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
1. Faça upload da planilha sintética do SINAPI (formato CSV)
2. Confirme as colunas detectadas automaticamente
3. Clique em "Gerar Curva ABC"
4. Visualize os resultados e baixe o arquivo
""")
st.markdown('</div>', unsafe_allow_html=True)

# Upload do arquivo
uploaded_file = st.file_uploader("Selecione a planilha sintética do SINAPI", type=["csv"])

if uploaded_file is not None:
    # Exibir informações do arquivo
    st.write(f"Arquivo: **{uploaded_file.name}**")
    
    with st.spinner('Processando o arquivo...'):
        # Processar o arquivo
        df, delimitador = processar_arquivo(uploaded_file)
        
        if df is not None:
            st.success(f"Arquivo carregado com sucesso! Delimitador detectado: '{delimitador}'")
            
            # Amostra dos dados
            with st.expander("Visualizar amostra dos dados"):
                st.dataframe(df.head(10))
            
            # Identificar colunas
            col_codigo, col_descricao, col_valor = identificar_colunas(df)
            
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
                        st.experimental_rerun()
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
    st.dataframe(df_filtrado, height=400)
    
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
