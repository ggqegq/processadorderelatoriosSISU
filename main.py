"""
app.py - Processador de Relatórios SISU - IQ/UFF
Versão simplificada e compatível com Python 3.11
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re

# Configuração da página - DEVE SER O PRIMEIRO COMANDO STREAMLIT
st.set_page_config(
    page_title="Processador SISU - IQ/UFF",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
    }
    .stApp {
        background-color: #f8f9fa;
    }
    .info-box {
        background-color: #e7f3fe;
        border-left: 6px solid #2196F3;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 6px solid #28a745;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Dicionário de motivos de cancelamento
MOTIVOS_CANCELAMENTO = {
    'Solicitação Oficial': 'Solicitação Oficial',
    'Abandono': 'Abandono',
    'Insuficiência de Aproveitamento': 'Insuficiência de Aproveitamento',
    'Ingressante - Insuf. Aproveit.': 'Ingressante - Insuf. Aproveit.',
    'Mudança de Curso': 'Mudança de Curso'
}

def normalizar_situacao(situacao):
    """Normaliza as legendas de situação do aluno"""
    if pd.isna(situacao):
        return 'Não informado'
    
    situacao = str(situacao).lower().strip()
    
    if 'pendente' in situacao:
        return 'Pendente'
    elif 'inscrito' in situacao:
        return 'Inscrito'
    elif 'concluinte' in situacao:
        return 'Concluinte'
    elif 'trancado' in situacao:
        return 'Trancado'
    elif 'formado' in situacao:
        return 'Formado'
    elif 'cancelamento' in situacao:
        return 'Cancelado'
    else:
        return situacao.capitalize()

def classificar_modalidade(codigo):
    """Classifica a modalidade de ingresso"""
    if pd.isna(codigo):
        return 'NÃO CLASSIFICADO'
    
    codigo = str(codigo).strip().upper()
    
    if codigo.startswith('A'):
        return 'AMPLA CONCORRÊNCIA'
    elif codigo.startswith('L'):
        return 'AÇÕES AFIRMATIVAS'
    else:
        return 'OUTRAS MODALIDADES'

def classificar_curso(titulacao):
    """Classifica o curso baseado na titulação"""
    if pd.isna(titulacao):
        return 'Não identificado'
    
    titulacao = str(titulacao).lower()
    
    if 'licenciatura' in titulacao:
        return 'Licenciatura Química'
    elif 'bacharel' in titulacao:
        if 'industrial' in titulacao:
            return 'Bacharel Q Industrial'
        else:
            return 'Bacharel Química'
    else:
        return 'Outros'

def classificar_motivo_cancelamento(situacao):
    """Classifica o motivo do cancelamento"""
    if pd.isna(situacao):
        return 'Outros'
    
    situacao_lower = str(situacao).lower()
    
    for motivo in MOTIVOS_CANCELAMENTO.values():
        if motivo.lower() in situacao_lower:
            return motivo
    
    return 'Outros'

def extrair_periodo(data_desvinculacao):
    """Extrai período (ANO/SEMESTRE) da data de desvinculação"""
    if pd.isna(data_desvinculacao):
        return None
    
    data_str = str(data_desvinculacao)
    match = re.search(r'(\d{4})\s*/\s*(\d+)[º°]?', data_str)
    if match:
        return f"{match.group(1)}.{match.group(2)}"
    return None

def processar_relatorio(df, periodo):
    """Processa o DataFrame do relatório"""
    
    # Identificar colunas
    colunas = {}
    for col in df.columns:
        col_lower = str(col).lower()
        if 'situação' in col_lower or 'situacao' in col_lower:
            colunas['situacao'] = col
        elif 'modalidade' in col_lower:
            colunas['modalidade'] = col
        elif 'desvinculado' in col_lower:
            colunas['desvinculado'] = col
        elif 'curso' in col_lower or 'titulação' in col_lower:
            colunas['curso'] = col
        elif 'matrícula' in col_lower or 'matricula' in col_lower:
            colunas['matricula'] = col
    
    # Verificar colunas essenciais
    if 'situacao' not in colunas or 'modalidade' not in colunas:
        st.error("❌ Colunas obrigatórias não encontradas. Verifique o formato do arquivo.")
        return None
    
    # Limpar dados
    df = df.dropna(subset=[colunas['modalidade']], how='all')
    df = df[~df[colunas['situacao']].astype(str).str.contains('Alunos de|Total|Resumo', na=False, case=False)]
    
    # Aplicar classificações
    df['SITUACAO_NORMALIZADA'] = df[colunas['situacao']].apply(normalizar_situacao)
    df['MODALIDADE_CLASSIFICADA'] = df[colunas['modalidade']].apply(classificar_modalidade)
    df['CURSO_CLASSIFICADO'] = df[colunas['curso']].apply(classificar_curso) if 'curso' in colunas else 'Não identificado'
    
    # Data de desvinculação
    if 'desvinculado' in colunas:
        df['PERIODO_DESVINCULACAO'] = df[colunas['desvinculado']].apply(extrair_periodo)
    else:
        df['PERIODO_DESVINCULACAO'] = None
    
    # Classificações adicionais
    df['E_CANCELADO'] = df[colunas['situacao']].astype(str).str.contains('Cancelamento', case=False, na=False)
    df['MOTIVO_CANCELAMENTO'] = df[colunas['situacao']].apply(classificar_motivo_cancelamento)
    df['E_TRANCADO'] = df[colunas['situacao']].astype(str).str.contains('Trancado', case=False, na=False)
    df['E_FORMADO'] = df[colunas['situacao']].astype(str).str.contains('Formado', case=False, na=False)
    
    # Matrículas ativas
    df['E_ATIVO'] = df['SITUACAO_NORMALIZADA'].isin(['Inscrito', 'Pendente', 'Concluinte'])
    
    # Adicionar período como coluna
    df['PERIODO_REFERENCIA'] = periodo
    
    return df

def gerar_dados_evasao(df, periodo):
    """Gera dados para planilha de evasão"""
    
    if df is None or df.empty:
        return None
    
    dados = []
    
    for curso in df['CURSO_CLASSIFICADO'].unique():
        for modalidade in ['AMPLA CONCORRÊNCIA', 'AÇÕES AFIRMATIVAS']:
            
            df_filtro = df[
                (df['CURSO_CLASSIFICADO'] == curso) & 
                (df['MODALIDADE_CLASSIFICADA'] == modalidade)
            ]
            
            if len(df_filtro) == 0:
                continue
            
            # Cancelamentos do período
            cancelamentos = df_filtro[
                (df_filtro['E_CANCELADO']) & 
                (df_filtro['PERIODO_DESVINCULACAO'] == periodo)
            ]
            
            # Matrículas ativas
            matriculas_ativas = df_filtro[
                (df_filtro['E_ATIVO']) & 
                (~df_filtro['E_CANCELADO']) & 
                (~df_filtro['E_FORMADO'])
            ]
            
            # Contagem por motivo
            cancel_por_motivo = {}
            for motivo in MOTIVOS_CANCELAMENTO.values():
                count = len(cancelamentos[cancelamentos['MOTIVO_CANCELAMENTO'] == motivo])
                cancel_por_motivo[motivo] = count
            
            dados.append({
                'Curso': curso,
                'Modalidade': modalidade,
                'Ingressantes': len(df_filtro),
                'Solicitação Oficial': cancel_por_motivo['Solicitação Oficial'],
                'Abandono': cancel_por_motivo['Abandono'],
                'Insuf. Aproveitamento': cancel_por_motivo['Insuficiência de Aproveitamento'],
                'Ingressante - Insuf.': cancel_por_motivo['Ingressante - Insuf. Aproveit.'],
                'Mudança de Curso': cancel_por_motivo['Mudança de Curso'],
                'Total Cancelamentos': len(cancelamentos),
                'Matrículas Ativas': len(matriculas_ativas) + len(df_filtro[df_filtro['E_TRANCADO']]),
                'Trancados': len(df_filtro[df_filtro['E_TRANCADO']]),
                'Formados': len(df_filtro[df_filtro['E_FORMADO']])
            })
    
    return pd.DataFrame(dados)

def main():
    """Função principal"""
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 style="color: white; margin: 0;">🧪 Processador de Relatórios SISU</h1>
        <p style="color: white; margin: 0; opacity: 0.9;">Instituto de Química - Universidade Federal Fluminense</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📁 Upload do Relatório")
        
        uploaded_file = st.file_uploader(
            "Carregar arquivo de alunos",
            type=['xlsx', 'xls'],
            help="Arquivo Excel com a listagem de alunos"
        )
        
        # Período de referência
        ano_atual = datetime.now().year
        semestre_atual = 1 if datetime.now().month <= 6 else 2
        
        periodo = st.text_input(
            "Período de referência",
            value=f"{ano_atual}.{semestre_atual}",
            help="Formato: AAAA.S (ex: 2025.1)"
        )
        
        processar = st.button(
            "🚀 Processar Relatório",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None
        )
        
        st.markdown("---")
        st.markdown("""
        ### 📌 Instruções
        1. Faça upload do relatório
        2. Confirme o período
        3. Clique em Processar
        4. Copie os dados gerados
        """)
    
    # Área principal
    if uploaded_file and processar:
        try:
            with st.spinner("🔄 Processando relatório..."):
                # Ler arquivo
                df = pd.read_excel(uploaded_file, header=5)
                
                # Processar
                df_processado = processar_relatorio(df, periodo)
                
                if df_processado is not None:
                    st.success("✅ Relatório processado com sucesso!")
                    
                    # Métricas
                    col1, col2, col3, col4, col5 = st.columns(5)
                    
                    with col1:
                        st.metric("Total de Alunos", len(df_processado))
                    with col2:
                        st.metric("Matrículas Ativas", df_processado['E_ATIVO'].sum())
                    with col3:
                        st.metric("Cancelamentos", df_processado['E_CANCELADO'].sum())
                    with col4:
                        st.metric("Trancados", df_processado['E_TRANCADO'].sum())
                    with col5:
                        st.metric("Formados", df_processado['E_FORMADO'].sum())
                    
                    # Resumo por curso
                    st.markdown("### 📊 Resumo por Curso")
                    
                    resumo = df_processado.groupby(['CURSO_CLASSIFICADO', 'MODALIDADE_CLASSIFICADA']).size().reset_index(name='Quantidade')
                    st.dataframe(resumo, use_container_width=True)
                    
                    # Dados para evasão
                    st.markdown("### 📋 Dados para Planilha de Evasão")
                    
                    dados_evasao = gerar_dados_evasao(df_processado, periodo)
                    
                    if dados_evasao is not None and not dados_evasao.empty:
                        st.dataframe(dados_evasao, use_container_width=True)
                        
                        # Consolidado
                        st.markdown("### 🎯 Consolidado")
                        
                        consolidado = pd.DataFrame([{
                            'Período': periodo,
                            'Total Ingressantes': dados_evasao['Ingressantes'].sum(),
                            'Total Cancelamentos': dados_evasao['Total Cancelamentos'].sum(),
                            'Total Matrículas Ativas': dados_evasao['Matrículas Ativas'].sum(),
                            'Total Formados': dados_evasao['Formados'].sum(),
                            '% Evasão': round(
                                (dados_evasao['Total Cancelamentos'].sum() / dados_evasao['Ingressantes'].sum() * 100)
                                if dados_evasao['Ingressantes'].sum() > 0 else 0, 2
                            )
                        }])
                        
                        st.dataframe(consolidado, use_container_width=True, hide_index=True)
                        
                        # Instruções
                        st.markdown(f"""
                        <div class="info-box">
                            <h4 style="margin-top: 0;">📌 Como atualizar a planilha principal</h4>
                            <p><strong>Período: {periodo}</strong></p>
                            <ol>
                                <li>Abra a planilha "Cópia de Evasão Cursos de Química IQ_SISU_versão2025_.xlsx"</li>
                                <li>Vá para a aba "Acumulado de 2025.1 a 2015.1"</li>
                                <li>Localize a coluna do período <strong>{periodo}</strong></li>
                                <li>Copie os valores:</li>
                                <ul>
                                    <li><strong>Ingressantes:</strong> {consolidado['Total Ingressantes'].values[0]}</li>
                                    <li><strong>Cancelamentos:</strong> {consolidado['Total Cancelamentos'].values[0]}</li>
                                    <li><strong>Matrículas Ativas:</strong> {consolidado['Total Matrículas Ativas'].values[0]}</li>
                                    <li><strong>Formados:</strong> {consolidado['Total Formados'].values[0]}</li>
                                    <li><strong>% Evasão:</strong> {consolidado['% Evasão'].values[0]}%</li>
                                </ul>
                            </ol>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Botão download
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            dados_evasao.to_excel(writer, sheet_name='Dados_Evasao', index=False)
                            consolidado.to_excel(writer, sheet_name=f'Resumo_{periodo}', index=False)
                        
                        st.download_button(
                            label="⬇️ Download Planilha Processada",
                            data=output.getvalue(),
                            file_name=f"dados_evasao_{periodo.replace('.', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ Nenhum dado de evasão encontrado para o período especificado.")
                
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            st.info("""
            **Dicas:**
            - Verifique se o arquivo está no formato correto
            - Confirme se o cabeçalho está na linha 6
            - As colunas de Situação e Modalidade são obrigatórias
            """)

if __name__ == "__main__":
    main()
