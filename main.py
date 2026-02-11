"""
app.py - Processador de Relatórios SISU - IQ/UFF
Versão ultra simplificada para Streamlit Cloud
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

# Configuração da página - PRIMEIRO COMANDO
st.set_page_config(
    page_title="Processador SISU - IQ/UFF",
    page_icon="🧪",
    layout="wide"
)

# CSS simples
st.markdown("""
<style>
    .stApp {
        background-color: #f0f2f6;
    }
    .main-title {
        background: linear-gradient(90deg, #1e3c72, #2a5298);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        padding: 15px;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

def processar_arquivo(df, periodo):
    """Processa o arquivo e extrai as informações"""
    
    resultados = {
        'total_alunos': 0,
        'ativos': 0,
        'cancelados': 0,
        'trancados': 0,
        'formados': 0,
        'por_curso': {},
        'cancelamentos_periodo': 0
    }
    
    try:
        # Encontrar colunas
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
        
        if 'situacao' not in colunas or 'modalidade' not in colunas:
            return None, "Colunas obrigatórias não encontradas"
        
        # Limpar dados
        df = df.dropna(subset=[colunas['modalidade']], how='all')
        df = df[~df[colunas['situacao']].astype(str).str.contains('Alunos de|Total', na=False, case=False)]
        
        resultados['total_alunos'] = len(df)
        
        # Processar cada linha
        for idx, row in df.iterrows():
            situacao = str(row[colunas['situacao']]).lower()
            modalidade = str(row[colunas['modalidade']]).upper()
            curso = str(row.get(colunas.get('curso', ''), '')).lower() if 'curso' in colunas else 'bacharel química'
            
            # Classificar curso
            if 'licenciatura' in curso:
                curso_nome = 'Licenciatura Química'
            elif 'bacharel' in curso:
                if 'industrial' in curso:
                    curso_nome = 'Bacharel Q Industrial'
                else:
                    curso_nome = 'Bacharel Química'
            else:
                curso_nome = 'Bacharel Química'  # padrão
            
            # Classificar modalidade
            if modalidade.startswith('A'):
                modalidade_nome = 'AMPLA CONCORRÊNCIA'
            elif modalidade.startswith('L'):
                modalidade_nome = 'AÇÕES AFIRMATIVAS'
            else:
                modalidade_nome = 'OUTROS'
            
            # Contabilizar matrículas ativas (Pendente, Inscrito, Concluinte)
            if any(x in situacao for x in ['pendente', 'inscrito', 'concluinte']):
                resultados['ativos'] += 1
            
            # Contabilizar trancados
            if 'trancado' in situacao:
                resultados['trancados'] += 1
            
            # Contabilizar formados
            if 'formado' in situacao:
                resultados['formados'] += 1
            
            # Contabilizar cancelamentos
            if 'cancelamento' in situacao:
                resultados['cancelados'] += 1
                
                # Verificar se é do período atual
                if 'desvinculado' in colunas and pd.notna(row[colunas['desvinculado']]):
                    data = str(row[colunas['desvinculado']])
                    if periodo.replace('.', '/') in data or periodo in data:
                        resultados['cancelamentos_periodo'] += 1
            
            # Agrupar por curso e modalidade
            chave = f"{curso_nome}|{modalidade_nome}"
            if chave not in resultados['por_curso']:
                resultados['por_curso'][chave] = {
                    'curso': curso_nome,
                    'modalidade': modalidade_nome,
                    'ingressantes': 0,
                    'cancelamentos': 0,
                    'ativos': 0,
                    'trancados': 0,
                    'formados': 0
                }
            
            resultados['por_curso'][chave]['ingressantes'] += 1
            
            if any(x in situacao for x in ['pendente', 'inscrito', 'concluinte']):
                resultados['por_curso'][chave]['ativos'] += 1
            
            if 'trancado' in situacao:
                resultados['por_curso'][chave]['trancados'] += 1
            
            if 'formado' in situacao:
                resultados['por_curso'][chave]['formados'] += 1
            
            if 'cancelamento' in situacao:
                resultados['por_curso'][chave]['cancelamentos'] += 1
        
        return resultados, None
        
    except Exception as e:
        return None, str(e)

def main():
    # Header
    st.markdown("""
    <div class="main-title">
        <h1 style="color: white; margin:0;">🧪 Processador de Relatórios SISU</h1>
        <p style="color: white; margin:0; opacity:0.9;">Instituto de Química - UFF</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📁 Upload do Relatório")
        
        uploaded_file = st.file_uploader(
            "Carregar arquivo Excel",
            type=['xlsx', 'xls'],
            help="Arquivo com a listagem de alunos"
        )
        
        # Período
        periodo = st.text_input(
            "Período (ex: 2025.1)",
            value="2025.1"
        )
        
        processar = st.button(
            "🚀 Processar",
            type="primary",
            use_container_width=True,
            disabled=uploaded_file is None
        )
        
        st.markdown("---")
        st.markdown("""
        ### 📌 Instruções
        1. Upload do relatório
        2. Confirmar período
        3. Processar
        4. Copiar dados
        """)
    
    # Área principal
    if uploaded_file and processar:
        with st.spinner("🔄 Processando..."):
            try:
                # Tentar diferentes headers
                df = None
                for header in [5, 4, 6, 0]:
                    try:
                        df = pd.read_excel(uploaded_file, header=header)
                        if len(df.columns) > 3:
                            break
                    except:
                        continue
                
                if df is None:
                    st.error("❌ Não foi possível ler o arquivo. Verifique o formato.")
                    return
                
                # Processar
                resultados, erro = processar_arquivo(df, periodo)
                
                if erro:
                    st.error(f"❌ Erro: {erro}")
                    return
                
                # MÉTRICAS PRINCIPAIS
                st.success("✅ Processamento concluído!")
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("Total Alunos", resultados['total_alunos'])
                with col2:
                    st.metric("Matrículas Ativas", resultados['ativos'])
                with col3:
                    st.metric("Cancelamentos", resultados['cancelados'])
                with col4:
                    st.metric("Cancelamentos Período", resultados['cancelamentos_periodo'])
                with col5:
                    st.metric("Formados", resultados['formados'])
                
                # TABELA POR CURSO
                st.markdown("### 📊 Dados por Curso e Modalidade")
                
                dados_tabela = []
                for chave, dados in resultados['por_curso'].items():
                    dados_tabela.append({
                        'Curso': dados['curso'],
                        'Modalidade': dados['modalidade'],
                        'Ingressantes': dados['ingressantes'],
                        'Ativos': dados['ativos'],
                        'Trancados': dados['trancados'],
                        'Cancelados': dados['cancelamentos'],
                        'Formados': dados['formados']
                    })
                
                df_tabela = pd.DataFrame(dados_tabela)
                st.dataframe(df_tabela, use_container_width=True)
                
                # DADOS PARA PLANILHA DE EVASÃO
                st.markdown("### 📋 Dados para Planilha de Evasão")
                
                # Consolidado
                total_ingressantes = df_tabela['Ingressantes'].sum()
                total_cancelamentos = resultados['cancelamentos_periodo']
                total_ativos = df_tabela['Ativos'].sum() + df_tabela['Trancados'].sum()
                total_formados = df_tabela['Formados'].sum()
                taxa_evasao = (total_cancelamentos / total_ingressantes * 100) if total_ingressantes > 0 else 0
                
                consolidado = pd.DataFrame([{
                    'Período': periodo,
                    'Ingressantes': total_ingressantes,
                    'Cancelamentos': total_cancelamentos,
                    'Matrículas Ativas': total_ativos,
                    'Formados': total_formados,
                    '% Evasão': round(taxa_evasao, 2)
                }])
                
                st.dataframe(consolidado, use_container_width=True, hide_index=True)
                
                # INSTRUÇÕES
                st.markdown(f"""
                <div class="success-box">
                    <h4 style="margin-top:0;">📌 Como atualizar a planilha principal</h4>
                    <p><strong>Período: {periodo}</strong></p>
                    <table style="width:100%; border-collapse: collapse;">
                        <tr style="background-color: #28a745; color: white;">
                            <th style="padding: 8px; text-align: left;">Campo</th>
                            <th style="padding: 8px; text-align: left;">Valor</th>
                            <th style="padding: 8px; text-align: left;">Local na Planilha</th>
                        </tr>
                        <tr style="background-color: white;">
                            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Ingressantes</strong></td>
                            <td style="padding: 8px; border: 1px solid #ddd;">{total_ingressantes}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">Linha "Total de Ingressantes"</td>
                        </tr>
                        <tr style="background-color: #f9f9f9;">
                            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Cancelamentos</strong></td>
                            <td style="padding: 8px; border: 1px solid #ddd;">{total_cancelamentos}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">Linha "TOTAL CANCELAMENTOS"</td>
                        </tr>
                        <tr style="background-color: white;">
                            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Matrículas Ativas</strong></td>
                            <td style="padding: 8px; border: 1px solid #ddd;">{total_ativos}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">Linha "TOTAL MATRIC. ATIVAS"</td>
                        </tr>
                        <tr style="background-color: #f9f9f9;">
                            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Formados</strong></td>
                            <td style="padding: 8px; border: 1px solid #ddd;">{total_formados}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">Linha "Alunos Formados"</td>
                        </tr>
                        <tr style="background-color: white;">
                            <td style="padding: 8px; border: 1px solid #ddd;"><strong>% Evasão</strong></td>
                            <td style="padding: 8px; border: 1px solid #ddd;">{round(taxa_evasao, 2)}%</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">Linha "% Cancelamento"</td>
                        </tr>
                    </table>
                </div>
                """, unsafe_allow_html=True)
                
                # Download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_tabela.to_excel(writer, sheet_name='Detalhado', index=False)
                    consolidado.to_excel(writer, sheet_name=f'Resumo_{periodo}', index=False)
                
                st.download_button(
                    label="⬇️ Download Planilha",
                    data=output.getvalue(),
                    file_name=f"dados_evasao_{periodo.replace('.', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")
                st.info("""
                **Dicas:**
                - Verifique se o arquivo está no formato correto
                - As colunas de Situação e Modalidade são obrigatórias
                - Tente usar um arquivo com menos linhas
                """)

if __name__ == "__main__":
    main()
