# %% [markdown]
# # Processador de Relatórios de Alunos - IQ/UFF
# ## Versão para Google Colab

# %%
import pandas as pd
import numpy as np
from google.colab import files
import re
from datetime import datetime

# %%
# Dicionário de mapeamento de motivos de cancelamento
MOTIVOS_CANCELAMENTO = {
    'Solicitação Oficial': 'Solicitação Oficial',
    'Abandono': 'Abandono',
    'Insuficiência de Aproveitamento': 'Insuficiência de Aproveitamento',
    'Ingressante - Insuf. Aproveit.': 'Ingressante - Insuf. Aproveit.',
    'Mudança de Curso': 'Mudança de Curso'
}

def normalizar_situacao(situacao):
    """Normaliza as legendas de situação do aluno"""
    situacao = str(situacao).lower()
    if any(termo in situacao for termo in ['inscrito', 'pendente', 'concluinte']):
        if 'pendente' in situacao:
            return 'Pendente'
        elif 'concluinte' in situacao:
            return 'Concluinte'
        elif 'inscrito' in situacao:
            return 'Inscrito'
    return situacao

def classificar_modalidade(codigo):
    """Classifica a modalidade de ingresso"""
    codigo = str(codigo).strip().upper()
    if codigo.startswith('A'):
        return 'AMPLA CONCORRÊNCIA'
    elif codigo.startswith('L'):
        return 'AÇÕES AFIRMATIVAS'
    else:
        return 'NÃO CLASSIFICADO'

def classificar_curso(titulacao):
    """Classifica o curso baseado na titulação"""
    titulacao = str(titulacao).lower()
    if 'licenciatura' in titulacao:
        return 'Licenciatura Química'
    elif 'bacharel' in titulacao:
        if 'industrial' in titulacao:
            return 'Bacharel Q Industrial'
        else:
            return 'Bacharel Química'
    return 'Outros'

def extrair_periodo(data_desvinculacao):
    """Extrai período (ANO/SEMESTRE) da data de desvinculação"""
    if pd.isna(data_desvinculacao) or data_desvinculacao == '':
        return None
    data_str = str(data_desvinculacao)
    match = re.search(r'(\d{4})\s*/\s*(\d+)', data_str)
    if match:
        ano = match.group(1)
        semestre = match.group(2)
        return f"{ano}.{semestre}"
    return None

def processar_relatorio_colab():
    """Função principal para processar relatório no Colab"""
    
    print("📤 Faça upload do arquivo de relatório de alunos (.xlsx ou .xls)")
    uploaded = files.upload()
    
    if not uploaded:
        print("❌ Nenhum arquivo foi carregado!")
        return
    
    filename = list(uploaded.keys())[0]
    print(f"✅ Arquivo carregado: {filename}")
    
    # Ler o arquivo
    df = pd.read_excel(filename, header=5)
    
    # Identificar colunas
    col_situacao = df.columns[df.columns.str.contains('Situação', case=False, na=False)][0]
    col_modalidade = df.columns[df.columns.str.contains('Modalidade', case=False, na=False)][0]
    col_desvinculado = df.columns[df.columns.str.contains('Desvinculado', case=False, na=False)][0]
    col_curso = df.columns[df.columns.str.contains('Curso|Titulação', case=False, na=False)][0]
    
    # Limpar dados
    df = df.dropna(subset=[col_modalidade, col_situacao], how='all')
    df = df[~df[col_situacao].astype(str).str.contains('Alunos de', na=False)]
    
    # Aplicar normalizações
    df['SITUACAO_NORMALIZADA'] = df[col_situacao].apply(normalizar_situacao)
    df['MODALIDADE_CLASSIFICADA'] = df[col_modalidade].apply(classificar_modalidade)
    df['CURSO_CLASSIFICADO'] = df[col_curso].apply(classificar_curso)
    df['PERIODO_DESVINCULACAO'] = df[col_desvinculado].apply(extrair_periodo)
    
    # Classificar cancelamentos
    df['E_CANCELADO'] = df[col_situacao].astype(str).str.contains('Cancelamento', case=False, na=False)
    df['MOTIVO_CANCELAMENTO'] = df[col_situacao].apply(
        lambda x: next((v for k, v in MOTIVOS_CANCELAMENTO.items() if k.lower() in str(x).lower()), 'Outros')
    )
    
    # Classificar situação
    df['E_ATIVO'] = df['SITUACAO_NORMALIZADA'].isin(['Pendente', 'Inscrito', 'Concluinte'])
    df['E_TRANCADO'] = df[col_situacao].astype(str).str.contains('Trancado', case=False, na=False)
    df['E_FORMADO'] = df[col_situacao].astype(str).str.contains('Formado', case=False, na=False)
    
    return df

# %%
# Executar o processamento
print("🔄 Iniciando processamento do relatório...")
df_processado = processar_relatorio_colab()

if df_processado is not None:
    print("\n" + "="*60)
    print("📊 RESULTADOS DO PROCESSAMENTO")
    print("="*60)
    
    print(f"\n📈 Total de alunos processados: {len(df_processado)}")
    
    # Resumo por curso
    print("\n🏛️ RESUMO POR CURSO E MODALIDADE:")
    print("-"*80)
    
    resumo = df_processado.groupby(['CURSO_CLASSIFICADO', 'MODALIDADE_CLASSIFICADA']).size().unstack(fill_value=0)
    print(resumo)
    
    # Cancelamentos
    print("\n🚫 CANCELAMENTOS POR MOTIVO:")
    print("-"*80)
    
    cancelamentos = df_processado[df_processado['E_CANCELADO']]
    if len(cancelamentos) > 0:
        cancel_resumo = cancelamentos.groupby(
            ['CURSO_CLASSIFICADO', 'MODALIDADE_CLASSIFICADA', 'MOTIVO_CANCELAMENTO']
        ).size().reset_index(name='Quantidade')
        print(cancel_resumo.to_string(index=False))
    else:
        print("Nenhum cancelamento encontrado.")
    
    # Dados para planilha de evasão
    print("\n📋 DADOS PARA PLANILHA DE EVASÃO:")
    print("-"*80)
    
    periodo = input("\n📅 Digite o período de referência (ex: 2025.1): ")
    
    dados_evasao = []
    for curso in df_processado['CURSO_CLASSIFICADO'].unique():
        for modalidade in ['AMPLA CONCORRÊNCIA', 'AÇÕES AFIRMATIVAS']:
            df_filtro = df_processado[
                (df_processado['CURSO_CLASSIFICADO'] == curso) & 
                (df_processado['MODALIDADE_CLASSIFICADA'] == modalidade)
            ]
            
            if len(df_filtro) > 0:
                cancel_periodo = df_filtro[
                    (df_filtro['E_CANCELADO']) & 
                    (df_filtro['PERIODO_DESVINCULACAO'] == periodo)
                ]
                
                matriculas_ativas = df_filtro[
                    (df_filtro['E_ATIVO']) & 
                    (~df_filtro['E_CANCELADO']) & 
                    (~df_filtro['E_FORMADO'])
                ]
                
                dados_evasao.append({
                    'Período': periodo,
                    'Curso': curso,
                    'Modalidade': modalidade,
                    'Ingressantes': len(df_filtro),
                    'Cancelamentos': len(cancel_periodo),
                    'Matrículas Ativas': len(matriculas_ativas) + len(df_filtro[df_filtro['E_TRANCADO']]),
                    'Formados': len(df_filtro[df_filtro['E_FORMADO']]),
                })
    
    df_evasao = pd.DataFrame(dados_evasao)
    print(df_evasao.to_string(index=False))
    
    # Resumo consolidado
    print("\n🎯 RESUMO CONSOLIDADO PARA PLANILHA PRINCIPAL:")
    print("-"*80)
    
    consolidado = pd.DataFrame([{
        'Período': periodo,
        'Ingressantes': df_evasao['Ingressantes'].sum(),
        'Cancelamentos': df_evasao['Cancelamentos'].sum(),
        'Matrículas Ativas': df_evasao['Matrículas Ativas'].sum(),
        'Formados': df_evasao['Formados'].sum(),
        '% Evasão': (df_evasao['Cancelamentos'].sum() / df_evasao['Ingressantes'].sum() * 100) if df_evasao['Ingressantes'].sum() > 0 else 0
    }])
    
    print(consolidado.to_string(index=False))
    
    # Download dos resultados
    print("\n💾 Gerando arquivo de resultados...")
    
    output_filename = f"resultados_evasao_{periodo.replace('.', '_')}.xlsx"
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_evasao.to_excel(writer, sheet_name='Detalhado', index=False)
        consolidado.to_excel(writer, sheet_name=f'Resumo_{periodo}', index=False)
        
        if len(cancelamentos) > 0:
            cancel_resumo.to_excel(writer, sheet_name='Cancelamentos', index=False)
    
    files.download(output_filename)
    print(f"✅ Arquivo '{output_filename}' gerado e download iniciado!")
    
    print("\n✨ Processamento concluído com sucesso!")
