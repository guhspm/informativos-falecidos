import pandas as pd
import glob
import re
import sys
import numpy as np
import calendar
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

def clean_cpf(cpf):
    if pd.isna(cpf): return ''
    if isinstance(cpf, float): cpf = str(int(cpf))
    return re.sub(r'\D', '', str(cpf)).zfill(11)

# ==========================================
# ATENÇÃO: COLOQUE AQUI O VALOR FIXO DO AERO
VALOR_FIXO_AERO = 2.07 
# ==========================================

arquivo_oficial = 'INFORMATIVOS_FALECIDOS.2026.xlsx'

print("Iniciando a Automação: Preenchendo a Planilha Original...")

try:
    open(arquivo_oficial, 'r+').close()
except PermissionError:
    print(f"\n❌ ERRO: A planilha '{arquivo_oficial}' está ABERTA no Excel!")
    print("Por favor, feche o arquivo no Excel e rode o script novamente.")
    sys.exit()

df_emissao = pd.read_excel(arquivo_oficial, sheet_name='EMISSÃO', skiprows=1)
df_emissao['CPF_CLEAN'] = df_emissao['CPF'].apply(clean_cpf)

colunas_emissao = df_emissao.columns.tolist()

# ==========================================
# PARTE A: OS PROCVs DOS DÉBITOS
# ==========================================
print("Buscando Passivo ZETRA (Até Nov/2021)...")
try:
    arquivo_zetra = glob.glob('CONTROLE CONSOLIDADO PASSIVO GESTORA ZETRA*.xlsx')[0]
    df_zetra = pd.read_excel(arquivo_zetra, sheet_name='INADIMPLÊNCIA - PBH', skiprows=1)
    df_zetra['CPF_CLEAN'] = df_zetra['CPF'].apply(clean_cpf)
    df_zetra['SALDO FINAL'] = pd.to_numeric(df_zetra['SALDO FINAL'], errors='coerce').fillna(0)
    zetra_grp = df_zetra.groupby('CPF_CLEAN', as_index=False)['SALDO FINAL'].sum()
    df_emissao = pd.merge(df_emissao, zetra_grp, on='CPF_CLEAN', how='left')
    col_zetra = [c for c in colunas_emissao if 'Nov/2021' in str(c)][0]
    df_emissao[col_zetra] = df_emissao['SALDO FINAL'].fillna(0)
    df_emissao.drop(columns=['SALDO FINAL'], inplace=True)
except Exception as e: print("Aviso: Não encontrei base Zetra ou houve erro:", e)

print("Buscando Passivo Consolidado (UNIMED 12/2021 a ATUAL)...")
try:
    arquivo_passivo = glob.glob('CONTROLE CONSOLIDADO DE INADIMPLENTES*.xlsx')[0]
    df_passivo = pd.read_excel(arquivo_passivo, sheet_name='UNIMED', skiprows=1)
    df_passivo['CPF_CLEAN'] = df_passivo['CPF'].apply(clean_cpf)
    passivo_grp = df_passivo.groupby('CPF_CLEAN', as_index=False)['SALDO CONSOLIDADO2'].sum()
    df_emissao = pd.merge(df_emissao, passivo_grp, on='CPF_CLEAN', how='left')
    col_passivo = [c for c in colunas_emissao if '12/2021' in str(c)][0]
    df_emissao[col_passivo] = df_emissao['SALDO CONSOLIDADO2'].fillna(0)
    df_emissao.drop(columns=['SALDO CONSOLIDADO2'], inplace=True)
except: print("Sem atualizações de Passivo Unimed.")

print("Buscando Boletos Em Aberto...")
for arquivo in glob.glob('PBH - MNC BOLETO*.xlsx'):
    try:
        match = re.search(r'BOLETO (\d{2})\.(\d{4})', arquivo)
        if match:
            mes, ano = match.groups()
            ref_str = f"ref. {mes}/{ano[-2:]}"
            col_destino = [c for c in colunas_emissao if 'Boleto' in str(c) and ref_str in str(c)]
            if col_destino:
                col_destino = col_destino[0]
                abas = pd.ExcelFile(arquivo).sheet_names
                aba_alvo = [a for a in abas if 'MNC' in a.upper()][0]
                df_bol = pd.read_excel(arquivo, sheet_name=aba_alvo)
                df_bol['CPF_CLEAN'] = df_bol['CPF Titular'].apply(clean_cpf)
                bol_grp = df_bol.groupby('CPF_CLEAN', as_index=False)['TOTAL GERAL'].sum()
                df_emissao = pd.merge(df_emissao, bol_grp, on='CPF_CLEAN', how='left')
                df_emissao[col_destino] = df_emissao['TOTAL GERAL'].fillna(0)
                df_emissao.drop(columns=['TOTAL GERAL'], inplace=True)
    except: pass

print("Buscando Enviados à Folha...")
for arquivo in glob.glob('PBH - ENVIADO À FOLHA*.xlsx'):
    try:
        match = re.search(r'FOLHA (\d{2})\.(\d{4})', arquivo)
        if match:
            mes, ano = match.groups()
            ref_str = f"ref. {mes}/{ano[-2:]}"
            col_destino = [c for c in colunas_emissao if ('Valores' in str(c) or 'plano' in str(c)) and ref_str in str(c)]
            if col_destino:
                col_destino = col_destino[0]
                abas = pd.ExcelFile(arquivo).sheet_names
                abas_validas = [a for a in abas if 'CONSOLIDADO' in a.upper() or 'FOLHA' in a.upper()]
                aba_alvo = abas_validas[0] if abas_validas else abas[0]
                df_folha = pd.read_excel(arquivo, sheet_name=aba_alvo)
                df_folha['CPF_CLEAN'] = df_folha['CPF'].apply(clean_cpf)
                folha_grp = df_folha.groupby('CPF_CLEAN', as_index=False)['VALOR'].sum()
                df_emissao = pd.merge(df_emissao, folha_grp, on='CPF_CLEAN', how='left')
                df_emissao[col_destino] = df_emissao['VALOR'].fillna(0)
                df_emissao.drop(columns=['VALOR'], inplace=True)
    except Exception as e: 
        print(f"Aviso: Erro ao ler a Folha {arquivo} -> {e}")
        
cols_debito = [c for c in colunas_emissao if 'Débito passivo' in str(c) or 'Boleto' in str(c) or 'Valores de plano' in str(c)]
df_emissao[' TOTAL DÉBITOS'] = df_emissao[cols_debito].sum(axis=1)

# ==========================================
# PARTE B: CÁLCULOS AVANÇADOS (MÊS A MÊS + AERO)
# ==========================================
print("Lendo Demonstrativos para Rateio Mês a Mês...")
df_emissao['DATA_OBITO'] = pd.to_datetime(df_emissao['DATA ÓBITO'], errors='coerce', dayfirst=True)
df_emissao['DATA_CANCELAMENTO'] = pd.to_datetime(df_emissao['DATA CANCELAMENTO'], errors='coerce', dayfirst=True)

xls_dem = pd.ExcelFile('DEMONSTRATIVOS (POR ANO).xlsx')
abas_anos = [aba for aba in xls_dem.sheet_names if str(aba).isdigit() and len(str(aba)) == 4]
df_dem = pd.concat([pd.read_excel(xls_dem, sheet_name=aba) for aba in abas_anos], ignore_index=True)
df_dem['CPF_CLEAN'] = df_dem['CPF BENEFICIÁRIO'].apply(clean_cpf)
df_dem['COMPETÊNCIA'] = pd.to_datetime(df_dem['COMPETÊNCIA'], errors='coerce')
df_dem = df_dem[df_dem['TIPO LANÇAMENTO'].str.upper().str.contains('MENSALIDADE', na=False)]
df_dem['CLASSIFICAÇÃO'] = df_dem['CLASSIFICAÇÃO'].str.upper().str.replace('Ú', 'U')

df_dem_grp = df_dem.groupby(['CPF_CLEAN', 'COMPETÊNCIA', 'CLASSIFICAÇÃO'], as_index=False)[['VALOR OPERADORA', 'SUBSDIO']].sum()
df_dem_pivot = df_dem_grp.pivot(index=['CPF_CLEAN', 'COMPETÊNCIA'], columns='CLASSIFICAÇÃO', values=['VALOR OPERADORA', 'SUBSDIO']).reset_index()
df_dem_pivot.columns = [f"{col[0]}_{col[1]}" if col[1] else col[0] for col in df_dem_pivot.columns]
df_dem_pivot.rename(columns={
    'VALOR OPERADORA_SAUDE': 'MENSALIDADE_SAUDE', 'SUBSDIO_SAUDE': 'SUBSIDIO_SAUDE',
    'VALOR OPERADORA_ODONTO': 'MENSALIDADE_ODONTO', 'SUBSDIO_ODONTO': 'SUBSIDIO_ODONTO'
}, inplace=True)

def calcular_rateio_mes_a_mes(row):
    cpf = row['CPF_CLEAN']
    dt_obito = row['DATA_OBITO']
    dt_exclusao = row['DATA_CANCELAMENTO']
    tem_aero = str(row.get('AERO', 'NÃO')).strip().upper() == 'SIM'
    
    if pd.isna(dt_obito) or pd.isna(dt_exclusao) or dt_exclusao <= dt_obito:
        return pd.Series({'ESPERADO_SAUDE': 0.0, 'PCT_PBH_SAUDE': 0.0, 'PCT_PBH_ODONTO': 0.0, 'PERIODO_SUBSIDIADO_STR': '-'})
        
    mes_atual = dt_obito.replace(day=1)
    mes_fim = dt_exclusao.replace(day=1)
    df_dem_cpf = df_dem_pivot[df_dem_pivot['CPF_CLEAN'] == cpf]
    
    esp_saude_total = 0.0
    esp_odonto_total = 0.0
    pbh_saude_acumulado = 0.0
    pbh_odonto_acumulado = 0.0
    meses_subsidiados = [] # NOVA LÓGICA: Lista para guardar os meses em que a PBH de fato subsidiou
    
    while mes_atual <= mes_fim:
        dem_mes = df_dem_cpf[df_dem_cpf['COMPETÊNCIA'] == mes_atual]
        
        if not dem_mes.empty:
            # Pessoa está no arquivo, pega as informações reais
            m_saude = float(dem_mes['MENSALIDADE_SAUDE'].iloc[0]) if pd.notna(dem_mes['MENSALIDADE_SAUDE'].iloc[0]) else 0
            s_saude = float(dem_mes['SUBSIDIO_SAUDE'].iloc[0]) if pd.notna(dem_mes['SUBSIDIO_SAUDE'].iloc[0]) else 0
            m_odonto = float(dem_mes['MENSALIDADE_ODONTO'].iloc[0]) if pd.notna(dem_mes['MENSALIDADE_ODONTO'].iloc[0]) else 0
            s_odonto = float(dem_mes['SUBSIDIO_ODONTO'].iloc[0]) if pd.notna(dem_mes['SUBSIDIO_ODONTO'].iloc[0]) else 0
        else:
            # Pessoa sumiu do demonstrativo (PBH cortou). 
            # Pega o valor do plano na época do óbito (para auditar a Unimed)
            dem_obito = df_dem_cpf[df_dem_cpf['COMPETÊNCIA'] == dt_obito.replace(day=1)]
            if not dem_obito.empty:
                m_saude = float(dem_obito['MENSALIDADE_SAUDE'].iloc[0]) if pd.notna(dem_obito['MENSALIDADE_SAUDE'].iloc[0]) else 0
                m_odonto = float(dem_obito['MENSALIDADE_ODONTO'].iloc[0]) if pd.notna(dem_obito['MENSALIDADE_ODONTO'].iloc[0]) else 0
            else:
                m_saude = m_odonto = 0.0
            # MAS O SUBSÍDIO É OBRIGATORIAMENTE ZERO, POIS A PBH NÃO PAGOU MAIS NADA!
            s_saude = 0.0
            s_odonto = 0.0
        
        # Registra se neste mês houve algum tipo de subsídio retido
        if s_saude > 0 or s_odonto > 0:
            meses_subsidiados.append(mes_atual)
            
        if tem_aero: m_saude += VALOR_FIXO_AERO
            
        dias_no_mes = calendar.monthrange(mes_atual.year, mes_atual.month)[1]
        
        if mes_atual == dt_obito.replace(day=1):
            dias_a_devolver = dias_no_mes - dt_obito.day
        else:
            dias_a_devolver = dias_no_mes 
            
        fracao_mes = dias_a_devolver / dias_no_mes
        esp_saude_mes = m_saude * fracao_mes
        esp_odonto_mes = m_odonto * fracao_mes
        
        esp_saude_total += esp_saude_mes
        esp_odonto_total += esp_odonto_mes
        
        pct_mes_pbh_saude = (s_saude / m_saude) if m_saude > 0 else 0.0
        pct_mes_pbh_odonto = (s_odonto / m_odonto) if m_odonto > 0 else 0.0
        
        pbh_saude_acumulado += (esp_saude_mes * pct_mes_pbh_saude)
        pbh_odonto_acumulado += (esp_odonto_mes * pct_mes_pbh_odonto)
        
        if mes_atual.month == 12: mes_atual = mes_atual.replace(year=mes_atual.year+1, month=1)
        else: mes_atual = mes_atual.replace(month=mes_atual.month+1)
            
    pct_global_saude = (pbh_saude_acumulado / esp_saude_total) if esp_saude_total > 0 else 0.0
    pct_global_odonto = (pbh_odonto_acumulado / esp_odonto_total) if esp_odonto_total > 0 else 0.0
    
    # Formata o texto final para a linha de "Período Subsidiado"
    meses_dict = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
    if not meses_subsidiados:
        str_sub = "-"
    else:
        p_mes = meses_subsidiados[0]
        u_mes = meses_subsidiados[-1]
        str_p = f"{meses_dict[p_mes.month]}/{str(p_mes.year)[-2:]}"
        str_u = f"{meses_dict[u_mes.month]}/{str(u_mes.year)[-2:]}"
        if str_p == str_u:
            str_sub = str_p
        else:
            str_sub = f"{str_p} a {str_u}"
    
    return pd.Series({
        'ESPERADO_SAUDE': round(esp_saude_total, 2),
        'PCT_PBH_SAUDE': pct_global_saude,
        'PCT_PBH_ODONTO': pct_global_odonto,
        'PERIODO_SUBSIDIADO_STR': str_sub
    })

print("Executando matemática mês a mês corrigida (com zeramento de subsídio)...")
resultados_rateio = df_emissao.apply(calcular_rateio_mes_a_mes, axis=1)
df_emissao['ESPERADO_SAUDE'] = resultados_rateio['ESPERADO_SAUDE']
# Pega o novo texto de período real
df_emissao['PERÍODO SUBSIDIADO STR'] = resultados_rateio['PERIODO_SUBSIDIADO_STR']

print("Puxando Créditos da Operadora...")
try:
    arquivo_creditos = glob.glob('CREDITOS*.xlsx')[0] 
    df_cred_saude = pd.read_excel(arquivo_creditos, sheet_name='CRED_SUB_SAÚDE')
    df_cred_saude['CPF_CLEAN'] = df_cred_saude['CPF_Beneficiario'].apply(clean_cpf)
    cred_saude = df_cred_saude.groupby('CPF_CLEAN', as_index=False)['Crédito abatido no Subsidio'].sum().rename(columns={'Crédito abatido no Subsidio': 'CRED_OPERADORA_SAUDE'})

    df_cred_odonto = pd.read_excel(arquivo_creditos, sheet_name='CRED_SUB_ODONTO')
    df_cred_odonto['CPF_CLEAN'] = df_cred_odonto['CPF_Beneficiario'].apply(clean_cpf)
    cred_odonto = df_cred_odonto.groupby('CPF_CLEAN', as_index=False)['Crédito abatido no Subsidio'].sum().rename(columns={'Crédito abatido no Subsidio': 'CRED_OPERADORA_ODONTO'})

    try:
        df_cred_generico = pd.read_excel(arquivo_creditos, sheet_name='CRED_SUB_GENÉRICO')
        df_cred_generico['CPF_CLEAN'] = df_cred_generico['CPF_Beneficiario'].apply(clean_cpf)
        cred_generico = df_cred_generico.groupby('CPF_CLEAN', as_index=False)['Crédito abatido no Subsidio'].sum().rename(columns={'Crédito abatido no Subsidio': 'CRED_OPERADORA_GENERICO'})
    except: cred_generico = pd.DataFrame(columns=['CPF_CLEAN', 'CRED_OPERADORA_GENERICO'])

    df_emissao = pd.merge(df_emissao, cred_saude, on='CPF_CLEAN', how='left')
    df_emissao = pd.merge(df_emissao, cred_odonto, on='CPF_CLEAN', how='left')
    df_emissao = pd.merge(df_emissao, cred_generico, on='CPF_CLEAN', how='left')
except IndexError:
    print("\n❌ ERRO: O arquivo de Créditos da Operadora não foi encontrado na pasta!")
    sys.exit()

df_emissao['CRED_OPERADORA_SAUDE'] = df_emissao['CRED_OPERADORA_SAUDE'].fillna(0)
df_emissao['CRED_OPERADORA_ODONTO'] = df_emissao['CRED_OPERADORA_ODONTO'].fillna(0)
df_emissao['CRED_OPERADORA_GENERICO'] = df_emissao['CRED_OPERADORA_GENERICO'].fillna(0)

def checar_status(row):
    if pd.isna(row['ESPERADO_SAUDE']) or row['ESPERADO_SAUDE'] == 0: return '-'
    if abs(row['CRED_OPERADORA_SAUDE'] - row['ESPERADO_SAUDE']) <= 0.10: return 'OK'
    return f'ERRO: Op mandou R${row["CRED_OPERADORA_SAUDE"]:.2f}, calc R${row["ESPERADO_SAUDE"]:.2f}'

df_emissao['STATUS_AUDITORIA'] = df_emissao.apply(checar_status, axis=1)

df_emissao['DEVOLVER_PBH_SAUDE'] = (df_emissao['CRED_OPERADORA_SAUDE'] * resultados_rateio['PCT_PBH_SAUDE']).round(2)
df_emissao['CRÉDITO FALECIDO SAÚDE'] = (df_emissao['CRED_OPERADORA_SAUDE'] - df_emissao['DEVOLVER_PBH_SAUDE']).round(2)
df_emissao['DEVOLVER_PBH_ODONTO'] = (df_emissao['CRED_OPERADORA_ODONTO'] * resultados_rateio['PCT_PBH_ODONTO']).round(2)
df_emissao['CRÉDITO FALECIDO ODONTO'] = (df_emissao['CRED_OPERADORA_ODONTO'] - df_emissao['DEVOLVER_PBH_ODONTO']).round(2)

print("Desenhando e salvando direto no Excel Original...")
wb = load_workbook(arquivo_oficial)
ws = wb['EMISSÃO']

df_emissao['DATA_OBITO_STR'] = df_emissao['DATA_OBITO'].dt.strftime('%d/%m/%Y').fillna('')
df_emissao['DATA_CANCELAMENTO_STR'] = df_emissao['DATA_CANCELAMENTO'].dt.strftime('%d/%m/%Y').fillna('')

borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
fundo_cabecalho = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
fonte_branca_negrito = Font(color="FFFFFF", bold=True)
formato_moeda = r'_-"R$ "* #,##0.00_-;\-"R$ "* #,##0.00_-;_-"R$ "* "-"??_-;_-@_-'

# Agora adicionamos a Coluna Nova para o PDF conseguir ler depois
colunas_para_escrever = [
    ('CRÉDITO ASSISTENCIAL', 'CRED_OPERADORA_SAUDE', formato_moeda),
    ('CRÉDITO ODONTO', 'CRED_OPERADORA_ODONTO', formato_moeda),
    ('CRÉDITO GENÉRICO', 'CRED_OPERADORA_GENERICO', formato_moeda), 
    ('ESPERADO SAÚDE', 'ESPERADO_SAUDE', formato_moeda),
    ('STATUS AUDITORIA', 'STATUS_AUDITORIA', None),
    ('SUBSÍDIO PBH SAÚDE', 'DEVOLVER_PBH_SAUDE', formato_moeda),
    ('SUBSÍDIO PBH ODONTO', 'DEVOLVER_PBH_ODONTO', formato_moeda),
    ('CRÉDITO FALECIDO SAÚDE', 'CRÉDITO FALECIDO SAÚDE', formato_moeda),
    ('CRÉDITO FALECIDO ODONTO', 'CRÉDITO FALECIDO ODONTO', formato_moeda),
    ('PERÍODO SUBSIDIADO', 'PERÍODO SUBSIDIADO STR', None), # A Coluna Mágica
]

headers = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]

for index, row in df_emissao.iterrows():
    row_excel = index + 3 
    if 'DATA ÓBITO' in headers: ws.cell(row=row_excel, column=headers.index('DATA ÓBITO') + 1, value=row['DATA_OBITO_STR'])
    if 'DATA CANCELAMENTO' in headers: ws.cell(row=row_excel, column=headers.index('DATA CANCELAMENTO') + 1, value=row['DATA_CANCELAMENTO_STR'])
    
    for col_nome in cols_debito + [' TOTAL DÉBITOS']:
        if col_nome in headers:
            idx = headers.index(col_nome) + 1
            celula = ws.cell(row=row_excel, column=idx, value=row[col_nome])
            celula.number_format = formato_moeda
            celula.border = borda_fina

coluna_inicial_novas = ws.max_column + 1
col_atual = coluna_inicial_novas
indices_novas = {}

for nome_col_excel, col_df, formato in colunas_para_escrever:
    if nome_col_excel in headers:
        c_idx = headers.index(nome_col_excel) + 1
    else:
        c_idx = col_atual
        celula_cab = ws.cell(row=2, column=c_idx, value=nome_col_excel)
        celula_cab.fill = fundo_cabecalho
        celula_cab.font = fonte_branca_negrito
        celula_cab.border = borda_fina
        celula_cab.alignment = Alignment(horizontal='center', vertical='center')
        col_atual += 1
    indices_novas[nome_col_excel] = (c_idx, col_df, formato)

for index, row in df_emissao.iterrows():
    row_excel = index + 3
    for nome_col_excel, config in indices_novas.items():
        c_idx, col_df, formato = config
        if col_df is not None and col_df in row.index:
            valor = row[col_df]
            if pd.notna(valor):
                celula = ws.cell(row=row_excel, column=c_idx, value=valor)
                if formato: celula.number_format = formato
                celula.border = borda_fina

wb.save(arquivo_oficial)
print("\n✅ Planilha preenchida com a Nova Lógica e Cálculo Exato de Período!")