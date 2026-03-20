import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os
import re

# --- FUNÇÕES AUXILIARES ---
def formata_moeda(valor):
    if pd.isna(valor): return "R$ 0,00"
    valor_formatado = f"{valor:,.2f}"
    valor_formatado = valor_formatado.replace(',', '_').replace('.', ',').replace('_', '.')
    return f"R$ {valor_formatado}"

def gera_periodo(data_obito, data_exclusao):
    try:
        d_obito = pd.to_datetime(data_obito, dayfirst=True, errors='coerce')
        d_exclusao = pd.to_datetime(data_exclusao, dayfirst=True, errors='coerce')
        if pd.isna(d_obito) or pd.isna(d_exclusao): return "-"
        meses = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 
                 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        str_obito = f"{meses[d_obito.month]}/{str(d_obito.year)[-2:]}"
        str_exclusao = f"{meses[d_exclusao.month]}/{str(d_exclusao.year)[-2:]}"
        if str_obito == str_exclusao: return str_obito
        return f"{str_obito} a {str_exclusao}"
    except: return "-"

def limpa_nome_arquivo(nome):
    return re.sub(r'[\\/*?:"<>|]', "", str(nome)).strip()

def clean_cpf(cpf):
    if pd.isna(cpf): return ''
    if isinstance(cpf, float): cpf = str(int(cpf))
    return re.sub(r'\D', '', str(cpf)).zfill(11)

if not os.path.exists("PDFs_Gerados"):
    os.makedirs("PDFs_Gerados")

print("Lendo a planilha oficial (Aba EMISSÃO)...")
try:
    df = pd.read_excel('INFORMATIVOS_FALECIDOS.2026.xlsx', sheet_name='EMISSÃO', skiprows=1)
except Exception as e:
    print("❌ ERRO: Não consegui ler a planilha. Verifique se ela está salva na pasta e fechada no Excel.")
    exit()

colunas_debito = {}
for col in df.columns:
    nome_col_min = str(col).lower()
    if 'passivo' in nome_col_min:
        colunas_debito[col] = 'Inadimplente'
    elif 'boleto' in nome_col_min:
        colunas_debito[col] = 'Em aberto'
    elif 'valores' in nome_col_min or 'plano' in nome_col_min:
        colunas_debito[col] = 'Enviado à Folha'

data_hoje = datetime.now().strftime("%d/%m/%Y")
pessoas_geradas = 0

print("Iniciando a geração dos PDFs...")

for index, row in df.iterrows():
    cpf_limpo = clean_cpf(row.get('CPF', ''))
    if not cpf_limpo: continue
    
    bruto_saude = float(row.get('CRÉDITO ASSISTENCIAL', 0)) if pd.notna(row.get('CRÉDITO ASSISTENCIAL', 0)) else 0
    bruto_odonto = float(row.get('CRÉDITO ODONTO', 0)) if pd.notna(row.get('CRÉDITO ODONTO', 0)) else 0
    bruto_generico = float(row.get('CRÉDITO GENÉRICO', 0)) if pd.notna(row.get('CRÉDITO GENÉRICO', 0)) else 0
    
    if (bruto_saude + bruto_odonto + bruto_generico) <= 0: continue
        
    pessoas_geradas += 1
    
    cpf_formatado = f"{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}"
    nome = str(row['SOLICITANTE']).strip()
    nome_arquivo = limpa_nome_arquivo(nome)
    matricula = str(row['MATRICULA']).replace('.0', '')
    
    # O Período de Crédito continua puxando das datas da operadora (Óbito a Exclusão)
    periodo_texto_credito = gera_periodo(row['DATA ÓBITO'], row['DATA CANCELAMENTO'])
    
    # NOVA LEITURA: O Período Subsidiado agora puxa EXATAMENTE o que a PBH pagou, lendo a nova coluna do Excel
    periodo_texto_subsidiado = str(row.get('PERÍODO SUBSIDIADO', '-'))
    # Prevenção: caso o valor seja nulo no pandas ("nan")
    if periodo_texto_subsidiado == 'nan' or periodo_texto_subsidiado == 'None':
        periodo_texto_subsidiado = '-'

    # ==========================================
    # 1. GERAR INFORMATIVO DE SALDO
    # ==========================================
    pdf = FPDF()
    pdf.add_page()
    
    try: pdf.image('logo.png', x=77.5, y=10, w=55)
    except: pass
    
    pdf.set_y(45)
        
    pdf.set_font("Arial", size=9)
    pdf.cell(0, 5, f"Emissão: {data_hoje}", ln=True, align='R')
    pdf.ln(5)
    
    pdf.set_font("Arial", style='B', size=12)
    pdf.cell(0, 10, "INFORMATIVO DE SALDO DO PLANO DE SAÚDE", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", style='B', size=10)
    pdf.cell(30, 6, "TITULAR:", border=0)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, nome, border=0, ln=True)
    
    pdf.set_font("Arial", style='B', size=10)
    pdf.cell(30, 6, "MATRÍCULA:", border=0)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, matricula, border=0, ln=True)
    
    pdf.set_font("Arial", style='B', size=10)
    pdf.cell(30, 6, "CPF:", border=0)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, cpf_formatado, border=0, ln=True)
    pdf.ln(5)
    
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, "Informamos que consta o saldo abaixo para o(a) ex-beneficiário(a):", ln=True)
    pdf.ln(2)
    
    pdf.set_font("Arial", style='B', size=9)
    pdf.cell(40, 6, "SITUAÇÃO", border=1)
    pdf.cell(110, 6, "DÉBITOS - DESCRIÇÃO", border=1)
    pdf.cell(40, 6, "VALOR", border=1, align='R', ln=True)
    
    pdf.set_font("Arial", size=9)
    total_debitos = 0
    for coluna, situacao in colunas_debito.items():
        if coluna in row and pd.notna(row[coluna]):
            valor_deb = float(row[coluna])
            if valor_deb > 0:
                pdf.cell(40, 6, situacao, border=1)
                pdf.cell(110, 6, coluna, border=1)
                pdf.cell(40, 6, formata_moeda(valor_deb), border=1, align='R', ln=True)
                total_debitos += valor_deb
                
    pdf.set_font("Arial", style='B', size=9)
    pdf.cell(150, 6, "TOTAL DE DÉBITOS", border=1, align='R')
    pdf.cell(40, 6, formata_moeda(total_debitos), border=1, align='R', ln=True)
    pdf.ln(5)
    
    pdf.cell(40, 6, "SITUAÇÃO", border=1)
    pdf.cell(110, 6, "CRÉDITOS - DESCRIÇÃO", border=1)
    pdf.cell(40, 6, "VALOR", border=1, align='R', ln=True)
    
    pdf.set_font("Arial", size=9)
    credito_saude = float(row.get('CRÉDITO FALECIDO SAÚDE', 0)) if pd.notna(row.get('CRÉDITO FALECIDO SAÚDE', 0)) else 0
    credito_odonto = float(row.get('CRÉDITO FALECIDO ODONTO', 0)) if pd.notna(row.get('CRÉDITO FALECIDO ODONTO', 0)) else 0
    
    if credito_saude > 0:
        pdf.cell(40, 6, "A restituir", border=1)
        pdf.cell(110, 6, "Restituição de crédito pro-rata por óbito saúde", border=1)
        pdf.cell(40, 6, formata_moeda(credito_saude), border=1, align='R', ln=True)
        
    if credito_odonto > 0:
        pdf.cell(40, 6, "A restituir", border=1)
        pdf.cell(110, 6, "Restituição de crédito pro-rata por óbito odonto", border=1)
        pdf.cell(40, 6, formata_moeda(credito_odonto), border=1, align='R', ln=True)
        
    if bruto_generico > 0:
        pdf.cell(40, 6, "A restituir", border=1)
        pdf.cell(110, 6, "Restituição de crédito genérico/ajuste", border=1)
        pdf.cell(40, 6, formata_moeda(bruto_generico), border=1, align='R', ln=True)
        
    total_creditos = credito_saude + credito_odonto + bruto_generico
    pdf.set_font("Arial", style='B', size=9)
    pdf.cell(150, 6, "TOTAL DE CRÉDITOS", border=1, align='R')
    pdf.cell(40, 6, formata_moeda(total_creditos), border=1, align='R', ln=True)
    
    saldo_final = total_creditos - total_debitos
    pdf.ln(2)
    pdf.set_font("Arial", style='B', size=9)
    
    texto_situacao = "CRÉDITO" if saldo_final >= 0 else "DÉBITO"
    valor_saldo_str = formata_moeda(abs(saldo_final))
    if saldo_final < 0: valor_saldo_str = "-" + valor_saldo_str
    
    pdf.cell(40, 6, texto_situacao, border=1)
    pdf.cell(110, 6, "TOTAL GERAL", border=1, align='C')
    pdf.cell(40, 6, valor_saldo_str, border=1, align='R', ln=True)
    
    pdf.ln(8)
    pdf.set_font("Arial", size=9)
    msg1 = "Novos valores de coparticipação poderão ser faturados via boleto bancário diretamente ao ex-beneficiário(a), mesmo após a exclusão do plano, se até então não tiverem sido informados pelos prestadores à operadora."
    pdf.multi_cell(0, 5, txt=msg1)
    pdf.ln(2)
    msg2 = "Caso o titular apresente comprovantes de pagamento dos débitos apontados neste informativo, os mesmos deverão ser desconsiderados."
    pdf.multi_cell(0, 5, txt=msg2)
    
    pdf.output(f"PDFs_Gerados/{nome_arquivo} - Informativo.pdf")

    # ==========================================
    # 2. GERAR MEMÓRIA DE CÁLCULO
    # ==========================================
    pdf_mem = FPDF()
    pdf_mem.add_page()
    
    try: pdf_mem.image('logo.png', x=77.5, y=10, w=55)
    except: pass
    
    pdf_mem.set_y(45)
        
    pdf_mem.set_font("Arial", size=9)
    pdf_mem.cell(0, 5, f"Emissão: {data_hoje}", ln=True, align='R')
    pdf_mem.ln(5)
    
    pdf_mem.set_font("Arial", style='B', size=12)
    pdf_mem.cell(0, 10, "MEMÓRIA DE CÁLCULO REFERENTE CRÉDITOS", ln=True, align='C')
    pdf_mem.ln(5)
    
    pdf_mem.set_font("Arial", style='B', size=10)
    pdf_mem.cell(30, 6, "TITULAR:", border=0)
    pdf_mem.set_font("Arial", size=10)
    pdf_mem.cell(0, 6, nome, border=0, ln=True)
    
    pdf_mem.set_font("Arial", style='B', size=10)
    pdf_mem.cell(30, 6, "MATRÍCULA:", border=0)
    pdf_mem.set_font("Arial", size=10)
    pdf_mem.cell(0, 6, matricula, border=0, ln=True)
    
    pdf_mem.set_font("Arial", style='B', size=10)
    pdf_mem.cell(30, 6, "CPF:", border=0)
    pdf_mem.set_font("Arial", size=10)
    pdf_mem.cell(0, 6, cpf_formatado, border=0, ln=True)
    pdf_mem.ln(10)
    
    pdf_mem.set_font("Arial", style='B', size=8)
    pdf_mem.cell(45, 10, "Composição de crédito", border=1, align='C')
    pdf_mem.cell(45, 10, "Valor bruto (Operadora)*", border=1, align='C')
    pdf_mem.cell(45, 10, "Subsídio retido (PBH)**", border=1, align='C')
    pdf_mem.cell(55, 10, "Destinado ao ex-beneficiário", border=1, align='C', ln=True)
    
    pdf_mem.set_font("Arial", size=9)
    
    pbh_saude = float(row.get('SUBSÍDIO PBH SAÚDE', 0)) if pd.notna(row.get('SUBSÍDIO PBH SAÚDE', 0)) else 0
    pdf_mem.cell(45, 8, "Crédito por óbito - saúde/aero", border=1)
    pdf_mem.cell(45, 8, formata_moeda(bruto_saude), border=1, align='C')
    pdf_mem.cell(45, 8, formata_moeda(pbh_saude), border=1, align='C')
    pdf_mem.cell(55, 8, formata_moeda(credito_saude), border=1, align='C', ln=True)
    
    pbh_odonto = float(row.get('SUBSÍDIO PBH ODONTO', 0)) if pd.notna(row.get('SUBSÍDIO PBH ODONTO', 0)) else 0
    pdf_mem.cell(45, 8, "Crédito por óbito - odonto", border=1)
    pdf_mem.cell(45, 8, formata_moeda(bruto_odonto), border=1, align='C')
    pdf_mem.cell(45, 8, formata_moeda(pbh_odonto), border=1, align='C')
    pdf_mem.cell(55, 8, formata_moeda(credito_odonto), border=1, align='C', ln=True)
    
    if bruto_generico > 0:
        pdf_mem.cell(45, 8, "Crédito Genérico/Ajuste", border=1)
        pdf_mem.cell(45, 8, formata_moeda(bruto_generico), border=1, align='C')
        pdf_mem.cell(45, 8, "R$ 0,00", border=1, align='C')
        pdf_mem.cell(55, 8, formata_moeda(bruto_generico), border=1, align='C', ln=True)
    
    pdf_mem.set_font("Arial", style='B', size=9)
    pdf_mem.cell(45, 8, "TOTAL", border=1)
    pdf_mem.cell(45, 8, formata_moeda(bruto_saude + bruto_odonto + bruto_generico), border=1, align='C')
    pdf_mem.cell(45, 8, formata_moeda(pbh_saude + pbh_odonto), border=1, align='C')
    pdf_mem.cell(55, 8, formata_moeda(credito_saude + credito_odonto + bruto_generico), border=1, align='C', ln=True)
    pdf_mem.ln(10)
    
    pdf_mem.set_font("Arial", size=9)
    pdf_mem.cell(0, 5, f"* período crédito: {periodo_texto_credito}", ln=True)
    # A MÁGICA ACONTECE AQUI:
    pdf_mem.cell(0, 5, f"** período subsidiado: {periodo_texto_subsidiado}", ln=True) 
    
    pdf_mem.output(f"PDFs_Gerados/{nome_arquivo} - Mem Calc.pdf")

print(f"\n✅ {pessoas_geradas} PDFs gerados com sucesso na pasta 'PDFs_Gerados'!")