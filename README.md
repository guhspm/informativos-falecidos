# 📄 Gerador de Informativos — Falecidos UNIMED
> Automação completa para emissão de informativos de saldo e memórias de cálculo em PDF para ex-beneficiários falecidos — geração em lote a partir de planilha Excel.

![Python](https://img.shields.io/badge/Python-3.10+-8b5cf6?style=flat-square&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-2.x-8b5cf6?style=flat-square&logo=pandas&logoColor=white)
![fpdf2](https://img.shields.io/badge/fpdf2-2.x-8b5cf6?style=flat-square&logoColor=white)
![Status](https://img.shields.io/badge/Status-Em%20Produção-22c55e?style=flat-square)

---

## 📌 Sobre o Projeto

Sistema de emissão automática de documentos oficiais para o processo de baixa de falecidos em planos de saúde.

O sistema é composto por dois scripts que trabalham em sequência:

1. **`preencher_planilha.py`** — pré-processa a planilha oficial, calcula créditos pro-rata, subsídios PBH e preenche automaticamente as colunas de cálculo
2. **`etapa2_gerar_pdfs.py`** — lê a planilha preenchida e gera dois PDFs por beneficiário: **Informativo de Saldo** e **Memória de Cálculo**

**Problema resolvido:** a emissão manual de informativos para dezenas de falecidos por mês consumia horas de trabalho repetitivo e estava sujeita a erros de cálculo. O sistema gera todos os documentos em segundos, com formatação executiva e cálculos auditáveis.

---

## 🚀 Funcionalidades

- ✅ **Preenchimento automático da planilha** — calcula crédito pro-rata, subsídio PBH saúde/odonto e valor líquido destinado ao beneficiário
- ✅ **Geração de PDFs em lote** — um Informativo de Saldo + uma Memória de Cálculo por pessoa
- ✅ **Layout executivo profissional** — logo da empresa, tabelas com bordas, formatação monetária R$
- ✅ **Leitura inteligente de datas** — extrai período de óbito a exclusão e período subsidiado PBH
- ✅ **Tratamento de dados sujos** — CPF em formato float, células nulas, encoding
- ✅ **Validação de arquivo aberto** — avisa se a planilha está aberta no Excel antes de sobrescrever

---

## 🛠️ Stack

| Tecnologia | Uso |
|---|---|
| Python 3.10+ | Lógica principal |
| Pandas | Leitura e processamento da planilha |
| fpdf2 | Geração dos PDFs com layout customizado |
| OpenPyXL | Escrita de volta na planilha original |
| NumPy | Cálculos financeiros |
| calendar | Cálculo pro-rata por dias no mês |

---

## 📁 Estrutura

```
informativos-falecidos/
├── preencher_planilha.py        # Etapa 1 — preenche cálculos na planilha
├── etapa2_gerar_pdfs.py         # Etapa 2 — gera PDFs em lote
├── requirements.txt
├── .gitignore
└── README.md

# Arquivos esperados na mesma pasta (não versionados):
├── INFORMATIVOS_FALECIDOS.2026.xlsx   # Planilha oficial (input)
├── logo.png                            # Logo para os PDFs
└── PDFs_Gerados/                       # Pasta de output (gerada automaticamente)
```

---

## ⚙️ Como Executar

```bash
pip install -r requirements.txt
```

**Etapa 1 — Preencher planilha:**
```bash
python preencher_planilha.py
# Preenche automaticamente os campos de crédito na planilha oficial
```

**Etapa 2 — Gerar PDFs:**
```bash
python etapa2_gerar_pdfs.py
# Gera os PDFs em PDFs_Gerados/
```

---

## 📈 Exemplo de Output

```
Lendo a planilha oficial (Aba EMISSÃO)...
Iniciando a geração dos PDFs...

✅ 47 PDFs gerados com sucesso na pasta 'PDFs_Gerados'!

Arquivos gerados (2 por pessoa):
  → João Silva - Informativo.pdf
  → João Silva - Mem Calc.pdf
  → Maria Santos - Informativo.pdf
  → Maria Santos - Mem Calc.pdf
  ...
```

---

## 📋 Documentos Gerados

### Informativo de Saldo
Documento oficial com:
- Dados do titular (nome, matrícula, CPF)
- Tabela de débitos por situação (Passivo, Boleto, Folha)
- Tabela de créditos (saúde, odonto, genérico)
- Saldo final (crédito ou débito)

### Memória de Cálculo
Documento de transparência com:
- Valor bruto da operadora
- Subsídio retido pela PBH
- Valor líquido destinado ao ex-beneficiário
- Período de crédito e período subsidiado

---

## 👤 Autor

**Gustavo** — Dev & Founder · Inside.co

[![LinkedIn](https://img.shields.io/badge/LinkedIn-8b5cf6?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/gustavo-henriquesp/)
[![Portfolio](https://img.shields.io/badge/Portfolio-8b5cf6?style=flat-square&logo=netlify&logoColor=white)](https://seusite.netlify.app)
[![Email](https://img.shields.io/badge/Email-8b5cf6?style=flat-square&logo=gmail&logoColor=white)](mailto:ghspdm@gmail.com)

---
> *"Construo soluções que outros apenas descrevem em planilhas."*
