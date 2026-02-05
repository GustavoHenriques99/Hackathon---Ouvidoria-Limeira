import os
import time
import re
import pandas as pd
import camelot

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# ============================================
# CONFIGURAÇÃO DOS ARQUIVOS
# ============================================
file_name = "eOuve - Limeira"
pdf_dir = os.path.abspath("pdf")
pdf_path = os.path.join(pdf_dir, f"{file_name}.pdf")

excel_output = os.path.abspath(f"{file_name}.xlsx")

if not os.path.exists(pdf_dir):
    os.makedirs(pdf_dir)

# ============================================
# CONFIGURAÇÃO DO SELENIUM
# ============================================
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": pdf_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options)

# ============================================
# 1) BAIXAR PDF AUTOMATICAMENTE
# ============================================

print("\n🌐 Acessando página para download...")
driver.get("https://www.ilovepdf.com/pt/pdf_para_excel")
time.sleep(2)

# Envia PDF
upload_btn = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
upload_btn.send_keys(pdf_path)

print("📤 Enviando PDF para conversão...")
time.sleep(4)

# Clica no botão converter
convert_btn = driver.find_element(By.ID, "processTask")
convert_btn.click()

print("⚙ Convertendo...")

# Esperar download do Excel
downloaded_excel = None
while not downloaded_excel:
    for f in os.listdir(pdf_dir):
        if f.endswith(".xlsx"):
            downloaded_excel = f
            break
    time.sleep(1)

driver.quit()

excel_path = os.path.join(pdf_dir, downloaded_excel)
print(f"📥 Excel convertido baixado: {excel_path}")

# ============================================
# 2) EXTRAIR TABELAS DIRETO DO PDF ORIGINAL
# ============================================

print("\n🔍 Extraindo tabelas do PDF com Camelot...")

regex_titulo = re.compile(r"^\d+(\.\d+)+\s*-\s*.+", re.IGNORECASE)

def limpar_texto(txt):
    if not txt:
        return ""
    txt = re.sub(r'\beOuve - Limeira\b', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'(https?://\S+|www\.\S+)', '', txt)
    txt = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', '', txt)
    txt = re.sub(r'\b\d{2}:\d{2}\b', '', txt)
    txt = re.sub(r'\s+', ' ', txt).rstrip()
    return txt

def linha_e_descricao(row):
    texto = " ".join([c for c in row if c])
    if not texto:
        return False
    if texto.strip().startswith("-"):
        return True
    if len([c for c in row if c]) == 1 and not re.search(r'\d', texto):
        return True
    if not re.search(r'\d', texto):
        return True
    return False

def linha_e_valores(row):
    return any(re.search(r'\d', c) for c in row if c)

tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream', edge_tol=150)

tabelas = {}
tabela_atual = None
buffer = []
descricao_parcial = ""

def salvar_tabela():
    global buffer, tabela_atual
    if tabela_atual and buffer:
        df = pd.DataFrame(buffer)
        tabelas[tabela_atual] = df.copy()

for table in tables:
    df = table.df

    for idx, row in df.iterrows():
        row = [limpar_texto(x) for x in row.values]
        linha = " ".join([x for x in row if x])

        if not linha:
            continue

        if "%" in linha:
            continue

        # Detectar título de tabela
        if regex_titulo.match(linha):
            salvar_tabela()
            buffer = []
            tabela_atual = linha.strip()
            descricao_parcial = ""
            continue

        if not tabela_atual:
            continue

        # Descrição quebrada
        if linha_e_descricao(row):
            descricao_parcial += " " + linha
            descricao_parcial = descricao_parcial.strip()
            continue

        # Linha de valores
        if linha_e_valores(row):
            if descricao_parcial:
                nova_linha = [descricao_parcial] + row[1:]
                buffer.append(nova_linha)
                descricao_parcial = ""
            else:
                buffer.append(row)
            continue

        # Linha comum
        buffer.append(row)

salvar_tabela()

# ============================================
# 3) EXPORTAR ORGANIZADO EM EXCEL FINAL
# ============================================

print("\n📊 Exportando tabelas extraídas para Excel final...")

with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
    for titulo, df in tabelas.items():
        sheet = titulo[:31]
        df.to_excel(writer, sheet_name=sheet, index=False)

print(f"\n🎉 Arquivo final gerado com sucesso: {excel_output}")
