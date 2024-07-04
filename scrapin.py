import PyPDF2
import pandas as pd
import re
import os

def extract_invoice_info(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        
        # Expressões regulares para extrair as informações
        numero_nota = re.search(r'Nº.\s*(\d+\.\d+\.\d+)', text)
        serie = re.search(r'Série\s*(\d+)', text)
        chave_acesso = re.search(r'CHAVE DE ACESSO\s*([\d\s]+)', text)
        operacao = re.search(r'NATUREZA DA OPERAÇÃO\s*([\w\s]+)\s*PROTOCOLO DE AUTORIZAÇÃO DE USO', text)
        data = re.search(r'DATA DA (?:EMISSÃO|SAÍDA/ENTRADA)\s*(\d{2}/\d{2}/\d{4})', text)
        valor_produtos = re.search(r'V\. TOTAL PRODUTOS\s*([\d.,]+)', text)
        valor_total = re.search(r'V\. TOTAL DA NOTA\s*([\d.,]+)', text)

        # Dicionário com as informações extraídas
        invoice_info = {
            'Número da Nota': numero_nota.group(1) if numero_nota else None,
            'Série': serie.group(1) if serie else None,
            'Chave de Acesso': chave_acesso.group(1) if chave_acesso else None,
            'Operação': operacao.group(1).strip() if operacao else None,
            'Data': data.group(1) if data else None,
            'Valor Produtos': valor_produtos.group(1).replace('.', '').replace(',', '.') if valor_produtos else None,
            'Valor Total': valor_total.group(1).replace('.', '').replace(',', '.') if valor_total else None,
        }
        print(f"Extraído de {pdf_path}: {invoice_info}")
        return invoice_info

def process_invoices(pdf_paths):
    invoices_data = []

    for pdf_path in pdf_paths:
        invoice_info = extract_invoice_info(pdf_path)
        invoices_data.append(invoice_info)

    df = pd.DataFrame(invoices_data)
    print(f"DataFrame criado com {len(df)} linhas.")
    print(df.head())
    output_file = 'notas_fiscais.xlsx'
    try:
        df.to_excel(output_file, index=False)
        current_dir = os.getcwd()  # Obter o diretório atual
        output_path = os.path.join(current_dir, output_file)
        print(f"Arquivo Excel salvo em: {output_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
    

# Obter todos os arquivos PDF do diretório especificado
directory = '/home/pedro/meu_ambiente/junho'  # Substitua pelo caminho do diretório onde estão os PDFs
pdf_paths = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.pdf')]

process_invoices(pdf_paths)
