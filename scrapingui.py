import PyPDF2
import pandas as pd
import re
import os
import PySimpleGUI as sg

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

def process_invoices(pdf_paths, output_dir):
    invoices_data = []
    total_files = len(pdf_paths)

    for i, pdf_path in enumerate(pdf_paths, start=1):
        invoice_info = extract_invoice_info(pdf_path)
        invoices_data.append(invoice_info)
        
        # Atualizar a barra de progresso
        sg.OneLineProgressMeter('Processando PDFs', i, total_files, 'key')

    df = pd.DataFrame(invoices_data)
    print(f"DataFrame criado com {len(df)} linhas.")
    print(df.head())
    output_file = os.path.join(output_dir, 'notas_fiscais.xlsx')
    try:
        df.to_excel(output_file, index=False)
        print(f"Arquivo Excel salvo em: {output_file}")
        return output_file
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        return None

def main():
    sg.theme('LightBlue2')

    layout = [
        [sg.Text('Selecione o diretório com os arquivos PDF:')],
        [sg.Input(), sg.FolderBrowse(key='-FOLDER-')],
        [sg.Submit('Processar'), sg.Cancel('Cancelar')],
        [sg.Text('', size=(40, 2), key='-OUTPUT-')]
    ]

    window = sg.Window('Processador de Notas Fiscais', layout)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, 'Cancelar'):
            break
        if event == 'Processar':
            directory = values['-FOLDER-']
            if directory:
                pdf_paths = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.pdf')]
                output_file = process_invoices(pdf_paths, directory)
                if output_file:
                    window['-OUTPUT-'].update(f'Arquivo Excel salvo em: {output_file}')
                else:
                    window['-OUTPUT-'].update('Erro ao salvar o arquivo Excel.')
            else:
                window['-OUTPUT-'].update('Por favor, selecione um diretório válido.')

    window.close()

if __name__ == '__main__':
    main()
