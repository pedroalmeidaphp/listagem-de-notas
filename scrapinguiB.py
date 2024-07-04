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
        numero_nota = re.search(r'Nº\s*(\d+)', text)
        serie = re.search(r'Série\s*(\d+)', text)
        chave_acesso = re.search(r'Chave\s*de\s*acesso\s*([\d\s\t\n]+)', text)
        operacao = re.search(r'Natureza\s*da\s*operação\s*([^\n]+)', text)
        data = re.search(r'Data\s*(?:emissão|saída)\s*(\d{2}/\d{2}/\d{4})', text)
        valor_produtos = re.search(r'Valor\s*total\s*dos\s*produtos\s*([\d.,]+)', text)
        valor_total = re.search(r'Valor\s*total\s*da\s*nota\s*([\d.,]+)', text)

        # Processar a chave de acesso para remover tabulações e novas linhas
        chave_acesso_formatada = chave_acesso.group(1).replace('\t', ' ').replace('\n', '').strip() if chave_acesso else None
        # Processar a operação para remover tabulações e novas linhas
        operacao_formatada = operacao.group(1).replace('\t', ' ').strip().split('\n')[0] if operacao else None

        # Dicionário com as informações extraídas
        invoice_info = {
            'Número da Nota': numero_nota.group(1) if numero_nota else None,
            'Série': serie.group(1) if serie else None,
            'Chave de Acesso': chave_acesso_formatada,
            'Operação': operacao_formatada,
            'Data': data.group(1) if data else None,
            'Valor Produtos': valor_produtos.group(1).replace('.', '').replace(',', '.') if valor_produtos else None,
            'Valor Total': valor_total.group(1).replace('.', '').replace(',', '.') if valor_total else None,
        }
        print(f"Extraído de {pdf_path}: {invoice_info}")
        return invoice_info

def process_invoices(pdf_paths, output_filename, window):
    invoices_data = []

    for i, pdf_path in enumerate(pdf_paths):
        invoice_info = extract_invoice_info(pdf_path)
        invoices_data.append(invoice_info)
        
        # Atualizar a barra de progresso
        window['-PROGRESS-'].update(current_count=i+1)

    df = pd.DataFrame(invoices_data)
    print(f"DataFrame criado com {len(df)} linhas.")
    print(df.head())
    
    try:
        df.to_excel(output_filename, index=False)
        current_dir = os.getcwd()  # Obter o diretório atual
        output_path = os.path.join(current_dir, output_filename)
        print(f"Arquivo Excel salvo em: {output_path}")
        sg.popup(f"Arquivo Excel salvo em: {output_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        sg.popup(f"Erro ao salvar o arquivo Excel: {e}")

# Layout da interface gráfica
layout = [
    [sg.Text('Selecione o diretório dos PDFs:')],
    [sg.Input(), sg.FolderBrowse()],
    [sg.Text('Nome do arquivo Excel (sem extensão):')],
    [sg.Input(key='-FILENAME-')],
    [sg.Button('Processar')],
    [sg.Text('', size=(40, 1), key='-OUTPUT-')],
    [sg.ProgressBar(max_value=100, orientation='h', size=(40, 20), key='-PROGRESS-')]
]

# Janela da interface gráfica
window = sg.Window('Processador de Notas Fiscais', layout)

while True:
    event, values = window.read()
    
    if event == sg.WINDOW_CLOSED:
        break
    if event == 'Processar':
        directory = values[0]
        output_filename = values['-FILENAME-']
        if not directory:
            sg.popup('Por favor, selecione um diretório.')
        elif not output_filename:
            sg.popup('Por favor, insira o nome do arquivo Excel.')
        else:
            # Adicionar a extensão .xlsx ao nome do arquivo se não estiver presente
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
                
            pdf_paths = [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.pdf')]
            if not pdf_paths:
                sg.popup('Nenhum arquivo PDF encontrado no diretório selecionado.')
            else:
                # Atualizar a barra de progresso para o número de arquivos PDF
                window['-PROGRESS-'].update(max=len(pdf_paths), current_count=0)
                # Processar as notas fiscais
                process_invoices(pdf_paths, output_filename, window)

window.close()
