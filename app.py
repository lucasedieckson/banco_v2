from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import zipfile
import io
from docxtpl import DocxTemplate
from datetime import datetime

app = Flask(__name__)

# Função para verificar a extensão do arquivo
def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

# Função para preencher modelo do Word com dados do Excel
def fill_word_template(excel_data, word_template_path):
    # Carrega o arquivo Excel
    tb_df = pd.read_excel(excel_data)
    
    # Formata a coluna 'DATA CONTRATACAO' para o formato '%d/%m/%Y'
    tb_df['DATA CONTRATACAO'] = tb_df['DATA CONTRATACAO'].dt.strftime('%d/%m/%Y')
    
    output_files = []

    for i in range(len(tb_df)):
        # Carrega o modelo do Word
        doc = DocxTemplate(word_template_path)

        # Dicionário com os dados para preenchimento
        context = {
            'CONTRATACAO': tb_df.loc[i, 'DATA CONTRATACAO'],
            'REDE': tb_df.loc[i, 'REDE'],
            'PDV': tb_df.loc[i, 'NOME DO PDV'],
            'ENDEREÇO_PDV': tb_df.loc[i, 'ENDEREÇO PDV'],
            'NOME': tb_df.loc[i, 'NOME DO COLABORADOR'],
            'RG': tb_df.loc[i, 'RG'],
            'CPF': tb_df.loc[i, 'CPF'],
            'CTPS': tb_df.loc[i, 'CTPS'],
            'SÉRIE': tb_df.loc[i, 'SÉRIE'],
            'ENDEREÇO': tb_df.loc[i, 'ENDEREÇO DO COLABORADOR'],
            'FUNÇÃO': tb_df.loc[i, 'FUNÇÃO'],
            'CLIENTE': tb_df.loc[i, 'CLIENTE'],
            'EMPRESA': tb_df.loc[i, 'EMPRESA']
        }
        
        # Renderiza o documento com os dados
        doc.render(context)
        
        # Gera o nome do arquivo com 'NOME DO COLABORADOR' e 'PDV'
        nome = str(tb_df.loc[i, 'NOME DO COLABORADOR'])
        pdv = str(tb_df.loc[i, 'NOME DO PDV'])
        output_filename = f"{nome}_{pdv}.docx"

        # Salva o documento preenchido na memória
        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)
        output_files.append((output_filename, output_stream))
        
        if tb_df.iloc[i].isnull().all():
            break    

    return output_files

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verifica se os arquivos foram enviados
        if 'excel_file' not in request.files:
            return render_template('index.html', message='Por favor, selecione o arquivo Excel.')

        # Obtem o arquivo enviado
        excel_file = request.files['excel_file']

        # Verifica se o arquivo é válido
        if excel_file.filename == '':
            return render_template('index.html', message='Por favor, selecione o arquivo Excel.')

        # Verifica a extensão do arquivo
        if not allowed_file(excel_file.filename, {'xlsx', 'xls'}):
            return render_template('index.html', message='Formato de arquivo do Excel inválido.')

        # Obtem o modelo de Word selecionado
        word_template = request.form.get('word_template')

        # Caminho do modelo do Word
        word_template_path = os.path.join(app.root_path, 'templates', f'{word_template}.docx')

        # Preenche um documento Word para cada linha do arquivo Excel
        output_files = fill_word_template(excel_file, word_template_path)

        # Cria um arquivo ZIP em memória com todos os documentos preenchidos
        zip_stream = io.BytesIO()
        with zipfile.ZipFile(zip_stream, 'w') as zipf:
            for filename, file_stream in output_files:
                # Adiciona cada documento ao arquivo ZIP
                zipf.writestr(filename, file_stream.read())
        zip_stream.seek(0)
        
        # Obter a data atual
        data_atual = datetime.now()

        # Formatar a data no formato desejado para incluir no nome do arquivo ZIP
        nome_arquivo_zip = f'Cartas_Prontas_{data_atual.strftime("%d-%m")}.zip'

        # Envia o arquivo ZIP para download
        return send_file(
            zip_stream,
            mimetype='application/zip',
            as_attachment=True,
            download_name=nome_arquivo_zip
        )
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0')