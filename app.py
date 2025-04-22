from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import zipfile
from generate_certificate import EditarCertificado

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CERTIFICATES_FOLDER = 'certificado'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CERTIFICATES_FOLDER'] = CERTIFICATES_FOLDER

# Certifique-se de que as pastas de uploads e certificados existem
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CERTIFICATES_FOLDER, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400

    file = request.files['file']
    if file.filename == '':
        return "Nenhum arquivo selecionado", 400

    if not file.filename.endswith('.xlsx'):
        return "Formato de arquivo inválido. Envie um arquivo .xlsx", 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        # Gera os certificados usando o arquivo enviado
        generator = EditarCertificado()
        generator.dados = pd.read_excel(filepath)

        # Verifica se a coluna 'Nome' existe no arquivo Excel
        if 'Nome' not in generator.dados.columns:
            return "O arquivo enviado não contém a coluna 'Nome'.", 400

        # Gera os certificados e salva na pasta CERTIFICATES_FOLDER
        certificados_gerados = []
        for _, linha in generator.dados.iterrows():
            nome = linha['Nome']
            caminho_certificado = os.path.join(app.config['CERTIFICATES_FOLDER'], f"{nome}_certificado.pdf")
            caminho_gerado = generator.criar_certificado(nome)  # Gera o certificado
            if os.path.exists(caminho_gerado):  # Verifica se o arquivo foi criado
                certificados_gerados.append(caminho_gerado)
            else:
                print(f"Erro: Certificado não encontrado para {nome}")

        # Compacta os certificados num arquivo .zip
        zip_path = os.path.join(app.config['CERTIFICATES_FOLDER'], 'certificados.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for certificado in certificados_gerados:
                zipf.write(certificado, os.path.basename(certificado))  # Adiciona ao .zip com o nome correto

        # Retorna o arquivo .zip para download
        return send_file(zip_path, as_attachment=True, download_name='certificados.zip')

    except Exception as e:
        print(f"Erro ao gerar certificados: {str(e)}")
        return f"Erro ao gerar certificados: {str(e)}", 500


if __name__ == '__main__':
    app.run(debug=True)