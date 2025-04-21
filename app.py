from flask import Flask, render_template, request, redirect, url_for
import os
import pandas as pd
from generate_certificate import editeCertificate

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Certifique-se de que a pasta de uploads existe
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
        generator = editeCertificate()
        generator.df = pd.read_excel(filepath)  # Substitui o DataFrame carregado
        generator.generate_certificates()
    except Exception as e:
        return f"Erro ao gerar certificados: {str(e)}", 500

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)