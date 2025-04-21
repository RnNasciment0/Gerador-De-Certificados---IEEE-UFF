from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

class editeCertificate:
    def __init__(self):
        self.read_data()

    def read_data(self):
        self.df = pd.read_excel('lista_alunos.xlsx')

    def abbreviate_name(self, name, max_length):
        # Divide o nome em partes
        parts = name.split()
        if len(name) <= max_length:
            return name

        # Mantém o primeiro e o último nome, abreviando os do meio
        abbreviated = [parts[0]] + [p[0] + '.' for p in parts[1:-1]] + [parts[-1]]
        abbreviated_name = ' '.join(abbreviated)

        # Garante que o nome abreviado não ultrapasse o limite
        while len(abbreviated_name) > max_length:
            if len(abbreviated) > 2:
                abbreviated.pop(-2)  # Remove o penúltimo nome abreviado
                abbreviated_name = ' '.join(abbreviated)
            else:
                break

        return abbreviated_name

    def create_certificate(self, name):
        # Verifica se o arquivo de fonte existe
        font_path = os.path.join('static', 'Roboto-Italic-VariableFont_wdth,wght.ttf')
        if not os.path.exists(font_path):
            raise FileNotFoundError(f"Arquivo de fonte não encontrado: {font_path}")

        # Abrevia o nome se necessário
        max_length = 30  # Defina o limite de caracteres
        name = self.abbreviate_name(name, max_length)

        # Cria a imagem do certificado
        imagem = Image.open('static/certificado.jpg')
        draw = ImageDraw.Draw(imagem)
        font = ImageFont.truetype(font_path, 27)

        draw.text((668, 578), name, font=font, fill=(0, 0, 0))
        pdf_path = f'{name}_certificado.pdf'
        imagem.convert('RGB').save(pdf_path, "PDF", resolution=100)
        return pdf_path

    def generate_certificates(self):
        for _, row in self.df.iterrows():
            name = row['Nome']
            self.create_certificate(name)
            print(f'Certificado gerado para {name}!')


start = editeCertificate()
start.generate_certificates()
