from PIL import Image, ImageDraw, ImageFont
import pandas as pd

class editeCertificate:
    def __init__(self):
        self.read_data()

    def read_data(self):
        self.df = pd.read_excel('lista_alunos.xlsx')

    def create_certificate(self, name):
        imagem = Image.open('certificado.jpg')
        draw = ImageDraw.Draw(imagem)
        font = ImageFont.truetype("Roboto-Italic-VariableFont_wdth,wght.ttf", 27)

        draw.text((668, 578), name, font=font, fill=(0, 0, 0))
        certificate_image = f'{name}_certificado.png'
        imagem.save(certificate_image)
        return certificate_image

    def generate_certificates(self):
        for _, row in self.df.iterrows():
            name = row['Nome']
            self.create_certificate(name)
            print(f'Certificado gerado para {name}!')

start = editeCertificate()
start.generate_certificates()