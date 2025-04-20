from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import yagmail
from config import pass_gmail

class editeCertificate:
    def __init__(self):
        self.read_data()
        self.send_email_with_certificate()

    def read_data(self):
        self.df = pd.read_excel('lista_alunos.xlsx')

    def send_email_with_certificate(self):
        for _, row in self.df.iterrows():
            name = row['Nome']
            email = row['Email']
            certificate_image = self.create_certificate(name)
            self.send_email_generic(name, email, certificate_image)

    def create_certificate(self, name):
        imagem = Image.open('certificado.jpg')
        draw = ImageDraw.Draw(imagem)
        font = ImageFont.truetype("Roboto-Italic-VariableFont_wdth,wght.ttf", 27)

        draw.text((668, 578), name, font=font, fill=(0, 0, 0))
        certificate_image = f'{name}_certificado.png'
        imagem.save(certificate_image)
        return certificate_image


    def send_email_generic(self, name, email, certificate_image):
        usuario = yagmail.SMTP(user='ieee.uff@gmail.com, password=pass_gmail')
        assunto = 'Certificado IEEE UFF'
        conteudo = f'Olá {name},\n\nSegue em anexo o seu certificado.\n\nAtenciosamente,\nIEEE UFF'
        usuario.send(
            to=email,
            subject=assunto,
            contents=conteudo,
            attachments=certificate_image
        )
        print(f'Email enviado para {name} com sucesso!')


start = editeCertificate()
