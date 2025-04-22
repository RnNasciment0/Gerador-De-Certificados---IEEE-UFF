from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


class EditarCertificado:
    def __init__(self):
        self.ler_dados()

    def ler_dados(self):
        self.dados = pd.read_excel('lista_alunos.xlsx')

    def abreviar_nome(self, nome, tamanho_maximo):
        # Lista de preposições que devem ser removidas
        preposicoes = {"DA", "DE", "DO", "DAS", "DOS", "E"}

        # Divide o nome em partes
        partes = nome.split()
        # Remove preposições
        partes = [parte for parte in partes if parte.upper() not in preposicoes]

        if len(' '.join(partes)) <= tamanho_maximo:
            return ' '.join(partes)

        # Mantém o primeiro e o último nome, abreviando os do meio
        abreviado = [partes[0]] + [parte[0] + '.' for parte in partes[1:-1]] + [partes[-1]]
        nome_abreviado = ' '.join(abreviado)

        # Garante que o nome abreviado não ultrapasse o limite
        while len(nome_abreviado) > tamanho_maximo:
            if len(abreviado) > 2:
                abreviado.pop(-2)  # Remove o penúltimo nome abreviado
                nome_abreviado = ' '.join(abreviado)
            else:
                break

        return nome_abreviado

    def criar_certificado(self, nome):
        # Verifica se o arquivo de fonte existe
        caminho_fonte = os.path.join('static', 'BeVietnamPro-Regular.ttf')
        if not os.path.exists(caminho_fonte):
            raise FileNotFoundError(f"Arquivo de fonte não encontrado: {caminho_fonte}")

        # Abrevia o nome se necessário
        tamanho_maximo = 30  # Defina o limite de caracteres
        nome = self.abreviar_nome(nome, tamanho_maximo)

        # Cria a imagem do certificado
        imagem = Image.open('static/certificado.jpg')
        desenhar = ImageDraw.Draw(imagem)
        fonte = ImageFont.truetype(caminho_fonte, 27)

        desenhar.text((675, 578), nome, font=fonte, fill=(0, 0, 0))
        caminho_pdf = os.path.join('certificado', f"{nome}_certificado.pdf")
        os.makedirs(os.path.dirname(caminho_pdf), exist_ok=True)  # Garante que a pasta exista
        imagem.convert('RGB').save(caminho_pdf, "PDF", resolution=100)
        return caminho_pdf

    def enviar_email(self, email_destinatario, caminho_certificado):
        # Configurações do servidor SMTP
        servidor_smtp = "smtp.gmail.com"
        porta_smtp = 587
        email_remetente = "insiraSeuEmail"
        senha_remetente = "insiraSuaSenha"

        # Configura o e-mail
        assunto = "Seu Certificado"
        corpo = "Segue em anexo o seu certificado. Obrigado por participar!"

        mensagem = MIMEMultipart()
        mensagem["From"] = email_remetente
        mensagem["To"] = email_destinatario
        mensagem["Subject"] = assunto

        mensagem.attach(MIMEText(corpo, "plain"))

        # Anexa o certificado
        with open(caminho_certificado, "rb") as anexo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(caminho_certificado)}",
        )
        mensagem.attach(parte)

        # Envia o e-mail
        try:
            with smtplib.SMTP(servidor_smtp, porta_smtp) as servidor:
                servidor.starttls()
                servidor.login(email_remetente, senha_remetente)
                servidor.sendmail(email_remetente, email_destinatario, mensagem.as_string())
            print(f"E-mail enviado para {email_destinatario} com sucesso!")
        except Exception as e:
            print(f"Erro ao enviar e-mail para {email_destinatario}: {str(e)}")

    def gerar_certificados(self):
        for _, linha in self.dados.iterrows():
            nome = linha['Nome']
            self.criar_certificado(nome)
            print(f'Certificado gerado para {nome}!')


inicio = EditarCertificado()
inicio.gerar_certificados()
