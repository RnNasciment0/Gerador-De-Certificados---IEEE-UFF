# Gerador de Certificados - IEEE-UFF

Este é um projeto para gerar certificados personalizados em formato PDF a partir de uma lista de participantes fornecida em um arquivo Excel (.xlsx). Os certificados gerados são compactados em um arquivo `.zip` e disponibilizados para download. Além disso, o sistema permite o envio dos certificados por e-mail.

## Funcionalidades

- Upload de arquivos Excel contendo a lista de participantes.
- Geração automática de certificados personalizados em PDF.
- Compactação dos certificados em um arquivo `.zip`.
- Envio de certificados por e-mail para os participantes.

## Tecnologias Utilizadas

- **Linguagem**: Python
- **Framework Web**: Flask
- **Manipulação de Dados**: Pandas
- **Geração de PDFs**: PIL (Pillow)
- **Envio de E-mails**: smtplib
- **Frontend**: HTML e CSS

## Estrutura do Projeto

- `app.py`: Arquivo principal que contém a aplicação Flask.
- `generate_certificate.py`: Contém a lógica para geração de certificados e envio de e-mails.
- `templates/index.html`: Página inicial para upload do arquivo Excel.
- `static/`: Contém arquivos estáticos como imagens e fontes.
- `uploads/`: Diretório para armazenar os arquivos Excel enviados.
- `certificado/`: Diretório onde os certificados gerados e o arquivo `.zip` são armazenados.

## Como Executar o Projeto

1. **Pré-requisitos**:
   - Python 3.8 ou superior.
   - Instale as dependências listadas no arquivo `requirements.txt`:
     ```bash
     pip install -r requirements.txt
     ```

2. **Configuração**:
   - Certifique-se de que as pastas `uploads` e `certificado` existem no diretório do projeto.
   - Configure as variáveis de ambiente para envio de e-mails:
     - `EMAIL_REMETENTE`: E-mail do remetente.
     - `SENHA_REMETENTE`: Senha ou senha de aplicativo do e-mail.

3. **Executar o servidor**:
   - Inicie o servidor Flask:
     ```bash
     python app.py
     ```
   - Acesse a aplicação no navegador em `http://127.0.0.1:5000`.

4. **Uso**:
   - Faça o upload de um arquivo Excel contendo uma coluna chamada `Nome` com os nomes dos participantes.
   - Baixe o arquivo `.zip` com os certificados gerados.

## Estrutura do Arquivo Excel

O arquivo Excel enviado deve conter, no mínimo, uma coluna chamada `Nome`, que será usada para personalizar os certificados.

Exemplo:

| Nome           |
|-----------------|
| João da Silva   |
| Maria Oliveira  |
| Carlos Santos   |

## Personalização

- **Imagem do Certificado**: Substitua o arquivo `static/certificado.jpg` pela imagem do certificado desejada.
- **Fonte**: Substitua o arquivo `static/BeVietnamPro-Regular.ttf` pela fonte desejada.

## Licença

Este projeto é de uso interno e não possui uma licença específica. Entre em contato com os desenvolvedores para mais informações.

## Créditos

Desenvolvido por Renan Nascimento.