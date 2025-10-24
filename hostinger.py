import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import threading
import pandas as pd

# Configurações de múltiplos servidores SMTP
smtp_servers = [
    ('smtp.hostinger.com', 587, 'prospeccao@mehsolucoes.com', 'Lum@1001'),
    ('smtp.hostinger.com', 587, 'vendas@mehsolucoes.com', 'Lum@1001'),
    ('smtp.hostinger.com', 587, 'euluccas@mehsolucoes.com', 'Lum@1001'),
    ('smtp.hostinger.com', 587, 'marketing@mehsolucoes.com', 'Lum@1001'),
    ('smtp.hostinger.com', 587, 'floresluccas@mehsolucoes.com', 'Lum@1001')
]

# Corpo do email em HTML
corpo_email = """
<html>
  
</html>
"""

# Função para enviar e-mail
def enviar_email(destinatario, nome, corpo_email, current_index):
    try:
        smtp_server, smtp_port, smtp_user, smtp_password = smtp_servers[current_index]

        # Conectar ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)

        # Criar e-mail
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = destinatario
        msg['Subject'] = 'Um robô pode ser a sua solução!'
        msg.attach(MIMEText(corpo_email.format(nome=nome), 'html'))

        # Anexar a imagem de assinatura
        try:
            with open("assinatura.png", 'rb') as img_file:
                img = MIMEImage(img_file.read())
                img.add_header('Content-ID', '<assinatura>')
                img.add_header('Content-Disposition', 'inline', filename='assinatura.png')
                msg.attach(img)
        except FileNotFoundError:
            print("Aviso: Arquivo 'assinatura.png' não encontrado. Email será enviado sem a imagem.")

        # Enviar e-mail
        server.sendmail(smtp_user, destinatario, msg.as_string())
        server.quit()
        print(f"Email enviado para {destinatario} usando {smtp_user}")
        return True, current_index
    except Exception as e:
        print(f'Erro ao enviar email para {destinatario}: {str(e)}')

        # Verificar erro de limite e trocar de servidor
        if "ratelimit" in str(e).lower():
            next_index = (current_index + 1) % len(smtp_servers)
            print(f'Trocando para próximo servidor: {smtp_servers[next_index][2]}')
            return enviar_email(destinatario, nome, corpo_email, next_index)

        return False, current_index

# Função para enviar emails para todos os leads
def send_emails():
    df = pd.read_excel('lead.xlsx')  # Supondo que o arquivo tenha colunas 'Email' e 'Nome'
    current_index = 0  # Índice do servidor SMTP atual

    for index, row in df.iterrows():
        destinatario = row['Email']

        nome = row['Nome']

        success, current_index = enviar_email(destinatario, nome, corpo_email, current_index)
        if not success:
            print(f"Não foi possível enviar e-mail para {destinatario}.")

# Função para iniciar envio de e-mails em uma thread separada
def start_sending():
    threading.Thread(target=send_emails).start()

# Iniciar o envio
start_sending()

