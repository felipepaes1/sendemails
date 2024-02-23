import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl

# Defina as credenciais e o endereço do servidor SMTP
smtp_server = "smtp.gmail.com"
port = 587
sender_email = ""
password = ""
recipient_emails = []  # Lista de e-mails
user_names = []
company_names = []

def take_mails():
    workbook = openpyxl.load_workbook('mailslist.xlsx')
    sheet = workbook.active

    # Limpe as listas existentes
    recipient_emails.clear()
    user_names.clear()
    company_names.clear()

    # Comece da segunda linha, já que a primeira contém os cabeçalhos
    for row in sheet.iter_rows(min_row=2):
        company_name = row[0].value  # Coluna A
        user_name = row[1].value  # Coluna B
        email = row[2].value  # Coluna C

        # Adicione os valores às listas
        recipient_emails.append(email)
        company_names.append(company_name)
        user_names.append(user_name)
def send_mail():
    try:
        # Crie a conexão com o servidor
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()  # Ative o modo TLS

        # Faça login na conta de e-mail
        server.login(sender_email, password)

        # Lê o template HTML
        with open('email_template_1.html', 'r', encoding='utf-8') as file:
            template_html = file.read()

        # Envie o e-mail para todos os destinatários
        for email, user_name, company_name in zip(recipient_emails, user_names, company_names):
            # Substitui os placeholders pelos valores reais
            personalized_html = template_html.replace("{user_name}", user_name).replace("{company_name}", company_name)

            # Crie a mensagem
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = email
            message["Subject"] = "Seu estoque de ferramentas está otimizado? Veja o que descobrimos"

            # Anexa a mensagem HTML personalizada
            message.attach(MIMEText(personalized_html, "html"))

            # Envie o e-mail
            server.sendmail(sender_email, email, message.as_string())

    except Exception as e:
        # Imprime qualquer erro que ocorrer
        print(f"Ocorreu um erro: {e}")

    finally:
        # Feche a conexão com o servidor, se possível
        server.quit()

def main():
    take_mails()
    send_mail()
    print("Emails enviados com sucesso")

if __name__ == "__main__":
    main()