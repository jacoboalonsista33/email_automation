import os
import time
import pandas as pd
import smtplib
from email.message import EmailMessage

TEMPLATE_DIR = 'templates'
ATTACHMENT = 'presentation.pdf'
CONTACT_FILE = 'contacts.xlsx'
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.example.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587'))
SMTP_USER = os.environ.get('SMTP_USER', 'user@example.com')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD', 'password')
FROM_ADDRESS = os.environ.get('FROM_ADDRESS', 'jaime.valero@example.com')


def load_template(name: str) -> str:
    path = os.path.join(TEMPLATE_DIR, name)
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


INITIAL_TEMPLATE = load_template('initial.txt')
FOLLOW1_TEMPLATE = load_template('follow1.txt')
FOLLOW2_TEMPLATE = load_template('follow2.txt')
FOLLOW3_TEMPLATE = load_template('follow3.txt')


def fill_template(template: str, nombre: str, empresa: str) -> str:
    return (
        template
        .replace('[Nombre]', nombre)
        .replace('[Persona]', nombre)
        .replace('[Empresa]', empresa)
    )


def send_email(to_address: str, subject: str, body: str, attachment: str | None = None) -> None:
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = FROM_ADDRESS
    msg['To'] = to_address
    msg.set_content(body)

    if attachment:
        with open(attachment, 'rb') as f:
            data = f.read()
            filename = os.path.basename(attachment)
            msg.add_attachment(data, maintype='application', subtype='octet-stream', filename=filename)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)


def load_contacts(path: str) -> list[dict]:
    df = pd.read_excel(path)
    contacts = []
    for _, row in df.iterrows():
        contacts.append({'nombre': row['Nombre'], 'empresa': row['Empresa'], 'email': row['Email']})
    return contacts


def send_initial_emails(contacts: list[dict]) -> None:
    for c in contacts:
        body = fill_template(INITIAL_TEMPLATE, c['nombre'], c['empresa'])
        subject = f"Propuesta estratégica para {c['empresa']}"
        send_email(c['email'], subject, body, ATTACHMENT)


def send_followup_emails(contacts: list[dict], template: str) -> None:
    for c in contacts:
        body = fill_template(template, c['nombre'], c['empresa'])
        subject = f"Seguimiento de propuesta para {c['empresa']}"
        send_email(c['email'], subject, body)


def main() -> None:
    contacts = load_contacts(CONTACT_FILE)
    send_initial_emails(contacts)

    for idx, template in enumerate([FOLLOW1_TEMPLATE, FOLLOW2_TEMPLATE, FOLLOW3_TEMPLATE], start=1):
        time.sleep(5 * 60)  # wait 5 minutes
        send_followup_emails(contacts, template)
        print(f"Follow-up {idx} sent")


if __name__ == '__main__':
    main()
