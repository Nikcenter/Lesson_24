from pathlib import Path
from docx import Document
import logging
import time
import csv
import smtplib
import ssl
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.text import MIMEText  # Для підтримки кирилиці
from email.header import Header

# Налаштування логування для відстеження процесу виконання
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
logging.debug('Початок виконання програми')

# Файли CSV зі списком регіонів та електронних адрес
QUERY_LIST_FILE_EMAILS = Path(f'./your_emails.csv')
QUERY_LIST_FILE_REGIONS = Path(f'./your_regions_list.csv')

# Текст, який буде міститися у тілі електронного листа
body = "Доброго дня,\nНадсилаю інформаційний запит згідно ЗУ Про доступ до публічної інформації.\nПрошу зареєструвати та повідомити вхідний номер мого запиту.\nДякую за розуміння та співпрацю!\nЗ повагою,\n..."

def get_keywords(query_file):
    """Зчитує дані з CSV-файлу та повертає список значень"""
    with open(query_file, 'r') as i_file:
        rows = csv.reader(i_file, delimiter=',')
        keywords = [row[0] for row in rows]
        return keywords

def get_para_data(output_doc_name, paragraph):
    """Копіює вміст та форматування абзацу у новий документ"""
    output_para = output_doc_name.add_paragraph()
    for run in paragraph.runs:
        output_run = output_para.add_run(run.text)
        output_run.bold = run.bold
        output_run.italic = run.italic
        output_run.underline = run.underline
        output_run.font.color.rgb = run.font.color.rgb
        output_run.style.name = run.style.name
    output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment

def sendEmail(sending_file, file_name, email):
    """Відправляє електронний лист із вкладеним файлом"""
    time.sleep(2)
    account = 'test.nikcenter.org'
    sender = 'test@nikcenter.org'
    sender_name = 'nikcenter'
    password = 'XXXXX'  # Тут потрібно вказати правильний пароль
    
    logging.debug(f'Вхід у поштовий сервер для відправки {email}')
    context = ssl.create_default_context()
    smtpObj = smtplib.SMTP_SSL('mail.nikcenter.org', 465, context=context)
    smtpObj.ehlo()
    smtpObj.login(account, password)
    
    msg = MIMEMultipart()
    msg['Subject'] = Header('Запит на доступ до публічної інформації', 'utf-8')
    msg.attach(MIMEText(body.encode('utf-8'), _charset='utf-8'))
    
    with open(sending_file, 'rb') as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {file_name}")
        msg.attach(part)
    
    logging.debug(f'Надсилання email до {email}')
    sendmailStatus = smtpObj.sendmail(sender, email, msg.as_string())
    
    if sendmailStatus != {}:
        logging.error(f'Помилка при відправці email до {email}: {sendmailStatus}')
    
    smtpObj.quit()

def formingDocx(i):
    """Створює Word-документ для конкретного регіону та надсилає його"""
    input_doc = Document('./zapyt.docx')  # Використовуємо шаблон
    output_doc = Document()
    stringToPut = f'До ГУНП в {REGIONS[i]} області'  # Додаємо регіон
    paragraph = output_doc.add_paragraph(stringToPut)
    paragraph.alignment = 2  # Вирівнюємо текст справа
    
    for para in input_doc.paragraphs:
        get_para_data(output_doc, para)
    
    output_name = f'zapytGUNP{str(i)}.docx'
    output_name_fullPath = f'./{output_name}'
    output_doc.save(output_name_fullPath)
    logging.debug(f"Документ збережено: {output_name}")
    
    sendEmail(output_name_fullPath, output_name, EMAILS[i])

# Завантаження списків регіонів та email-адрес
REGIONS = get_keywords(QUERY_LIST_FILE_REGIONS)
EMAILS = get_keywords(QUERY_LIST_FILE_EMAILS)

# Генерація документів та відправка листів
for i in range(len(EMAILS)):
    formingDocx(i)

logging.debug("Виконання завершено!")
print("Готово!")

