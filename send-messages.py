import os
import smtplib
import time
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename

from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv()


def get_values(
    col: list,  # Колонка excel-файла
) -> list:
    """Программа для получения данных из колонок excel-файла."""
    storage_list = []
    for cell in col:
        if cell.value is not None:
            storage_list.append(cell.value)
    return storage_list


def send_msg(
    from_addr: str,  # Ящик, с которого будет отправляться почта
    from_passw: str,  # Пароль от ящика
    smtp_server: str,  # Смтп-сервер, к которому относится ящик для отправки
    smtp_port: int,  # Порт для соединения с смтп-сервреом
    subject: str,  # Значение для поля "Тема"
    emails: list,  # Список эмейлов, на которые будет отправлено письмо
    text: str,  # Текст письма
    files=None  # Список путей для файлов вложений
) -> None:
    """
    Скрипт для рассылки почты по списку из excel-файла.
    Можно добавлять вложения, один или несколько файлов,
    в любом случае пути до файлов должны передаваться списком.
    """
    print(f'Соединяюсь с сервером {smtp_server}...')
    try:
        mail_lib = smtplib.SMTP_SSL(smtp_server, smtp_port)
        print('ОК')
    except Exception as e:
        print(f'Не могу соединиться с сервером. Ошибка {e}')
    print('Авторизуюсь на сервере...')
    try:
        mail_lib.login(from_addr, from_passw)
        print('ОК')
    except Exception as e:
        print('Не могу авторизоваться на сервере '
              f'по предоставленным данным. Ошибка {e}')
    # Создаём сообщение
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = ''
    msg['Subject'] = Header(
        subject,
        'utf-8'
    ).encode()

    msg.attach(MIMEText(text))
    # Принимаем файлы для вложений
    for file in files or []:
        with open(file, 'rb') as work_file:
            part = MIMEApplication(
                work_file.read(),
                name=basename(file)
            )
        part[
            'Content-Disposition'
        ] = 'attachment; filename="%s"' % basename(file)
        msg.attach(part)
    sent_emails = 0  # Счётчик отправленных сообщений
    for email in emails:
        # Смтп-сервер Яндекса банит,
        # если количество запросов в минуту превышает 20
        if (sent_emails + 1) % 20 == 0:
            print(
                f'Количество сообщений достигло {sent_emails}. '
                'Сплю минуту.'
            )
            time.sleep(60)
        print(
            f'Отправляю сообщение на адрес {email}. '
            f'Отправлено {sent_emails + 1} из {len(emails)} сообщений.'
        )
        mail_lib.sendmail(from_addr, email, msg.as_string())
        sent_emails += 1
    print('Все сообщения отправлены')
    mail_lib.close()


if __name__ == '__main__':
    # Подключаем нашу таблицу Excel
    xl_file = load_workbook('file_data/emails.xlsx')
    sheet = xl_file.active
    emails_list = get_values(sheet['A'])  # Наши эмейлы
    send_msg(
        from_addr=os.environ.get('EMAIL'),
        from_passw=os.environ.get('PASSWORD'),
        smtp_server='',  # smtp-server address
        smtp_port=,  # smt
        subject='Инструкция по переходу на новый почтовый сервис',
        emails=emails_list,
        text=('Добрый день, высылаю Вам инструкцию '
                'для перехода на новый почтовый сервис.'),
        files=['file_data/syncinstr.pdf', ]
    )
