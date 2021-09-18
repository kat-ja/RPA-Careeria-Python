"""Read Exchange mail messages, make a table of specific messages and send that information to user."""

from RPA.Email.Exchange import Exchange
from RPA.Robocorp.Vault import Vault
from RPA.Excel.Files import Files
from RPA.Tables import Tables

from ParseDate import ParseDate
from ParseTime import ParseTime

mail = Exchange()
vault = Vault()
files = Files()
tables = Tables()
pd = ParseDate()
pt = ParseTime()


# authenticate to mail account


def auth():
    _secret = vault.get_secret("credentials")
    account = _secret["account"]
    password = _secret["password"]
    mail.authorize(account, password)


# get hoks-messages from inbox


def get_teacher_messages():
    teacher_messages = []
    messages = mail.list_messages()
    for item in messages:
        if item["sender"]["name"] == "Katja Valanne":
            teacher_messages.append(item)
    return teacher_messages


# get only necessary information from message, adjust data to suit excel table


def parse_teacher_messages(t_messages):
    parsed_teacher_messages = []
    for item in t_messages:

        # hoks-asian käsittely
        hoks = ''
        if "HOKS" in item["subject"] or "HOPS" in item["subject"]:
            hoks = 'hoks'

        # liitteiden käsittely
        attachments = []
        str_attachments = ''
        if len(item["attachments_object"]) > 0:
            for i in item["attachments_object"]:
                attachments.append(getattr(i, 'name'))
            if len(attachments) > 1:
                str_attachments = ', '.join(attachments)
            else:
                str_attachments = ''.join(attachments)
        else:
            str_attachments = 'Ei liitteitä.'

        # lista kustakin rivistä
        parsed_teacher_messages.append([pd.parse_date(str(item['datetime_received'])), pt.parse_time(
            str(item['datetime_received'])), item['subject'], item['text_body'], str_attachments, hoks])

    return parsed_teacher_messages


# get the amount of HOKS-messages


def get_hoks_count(t_messages):
    hoks_count = 0
    for item in t_messages:
        if "HOKS" in item["subject"] or "HOPS" in item["subject"]:
            hoks_count += 1
    return hoks_count


# make table for excel


def make_table(parsed_t_messages):
    columns = ['date', 'time', 'subject', 'mailbody', 'attachments', 'hoks']
    table_for_excel = tables.create_table(data=parsed_t_messages, columns=columns)
    return table_for_excel


# make excel file with table


def make_excel_with_table(table_for_excel):
    files.create_workbook()
    files.append_rows_to_worksheet(content=table_for_excel, header=True)
    files.save_workbook('messages.xlsx')


# viestin muotoilu


def format_message(parsed_t_messages, hoks_c):
    str_message = (
        f'<p>Opettaja Katja Valanne on lähettänyt minulle Katja Valanne {len(parsed_t_messages)} kpl sähköpostiviestejä. Tässä lista viesteistä: </p>'
        f'<u>HOKS-viestejä {hoks_c} kappaletta</u></p>'
    )
    hoks_messages = []
    other_messages = []
    count_h = 0
    count_o = 0
    for item in parsed_t_messages:
        str_item = f'Päivämäärä: {item[0]}, Aihe: {item[2]}, Liitetiedostot: {item[4]}'
        if "HOKS" in item[2] or "HOPS" in item[2]:
            count_h += 1
            hoks_messages.append(f'- Viesti {count_h}: {str_item}<br>')
        else:
            count_o += 1
            other_messages.append(f'- Viesti {hoks_c + count_o}: {str_item}<br>')
    str_message += ''.join(hoks_messages)
    str_message += f'<p><u> Muita viestejä {len(parsed_t_messages)-hoks_c} kappaletta: </u></p>'
    str_message += ''.join(other_messages)
    return str_message


# viestin lähetys


def send_message(str_message):
    mail.send_message(
        recipients = "katja.valanne@gmail.com",
        subject = "Message from RPA Python",
        body = str_message,
        html = True
    )


def main():
    auth()
    t_messages = get_teacher_messages()
    parsed_t_messages = parse_teacher_messages(t_messages)
    hoks_c = get_hoks_count(t_messages)
    table_for_excel = make_table(parsed_t_messages)
    make_excel_with_table(table_for_excel)
    str_message = format_message(parsed_t_messages, hoks_c)
    send_message(str_message)


if __name__ == "__main__":
    main()
