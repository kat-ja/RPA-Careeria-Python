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


teacher_messages = []
parsed_teacher_messages = []
HOKS_messages = []

'''
'subject': 'Olet liittynyt ryhmään Testi',
'sender': {'name': 'Testi', 'email_address': 'Testi652@iukky.onmicrosoft.com'}, 
'datetime_received': EWSDateTime(2021, 8, 31, 9, 45, 5, tzinfo=EWSTimeZone(key='UTC')),
'folder': 'Inbox (Saapuneet)', 
'body': '<!doctype html><html><head>\r\n<meta http-equiv="Content-Type" content="text/html;

sender[name]
tarvitaan taulukkoon:
datetime_received
subject
text_body
'attachments_object'[name]
hoks
'''

# authenticate to mail account
def auth():  
    _secret = vault.get_secret("credentials")
    account = _secret["account"]
    password = _secret["password"]
    mail.authorize(account, password)

# get hoks-messages from inbox
def get_teacher_messages():  
    messages = mail.list_messages()
    for item in messages:
        if item["sender"]["name"] == "Katja Valanne":
            teacher_messages.append(item)
    return teacher_messages

# get only necessary information from message, adjust data to suit excel table
def parse_teacher_messages(t_messages): 
    for item in t_messages:

        # hoks-asian käsittely
        hoks = ''
        if "HOKS" in item["subject"] or "HOPS" in item["subject"]:
            #HOKS_messages.append(item)
            hoks = 'hoks'

        # liitteiden käsittely
        attachments = []
        str_attachments = ''
        if len(item["attachments_object"]) > 0:
            for i in item["attachments_object"]:
                attachments.append(getattr(i, 'name'))
            str_attachments = ''.join(attachments)
        else:
            #attachments.append('Ei liitteitä.')
            str_attachments = 'Ei liitteitä.'
        
        parsed_teacher_messages.append([pd.parse_date(str(item['datetime_received'])), pt.parse_time(
            str(item['datetime_received'])), item['subject'], item['text_body'], str_attachments, hoks])

    return parsed_teacher_messages

# make table for excel
def make_table(parsed_t_messages):
    columns = ['date', 'time', 'subject', 'mailbody', 'attachments', 'hoks']
    table_for_excel = tables.create_table(data = parsed_teacher_messages, columns = columns)
    return table_for_excel

# make excel file with table
def make_excel_with_table(table_for_excel):
    files.create_workbook()
    files.append_rows_to_worksheet(content = table_for_excel, header = True)
    files.save_workbook('messages.xlsx')


def get_hoks_messages(t_messages):
    for item in t_messages:
        if "HOKS" in item["subject"] or "HOPS" in item["subject"]:
            HOKS_messages.append(item)


def main():
    auth()
    t_messages = get_teacher_messages()
    parsed_t_messages = parse_teacher_messages(t_messages)
    table_for_excel = make_table(parsed_t_messages)
    make_excel_with_table(table_for_excel)


if __name__ == "__main__":
    main()
