"""Read Exchange mail messages, make a table of specific messages and send that information to user."""

from RPA.Email.Exchange import Exchange
from RPA.Robocorp.Vault import Vault
from RPA.Excel.Files import Files

from ParseDate import ParseDate
from ParseTime import ParseTime

mail = Exchange()
vault = Vault()
files = Files()
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


def auth():  # authenticate to mail account
    _secret = vault.get_secret("credentials")
    account = _secret["account"]
    password = _secret["password"]
    mail.authorize(account, password)


def get_teacher_messages():  # get hoks-messages from inbox
    messages = mail.list_messages()
    for item in messages:
        if item["sender"]["name"] == "Katja Valanne":
            teacher_messages.append(item)
    return teacher_messages


def parse_teacher_messages(t_messages):
    for item in t_messages:
        # liitteiden käsittely
        attachments = []
        if len(item["attachments_object"]) > 0:
            for i in item["attachments_object"]:
                attachments.append(getattr(i, 'name'))
        else:
            attachments.append('Ei liitteitä.')
        # todo: add hoks data
        parsed_teacher_messages.append([pd.parse_date(str(item['datetime_received'])), pt.parse_time(
            str(item['datetime_received'])), item['subject'], item['text_body'], attachments])

    return parsed_teacher_messages

# todo


def get_hoks_messages(t_messages):
    for item in t_messages:
        if "HOKS" in item["subject"] or "HOPS" in item["subject"]:
            HOKS_messages.append(item)


def main():
    auth()
    t_messages = get_teacher_messages()
    parsed_t_messages = parse_teacher_messages(t_messages)
    print(parsed_t_messages)
    # get_hoks_messages()


if __name__ == "__main__":
    main()
