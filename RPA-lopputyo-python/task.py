"""Read Exchange mail messages, make a table of specific messages and send that information to user."""

from RPA.Email.Exchange import Exchange
from RPA.Robocorp.Vault import Vault

mail = Exchange()
vault = Vault()

'''
'subject': 'Olet liittynyt ryhmään Testi',
'sender': {'name': 'Testi', 'email_address': 'Testi652@iukky.onmicrosoft.com'}, 
'datetime_received': EWSDateTime(2021, 8, 31, 9, 45, 5, tzinfo=EWSTimeZone(key='UTC')),
'folder': 'Inbox (Saapuneet)', 
'body': '<!doctype html><html><head>\r\n<meta http-equiv="Content-Type" content="text/html;
'''

def auth(): # authenticate to mail account
    _secret = vault.get_secret("credentials")
    account = _secret["account"]
    password = _secret["password"]
    mail.authorize(account, password)

def get_hoks_messages(): # get hoks-messages from inbox
    messages = mail.list_messages()
    #print(messages[0])
    #print(messages[1]["subject"])
    for item in messages:
        if item["sender"]["name"] == "Katja Valanne":
            print(item["subject"])

    # tulostetaan ensimmäinen listasta: nähdään rakenne
    # iterator = iter(messages)
    # print(next(iterator))
    # 
def main():
    auth()
    get_hoks_messages()

if __name__ == "__main__":
    main()
