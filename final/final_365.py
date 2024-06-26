import os
from dotenv import load_dotenv
from O365 import FileSystemTokenBackend, Account
from datetime import datetime
from processa_pdf import process_attachments
from process_viterra import process_viterra
from process_coamo import process_coamo
from save_to_excel import save_to_excel

def connect_o365(client_id, secret_id, tenant_id):
    token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
    credentials = (client_id, secret_id)
    account = Account(credentials, tenant_id=tenant_id, token_backend=token_backend)

    if not account.is_authenticated:
        result = account.authenticate(scopes=['basic', 'message_all'])
        if not result:
            raise RuntimeError('Falha na autenticação com O365.')

    return account

def fetch_emails_o365(account, since_date, before_date):
    mailbox = account.mailbox()
    inbox = mailbox.inbox_folder()

    since_date_str = since_date.strftime('%Y-%m-%dT%H:%M:%SZ')
    before_date_str = before_date.strftime('%Y-%m-%dT%H:%M:%SZ')

    messages = inbox.get_messages(query=f'ReceivedDateTime ge {since_date_str} and ReceivedDateTime lt {before_date_str}',
                                  download_attachments=True)

    return messages


if __name__ == "__main__":

    load_dotenv()
    client_id = os.getenv('O365_CLIENT_ID')
    secret_id = os.getenv('O365_SECRET_ID')
    tenant_id = os.getenv('O365_TENANT_ID')

    account = connect_o365(client_id, secret_id, tenant_id)

    since_date = datetime(2024, 6, 1)
    before_date = datetime(2024, 6, 30)

    messages = fetch_emails_o365(account, since_date, before_date)

    nf_excel_map = []
    nf_pdf_map = {}
    nf_zip_map = {}

    for message in messages:
        #check_email_content(message, pattern_to_search)
        process_attachments(message, './attachments', nf_pdf_map)
        process_viterra(message,'./attachments', nf_pdf_map)
        process_coamo(message,'./attachments', nf_zip_map)

    save_to_excel(nf_pdf_map,nf_zip_map, 'notas_fiscais365.xlsx')


