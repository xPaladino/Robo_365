import os
from dotenv import load_dotenv
from O365 import FileSystemTokenBackend, Account
from datetime import datetime
from final.back.clientes.process_adm import process_adm
from final.back.clientes.process_chs import process_chs
from final.back.clientes.process_btg import process_btg
from final.back.clientes.process_bunge import process_bunge
from final.back.clientes.process_viterra import process_viterra
from final.back.clientes.process_cofco import process_cofco
from final.back.clientes.process_coamo import process_coamo
from final.back.clientes.process_royal import process_royal
from final.back.clientes.process_olam import process_olam
from final.back.clientes.process_cj import process_cj
from final.back.clientes.process_ldc import process_ldc

from final.back.SQL.save_to_excel import save_to_excel


def connect_o365(client_id, secret_id, tenant_id):
    token_backend = FileSystemTokenBackend(token_path='.', token_filename='aut/o365_token.txt')
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
    print(f"Busca entre {since_date_str} e {before_date_str}")

    messages = inbox.get_messages(limit=None,
                                  query=f'ReceivedDateTime ge {since_date_str} and ReceivedDateTime lt {before_date_str}',
                                  download_attachments=True)

    return messages


def main(output_dir, search_date, search_date_end):
    """
    Função principal, responsável pela conexão à conta do Office 365, buscando e processando os e-mails dentro de um
    intervalo de data especificado, e salvando os resultados em um arquivo Excel.

    Retorna:

    - nf_pdf_count(int): Contagem de PDF's processados.
    - resultados_pdf_count(int): Contagem de resultados de PDF's.
    - nf_zip_count(int): Contagem de arquivos ZIP's processados.
    - resultados_zip_count(int): Contagem de resultados de ZIP's.
    - nf_excel_count(int): Contagem de arquivos Excel processados.
    - resultados_excel_count(int): Contagem de resultados de Excel
    - Em caso de erro, retorna None
    :param output_dir: O diretório onde o arquivo Excel resultante será salvo.
    :param search_date: Data de início para busca dos e-mails, formato YYYY-MM-DD.
    :param search_date_end: Data de fim para busca dos e-mails, formato YYYY-MM-DD.

    """
    load_dotenv()
    client_id = os.getenv('O365_CLIENT_ID')
    secret_id = os.getenv('O365_SECRET_ID')
    tenant_id = os.getenv('O365_TENANT_ID')

    account = connect_o365(client_id, secret_id, tenant_id)

    since_date = datetime.strptime(search_date, '%Y-%m-%d')
    before_date = datetime.strptime(search_date_end, '%Y-%m-%d')

    messages = fetch_emails_o365(account, since_date, before_date)

    nf_excel_map = []
    nf_pdf_map = {}
    nf_zip_map = {}

    for message in messages:
        process_adm(message, './attachments', nf_pdf_map)
        process_btg(message, './attachments', nf_zip_map)
        process_bunge(message, './attachments', nf_excel_map)
        process_chs(message, './attachments', nf_pdf_map)
        process_cj(message, '.attachments', nf_pdf_map, nf_excel_map)
        process_coamo(message, './attachments', nf_pdf_map, nf_zip_map)
        process_cofco(message, './attachments', nf_pdf_map)
        process_ldc(message, './attachments', nf_zip_map)
        process_olam(message, './attachments', nf_excel_map)
        process_royal(message, './attachments', nf_pdf_map)
        process_viterra(message, './attachments', nf_pdf_map, nf_zip_map)

    try:
        nf_pdf_count, resultados_pdf_count, \
            nf_zip_count, resultados_zip_count, \
            nf_excel_count, resultados_excel_count \
            = save_to_excel(nf_pdf_map, nf_zip_map, nf_excel_map, os.path.join(output_dir, 'notas_fiscais365.xlsx'))
        return nf_pdf_count, resultados_pdf_count, nf_zip_count, resultados_zip_count, nf_excel_count, \
            resultados_excel_count
    except Exception as e:
        print(f'erro {e}')
        return None, None
