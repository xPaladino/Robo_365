import os
import re
import base64
import tempfile
import zipfile

from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_coamo(message, salve_folder, nf_zip_map):
    corpo_email = message.body
    print(f"{message.received} - {message.subject}")
    if re.search(r'@coamo\.com\.br', corpo_email):
        print('tem coamo')
        if message.attachments:
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    if file_extension == ".zip":
                        decoded_content = base64.b64decode(attachment.content)
                        print(f"Processando: {attachment.name}")
                        with tempfile.TemporaryDirectory() as tmp_dir:
                            zip_file_path = os.path.join(tmp_dir,'attachment.zip')
                            with open(zip_file_path,'wb') as f:
                                f.write(decoded_content)
                            with zipfile.ZipFile(zip_file_path,'r') as zip_ref:
                                zip_file_names = zip_ref.namelist()
                                zip_ref.extractall(tmp_dir)
                            for file_name in os.listdir(tmp_dir):
                                file_path = os.path.join(tmp_dir, file_name)
                                if file_path.endswith('.pdf'):
                                    with open(file_path,'rb') as pdf_file:
                                        pdf_reader = PdfReader(pdf_file)
                                        pdf_text = ''
                                        for page in pdf_reader.pages:
                                            pdf_text += page.extract_text()
                                            notas_fiscais = []
                                            serie_matches = []
                                            segundo_cnpj = []
                                            cnpj_match = []
                                            nfe_match = []
                                            notas_fiscais.extend(
                                                re.finditer(r'(?:número:|REF\s+A\s+NOTA'
                                                            r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
                                                            r'|referente\s+a\s+NF|'
                                                            r'NF:)'
                                                            r'\s*(?:\d+\s*,\s*)?(\d{4,})'
                                                            r'|Nota\s+Fiscal:\s*(?:\d+\s*,\s*)?(\d{4,})', pdf_text,
                                                            re.IGNORECASE))
                                            serie_matches.extend(
                                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})'
                                                            , pdf_text, re.IGNORECASE))

                                            if not serie_matches:
                                                serie_matches.append(0)


                                            # print(f'Series {serie_matches}')
                                            chave_acesso_match = re.findall(
                                                r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b',
                                                pdf_text, re.IGNORECASE)

                                            if not chave_acesso_match:
                                                chave_acesso_match.append(0)
                                            # print(f'Chave {chave_acesso_match}')
                                            cnpj_match.extend(
                                                re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_text,
                                                            re.IGNORECASE))
                                            # print(f'Cnpj {cnpj_match}')
                                            if len(cnpj_match) >= 2:
                                                segundo_cnpj.append(cnpj_match[1])
                                            if not segundo_cnpj:
                                                segundo_cnpj.append(0)
                                            # print(f'Segundo {segundo_cnpj}')
                                            nfe_match.extend(
                                                re.finditer(r'(?:NF-e\s+Nº\s+|NF-e\s+Nº.\s+)(\d{3}.\d{3}.\d{3})|'
                                                            r'\d{3}\.\d{3}\.\d{3}', pdf_text,
                                                            re.IGNORECASE))
                                            if not nfe_match:
                                                nfe_match.append(0)

                                            replica_nota = []
                                            # print(notas_fiscais)
                                            for i in notas_fiscais:
                                                if i != 0:
                                                    if i != 'e':
                                                        notas = None
                                                        for x in range(1, 3):
                                                            notas = i.group(x)
                                                            if notas is not None:
                                                                break
                                                        replica_nota.append(notas)

                                            #print(f'teste cnpj segundo {segundo_cnpj}')
                                            ##print(f'teste nfe match {nfe_match}')
                                            #print(f'teste series {serie_matches}')
                                            #print(f'teste chave {chave_acesso_match}')

                                            for nf, nfe, chave, cnpj, serie in zip(replica_nota, nfe_match,chave_acesso_match,segundo_cnpj, serie_matches):
                                                if isinstance(nfe.group(0),str):
                                                    #print('sim')
                                                    padraonfe = nfe.group(0)
                                                    nfe_pont = re.sub(r'[.]','',padraonfe)
                                                    nfe_ajust = nfe_pont if nfe_pont[0:3] != '000' else nfe_pont[3:]

                                                chaves = ''.join(chave.split())
                                                cnpj_tratado = re.sub(r'[./-]','',cnpj.group(1))

                                                for i in range(1,4):
                                                    series = serie.group(i)
                                                    if series is not None:
                                                        break

                                                nf_zip_map[nf] = {
                                                    'nota_fiscal': nf,
                                                    'data_email': message.received,
                                                    'chave_comp': chaves,
                                                    'chave_acesso': '0',
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': series,
                                                    'data_emissao': '0',
                                                    'cnpj': cnpj_tratado,
                                                    'nfe': nfe_ajust
                                                }
