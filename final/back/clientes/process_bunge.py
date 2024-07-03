import os
import re
import base64
import tempfile
import zipfile

import pandas as pd
from PyPDF2 import PdfReader

def process_bunge(message, salve_folder, nf_excel_map):
    corpo_email = message.body
    print(f"{message.received} - {message.subject}")
    if re.search(r'@bunge\.com', corpo_email):
        print('Tem Bunge')
        if message.attachments:
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                decoded_content = base64.b64decode(attachment.content)
                if file_extension == ".xlsx":
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_xlsx:
                        temp_xlsx.write(decoded_content)
                        pd.set_option('display.max_columns', None)
                        pd.set_option('display.max_rows', None)
                        df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=0)
                        df.columns = df.columns.str.strip()

                        if 'NF ref.' in df.columns and 'Chave de acesso NF de ref.' in df.columns:
                            cnpj_bunge = '84046101028101'
                            nf_ref_cleaned = df['NF ref.'].dropna().astype(str)
                            nf_ref_cleaned = [nf.split('.')[0] for nf in nf_ref_cleaned]
                            chave_ref_clean = df['Chave de acesso NF de ref.'].dropna().astype(
                                str).str.extract(
                                r'(\d{44})')
                            chave_ref_clean = chave_ref_clean.dropna().values.tolist()
                            nfe_ref = df['NFE'].dropna().astype(str)
                            nfe_ref = [nf.split('.')[0] for nf in nfe_ref]
                            chave_nfe_ref = df['Chave de acesso de 44 posições'].dropna().astype(
                                str).str.extract(
                                r'(\d{44})')
                            chave_nfe_ref = chave_nfe_ref.dropna().values.tolist()
                            serie_nfe = df['Séries'].dropna().astype(str)
                            serie_nfe = [serie.split('.')[0] for serie in serie_nfe]
                            dif_peso = df['Diferença de peso'].dropna().astype(str)
                            dif_peso = [dif.split('.')[0] for dif in dif_peso]
                            cnpj_nfe = df['CNPJ'].dropna().astype(str)
                            cnpj_nfe = [cn.split('.')[0] for cn in cnpj_nfe]
                            replica_nfe = []
                            replica_chave_nfe_ref = []
                            replica_serie_nfe = []
                            replica_chave = []
                            replica_cnpj = []

                            for i, nf in enumerate(nf_ref_cleaned):
                                if i < len(nfe_ref):
                                    for match in chave_ref_clean:
                                        replica_nfe.append(nfe_ref[i % len(nfe_ref)])

                                        if i < len(chave_nfe_ref):
                                            replica_chave_nfe_ref.append(
                                                chave_nfe_ref[i % len(chave_nfe_ref)])
                                            replica_chave.append(match)

                                        replica_serie_nfe.append(serie_nfe[i % len(serie_nfe)])
                                        replica_cnpj.append(cnpj_nfe[i % len(cnpj_nfe)])


                            for nf, chave, nfe, chave_comp, peso_bungue, cnpj in zip(
                                    nf_ref_cleaned,
                                    chave_ref_clean, replica_nfe, replica_chave_nfe_ref, dif_peso, replica_cnpj):

                                chaves = ''.join(chave)
                                chaves_comp = ''.join(chave_comp)

                                nf_excel_map.append({'nota_fiscal': nf,
                                                     'data_email': message.received,
                                                     'chave_acesso': chaves,
                                                     'email_vinculado': message.subject,
                                                     'serie_nf': '0',
                                                     'data_emissao': '0',
                                                     'nfe': nfe,
                                                     'cnpj': cnpj_bunge,
                                                     'chave_comp': chaves_comp,
                                                     'peso_nfe': peso_bungue,
                                            'transportadora': 'BUNGE'})  # peso_bungue})

                    os.remove(temp_xlsx.name)
                else:
                    pass