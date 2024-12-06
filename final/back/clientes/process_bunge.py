import os
import re
import base64
import tempfile
import zipfile
from datetime import datetime
import pandas as pd
from PyPDF2 import PdfReader



def process_bunge(message, salve_folder, nf_excel_map, nf_zip_map):
    """
    Essa função é responsável por ler e tratar os dados vindos do Cliente BUNGE, extraído de um arquivo Excel,
    podendo ser alterado o estilo de captura da Nota Fiscal para se adequar ao projeto. A captura de um Excel é
    realizada pela coluna que é lida, diferente do PDF/ZIP que precisa achar um padrão.
    :param message: Variável pertencente a lista Messages.
    :param salve_folder: Local onde vai ser salvo o arquivo.
    :param nf_excel_map: Lista responsável por salvar os dados referente ao EXCEL.
    """
    corpo_email = message.body

    if re.search(r'@bunge\.com', corpo_email):
        print('Tem Bunge')
        if message.attachments:
            pdf_chave = []
            pdf_data = []
            pdf_series = []
            pdf_excel_chave = []
            pdf_excel_data = []
            pdf_excel_series = []
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                pdf_cont = 0
                zip_cont = 0
                if file_extension == ".zip":
                    decoded_content = base64.b64decode(attachment.content)
                    print('tem zip')
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        zip_file_path = os.path.join(tmp_dir, 'attachment.zip')
                        with open(zip_file_path, 'wb') as f:
                            f.write(decoded_content)

                        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)

                        for file_name in os.listdir(tmp_dir):
                            file_path = os.path.join(tmp_dir, file_name)

                            if os.path.isdir(file_path) and "rocha" in file_name.lower():
                                for file_dir in os.listdir(file_path):
                                    full_path = os.path.join(file_path, file_dir)
                                    if os.path.isfile(full_path):
                                        if file_dir.endswith('.pdf'):
                                            with open(full_path, 'rb') as dir_file:
                                                pdf_reader = PdfReader(dir_file)
                                                pdf_text = ''
                                                for page in pdf_reader.pages:
                                                    pdf_text += page.extract_text()

                                                chave_acesso = re.findall(
                                                    r'(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b)',
                                                    pdf_text, re.IGNORECASE
                                                )
                                                data_match = []
                                                data_match.extend(re.finditer(
                                                    r'DATA\s+DA\s+EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                                    r'DATA\s+DA\sEMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                                    r'(?:DATA\s+DA\sEMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                                    , pdf_text, re.IGNORECASE))

                                                if not data_match:
                                                    data_match.append(0)
                                                serie_matches = []
                                                serie_matches.extend(
                                                    re.finditer(
                                                        r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|(?:NF-e\s+\d{3}\.\d{3}\.\d{3}\s+)(\d{1,3})|'
                                                        r'(?:Nº\s*SÉRIE:\s*\d+)(\s+\d{1,3})',
                                                        pdf_text, re.IGNORECASE))
                                                if not serie_matches:
                                                    serie_matches.append(0)

                                                for serie in serie_matches:
                                                    try:
                                                        for x in range(1, 6):
                                                            series = serie.group(x)
                                                            if series is not None:
                                                                break
                                                    except IndexError:
                                                        series = 0
                                                    pdf_series.append(series)
                                                for data in data_match:
                                                    try:
                                                        if data != 0:
                                                            for i in range(1, 3):
                                                                data_ajust = data.group(i)
                                                                if data_ajust is not None:
                                                                    break
                                                    except IndexError:
                                                        data_ajust = 0
                                                    pdf_data.append(data_ajust)

                                                for chave in chave_acesso:
                                                    chaves = ''.join(chave.split())
                                                    pdf_chave.append(chaves)

                                        elif file_dir.endswith('.xlsx'):
                                            print("Achei um excel")
                                            excel_path = full_path
                                            try:
                                                pd.set_option('display.max_columns', None)
                                                pd.set_option('display.max_rows', None)
                                                df = pd.read_excel(excel_path, engine='openpyxl', header=0)
                                                df.columns = df.columns.str.strip()
                                                if 'NOTA' in df.columns and 'CHAVE' in df.columns:
                                                    cnpj_bunge = '84046101028101'
                                                    nf_ref = df['NOTA'].dropna().astype(str)
                                                    nf_ref = [nf.split('.')[0] for nf in nf_ref]

                                                    chave_ref_clean = df['CHAVE DE REFERÊNCIA'].dropna().astype(
                                                        str).str.extract(
                                                        r'(\d{44})')
                                                    chave_ref_clean = chave_ref_clean.dropna().values.astype(str)
                                                    replica_chave_nfe_ref = []
                                                    replica_data_nfe = []
                                                    replica_series_nfe = []
                                                    nf_ref_cleaned = df['NOTA DE REFERÊNCIA'].dropna().astype(str)
                                                    nf_ref_cleaned = [nf.split('-')[0] for nf in nf_ref_cleaned]
                                                    if pdf_chave:
                                                        for chave_ref in chave_ref_clean:
                                                            for chave, data, series in zip(pdf_chave, pdf_data,
                                                                                           pdf_series):
                                                                if chave_ref == chave:
                                                                    replica_chave_nfe_ref.append(chave)
                                                                    replica_data_nfe.append(data)
                                                                    replica_series_nfe.append(series)

                                                    for nf, chave_comp, nfe, chave_nfe, data, serie in zip(nf_ref,
                                                                                                           chave_ref_clean,
                                                                                                           nf_ref_cleaned,
                                                                                                           replica_chave_nfe_ref,
                                                                                                           replica_data_nfe,
                                                                                                           replica_series_nfe):
                                                        nf_excel_map[nf] = {
                                                            'nota_fiscal': nf,
                                                            'data_email': message.received,
                                                            'chave_acesso': '0',
                                                            'email_vinculado': message.subject,
                                                            'serie_nf': serie,
                                                            'data_emissao': data,
                                                            'nfe': nfe.lstrip('0'),
                                                            'cnpj': cnpj_bunge,
                                                            'chave_comp': chave_nfe,
                                                            'peso_nfe': '0',
                                                            'serie_comp': '0',
                                                            'peso_comp': '0',
                                                            'transportadora': 'BUNGE'}  # peso_bungue})
                                                else:
                                                    nf_excel_map[file_dir] = {
                                                        'nota_fiscal': '0',
                                                        'data_email': message.received,
                                                        'chave_acesso': 'SEM LEITURA',
                                                        'email_vinculado': message.subject,
                                                        'serie_nf': 'SEM LEITURA',
                                                        'data_emissao': 'SEM LEITURA',
                                                        'cnpj': 'SEM LEITURA',
                                                        'nfe': '0',
                                                        'chave_comp': 'SEM LEITURA',
                                                        'transportadora': 'BUNGE',
                                                        'peso_comp': 'SEM LEITURA',
                                                        'serie_comp': 'SEM LEITURA'
                                                    }

                                            except Exception as e:
                                                print(f'Erro ao abrir o Excel: {e}')
                if file_extension == ".pdf":
                    print(attachment.name)
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:

                        temp_pdf.write(decoded_content)
                        pdf_reader = PdfReader(temp_pdf.name)
                        pdf_text = ''
                        for page in pdf_reader.pages:
                            pdf_text += page.extract_text()

                        chave_acesso = re.findall(
                            r'(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b)',
                            pdf_text, re.IGNORECASE
                        )
                        data_match = []
                        data_match.extend(re.finditer(
                            r'DATA\s+DA\s+EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                            r'DATA\s+DA\sEMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                            r'(?:DATA\s+DA\sEMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                            , pdf_text, re.IGNORECASE))

                        if not data_match:
                            data_match.append(0)
                        serie_matches = []
                        serie_matches.extend(
                            re.finditer(
                                r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|(?:NF-e\s+\d{3}\.\d{3}\.\d{3}\s+)(\d{1,3})|'
                                r'(?:Nº\s*SÉRIE:\s*\d+)(\s+\d{1,3})',
                                pdf_text, re.IGNORECASE))
                        if not serie_matches:
                            serie_matches.append(0)

                        for serie in serie_matches:
                            try:
                                if isinstance(serie,re.Match):
                                    for x in range(1, 6):
                                        series = serie.group(x)
                                        if series is not None:
                                            break
                                else:
                                    series = 0
                            except (IndexError,AttributeError):
                                series = 0
                            pdf_excel_series.append(series)

                        for data in data_match:
                            try:
                                if isinstance(data, re.Match):
                                    if data != 0:
                                        for i in range(1, 3):
                                            data_ajust = data.group(i)
                                            if data_ajust is not None:
                                                break
                                else:
                                    data_ajust = 0
                            except (IndexError,AttributeError):
                                data_ajust = 0
                            pdf_excel_data.append(data_ajust)

                        for chave in chave_acesso:
                            chaves = ''.join(chave.split())
                            pdf_excel_chave.append(chaves)
                    os.remove(temp_pdf.name)

            for attachment2 in message.attachments:
                file_extension = os.path.splitext(attachment2.name)[1].lower()
                if file_extension == ".xlsx":
                    decoded_content = base64.b64decode(attachment2.content)
                    print(f'{message.received} - {message.subject}')
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_xlsx:
                        temp_xlsx.write(decoded_content)
                        pd.set_option('display.max_columns', None)
                        pd.set_option('display.max_rows', None)
                        df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=0)
                        df.columns = df.columns.str.strip()

                        if 'NF ref.' in df.columns and 'Chave de acesso NF de ref.' in df.columns:
                            cnpj_bunge = '84046101028101'
                            nf_ref_cleaned = df['NF ref.'].dropna().astype(str)
                            nf_ref_cleaned = [nf.split('.')[0].split('-')[0] for nf in nf_ref_cleaned]
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
                                nf_excel_map[nf] = {'nota_fiscal': nf,
                                                    'data_email': message.received,
                                                    'chave_acesso': chaves,
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': '0',
                                                    'data_emissao': '0',
                                                    'nfe': nfe,
                                                    'cnpj': cnpj_bunge,
                                                    'chave_comp': chaves_comp,
                                                    'peso_nfe': peso_bungue,
                                                    'serie_comp': '0',
                                                    'peso_comp': '0',
                                                    'transportadora': 'BUNGE'}

                        elif 'NOTA' in df.columns and 'CHAVE' in df.columns:
                            cnpj_bunge = '84046101028101'
                            nf_ref = df['NOTA'].dropna().astype(str)
                            nf_ref = [nf.split('.')[0] for nf in nf_ref]
                            if 'CHAVE DE REFERÊNCIA' in df.columns:
                                chave_ref_clean = df['CHAVE DE REFERÊNCIA'].dropna().astype(
                                    str).str.extract(
                                    r'(\d{44})')
                                chave_ref_clean = chave_ref_clean.dropna().values.tolist()
                            elif 'CHAVE DE REFÊRENCIA' in df.columns:
                                chave_ref_clean = df['CHAVE DE REFÊRENCIA'].dropna().astype(
                                    str).str.extract(
                                    r'(\d{44})')
                                chave_ref_clean = chave_ref_clean.dropna().values.tolist()
                            nf_ref_cleaned = df['NOTA DE REFERÊNCIA'].dropna().astype(str)
                            nf_ref_cleaned = [nf.split('-')[0] for nf in nf_ref_cleaned]
                            if pdf_excel_chave:
                                print('teste')
                            replica_nfe = []
                            replica_chave_nfe_ref = []
                            replica_serie_nfe = []
                            replica_chave = []
                            for nf, chave_comp, nfe in zip(
                                    nf_ref,
                                    chave_ref_clean, nf_ref_cleaned):
                                chaves_comp = ''.join(chave_comp)

                                nf_excel_map[nf] = {'nota_fiscal': nf,
                                                    'data_email': message.received,
                                                    'chave_acesso': '0',
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': '0',
                                                    'data_emissao': '0',
                                                    'nfe': nfe.lstrip('0'),
                                                    'cnpj': cnpj_bunge,
                                                    'chave_comp': chaves_comp,
                                                    'peso_nfe': '0',
                                                    'serie_comp': '0',
                                                    'peso_comp': '0',
                                                    'transportadora': 'BUNGE'}  # peso_bungue})
                        else:
                            nf_excel_map[file_dir] = {'nota_fiscal': '0',
                                                      'data_email': message.received,
                                                      'chave_acesso': 'SEM LEITURA',
                                                      'email_vinculado': message.subject,
                                                      'serie_nf': 'SEM LEITURA',
                                                      'data_emissao': 'SEM LEITURA',
                                                      'nfe': '0',
                                                      'cnpj': 'SEM LEITURA',
                                                      'chave_comp': 'SEM LEITURA',
                                                      'peso_nfe': '0',
                                                      'serie_comp': 'SEM LEITURA',
                                                      'peso_comp': 'SEM LEITURA',
                                                      'transportadora': 'BUNGE'}

                    os.remove(temp_xlsx.name)
                if file_extension == '.xlsm':
                    decoded_content = base64.b64decode(attachment2.content)
                    print(f'{message.received} - {message.subject}')
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
                            dif_peso = df['Diferença de peso rel'].dropna().astype(str)
                            dif_peso = [dif.split('.')[0] for dif in dif_peso]
                            if not dif_peso:
                                dif_peso = [0]

                            dta_emissao = df['Data documento'].dropna().astype(str)

                            data_doc = dta_emissao.apply(
                                lambda x: datetime.strptime(x, '%Y-%m-%d').strftime('%d/%m/%Y')
                                if len(x) == 10 and '-' in x else x)
                            serie_nfe = df['Séries'].dropna().astype(str)
                            serie_nfe = [serie.split('.')[0] for serie in serie_nfe]
                            cnpj_nfe = df['CNPJ'].dropna().astype(str)
                            cnpj_nfe = [cn.split('.')[0] for cn in cnpj_nfe]
                            replica_nfe = []
                            replica_peso = []
                            replica_chave_nfe_ref = []
                            replica_serie_nfe = []
                            replica_data = []
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
                                        replica_data.append(data_doc[i % len(data_doc)])
                                        replica_cnpj.append(cnpj_nfe[i % len(cnpj_nfe)])
                                        replica_peso.append(dif_peso[i % len(dif_peso)])

                            for nf, chave, nfe, chave_comp, peso_bungue, cnpj, data_doc,serie_nfe in zip(
                                    nf_ref_cleaned,
                                    chave_ref_clean, replica_nfe, replica_chave_nfe_ref, replica_peso, replica_cnpj, replica_data,
                            replica_serie_nfe):

                                chaves = ''.join(chave)
                                chaves_comp = ''.join(chave_comp)
                                nf_excel_map[nf] = {'nota_fiscal': nf,
                                                    'data_email': message.received,
                                                    'chave_acesso': chaves,
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': serie_nfe,
                                                    'data_emissao': data_doc,
                                                    'nfe': nfe,
                                                    'cnpj': cnpj_bunge,
                                                    'chave_comp': chaves_comp,
                                                    'peso_nfe': peso_bungue,
                                                    'serie_comp': '0',
                                                    'peso_comp': '0',
                                                    'transportadora': 'BUNGE'}  # peso_bunge})

                    os.remove(temp_xlsx.name)