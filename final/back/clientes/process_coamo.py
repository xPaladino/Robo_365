import os
import re
import base64
import tempfile
import zipfile

from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
def process_coamo(message, salve_folder, nf_pdf_map, nf_zip_map):
    if re.search(r'@coamo\.com\.br', message.body):
        print('tem coamo')
        if message.attachments:
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    if file_extension == ".pdf":
                        decoded_content = base64.b64decode(attachment.content)
                        #print(f"Processando: {attachment.name}")
                        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                            temp_pdf.write(decoded_content)
                            temp_pdf_path = temp_pdf.name
                            pdf_reader = extract_text(temp_pdf_path)
                            # print(pdf_reader)
                            notas_fiscais = []
                            notas_fiscais.extend(re.finditer(
                                r'(?:#NF:|Nota\s+Fiscal:|fiscais:|NF:|'  # padrao
                                r'Ref\s+NF|'  # adicionado royal
                                r'NF\s+n|'  # adicionado royal
                                r'Nfe\s+de\s+n\s+:|'
                                r'Referente\s+NF|'
                                r'REF\s+A\s+NOTA|'
                                r'REF\s+A\s+NOTA\s+N)'  # adicionado para usimat destilaria
                                r'\s*(?:\d+\s*,\s*)?(\d{3,8})|'  # padrao 
                                r'ORIGEM\s+NR\.: (\d+(\.\d+)?)|'  # adicionado para agricola gemelli
                                r'(:?REF\s+NFS\s+)(\d+/\d+/\d+)\s+\((.*?)\)|'
                                r'NR\./SERIE/DATA:\s*\*\s*(\d{4,})\s*\*|'
                                r'(?:NOTA\s+FISCAL\s+)?NR\./SERIE/DATA:\s*(\d{1,3}(?:\.\d{3})*)',
                                pdf_reader, re.IGNORECASE))
                            if not notas_fiscais:
                                notas_fiscais.append(0)

                            chave_acesso = re.findall(
                                r'(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b)',
                                pdf_reader,re.IGNORECASE
                            )

                            remessa = re.findall(
                                r'REM.P/FORM.LOTE\s+P/POSTER.EXPORTACAO',pdf_reader,re.IGNORECASE
                            )
                            if not remessa:
                                remessa.append(0)
                            serie_match = []
                            serie_match.extend(
                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})'
                                            , pdf_reader, re.IGNORECASE))

                            if not serie_match:
                                serie_match.append(0)
                            #print(serie_match)
                            nfe = []
                            nfe.extend(
                                re.finditer(r'(?:NF-e\s+Nº\s+|NF-e\s+Nº.\s+)(\d{3}.\d{3}.\d{3})|'
                                            r'\d{3}\.\d{3}\.\d{3}', pdf_reader,
                                            re.IGNORECASE))
                            if not nfe:
                                nfe.append(0)
                            data_match = []
                            data_match.extend(re.finditer(
                                r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                , pdf_reader, re.IGNORECASE))
                            if not data_match:
                                data_match.append(0)

                            for match, chave, rem, nfe_comp,serie, dta in zip(notas_fiscais, chave_acesso,remessa,nfe, serie_match,data_match):
                                chaves = ''.join(chave.split())
                                if match == 0:
                                    nf_pdf_map[attachment.name] = {
                                        'nota_fiscal': '0',
                                        'data_email': message.received,
                                        'chave_acesso': '0',
                                        'email_vinculado': message.subject,
                                        'serie_nf': 'SEM LEITURA',
                                        'data_emissao': 'SEM LEITURA',
                                        'cnpj': rem,
                                        'nfe': attachment.name,
                                        'chave_comp': chaves,
                                        'transportadora': 'COAMO',
                                        'peso_comp': rem,
                                        'serie_comp': rem
                                    }

                                else:
                                    try:
                                        if isinstance(nfe_comp.group(0), str):
                                            padraonfe = nfe_comp.group(0)
                                            nfe_pont = re.sub(r'[.]', '', padraonfe)
                                            nfe_ajust = nfe_pont if nfe_pont[0:3] != '000' else nfe_pont[3:]
                                    except IndexError:
                                        nfe_ajust = 0

                                    try:
                                        if dta != 0:
                                            for i in range(1, 3):
                                                data = dta.group(i)
                                                if data is not None:
                                                    break
                                    except IndexError:
                                        data = 0

                                    try:
                                        for x in range(1, 10):
                                            nota = match.group(x)
                                            if nota is not None:
                                                break
                                        ajust = nota.replace('.', '')
                                    except IndexError:
                                        ajust = 0

                                    try:
                                        for i in range(1, 4):
                                            series = serie.group(i)
                                            if series is not None:
                                                break
                                    except IndexError:
                                        series = 0
                                    cpnj_coamo = '75904383024810'
                                    nf_pdf_map[ajust] = {
                                        'nota_fiscal': ajust,
                                        'data_email': message.received,
                                        'chave_acesso': '0',
                                        'email_vinculado': message.subject,
                                        'serie_nf': series,
                                        'data_emissao': data,
                                        'cnpj': cpnj_coamo,
                                        'nfe': nfe_ajust,
                                        'chave_comp': chaves,
                                        'transportadora': 'COAMO',
                                        'peso_comp': '0',
                                        'serie_comp': '0'
                                    }

                        os.remove(temp_pdf_path)
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
                                                            r'|Nota\s+Fiscal:\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                                            r'NR\./SERIE/DATA:\s*\*\s*(\d{4,})\s*\*|'
                                                            r'(?:NOTA\s+FISCAL\s+)?NR\./SERIE/DATA:\s*(\d{1,3}(?:\.\d{3})*)', pdf_text,
                                                            re.IGNORECASE))
                                            for match in notas_fiscais:
                                                if match == 0:
                                                    nf_zip_map[attachment.name] = {
                                                        'nota_fiscal': '0',
                                                        'data_email': message.received,
                                                        'chave_acesso': 'SEM LEITURA',
                                                        'email_vinculado': message.subject,
                                                        'serie_nf': 'SEM LEITURA',
                                                        'data_emissao': 'SEM LEITURA',
                                                        'cnpj': 'SEM LEITURA',
                                                        'nfe': attachment.name,
                                                        'chave_comp': 'SEM LEITURA',
                                                        'transportadora': 'COAMO',
                                                        'peso_comp': 'SEM LEITURA',
                                                        'serie_comp': 'SEM LEITURA'

                                                    }
                                            serie_matches.extend(
                                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})'
                                                            , pdf_text, re.IGNORECASE))

                                            if not serie_matches:
                                                serie_matches.append(0)

                                            chave_acesso_match = re.findall(
                                                r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b',
                                                pdf_text, re.IGNORECASE)

                                            if not chave_acesso_match:
                                                chave_acesso_match.append(0)

                                            cnpj_match.extend(
                                                re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_text,
                                                            re.IGNORECASE))

                                            if len(cnpj_match) >= 2:
                                                segundo_cnpj.append(cnpj_match[1])
                                            if not segundo_cnpj:
                                                segundo_cnpj.append(0)

                                            nfe_match.extend(
                                                re.finditer(r'(?:NF-e\s+Nº\s+|NF-e\s+Nº.\s+)(\d{3}.\d{3}.\d{3})|'
                                                            r'\d{3}\.\d{3}\.\d{3}', pdf_text,
                                                            re.IGNORECASE))
                                            if not nfe_match:
                                                nfe_match.append(0)
                                            data_matches = []
                                            data_matches.extend(re.finditer(
                                                r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                                r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                                r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                                , pdf_text, re.IGNORECASE))
                                            if not data_matches:
                                                data_matches.append(0)

                                            replica_nota = []
                                            # print(notas_fiscais)
                                            for i in notas_fiscais:
                                                if i != 0:
                                                    if i != 'e':
                                                        notas = None
                                                        for x in range(1, 5):
                                                            notas = i.group(x)
                                                            if notas is not None:
                                                                break
                                                        ajust = notas.replace('.','')
                                                        replica_nota.append(ajust)

                                            for nf, nfe, chave, cnpj, serie, datas in zip(replica_nota, nfe_match,chave_acesso_match,segundo_cnpj, serie_matches, data_matches):
                                                try:
                                                    if isinstance(nfe.group(0),str):
                                                        padraonfe = nfe.group(0)
                                                        nfe_pont = re.sub(r'[.]','',padraonfe)
                                                        nfe_ajust = nfe_pont if nfe_pont[0:3] != '000' else nfe_pont[3:]
                                                except IndexError:
                                                    nfe_ajust = 0
                                                try:
                                                    if datas != 0:
                                                        for i in range(1, 3):
                                                            data = datas.group(i)
                                                            if data is not None:
                                                                break
                                                except IndexError:
                                                    data = 0
                                                try:
                                                    chaves = ''.join(chave.split())
                                                    cnpj_tratado = re.sub(r'[./-]','',cnpj.group(1))
                                                except:
                                                    cnpj_tratado = '75904383024810'
                                                try:
                                                    for i in range(1,4):
                                                        series = serie.group(i)
                                                        if series is not None:
                                                            break
                                                except:
                                                    series = 0

                                                nf_zip_map[nf] = {
                                                    'nota_fiscal': nf,
                                                    'data_email': message.received,
                                                    'chave_comp': chaves,
                                                    'chave_acesso': '0',
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': series,
                                                    'data_emissao': data,
                                                    'cnpj': cnpj_tratado,
                                                    'nfe': nfe_ajust,
                                                    'transportadora': 'COAMO',
                                                    'serie_comp': '0',
                                                    'peso_comp': '0'
                                                }
