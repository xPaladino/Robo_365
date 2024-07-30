import os, re, base64, tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_adm(message, save_folder, nf_pdf_map):

    if re.search(r'@adm\.com', message.body):
        print('tem ADM')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name
                        pdf_reader = extract_text(temp_pdf_path)

                        notas_fiscais = []
                        notas_fiscais.extend(re.finditer(
                            r'(?:#NF:|Nota\s+Fiscal:|fiscais:|NF:|'  # padrao
                            r'Ref\s+NF|'  # adicionado royal
                            r'NF\s+n|'  # adicionado royal
                            r'Nfe\s+de\s+n\s+:|'
                            r'REF.\s+NOTA\s+FISCAL|'
                            r'Referente\s+NF|'
                            r'REF\s+A\s+NOTA|'
                            r'REF\s+A\s+NOTA\s+N)'  # adicionado para usimat destilaria
                            r'\s*(?:\d+\s*,\s*)?(\d{3,8})|'  # padrao 
                            r'ORIGEM\s+NR\.: (\d+(\.\d+)?)|'
                            r'(:?REF\s+NFS\s+)(\d+/\d+/\d+)\s+\((.*?)\)|'
                            r'Referente\s+NF\s+(\d{2}\s+\d+)',
                            pdf_reader, re.IGNORECASE))
                        if not notas_fiscais:
                            notas_fiscais.append(0)

                        chave_acesso_match = (re.findall(
                            r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b|'
                            r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))',
                            pdf_reader, re.IGNORECASE))
                        if not chave_acesso_match:
                            chave_acesso_match.append(0)

                        chave_comp = []
                        chave_comp.extend(
                            re.finditer(
                                r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))|'
                                r'\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b|'
                                r'NFe\d{44}',
                                pdf_reader, re.IGNORECASE))

                        if not chave_comp:
                            chave_comp.append(0)

                        serie_matches = []
                        serie_matches.extend(
                            re.finditer(
                                r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|(?:NF-e\s+\d{3}\.\d{3}\.\d{3}\s+)(\d{1,3})|'
                                r'(?:Nº\s*SÉRIE:\s*\d+)(\s+\d{1,3})',
                                pdf_reader, re.IGNORECASE))
                        if not serie_matches:
                            serie_matches.append(0)

                        data_matches = []
                        data_matches.extend(re.finditer(
                            r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                            r'EMISSAO\s+(\d{2}/\d{2}/\d{2,})|'
                            r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})|'
                            r'(?:DATA\s+DA\s+EMISSÃO\s*\d{2}.\d{3}.\d{3}/\d{4}-\d{2}\s*)(\d{2}/\d{2}/\d{2,})'

                            , pdf_reader, re.IGNORECASE))
                        if not data_matches:
                            data_matches.append(0)

                        nfe_match = []
                        nfe_match.extend(
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº)\s+(\d+(?:\.\d+)*)|NF-e\s+Nº\s+(\d+\s+\d+\s+\d+)',
                                        pdf_reader,
                                        re.IGNORECASE))

                        pdf_reader2 = PdfReader(temp_pdf)
                        pdf_text = ''
                        for page in pdf_reader2.pages:
                            pdf_text += page.extract_text()
                            nfe = []
                            nfe.extend(
                                re.finditer(r'\d+No\.\s*(\d{6,9})|SÉRIE:\s*(\d{4,})|lado\s*(\d{3}\.\d{3}\.\d{3})|'
                                            r'NF-e\s+Nº\s+(\d+\s+\d+\s+\d+)',
                                            pdf_text, re.IGNORECASE))

                            if not nfe:
                                nfe.append(0)

                        cnpj_adm = '02003402007269'
                        notas_fiscais = sorted(list(set(notas_fiscais)))

                        for nf, serie_match, data_match, chave, nfe_comp in zip(notas_fiscais, serie_matches,
                                                                                data_matches, chave_comp, nfe):
                            if nf != 0:
                                try:
                                    for x in range(1, 10):
                                        nf_ajust = nf.group(x)
                                        if nf_ajust is not None:
                                            break
                                    nf_formatado = nf_ajust.replace(' ', '')
                                except IndexError:
                                    nf_formatado = 0

                                try:
                                    for x in range(1, 5):
                                        data_ajust = data_match.group(x)
                                        if data_ajust is not None:
                                            break
                                    data_emissao = data_ajust.replace('.', '/') if data_match else '0'
                                except IndexError:
                                    data_emissao = 0

                                try:
                                    for x in range(1, 6):
                                        serie = serie_match.group(x)
                                        if serie is not None:
                                            break
                                except IndexError:
                                    serie = 0

                                try:
                                    for x in range(0, 2):
                                        ch = chave.group(x)
                                        if ch is not None:
                                            break
                                    ch_trat = ''.join(ch.split())
                                    chave_trat = re.sub(r'[\D]', '', ch_trat)
                                except IndexError:
                                    chave_trat = 0

                                try:
                                    for x in range(1, 5):
                                        comp_nota = nfe_comp.group(x)
                                        if comp_nota is not None:
                                            break
                                    nota_comp = comp_nota.replace('.', '')
                                except IndexError:
                                    nota_comp = 0

                                nf_pdf_map[nf_formatado] = {
                                    'nota_fiscal': nf_formatado,
                                    'data_email': message.received,
                                    'chave_acesso': '0',
                                    'email_vinculado': message.subject,
                                    'serie_nf': serie,
                                    'data_emissao': data_emissao,
                                    'cnpj': cnpj_adm,
                                    'nfe': nota_comp,
                                    'chave_comp': chave_trat,
                                    'transportadora': 'ADM',
                                    'peso_comp': '0',
                                    'serie_comp': '0'
                                }
                            if nf == 0:
                                nf_pdf_map[attachment.name] = {
                                    'nota_fiscal': '0',
                                    'data_email': message.received,
                                    'chave_acesso': 'SEM LEITURA',
                                    'email_vinculado': message.subject,
                                    'serie_nf': 'SEM LEITURA',
                                    'data_emissao': 'SEM LEITURA',
                                    'cnpj': 'SEM LEITURA',
                                    'nfe': attachment.name,
                                    'chave_comp': 'SEM LEITURA',
                                    'transportadora': 'ADM',
                                    'peso_comp': 'SEM LEITURA',
                                    'serie_comp': 'SEM LEITURA'
                                }
                        os.remove(temp_pdf.name)
