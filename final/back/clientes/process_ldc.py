import os, re, base64, tempfile, zipfile
from PyPDF2 import PdfReader
def process_ldc(message, save_folder, nf_zip_map):
    print(f"{message.received} - {message.subject}")
    if re.search(r'@ldc\.com', message.body):
        print("TEM LDC")
        if message.attachments:
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".zip":
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        zip_file_path = os.path.join(tmp_dir, 'attachment.zip')
                        with open(zip_file_path,'wb') as f:
                            f.write(decoded_content)
                        with zipfile.ZipFile(zip_file_path,'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)
                        for file_name in os.listdir(tmp_dir):
                            file_path = os.path.join(tmp_dir, file_name)
                            if file_path.endswith('.pdf'):
                                with open(file_path,'rb') as pdf_file:
                                    pdf_reader = PdfReader(pdf_file)
                                    pdf_text = ''
                                    for page in pdf_reader.pages:
                                        pdf_text += page.extract_text()

                                        data_emissao = []
                                        notas_fiscais = []
                                        serie_matches = []
                                        segundo_cnpj = []
                                        cnpj_match = []
                                        nfe_match = []

                                        data_emissao.extend(re.finditer(
                                            r'(?:EMISSÃO\s+)(\d{2}/\d{2}/\d{2,})', pdf_text, re.IGNORECASE))
                                        if not data_emissao:
                                            data_emissao.append(0)

                                        notas_fiscais.extend(
                                            re.finditer(r'(?:número:|REF\s+A\s+NOTA'
                                                        r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
                                                        r'|referente\s+a\s+NF|'
                                                        r'NF:)'
                                                        r'\s*(?:\d+\s*,\s*)?(\d{4,})'
                                                        r'|Nota\s+Fiscal:\s*(?:\d+\s*,\s*)?(\d{4,})', pdf_text,
                                                        re.IGNORECASE))
                                        if not notas_fiscais:
                                            notas_fiscais.append(0)

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
                                        for nota, data_match, series, chaves, cnpj, nfe in zip(
                                                notas_fiscais, data_emissao, serie_matches, chave_acesso_match,
                                                segundo_cnpj,
                                                nfe_match):
                                            if nota == 0:
                                                """print('NÃO ENCONTREI NOTA FISCAL')
                                                nf_zip_map[message.subject,] = {
                                                    'nota_fiscal': '0',
                                                    'data_email': message.received,
                                                    'chave_comp': '0',
                                                    'chave_acesso': '0',
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': '0',
                                                    'data_emissao': '0',
                                                    'cnpj': '0',
                                                    'nfe': '0',
                                                    'transportadora': 'LDC'
                                                }"""
                                                pass
                                            else:
                                                nota_match = nota.group(1)
                                                pass
                                                if nota_match is None:

                                                    notas = nota.group(2)
                                                    emissao_match = data_match.group(1)
                                                    chave = ''.join(chaves.split())
                                                    series_match = series.group(3)
                                                    cnpj_segundo = cnpj.group(0)
                                                    nfe_formatado = nfe.group(0)
                                                    nfe_ponto = re.sub(r'[.]', '', nfe_formatado)  # remoção do ponto
                                                    nfe_tratada = nfe_ponto if nfe_ponto[0:3] != '000' else nfe_ponto[
                                                                                                            3:]
                                                    cn_tratada = re.sub(r'[./-]', '', cnpj_segundo)
                                                    nf_zip_map[notas] = {
                                                        'nota_fiscal': notas,
                                                        'data_email': message.received,
                                                        'chave_comp': chave,
                                                        # 'chave_acesso': chave,
                                                        'email_vinculado':  message.subject,
                                                        'serie_nf': series_match,
                                                        'data_emissao': emissao_match,
                                                        'cnpj': cn_tratada,
                                                        'nfe': nfe_tratada
                                                    }

                                                elif len(notas_fiscais) <= 2:
                                                    nota_match = nota.group(1)
                                                    emissao_match = data_match.group(1)
                                                    chave = ''.join(chaves.split())
                                                    series_match = series.group(3)
                                                    cnpj_segundo = cnpj.group(0)
                                                    nfe_formatado = nfe.group(1)
                                                    nfe_ponto = re.sub(r'[.]', '', nfe_formatado)  # remoção do ponto
                                                    nfe_tratada = nfe_ponto if nfe_ponto[0:3] != '000' \
                                                        else nfe_ponto[3:]
                                                    cn_tratada = re.sub(r'[./-]', '', cnpj_segundo)
                                                    nf_zip_map[nota_match] = {
                                                        'nota_fiscal': nota_match,
                                                        'data_email': message.received,
                                                        'chave_comp': chave,
                                                        'chave_acesso': '',
                                                        'email_vinculado':  message.subject,
                                                        'serie_nf': series_match,
                                                        'data_emissao': emissao_match,
                                                        'cnpj': cn_tratada,
                                                        'nfe': nfe_tratada,
                                                        'transportadora': 'LDC'
                                                    }
