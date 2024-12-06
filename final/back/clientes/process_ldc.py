import os, re, base64, tempfile, zipfile
from PyPDF2 import PdfReader


def process_ldc(message, save_folder, nf_zip_map):
    """
    Essa função é responsável por ler e tratar os dados vindos do Cliente LDC (Dreyfus), extraído de um arquivo ZIP,
    podendo ser alterado os padrões de captura da Nota Fiscal para se adequar à esse projeto, os padrões atuais de
    captura são:

    Padrão 1.0:

    notas_fiscais.extend(
                                                re.finditer(r'(?:número:|REF\s+A\s+NOTA'
                                                            r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
                                                            r'|referente\s+a\s+NF|'
                                                            r'NF:)'
                                                            r'\s*(?:\d+\s*,\s*)?(\d{4,})'
                                                            r'|Nota\s+Fiscal:\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                                            r'(?:NOTA\s+FISCAL\s+DE\s+ORIGEM\s+NR\.:\s+)(\d{2,4}.\d{3,6})',
                                                            pdf_text,
                                                            re.IGNORECASE))

    Caso seja realizado alguma alteração, favor documentar.
    :param message: Varíavel pertencente a lista Messages.
    :param save_folder: Local onde vai ser salvo o arquivo.
    :param nf_zip_map: Dicionário responsável por salvar os dados referente aos ZIP's.
    :return:
    """
    if re.search(r'@ldc\.com', message.body):
        print("TEM LDC")
        if message.attachments:
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".zip":
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        zip_file_path = os.path.join(tmp_dir, 'attachment.zip')
                        with open(zip_file_path, 'wb') as f:
                            f.write(decoded_content)
                        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)
                        for file_name in os.listdir(tmp_dir):
                            file_path = os.path.join(tmp_dir, file_name)
                            if file_path.endswith('.pdf'):
                                if file_name == '_JUNTO.pdf':
                                    pass
                                else:
                                    with open(file_path, 'rb') as pdf_file:
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
                                                            r'|Nota\s+Fiscal:\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                                            r'(?:NOTA\s+FISCAL\s+DE\s+ORIGEM\s+NR\.:\s+)(\d{2,4}.\d{3,6})',
                                                            pdf_text,
                                                            re.IGNORECASE))
                                            if not notas_fiscais:
                                                notas_fiscais.append(0)

                                            serie_matches.extend(
                                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|(?:Série\s+)(\d{1,})'
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
                                                re.finditer(r'(?:NF-e\s+Nº\s+|NF-e\s+Nº.\s+)((\d{3}.\d{3}.\d{3})|'
                                                            r'\d{3}\.\d{3}\.\d{3})', pdf_text,
                                                            re.IGNORECASE))
                                            if not nfe_match:
                                                nfe_match.append(0)
                                            cnpj_ldc = '47067525009911'

                                            vazio = [notas_fiscais, data_emissao, serie_matches, chave_acesso_match, nfe_match]
                                            if None in vazio:
                                                print(f'Encontrado vazio em um dos valores da LDC\n'
                                                      f'Nota: {notas_fiscais}\n'
                                                      f'NFE: {nfe_match}\n'
                                                      f'Data: {data_emissao}\n'
                                                      f'Chave: {chave_acesso_match}\n'
                                                      f'Serie: {serie_matches}\n')
                                                break
                                            else:
                                                for nota, data_match, series, chaves, nfe in zip(
                                                        notas_fiscais, data_emissao, serie_matches, chave_acesso_match,
                                                        nfe_match):
                                                    if nota != 0:
                                                        try:
                                                            for x in range(1, 5):
                                                                nf_ajust = nota.group(x)
                                                                if nf_ajust is not None:
                                                                    break
                                                            nf_formatado = nf_ajust.replace('.', '')
                                                        except IndexError:
                                                            nf_formatado = 0

                                                        try:
                                                            for x in range(1, 2):
                                                                data_ajust = data_match.group(x)
                                                                if data_ajust is not None:
                                                                    break
                                                        except IndexError:
                                                            data_ajust = 0

                                                        try:
                                                            for x in range(1, 4):
                                                                serie = series.group(x)
                                                                if serie is not None:
                                                                    break

                                                        except IndexError:
                                                            serie = 0

                                                        try:
                                                            chave = ''.join(chaves.split())
                                                        except IndexError:
                                                            chave = 0

                                                        try:
                                                            for x in range(1, 4):
                                                                nfe_comp = nfe.group(x)
                                                                if nfe_comp is not None:
                                                                    break
                                                            nfe_ponto = re.sub(r'[.]', '', nfe_comp)
                                                            nfe_ajust = nfe_ponto.lstrip('0')
                                                        except IndexError:
                                                            nfe_comp = 0

                                                        nf_zip_map[nf_formatado] = {
                                                            'nota_fiscal': nf_formatado,
                                                            'data_email': message.received,
                                                            'chave_acesso': '',
                                                            'email_vinculado': message.subject,
                                                            'serie_nf': serie,
                                                            'data_emissao': data_ajust,
                                                            'cnpj': cnpj_ldc,
                                                            'nfe': nfe_ajust,
                                                            'chave_comp': chave,
                                                            'transportadora': 'LDC',
                                                            'peso_comp': '0',
                                                            'serie_comp': ''
                                                        }
                                                    if nota == 0:
                                                        print('NÃO ENCONTREI NOTA FISCAL')
                                                        nf_zip_map[file_name] = {
                                                            'nota_fiscal': '0',
                                                            'data_email': message.received,
                                                            'chave_acesso': 'SEM LEITURA',
                                                            'email_vinculado': message.subject,
                                                            'serie_nf': 'SEM LEITURA',
                                                            'data_emissao': 'SEM LEITURA',
                                                            'cnpj': 'SEM LEITURA',
                                                            'nfe': '0',
                                                            'chave_comp': 'SEM LEITURA',
                                                            'transportadora': 'LDC',
                                                            'peso_comp': 'SEM LEITURA',
                                                            'serie_comp': 'SEM LEITURA'

                                                        }
