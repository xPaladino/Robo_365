import os, re, base64, tempfile, zipfile
from PyPDF2 import PdfReader

def process_btg(message, save_folder, nf_zip_map):

    if re.search(r'@btgpactual\.com', message.body):
        print("TEM BTG")
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

                                        cnpj_btg = '14796754000961'
                                        notas_fiscais = []
                                        notas_fiscais.extend(
                                            re.finditer(r'NFs\s+de\s+(\d+/\d+/\d+)\s+\((.*?)\)', pdf_text,
                                                        re.IGNORECASE))
                                        nfe_comp = []
                                        nfe_comp.extend(
                                            re.finditer(r'SÉRIE(\d+)', pdf_text, re.IGNORECASE)
                                        )
                                        if not notas_fiscais:
                                            notas_fiscais.append(0)
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
                                                    'transportadora': 'BTG',
                                                    'peso_comp': 'SEM LEITURA',
                                                    'serie_comp': 'SEM LEITURA'
                                                }

                                        chave_acesso = re.findall(
                                            r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b',
                                            pdf_text, re.IGNORECASE)

                                        if not chave_acesso:
                                            chave_acesso.append(0)

                                        replica_nfe = []
                                        replica_chave = []

                                        for i, nf_ref in enumerate(notas_fiscais):
                                            nf = nf_ref.group(2)
                                            nf_split = nf.split(', ')
                                            for valor in nf_split:
                                                for match in nfe_comp:
                                                    nota = match.group(1)
                                                    if nota != "0":
                                                        replica_nfe.append(nota)
                                                        replica_chave.append(chave_acesso[0])
                                        replica_chave = [''.join(chave.split()) for chave in replica_chave]
                                        for match_nf, match_nf2, nfe_comp, chave in zip(notas_fiscais, notas_fiscais,
                                                                                        replica_nfe, replica_chave):

                                            data = match_nf2.group(1)
                                            nota = match_nf.group(2)
                                            nota_split = nota.split(', ')
                                            for valor in nota_split:
                                                nota_sem_hifen = re.findall(r'\b(\d+)-\d+\b', valor)
                                                nota_sem_hifen_str = ', '.join(nota_sem_hifen)
                                                serie = re.findall(r'-(\d+)\b', valor)
                                                serie_str = ', '.join(serie)
                                                print(f'{valor}\n{nota_sem_hifen_str}')
                                                nf_zip_map[nota_sem_hifen_str] = {
                                                    'nota_fiscal': nota_sem_hifen_str,
                                                    'data_email': message.received,
                                                    'email_vinculado': message.subject,
                                                    'data_emissao': data,
                                                    'cnpj': cnpj_btg,
                                                    'chave_comp': chave,
                                                    'chave_acesso': '',
                                                    'nfe': nfe_comp,
                                                    'serie_nf': serie_str,
                                                    'transportadora': 'BTG',
                                                    'serie_comp': '0',
                                                    'peso_comp': '0'
                                                }
