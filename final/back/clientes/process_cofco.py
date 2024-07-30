import os, re, base64, tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_cofco(message, save_folder, nf_pdf_map):
    if re.search(r'@cofcointernational\.com',message.body):
        print('Tem COFCO')
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
                        cnpj_cofco = '06315338000542'


                        #print(pdf_reader)
                        notas_fiscais = []
                        notas_fiscais.extend(re.finditer(
                            r'NF:\s*(\d{4,8})(?:\s*,\s*(\d{4,8}))*|'
                            #r'nota\(s\)\s*fiscais:\s*(?:\d+\s*,\s*)?(\d{4,9})|'
                            r'nota\(s\)\s*fiscais:\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'PESO\s+REF\.\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'PESO\s+REF\..A\s*NF\s*(\d{4,8})|'
                            r'REF.\s+NF\s*(\d{4,8})',
                            pdf_reader, re.IGNORECASE))
                        if not notas_fiscais:
                            notas_fiscais.append(0)
                        for match in notas_fiscais:
                            if match == 0:
                                nf_pdf_map[attachment.name] = {
                                    'nota_fiscal': '0',
                                    'data_email': message.received,
                                    'chave_acesso': '0',
                                    'email_vinculado': message.subject,
                                    'serie_nf': '0',
                                    'data_emissao': '0',
                                    'cnpj': '0',
                                    'nfe': attachment.name,
                                    'chave_comp': '0',
                                    'transportadora': 'COFCO',
                                    'peso_comp': '0',
                                    'serie_comp': '0'
                                }
                                print('zerado')

                        replica_nota = []
                        serie_matches = []

                        serie_matches.extend(
                            re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,3})',
                                        pdf_reader, re.IGNORECASE))

                        chaves_comp = (re.findall(r'\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}',
                                        pdf_reader,re.IGNORECASE))

                        if not chaves_comp:
                            chaves_comp.append(0)

                        nfe_match = []
                        nfe_match.extend(
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|VALOR:\s+)\s+(\d+(?:\.\d+)*)|'
                                        r'SÉRIE\s*(\d{4,8})', pdf_reader,
                                        re.IGNORECASE))

                        pdf_reader2 = PdfReader(temp_pdf)
                        pdf_text = ''
                        for page in pdf_reader2.pages:
                            pdf_text += page.extract_text()
                            pesocomp = []
                            pesocomp.extend(
                                re.finditer(r'PESO\s+LÍQUIDO\s*(\d+(?:\.\d+)*(?:,\d{1,3})?)',pdf_text,re.IGNORECASE)
                            )


                        if not serie_matches:
                            serie_matches.append(0)
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
                                r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))',
                                pdf_reader, re.IGNORECASE))

                        data_emissao = []
                        data_emissao.extend(
                            re.finditer(r'(?:DATA\s+DA\s+EMISSÃO\s+|DATA\s+DE\s+EMISSÃO\s+)(\d{2}/\d{2}/\d{4})|'
                                        r'(?:DATA\s+DA\s+EMISSÃO\s+)(\d{2}.\d{2}.\d{4})', pdf_reader, re.IGNORECASE))

                        for i in notas_fiscais:
                            if i != 0:
                                numeros_notas_fiscais = re.findall(r'\d{4,9}', i.group())
                                for i in numeros_notas_fiscais:
                                    replica_nota.append(i)

                            else:
                                pass
                        for nf, chave,comp,nfe,serie,dta_comp in zip(replica_nota,chave_acesso_match,chaves_comp,nfe_match,
                                                            serie_matches,data_emissao):
                            try:
                                chaves = ''.join(chave.split())
                                ch_comp = ''.join(comp.split())
                            except IndexError:
                                ch_comp = 0
                            try:
                                for x in range(1, 3):
                                    data = dta_comp.group(x)
                                    if data is not None:
                                        break
                                data_emissao = re.sub(r'[.]', '/', data)
                            except IndexError:
                                data_emissao = 0

                            try:
                                for i in range(0,4):
                                    notas = nfe.group(i)
                                    if notas is not None:
                                        break
                                nota_trata = re.sub(r'[\D]', '', notas)
                            except IndexError:
                                nota_trata = 0
                            try:
                                for i in range(0,2):
                                    series = serie.group(i)
                                    if series is not None:
                                        break
                                serie_trata = re.sub(r'[\D]', '', series)
                            except IndexError:
                                serie_trata = 0

                            for nota in replica_nota:
                                nf_pdf_map[nota] = {
                                    'nota_fiscal': nota,
                                    'chave_acesso': chaves,
                                    'data_email': message.received,
                                    'email_vinculado': message.subject,
                                    'serie_nf': serie_trata,
                                    'serie_comp': '0',
                                    'data_emissao': data_emissao,
                                    'cnpj': cnpj_cofco,
                                    'nfe': nota_trata,
                                    'chave_comp': ch_comp,
                                    'transportadora': 'COFCO',
                                    'peso_comp': '0'

                                }

                    os.remove(temp_pdf.name)