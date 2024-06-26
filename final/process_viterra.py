import os
import re
import base64
import tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_viterra(message, salve_folder, nf_pdf_map):
    corpo_email = message.body
    print(f"{message.received} - {message.subject}")
    if re.search(r'@viterra\.com', corpo_email):
        print('tem viterra')
        if message.attachments:
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    if file_extension == ".pdf":
                        decoded_content = base64.b64decode(attachment.content)
                        print(f"Processando: {attachment.name}")
                        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                            temp_pdf.write(decoded_content)
                            temp_pdf_path = temp_pdf.name
                            pdf_reader = extract_text(temp_pdf.name)
                            nf = []
                            chave_comp = []
                            serie_matches = []
                            data_matches = []
                            nfe_match = []
                            segundo_cnpj = []
                            nf.extend(
                                re.finditer(r'REF\.\s+NF\s+n\s+(\d+(?:\s*(?:,\s*|\s+e\s+)?\d+)*)'
                                            r'|NF\s+REFERENCIADA:\s*(\d+)|'
                                            r'REF.\s+NFe:\s*(\d+)|'
                                            r'Complementar\s+à\s+NF\s+Núm\s+(\d+)|,\s*(\d{4,})\s+de|'
                                            #            r'REFERENTE\s+NF\s*((?:\d+\s*/\s*)*\d+)'
                                            r'REFERENTE\s+NF\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)'
                                            r'|NFe\s+de\s+N\s+:\s+((?:\d+\s*/\s*)*\d+)|'
                                            r'notas\s+fiscais:\s*(\d+(?:\s+\d+)*)|'
                                            r'REFERENTE\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                            r'NOTA\s+FISCAL\s+NR.\s*(\d+)|'
                                            r'REF.\s+NFe\s+de\s+N\s+:\s*((?:\d+\s*/\s*)*\d+)|'
                                            r'REF.\s+Nfe:\s*(\d+)', pdf_reader, re.IGNORECASE))
                            if not nf:
                                nf.append(0)


                            chave_acesso_match = (re.findall(
                                r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b|'
                                r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))',
                                pdf_reader, re.IGNORECASE))
                            if not chave_acesso_match:
                                chave_acesso_match.append(0)

                            chave_comp.extend(
                                re.finditer(
                                    r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))',
                                    pdf_reader, re.IGNORECASE))
                            if not chave_comp:
                                chave_comp.append(0)

                            serie_matches.extend(
                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})',
                                            pdf_reader, re.IGNORECASE))
                            if not serie_matches:
                                serie_matches.append(0)

                            data_matches.extend(re.finditer(
                                r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                r'EMISSAO\s+(\d{2}/\d{2}/\d{2,})|'
                                r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                , pdf_reader, re.IGNORECASE))
                            if not data_matches:
                                data_matches.append(0)

                            nfe_match.extend(
                                re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº)\s+(\d+(?:\.\d+)*)', pdf_reader,
                                            re.IGNORECASE))
                            if not nfe_match:
                                nfe_match.append(0)

                            pdf_reader2 = PdfReader(temp_pdf)
                            pdf_text = ''
                            for page in pdf_reader2.pages:
                                pdf_text += page.extract_text()
                                cnpj_match = []
                                cnpj_match.extend(
                                    re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_text,
                                                re.IGNORECASE))
                                if len(cnpj_match) >= 2:
                                    segundo_cnpj.append(cnpj_match[1])
                            if not segundo_cnpj:
                                segundo_cnpj.append(0)

                            replica_subject = []
                            replica_nfe = []
                            replica_chave = []
                            replica_serie = []
                            replica_data = []
                            replica_nota = []
                            replica_cnpj = []

                            for i, nota in enumerate(nf):
                                if nota != 0:
                                    notas = None
                                    for i in range(1, 13):
                                        notas = nota.group(i)
                                        if notas is not None:
                                            break

                                    notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|\s+', notas)
                                    for valor in notas_split:
                                        if valor != 'E':
                                            replica_nota.append(valor)
                                            for match, chave, data, serie, cnpj in zip(nfe_match, chave_comp, data_matches,
                                                                                 serie_matches,segundo_cnpj):
                                                datax = None
                                                if data != 0:
                                                    for i in range(1, 3):
                                                        datax = data.group(i)
                                                        if datax is not None:
                                                            break
                                                seriex = None
                                                for i in range(1, 4):
                                                    seriex = serie.group(i)
                                                    if seriex is not None:
                                                        break

                                                notax = match.group(1)

                                                chavex = chave.group(1)
                                                if notax != "0":

                                                    replica_nfe.append(notax)
                                                    replica_subject.append(message.subject)
                                                    replica_chave.append(chavex)
                                                    replica_data.append(datax)
                                                    replica_serie.append(seriex)
                                                    replica_cnpj.append(cnpj)

                            for nota,chave,nfe,serie, cnpj in zip(replica_nota,replica_chave,replica_nfe,replica_serie,replica_cnpj):
                                chaves = ''.join(chave.split())
                                cn = cnpj.group(1)
                                cn_tratada = re.sub(r'[./-]', '',
                                                    cn)
                                nf_pdf_map[nota] = {'nota_fiscal' : nota,
                                                    'data_email': message.received,
                                                    'data_emissao': '0',
                                                    'chave_comp': chaves,
                                                    'serie_nf': serie,
                                                    'cnpj': cn_tratada,
                                                    'nfe': nfe,
                                                    'email_vinculado': message.subject
                                                    }
                            """for nota, chaves, cnpj, nota_nf, chave, dta, serie in zip(replica_nota, chave_comp, segundo_cnpj,
                                                                                              replica_nfe, replica_subject,
                                                                                              replica_chave,
                                                                                              replica_serie,
                                                                                             ):
                                notas = None
                                for i in range(1, 13):
                                    notas = nota.group(i)
                                    if notas is not None:
                                        break
                                

                                notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|\s+', notas)
                                print(nota)
                                print(chaves)
                                print(cnpj)
                                print(nota_nf)
                                print(chave)
                                print(dta)
                                print(serie)

                                chaves = ''.join(chave.split())
                                if nota == 0:
                                    print('SEM NOTA')
                                    nf_pdf_map[message.subject] = {
                                        'nota_fiscal': 'NÃO ENCONTREI',
                                        'data_email': message.received,
                                        'data_emissao': '0',
                                        'chave_comp': '0',
                                        'serie_nf': '0',
                                        'cnpj': '0',
                                        'nfe': '0',
                                        'email_vinculado': message.subject
                                    }
                                for notax in notas_split:
                                    if notax != 'E':
                                nf_pdf_map[nota] = {'nota_fiscal': nota,
                                                             'data_email': message.received,
                                                             'data_emissao': dta,
                                                             'serie_nf': serie,
                                                             'chave_comp': chaves,
                                                             'cnpj': '0',
                                                             'email_vinculado': message.subject,
                                                             'nfe': nota_nf
                                                             }"""

    else:
        pass
