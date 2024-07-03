import os
import re
import base64
import tempfile, zipfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_viterra(message, salve_folder, nf_pdf_map, nf_zip_map):
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
                                                    'email_vinculado': message.subject,
                                                    'transportadora': 'VITERRA'
                                                    }
                    if file_extension == '.zip':
                        print('ZIP')
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
                                    with open(file_path, 'rb') as pdf_file:
                                        pdf_reader = PdfReader(pdf_file)
                                        num_page = len(pdf_reader.pages)
                                        # print(f' numero pagina{num_pages}')
                                        for page_numb in range(num_page):
                                            if page_numb == 0:
                                                # print(f'pagina {page_numb + 1}')
                                                page = pdf_reader.pages[page_numb]
                                                # print(f' teste pagina {page}')

                                                pdf_text = page.extract_text()
                                                nf = []
                                                serie_matches = []
                                                data_matches = []
                                                nfe_match = []
                                                cnpj_viterra = '04485210000330'
                                                segundo_cnpj = []
                                                # print(pdf_text)
                                                # print(pdf_text)
                                                chave_acesso_match = (re.findall(
                                                    r'(\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4})',
                                                    pdf_text, re.IGNORECASE))
                                                if not chave_acesso_match:
                                                    chave_acesso_match.append(0)

                                                data_matches.extend(re.finditer(
                                                    r'(?:EMISSÃO\s+)(\d{2}/\d{2}/\d{2,})|'
                                                    r'(?:\d{2}.\d{3}.\d{3}/\d{4}-\d{2}\s+)(\d{2}/\d{2}/\d{2,})',
                                                    pdf_text,
                                                    re.IGNORECASE))
                                                if not data_matches:
                                                    data_matches.append(0)
                                                # for i in chave_acesso_match:
                                                #   print(i)
                                                # print(data_matches)
                                                """
                                                for i in data_matches:
                                                    if i != 0:
                                                        data = None
                                                        for x in range(1, 3):
                                                            data = i.group(x)
                                                            if data is not None:
                                                                break
                                                        print(data)
                                                """
                                                nfe_match.extend(
                                                    re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº)\s+(\d+(?:\.d+)*)'
                                                                r'|(?:DE\s+ENTREGA)\s+(\d+)', pdf_text, re.IGNORECASE))
                                                if not nfe_match:
                                                    nfe_match.append(0)
                                                # print(nfe_match)
                                                """
                                                for i in nfe_match:
                                                    if i != 0:
                                                        nfe = None
                                                        for x in range(1, 3):
                                                            nfe = i.group(x)
                                                            if nfe is not None:
                                                                break
                                                        print(nfe)
                                                """
                                                serie_matches.extend(
                                                    re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|'
                                                                r'(?:DE\s+ENTREGA\s+\d+)\s+(\d+)',
                                                                pdf_text, re.IGNORECASE))
                                                # print(serie_matches)

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
                                                                r'REF.\s+Nfe:\s*(\d+)|'
                                                                r'Relacionadas\s*-\s*((?:\d+/\d+\s*;?\s*)+)', pdf_text,
                                                                re.IGNORECASE))

                                                if not nf:
                                                    nf.append(0)
                                                notau = set()
                                                print()
                                                # print(f'nota {nf} + data match {data_matches}')
                                                for i in nf:
                                                    if i != 0:
                                                        notas = None
                                                        for x in range(1, 14):
                                                            notas = i.group(x)
                                                            if notas is not None:
                                                                break
                                                        # print(f'teste nota {notas}')
                                                        if notas:
                                                            nota_encontrada = re.split(r'/\d{1};\s+', notas)
                                                            # notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|\s+', notas)
                                                            for nota in nota_encontrada:
                                                                if nota.strip():
                                                                    notau.add(nota.strip())

                                                replica_serie = []
                                                replica_chave = []
                                                replica_nfe = []
                                                replica_data = []

                                                for i, nfe, dta, chave in zip(serie_matches, nfe_match, data_matches,
                                                                              chave_acesso_match):
                                                    if i != 0:
                                                        serie = None
                                                        for x in range(1, 5):
                                                            serie = i.group(x)
                                                            if serie is not None:
                                                                break
                                                    if nfe != 0:
                                                        notas = None
                                                        for x in range(1, 14):
                                                            notas = nfe.group(x)
                                                            if notas is not None:
                                                                break
                                                    if dta != 0:
                                                        data = None
                                                        for x in range(1, 3):
                                                            data = dta.group(x)
                                                            if data is not None:
                                                                break

                                                        for nota in sorted(notau):
                                                            replica_data.append(data)
                                                            replica_serie.append(serie)
                                                            replica_nfe.append(notas)
                                                            replica_chave.append(chave)
                                                            #replica_subject.append(subject_tratado_zip)
                                                for nota, chave, nota_nf, serie, dta in zip(nf,
                                                                                                       chave_acesso_match,
                                                                                                       replica_nfe,
                                                                                                       replica_serie,
                                                                                                       replica_data):
                                                    chaves = ''.join(chave.split())
                                                    notas = None
                                                    for i in range(1, 13):
                                                        notas = nota.group(i)
                                                        if notas is not None:
                                                            break
                                                    for match in sorted(notau):
                                                        nf_zip_map[match] = {
                                                            'nota_fiscal': match,
                                                            'data_email': message.received,
                                                            'chave_comp': chaves,
                                                            'chave_acesso': '',
                                                            'email_vinculado': message.subject,
                                                            'serie_nf': serie,
                                                            'data_emissao': dta,
                                                            'cnpj': cnpj_viterra,
                                                            'nfe': nota_nf,
                                                            'transportadora': 'VITERRA'
                                                        }



    else:
        pass
