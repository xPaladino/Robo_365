import os, re, base64, tempfile, zipfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
import pandas as pd

def process_btg(message, save_folder, nf_zip_map, nf_pdf_map):
    """
    Essa função é responsável por ler e tratar os dados vindos do Cliente BTG, extraídos de um arquivo ZIP, podendo
    ser alterado os padrões de captura da Nota Fiscal para se adequar à esse projeto, os padrões atuais de captura são:

    Padrão 1.0:
    notas_fiscais.extend(re.finditer(r'NFs\s+de\s+(\d+/\d+/\d+)\s+\((.*?)\)', pdf_text, re.IGNORECASE))

    Caso seja realizado alguma alteração, favor documentar.

    :param message: Varíavel pertencente a lista Messages.
    :param save_folder: Local onde vai ser salvo o arquivo.
    :param nf_zip_map: Dicionário responsável por salvar os dados referente aos ZIP's.
    """
    attachment = message.attachments
    if re.search(r'@btgpactual\.com', message.body):
        print("TEM BTG")
        if message.attachments:
            xlsx_att = [attachment for attachment in attachment if os.path.splitext(attachment.name)[1].lower() == '.xlsx']
            if xlsx_att:
                for attachment in xlsx_att:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    print(file_extension)
                    print("achei um xlsx")
                    if file_extension == '.xlsx':
                        decoded_content = base64.b64decode(attachment.content)
                        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_xlsx:
                            temp_xlsx.write(decoded_content)
                        try:
                            with pd.ExcelFile(temp_xlsx.name, engine='openpyxl') as excel_file:
                                df = pd.read_excel(excel_file, engine='openpyxl',header=1)

                        except Exception as e:
                            print(f"erro as {e}")

            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                print(file_extension)
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
                                with open(file_path, 'rb') as pdf_file:
                                    pdf_reader = PdfReader(pdf_file)
                                    pdf_text = ''
                                    for page in pdf_reader.pages:
                                        pdf_text += page.extract_text()
                                        cnpj_btg = '14796754000961'
                                        notas_fiscais = []
                                        notas_fiscais.extend(
                                            re.finditer(r'NFs\s+de\s+(\d+/\d+/\d+)\s+\((.*?)\)|FISCO\s*\(\s*(\d{5}-\d{1,})|'
                                                        r'NFE\s*COMPLEMENTAR\s*a\s*NFE\s*(\d+)'
                                                        , pdf_text,
                                                        re.IGNORECASE | re.DOTALL))
                                        #INFORMAÇÕES COMPLEMENTARES RESERVADO AO FISCO(27618-111)
                                        nfe_comp = []
                                        nfe_comp.extend(
                                            re.finditer(r'SÉRIE(\d+)', pdf_text, re.IGNORECASE)
                                        )

                                        if not nfe_comp:
                                            nfe_comp.append(0)

                                        serie_matches = []
                                        serie_matches.extend(
                                            re.finditer(
                                                r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|(?:NF-e\s+\d{3}\.\d{3}\.\d{3}\s+)(\d{1,3})|'
                                                r'(?:Nº\s*SÉRIE:\s*\d+)(\s+\d{1,3})|(?:SÉRIE\s*\d+)(\s+\d{1,3})',
                                                pdf_text, re.IGNORECASE))
                                        if not serie_matches:
                                            serie_matches.append(0)

                                        for serie_match in serie_matches:
                                            try:
                                                for x in range(1, 7):
                                                    serie = serie_match.group(x)
                                                    if serie is not None:
                                                        break
                                            except IndexError:
                                                serie = 0

                                        if not notas_fiscais:
                                            notas_fiscais.append(0)

                                        data_matches = []
                                        data_matches.extend(re.finditer(
                                            r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                            r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                            r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})|\d{2}.\d{3}.\d{3}/\d{4}-\d{2}\s+(\d{2}/\d{2}/\d{2,})'
                                            , pdf_text, re.IGNORECASE))
                                        if not data_matches:
                                            data_matches.append(0)

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
                                                    'nfe': '0',
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
                                        replica_serie = []

                                        for i, nf_ref in enumerate(notas_fiscais):
                                            if i != 0:
                                                try:
                                                    for x in range(1, 15):
                                                        nf = nf_ref.group(x)
                                                        if nf is not None:
                                                            break

                                                except IndexError:
                                                    nf = 0
                                                if '\n' in nf: #TRATAMENTO PARA MAIS DE UMA NOTA
                                                    nf_ajust = nf.replace('\n','').replace(' ','')

                                                    nf_split = nf_ajust.split(',')

                                                    for valor in nf_split:
                                                        for match in nfe_comp:
                                                            nota = match.group(1)

                                                            if nota != "0":
                                                                replica_nfe.append(nota)
                                                                replica_chave.append(chave_acesso[0])
                                                                for serie_match in serie_matches:
                                                                    try:
                                                                        for x in range(1, 7):
                                                                            serie = serie_match.group(x)
                                                                            if serie is not None:
                                                                                break
                                                                    except IndexError:
                                                                        serie = 0
                                                                replica_serie.append(serie)
                                                else:
                                                    if len(nfe_comp) >= 2:
                                                        match = (nfe_comp[0])
                                                        if match != "0":
                                                            nfe_split = re.sub(r'[\D]', '', match[0])
                                                            replica_nfe.append(nfe_split)
                                                            replica_chave.append(chave_acesso[0])
                                                            for serie_match in serie_matches:
                                                                try:
                                                                    for i in range(1, 8):
                                                                        series = serie_match.group(i)
                                                                        if series is not None:
                                                                            break
                                                                    serie_split = re.sub(r'[\D]', '', series)
                                                                except IndexError:

                                                                    serie_split = 0
                                                            replica_serie.append(serie_split)



                                            replica_chave = [''.join(chave.split()) for chave in replica_chave]

                                        vazio = [notas_fiscais, replica_nfe, replica_chave, replica_serie,
                                                 replica_nfe]

                                        if None in vazio:
                                            print(f'Encontrado vazio em um dos valores\n'
                                                  f'Nota: {notas_fiscais}\n'
                                                  f'NFE: {replica_nfe}\n'
                                                  f'Data: {notas_fiscais}\n'
                                                  f'Chave: {replica_chave}\n'
                                                  f'Serie: {replica_serie}')
                                            break
                                        else:
                                            for match_nf, match_nf2, nfe_comp, chave, serie_comp in zip(notas_fiscais, notas_fiscais,
                                                                                            replica_nfe, replica_chave,replica_serie):

                                                data = match_nf2.group(1)
                                                notas = match_nf.group(2)
                                                if notas is None:
                                                    notas = match_nf.group(3)
                                                if notas is not None:
                                                    if '\n' in notas:
                                                        nf_ajust = notas.replace('\n', '').replace(' ', '')

                                                        nf_split = nf_ajust.split(',')
                                                        #nota_split = [nota.strip() for nota in notas.split(',')]
                                                        for valor in nf_split:

                                                            nota_sem_hifen = re.findall(r'\b(\d+)-\d+\b', valor)
                                                            nota_sem_hifen_str = ', '.join(nota_sem_hifen)
                                                            serie = re.findall(r'-(\d+)\b', valor)
                                                            serie_str = ', '.join(serie)
                                                            #print(f'{valor}\n{nota_sem_hifen_str}')
                                                            nf_zip_map[nota_sem_hifen_str] = {
                                                                'nota_fiscal': nota_sem_hifen_str,
                                                                'data_email': message.received,
                                                                'email_vinculado': message.subject,
                                                                'data_emissao': data,
                                                                'cnpj': cnpj_btg,
                                                                'chave_comp': chave,
                                                                'chave_acesso': '',
                                                                'nfe': nfe_comp,
                                                                'serie_nf': serie_comp,
                                                                'transportadora': 'BTG',
                                                                'serie_comp': '0',
                                                                'peso_comp': '0'
                                                            }
                                                else:
                                                    try:
                                                        for x in range(1, 10):
                                                            nf = match_nf.group(x)
                                                            if nf is not None:
                                                                break

                                                    except IndexError:
                                                        nf = 0
                                                    for data in data_matches:
                                                        if data != 0:
                                                            try:
                                                                for x in range(1, 6):
                                                                    data_emissao = data.group(x)
                                                                    if data_emissao is not None:
                                                                        break
                                                            except IndexError:

                                                                data_emissao = 0

                                                    nf_zip_map[nf] = {
                                                        'nota_fiscal': nf,
                                                        'data_email': message.received,
                                                        'email_vinculado': message.subject,
                                                        'data_emissao': data_emissao,
                                                        'cnpj': cnpj_btg,
                                                        'chave_comp': chave,
                                                        'chave_acesso': '',
                                                        'nfe': nfe_comp,
                                                        'serie_nf': serie_comp,
                                                        'transportadora': 'BTG',
                                                        'serie_comp': '0',
                                                        'peso_comp': '0'
                                                    }



                if file_extension == '.pdf':
                    print(file_extension)
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name
                        print('PDF')
                        pdf_reader = extract_text(temp_pdf_path)
                        notas_fiscais = []

                        notas_fiscais.extend(re.finditer
                                             (r'REF\.\s+NF\s+n\s+(\d+(?:\s*(?:,\s*|\s+e\s+)?\d+)*)'
                                              r'|NF\s+REFERENCIADA:\s*(\d+)|'
                                              r'REF.\s+NFe:\s*(\d+)|NOTA\s+FISCAL\s+Nº\s*(\d+)|'
                                              r'Complementar\s+à\s+NF\s+Núm\s+(\d+)|,\s*(\d{4,})\s+de|'
                                              r'REFERENTE\s+A\s+NF\s*(\d+)|'
                                              r'NFE\s*COMPLEMENTAR\s*a\s*NFE\s*(\d+)|'
                                              #            r'REFERENTE\s+NF\s*((?:\d+\s*/\s*)*\d+)'
                                              r'REFERENTE\s+NF\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)'
                                              r'|NFe\s+de\s+N\s+:\s+((?:\d+\s*/\s*)*\d+)|'
                                              r'notas\s+fiscais:\s*(([\d\s,]+)|(\d+(?:\s+\d+)*))|'
                                              r'NOTA\s+FISCAL\s+N\s+:\s*([\d\s/]+)|'
                                              r'NOTAS\s+FISCAIS\s+N\s+:\s*([\d\s/]+)|'
                                              r'Notas\s+de\s+referencia\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                              r'REFERENTE\s+AS\s+NF-e\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                              r'EXPORTACAO\s+NFE:(\d+)|'
                                              r'REF\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|-|\s*/\s*)?\d+)*)|'
                                              r'REF\s*A\s*NF\s*(\d+)|'
                                              # r'notas\s+fiscais:\s*([\d\s,]+)|'
                                              r'REFERENTE\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                              r'NOTA\s+FISCAL\s+NR.\s*(\d+)|'
                                              r'REF.\s+NFe\s+de\s+N\s+:\s*((?:\d+\s*/\s*)*\d+)|'
                                              r'REF.\s+Nfe:\s*(\d+)|(?:COMPLEMENTO\s+REF.\s+NOTA\s+FISCAL)\s*(?:\d+\s*,\s*)?(\d{3,8})|'
                                              r'NFe\s+Ref\.\:\s+Número\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)',
                                              pdf_reader, re.IGNORECASE))
                        if not notas_fiscais:
                            notas_fiscais.append(0)
                        print(notas_fiscais)
                        chave_comp = []
                        chave_comp.extend(
                            re.finditer(
                                r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))|'
                                r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))|'
                                r'(?:CHAVE\s+DE\s+ACESSO\s+da\s+NF-e\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))'
                                ,
                                pdf_reader, re.IGNORECASE))
                        if not chave_comp:
                            chave_comp.append(0)

                        nfe_comp = []
                        nfe_comp.extend(re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|NF-e\s*Nº.\s*SÉRIE|'
                                                    r'NF-e\s+N.º|NF-e\s*N.)\s+(\d+(?:\.\d+)*)'
                                                    , pdf_reader, re.IGNORECASE
                                                    ))
                        # print(pdf_reader)
                        # NF-eNF-e N.º 000.022.392
                        if not nfe_comp:
                            nfe_comp.append(0)

                        # print(f'{nfe_comp}-{attachment}')
                        data_matches, serie_matches = [], []
                        data_matches.extend(re.finditer(
                            r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                            r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                            r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                            , pdf_reader, re.IGNORECASE))
                        if not data_matches:
                            data_matches.append(0)

                        serie_matches.extend(
                            re.finditer(
                                r'(?:VALOR:\s+)(\d{9}\s+(\d+))|(?:Nº.\s*Série\s*\d+\s*)(\d+)|(?:NF-e\s*N.\s*\d+\s*SÉRIE\s*)(\d+)|'
                                r'(?:NF-e\s*N.º\s*\d{3}.\d{3}.\d{3}\s*SÉRIE\s*)(\d+)|(?:NF-e\s*\s*Nº\s*\d{3}.\d{3}.\d{3}\s*SÉRIE\s*)(\d+)',
                                pdf_reader, re.IGNORECASE))
                        if not serie_matches:
                            serie_matches.append(0)

                        vazio = [notas_fiscais, chave_comp, nfe_comp, data_matches, serie_matches]

                        if None in vazio:
                            print(f'Encontrado vazio em um dos valores\n'
                                  f'Nota: {notas_fiscais}\n'
                                  f'NFE: {nfe_comp}\n'
                                  f'Data: {data_matches}\n'
                                  f'Chave: {replica_chave}\n'
                                  f'Serie: {serie_matches}')
                            break
                        else:
                            cnpj_btg = '14796754000961'
                            for i, nota in enumerate(notas_fiscais):

                                if nota != 0:
                                    notas = None

                                    for i in range(1, 40):
                                        notas = nota.group(i)

                                        if notas is not None:
                                            break


                                    for chave, nfe, data, serie in zip(chave_comp,nfe_comp,data_matches,serie_matches):

                                        if chave != 0 and isinstance(chave,
                                                                     re.Match):
                                            chaves = None
                                            for i in range(1, 4):
                                                chaves = chave.group(i)
                                                if chaves is not None:
                                                    break
                                            chave_ajustada = ''.join(chaves.split()) if chaves else '0'
                                        else:
                                            chave_ajustada = '0'

                                        if nfe != 0:
                                            try:
                                                for i in range(1, 5):
                                                    nfe_ajust = nfe.group(i)
                                                    if nfe_ajust is not None:
                                                        break

                                                nfe_split = re.sub(r'[\D]', '', nfe_ajust)
                                            except IndexError:
                                                nfe_split = 0

                                        if data != 0:
                                            try:
                                                for x in range(1, 4):
                                                    data_emissao = data.group(x)
                                                    if data_emissao is not None:
                                                        break
                                            except IndexError:

                                                data_emissao = 0
                                        if serie != 0:
                                            try:
                                                for i in range(1, 8):
                                                    series = serie.group(i)
                                                    if series is not None:
                                                        break
                                                serie_split = re.sub(r'[\D]', '', series)
                                            except IndexError:

                                                serie_split = 0

                                        nf_pdf_map[notas] = {'nota_fiscal': notas,
                                                            'data_email': message.received,
                                                            'data_emissao': data_emissao,
                                                            'chave_comp': chave_ajustada,
                                                            'serie_nf': serie_split,
                                                            'cnpj': cnpj_btg,
                                                            'nfe': nfe_split,
                                                            'email_vinculado': message.subject,
                                                            'transportadora': 'BTG',
                                                            'peso_comp': '0',
                                                            'serie_comp': '0'
                                                            }
                    os.remove(temp_pdf_path)