import os
import re
import base64
import tempfile
import zipfile
from pdfminer.high_level import extract_text
from pdfminer.pdfparser import PDFSyntaxError
from PyPDF2 import PdfReader

def valida_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            header = f.read(4)
            return header == b'%PDF'
    except Exception as e:
        print(f"Erro ao verificar o PDF: {e}")
        return False

def process_amaggi(message, save_folder, nf_pdf_map,nf_zip_map):
    corpo_email = message.body
    if re.search(r'@amaggi\.com', corpo_email):
        print('Tem AMAGGI')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":

                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:

                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name

                        if valida_pdf(temp_pdf_path):
                            try:
                                pdf_reader = extract_text(temp_pdf_path)

                                notas_fiscais = []

                                notas_fiscais.extend(re.finditer
                                            (r'REF\.\s+NF\s+n\s+(\d+(?:\s*(?:,\s*|\s+e\s+)?\d+)*)'
                                            r'|NF\s+REFERENCIADA:\s*(\d+)|'
                                            r'REF.\s+NFe:\s*(\d+)|NOTA\s+FISCAL\s+Nº\s*(\d+)|'
                                            r'Complementar\s+à\s+NF\s+Núm\s+(\d+)|,\s*(\d{4,})\s+de|'
                                            r'REFERENTE\s+A\s+NF\s*(\d+)|'
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
                                            #r'notas\s+fiscais:\s*([\d\s,]+)|'
                                            r'REFERENTE\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                            r'NOTA\s+FISCAL\s+NR.\s*(\d+)|'
                                            r'REF.\s+NFe\s+de\s+N\s+:\s*((?:\d+\s*/\s*)*\d+)|'
                                            r'REF.\s+Nfe:\s*(\d+)|(?:COMPLEMENTO\s+REF.\s+NOTA\s+FISCAL)\s*(?:\d+\s*,\s*)?(\d{3,8})',
                                         pdf_reader, re.IGNORECASE))
                                if not notas_fiscais:
                                    notas_fiscais.append(0)
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
                                    ,pdf_reader, re.IGNORECASE
                                ))
                                peso_liq = []
                                peso_liq.extend(re.finditer(
                                    r'(?:Peso\s+Liquido\s*.*?|Peso\s+Líquido\s*.*?)(\b\d{1,4}\.\d{3},\d{3}\b|'
                                    r'\b\d{3},\d{3}\b)|'
                                    r'(?:TON\s*.*?|KG\s*.*?)(\b\d{1,3},\d{1,5}\b|'
                                    r'\b\d{1,4}\.\d{1,5},\d{1,5}\b)',
                                    pdf_reader, re.IGNORECASE))

                                # print(content)
                                final_value = None
                                final_group = None
                                if peso_liq:

                                    #for i, peso in enumerate(peso_liq):
                                    for peso in peso_liq:
                                        matched_value = None
                                        matched_group = None

                                        # Verifica os grupos em ordem de prioridade (1 a 4)
                                        for group_index in range(1, 5):
                                            matched_value = peso.group(group_index)
                                            if matched_value is not None:
                                                matched_group = group_index
                                                break

                                        # Atualiza o valor final apenas se ainda não tiver sido definido
                                        if final_value is None or matched_group == 1:
                                            final_value = matched_value
                                            final_group = matched_group

                                    # Exibe o valor final encontrado e seu grupo
                                    print(f"{attachment} - Valor encontrado: {final_value} - Grupo: {final_group}")
                                #NF-eNF-e N.º 000.022.392
                                if not nfe_comp:
                                    nfe_comp.append(0)

                                #print(f'{nfe_comp}-{attachment}')
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

                                #for i in serie_matches:
                                #    print(f'{i}- {attachment}')

                                replica_nota, replica_chave, replica_nfe = [], [], []
                                replica_serie, replica_data = [], []

                                for i, nota in enumerate(notas_fiscais):
                                    if nota != 0:
                                        notas = None
                                        for i in range(1, 20):
                                            notas = nota.group(i)
                                            if notas is not None:
                                                break

                                        notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|-|\s+', notas)
                                        for valor in notas_split:

                                            if valor != 'E':
                                                if valor != '':
                                                    replica_nota.append(valor)
                                                    for chave,nfe in zip(chave_comp,nfe_comp):
                                                        replica_chave.append(chave)
                                                        nfe_split = 0
                                                        if nfe != 0:
                                                            try:
                                                                for i in range(1, 5):
                                                                    nfe_ajust = nfe.group(i)
                                                                    if nfe_ajust is not None:
                                                                        break

                                                                nfe_split = re.sub(r'[\D]', '', nfe_ajust)
                                                            except IndexError:
                                                                nfe_split = 0
                                                        replica_nfe.append(nfe_split)
                                                    for serie, data in zip(serie_matches, data_matches):
                                                        if data != 0:
                                                            try:
                                                                for x in range(1, 4):
                                                                    data_emissao = data.group(x)
                                                                    if data_emissao is not None:
                                                                        break
                                                            except IndexError:

                                                                data_emissao = 0
                                                            replica_data.append(data_emissao)
                                                        if serie != 0:
                                                            try:
                                                                for i in range(1, 8):
                                                                    series = serie.group(i)
                                                                    if series is not None:
                                                                        break
                                                                serie_split = re.sub(r'[\D]', '', series)
                                                            except IndexError:

                                                                serie_split = 0
                                                            replica_serie.append(serie_split)


                                for nota, chave,nfe, data,serie in zip(replica_nota,replica_chave,replica_nfe,replica_data,replica_serie):
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

                                    cnpj_amaggi = '77294254000356'

                                    nf_pdf_map[nota] = {'nota_fiscal': nota,
                                                'data_email': message.received,
                                                'data_emissao': data,
                                                'chave_comp': chave_ajustada,
                                                'serie_nf': serie,
                                                'cnpj': cnpj_amaggi,
                                                'nfe': nfe.lstrip('0'),
                                                'email_vinculado': message.subject,
                                                'transportadora': 'AMAGGI',
                                                'peso_comp': '0',
                                                'serie_comp': '0'
                                                 }


                            except PDFSyntaxError:
                                print("Esse arquivo não é um PDF válido")
                                nf_pdf_map[attachment.name] = {
                                    'nota_fiscal': '0',
                                    'data_email': message.received,
                                    'chave_acesso': 'PDF INVÁLIDO',
                                    'email_vinculado': message.subject,
                                    'serie_nf': 'PDF INVÁLIDO',
                                    'data_emissao': 'PDF INVÁLIDO',
                                    'cnpj': 'PDF INVÁLIDO',
                                    'nfe': '0',
                                    'chave_comp': 'PDF INVÁLIDO',
                                    'transportadora': 'AMAGGI',
                                    'peso_comp': 'PDF INVÁLIDO',
                                    'serie_comp': 'PDF INVÁLIDO'
                                }

                            except Exception as e:
                                print(f"Erro durante o processamento {e}")
                        else:
                            print(f"{attachment} não é um arquivo PDF")
                            nf_pdf_map[attachment.name] = {
                                'nota_fiscal': '0',
                                'data_email': message.received,
                                'chave_acesso': 'PDF INVÁLIDO',
                                'email_vinculado': message.subject,
                                'serie_nf': 'PDF INVÁLIDO',
                                'data_emissao': 'PDF INVÁLIDO',
                                'cnpj': 'PDF INVÁLIDO',
                                'nfe': '0',
                                'chave_comp': 'PDF INVÁLIDO',
                                'transportadora': 'AMAGGI',
                                'peso_comp': 'PDF INVÁLIDO',
                                'serie_comp': 'PDF INVÁLIDO'
                            }
                    os.remove(temp_pdf_path)
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
                                                    notas_fiscais = []

                                                    notas_fiscais.extend(re.finditer
                                                                         (r'REF\.\s+NF\s+n\s+(\d+(?:\s*(?:,\s*|\s+e\s+)?\d+)*)'
                                                                          r'|NF\s+REFERENCIADA:\s*(\d+)|'
                                                                          r'REF.\s+NFe:\s*(\d+)|NOTA\s+FISCAL\s+Nº\s*(\d+)|'
                                                                          r'Complementar\s+à\s+NF\s+Núm\s+(\d+)|,\s*(\d{4,})\s+de|'
                                                                          r'REFERENTE\s+A\s+NF\s*(\d+)|'
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
                                                                          r'COMPLEMENTO\s+REFERENTE\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|-|\s*/\s*)?\d+)*)|'
                                                                          r'NOTA\s+COMPLEMENTAR\s+REF.\s+NFs\s+n\s+((?:\d+\s*)+)|'
                                                                          # r'notas\s+fiscais:\s*([\d\s,]+)|'
                                                                          r'REFERENTE\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                                                          r'NOTA\s+FISCAL\s+NR.\s*(\d+)|'
                                                                          r'REF.\s+NFe\s+de\s+N\s+:\s*((?:\d+\s*/\s*)*\d+)|'
                                                                          r'REF.\s+Nfe:\s*(\d+)|(?:COMPLEMENTO\s+REF.\s+NOTA\s+FISCAL)\s*(?:\d+\s*,\s*)?(\d{3,8})|'
                                                                          r'COMPLEMENTO\s+REFERENTE\s+A\s+NFe\s+N\s+(\d{2,3}.\d+)|PARANAGUA\s+NF:\s+(\d+)|'
                                                                          r'REF\s+as\s+NFs\s+(\d{1,3}.\d{3})|NOTA\s+DE\s+COMPLEMENTO\s+DE\s+VALOR\s+(\d+)|'
                                                                          r'COMPLEMENTO\s+a\s+Nota\s+Fiscal\s+((?:\d+\s*)+)|'
                                                                          r'NOTA\s+complementar\s+da\(s\)\s+nota\(s\):\s+((?:\d+\s*)+)|'
                                                                          r'NF\s+REFERENTE\s+AS\s+NFS:\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                                                          r'COMPLEMENTO\s+REF\s+SOBRA\s+na\s+DESCARGA\s+NFS\s*(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*|\s*/#\s*)?\d+)*)|'
                                                                          r'(?:ref\s+a\s+complemento\s+nf\s*|ref\s+a\s+complemento\s+a\s+nf\s*)((?:\d+\s*;\s*)*\d+)|'
                                                                          r'referente\s+a\(s\)\s+nota\(s\)\s+fiscal\(is\):\s+((?:\d+/\d{1}\s*)+|\d+(?:\s+\d+)*)|'
                                                                          r'VALOR\s+REFERNETE\s+AS\s+NOTAS\s*(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                                                          r'(?:REFERENTE\s+as\s+NOTAS\s+FISCAIS\s+NR.\s*|LOR\s+REFERENTE\s*as\s+NOTAS\s*|LOR\s+REF\s+as\s+NF\s*)(\d+(?:\s*(?:,\s*|\s+E\s+|-|\s*e\s*|\s*/\s*)?\d+)*)|'
                                                                          r'REF.\s+Compl.\s+NF:\s*(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)',
                                                                          #r'|(?:EMISSAO\s+ORIGINAL\s+NF-e:\s+\d+\s*)(\d+)',
                                                                          pdf_text, re.IGNORECASE))
                                                    if not notas_fiscais:
                                                        notas_fiscais.append(0)
                                                    for match in notas_fiscais:
                                                        if match == 0:
                                                            cnpj_amaggi = '77294254000356'
                                                            for chave, nfe in zip(chave_comp, nfe_comp):
                                                                if chave != 0 and isinstance(chave,
                                                                                             re.Match):
                                                                    chaves = None
                                                                    for i in range(1, 4):
                                                                        chaves = chave.group(i)
                                                                        if chaves is not None:
                                                                            break
                                                                    chave_ajustada = ''.join(
                                                                        chaves.split()) if chaves else '0'
                                                                else:
                                                                    chave_ajustada = '0'

                                                                try:
                                                                    for i in range(1, 5):
                                                                        nfe_ajust = nfe.group(i)
                                                                        if nfe_ajust is not None:
                                                                            break

                                                                    nfe_split = re.sub(r'[\D]', '',
                                                                                       nfe_ajust)
                                                                except IndexError:
                                                                    nfe_split = 0
                                                            for serie, data in zip(serie_matches, data_matches):
                                                                if data != 0:
                                                                    try:
                                                                        for x in range(1, 4):
                                                                            data_emissao = data.group(x)
                                                                            if data_emissao is not None:
                                                                                break
                                                                    except IndexError:

                                                                        data_emissao = 0
                                                                if '.' in data_emissao and '/' not in data_emissao:
                                                                    dataajust = data_emissao.replace('.', '/')
                                                                if serie != 0:
                                                                    try:
                                                                        for i in range(1, 8):
                                                                            series = serie.group(i)
                                                                            if series is not None:
                                                                                break
                                                                        serie_split = re.sub(r'[\D]', '',
                                                                                             series)
                                                                    except IndexError:

                                                                        serie_split = 0

                                                            nf_zip_map[file_dir] = {
                                                                'nota_fiscal': '0',
                                                                'data_email': message.received,
                                                                'chave_acesso': 'SEM LEITURA',
                                                                'email_vinculado': message.subject,
                                                                'serie_nf': serie_split,
                                                                'data_emissao': dataajust,
                                                                'cnpj': cnpj_amaggi,
                                                                'nfe': nfe_split.lstrip('0'),
                                                                'chave_comp': chave_ajustada,
                                                                'transportadora': 'AMAGGI',
                                                                'peso_comp': 'SEM LEITURA',
                                                                'serie_comp': 'SEM LEITURA'
                                                            }
                                                            print('zerado')
                                                    chave_comp = []
                                                    chave_comp.extend(
                                                        re.finditer(
                                                            r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))|'
                                                            r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))|'
                                                            r'(?:CHAVE\s+DE\s+ACESSO\s+da\s+NF-e\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))'
                                                            ,
                                                            pdf_text, re.IGNORECASE))
                                                    if not chave_comp:
                                                        chave_comp.append(0)

                                                    nfe_comp = []
                                                    nfe_comp.extend(
                                                        re.finditer(
                                                                    #r'(?:NF-e\s*No.|NF-e\s+Nº|NF-e\s*Nº.\s*SÉRIE|'
                                                                    #r'NF-e\s+N.º|NF-e\s*N.)\s+(\d+(?:\.\d+)*)'
                                                                    r'(?:FISCAL\s*ELETRÔNICA\s+)(\d+)'
                                                                    , pdf_text, re.IGNORECASE
                                                                    ))

                                                    if not nfe_comp:
                                                        nfe_comp.append(0)

                                                    data_matches, serie_matches = [], []
                                                    data_matches.extend(re.finditer(
                                                        r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                                        r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                                        r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                                        , pdf_text, re.IGNORECASE))
                                                    if not data_matches:
                                                        data_matches.append(0)

                                                    serie_matches.extend(
                                                        re.finditer(
                                                            #r'(?:VALOR:\s+)(\d{9}\s+(\d+))|(?:Nº.\s*Série\s*\d+\s*)(\d+)|(?:NF-e\s*N.\s*\d+\s*SÉRIE\s*)(\d+)|'
                                                            #r'(?:NF-e\s*N.º\s*\d{3}.\d{3}.\d{3}\s*SÉRIE\s*)(\d+)|(?:NF-e\s*\s*Nº\s*\d{3}.\d{3}.\d{3}\s*SÉRIE\s*)(\d+)',
                                                            r'SÉRIE:\s+(\d+)',
                                                            pdf_text, re.IGNORECASE))
                                                    if not serie_matches:
                                                        serie_matches.append(0)
                                                    replica_nota, replica_chave, replica_nfe = [], [], []
                                                    replica_serie, replica_data = [], []

                                                    for nota in notas_fiscais:
                                                        if nota != 0:
                                                            notas = None
                                                            grupo = None
                                                            for i in range(1, 40):
                                                                notas = nota.group(i)
                                                                if notas is not None:
                                                                    grupo = i
                                                                    break
                                                            if grupo:
                                                                pass #print(f'nota {notas} do grupo {grupo}')
                                                            if nota.group(18) is not None:
                                                                notajunt = ''.join(notas.split())
                                                                nf_formatado = notajunt.replace('\n', '')
                                                            elif nota.group(35) is not None:
                                                                notabarra = re.sub(r'/\d+','',notas)
                                                                nf_formatado = notabarra
                                                            elif nota.group (28) is not None or nota.group (26) is not None:
                                                                nf_formatado = notas.replace('.','')
                                                            elif nota.group(20) is not None or nota.group(30) is not None:
                                                                nf_formatado = notas
                                                            else:
                                                                nf_formatado = notas.replace(' ','')
                                                            notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|E|\s*;\s*|#|-|\s+',
                                                                                   nf_formatado)

                                                            for valor in notas_split:
                                                                if valor != '':
                                                                    replica_nota.append(valor)
                                                                    for chave, nfe in zip(chave_comp, nfe_comp):
                                                                        replica_chave.append(chave)
                                                                        nfe_split = 0
                                                                        if nfe != 0:
                                                                            try:
                                                                                for i in range(1, 5):
                                                                                    nfe_ajust = nfe.group(i)
                                                                                    if nfe_ajust is not None:
                                                                                        break

                                                                                nfe_split = re.sub(r'[\D]', '',
                                                                                                   nfe_ajust)
                                                                            except IndexError:
                                                                                nfe_split = 0
                                                                        replica_nfe.append(nfe_split)
                                                                    for serie, data in zip(serie_matches, data_matches):
                                                                        if data != 0:
                                                                            try:
                                                                                for x in range(1, 4):
                                                                                    data_emissao = data.group(x)
                                                                                    if data_emissao is not None:
                                                                                        break
                                                                            except IndexError:

                                                                                data_emissao = 0
                                                                            replica_data.append(data_emissao)
                                                                        if serie != 0:
                                                                            try:
                                                                                for i in range(1, 8):
                                                                                    series = serie.group(i)
                                                                                    if series is not None:
                                                                                        break
                                                                                serie_split = re.sub(r'[\D]', '',
                                                                                                     series)
                                                                            except IndexError:

                                                                                serie_split = 0
                                                                            replica_serie.append(serie_split)

                                                    for nota, chave, nfe, data, serie in zip(
                                                            replica_nota, replica_chave, replica_nfe,
                                                            replica_data, replica_serie):
                                                        if chave != 0 and isinstance(chave,
                                                                                     re.Match):
                                                            chaves = None
                                                            for i in range(1, 4):
                                                                chaves = chave.group(i)
                                                                if chaves is not None:
                                                                    break
                                                            chave_ajustada = ''.join(
                                                                chaves.split()) if chaves else '0'
                                                        else:
                                                            chave_ajustada = '0'
                                                        if '.' in data and '/' not in data:
                                                            data = data.replace('.','/')
                                                        cnpj_amaggi = '77294254000356'
                                                        nf_zip_map[nota] = {'nota_fiscal': nota,
                                                                            'data_email': message.received,
                                                                            'data_emissao': data,
                                                                            'chave_comp': chave_ajustada,
                                                                            'serie_nf': serie,
                                                                            'cnpj': cnpj_amaggi,
                                                                            'nfe': nfe.lstrip('0'),
                                                                            'email_vinculado': message.subject,
                                                                            'transportadora': 'AMAGGI',
                                                                            'peso_comp': '0',
                                                                            'serie_comp': '0'
                                                                            }

                                                    #print(f'{file_dir}-{notas_fiscais}')