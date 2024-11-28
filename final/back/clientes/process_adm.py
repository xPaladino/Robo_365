import os, re, base64, tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def valida_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            header = f.read(4)
            return header == b'%PDF'
    except Exception as e:
        print(f"Erro ao verificar o PDF: {e}")
        return False

def process_adm(message, save_folder, nf_pdf_map):
    """
    Essa função é responsável por ler e tratar os dados vindos do Cliente ADM, extraído de um arquivo PDF, podendo ser
    alterado os padrões de captura da Nota Fiscal para se adequar à esse projeto, os padrões atuais de captura são:

    Padrão 1.0:

    notas_fiscais.extend(
    re.finditer( r'(?:#NF:|Nota\s+Fiscal:|fiscais:|NF:|' r'Ref\s+NF|' r'NF\s+n|' r'Nfe\s+de\s+n\s+:|'
    r'REF.\s+NOTA\s+FISCAL|' r'Referente\s+NF|' r'REF\s+A\s+NOTA|' r'REF\s+A\s+NOTA\s+N)' r'\s*(?:\d+\s*,\s*)?(\d{3,
    8})|' r'ORIGEM\s+NR\.: (\d+(\.\d+)?)|' r'(:?REF\s+NFS\s+)(\d+/\d+/\d+)\s+\((.*?)\)|' r'Referente\s+NF\s+(\d{
    2}\s+\d+)', pdf_reader, re.IGNORECASE))

    Caso seja realizado alguma alteração, favor documentar.
    :param message: Variável pertencente a lista Messages.
    :param save_folder: Local onde vai ser salvo o arquivo.
    :param nf_pdf_map: Dicionário responsável por salvar os dados referente aos PDF's.

    """
    if re.search(r'@adm\.com', message.body):
        print('tem ADM')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":
                    decoded_content = base64.b64decode(attachment.content)
                    with (tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf):
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name

                        if valida_pdf(temp_pdf_path):
                            pdf_reader = extract_text(temp_pdf_path)
                            notas_fiscais = []
                            notas_fiscais.extend(re.finditer(
                                r'(?:#NF:|Nota\s+Fiscal:|NF:|'
                                r'Ref\s+NF|'
                                r'NF\s+n|'
                                r'Nfe\s+de\s+n\s+:|'
                                r'REF.\s+NOTA\s+FISCAL|'
                                r'Referente\s+NF|'
                                r'REF\s+A\s+NOTA|'
                                r'REF\s+A\s+NOTA\s+N|'
                                r'REFERENTE\s+A\s+NOTA\s+FISCAL\s+N\s+:)'
                                r'\s*(?:\d+\s*,\s*)?(\d{3,8})|'
                                r'ORIGEM\s+NR\.: (\d+(\.\d+)?)|'
                                r'(:?REF\s+NFS\s+)(\d+/\d+/\d+)\s+\((.*?)\)|'
                                r'Referente\s+NF\s+(\d{2}\s+\d+)|'
                                r'NOTAS\s+FISCAIS\s+N\s+:(\s*([\d\s]+))|'
                                r'REF\.?\s+NF\s+n[ºo]?\s*([\d\s/]+)|'                            
                                #r'(?:Ref\.\s*NF\s+n[ºo]?\s*(\d{3,8}))|(?:NFe[:\s]*(\d{3,8}))|'
                                r'REFERENTE\s+NFES\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                                r'REFERENTE\s+AS\s+NOTAS\s+FISCAIS:\s*([\d\s\,]+)|'
                                r'(?:Ref\.\s*NF\s+n[ºo]?\s*|NFe[:\s]*)(\d{3,8})|'
                                r'Ref.\s*as\s*NFS\s*n\s*(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                                r'Nota\s*de\s*complemento\s*ref\s*as\s*nfs\s*(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)',
                                #r'(?:NFe:\s*(\d{3,8}))',
                                pdf_reader, re.IGNORECASE))
                            if not notas_fiscais:
                                notas_fiscais.append(0)
                            print(notas_fiscais)

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
                                                r'NF-e\s+Nº\s+(\d+\s+\d+\s+\d+)|(?:NF-e\s+Nº)\s+(\d+(?:\.\d+)*)|NF-e\s*N°\s*(\d+\.\d+\.\d+)',
                                                pdf_text, re.IGNORECASE))

                                #N° 000.543.973
                                #print(pdf_text)
                                if not nfe:
                                    nfe.append(0)

                            cnpj_adm = '02003402007269'
                            #notas_fiscais = sorted(list(set(notas_fiscais)))
                            replica_nota = []
                            for nf in notas_fiscais:
                                print(nf)
                                if nf == 0:
                                    nf_pdf_map[attachment.name] = {
                                        'nota_fiscal': '0',
                                        'data_email': message.received,
                                        'chave_acesso': 'SEM LEITURA',
                                        'email_vinculado': message.subject,
                                        'serie_nf': 'SEM LEITURA',
                                        'data_emissao': 'SEM LEITURA',
                                        'cnpj': 'SEM LEITURA',
                                        'nfe': '0',
                                        'chave_comp': 'SEM LEITURA',
                                        'transportadora': 'ADM',
                                        'peso_comp': 'SEM LEITURA',
                                        'serie_comp': 'SEM LEITURA'
                                    }
                                    print('zerado')
                                elif nf != 0:

                                    try:

                                        for x in range(1, 30):
                                            nf_ajust = nf.group(x)
                                            if nf_ajust is not None:
                                                break

                                        # tive que fixar os grupos 9 e 10 para funcionar
                                        if nf.group(9) is not None:
                                            nf_formatado = nf_ajust.replace('\n', '')
                                        elif nf.group(10) is not None:
                                            nf_formatado = nf_ajust.replace('\n', '')

                                        elif nf.group(12) is not None:
                                            nf_formatado = nf_ajust.replace('\n', '')
                                        else:
                                            nf_formatado = nf_ajust.strip()# = nf_ajust.replace(' ', '')

                                        nf_padrao = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|\s*E\s*|\s+', nf_formatado.strip())

                                    except IndexError:
                                        nf_padrao = 0
                                if nf !=0:
                                    replica_nota.extend(nf_padrao)

                            replica_serie = []
                            replica_data = []
                            replica_chave_comp = []
                            replica_nfe = []
                            for nf in replica_nota:
                                if nf != 0:
                                    for serie_match, data_match, chave, nfe_comp in zip(serie_matches,
                                                                                    data_matches, chave_comp, nfe):
                                        data_emissao, serie, chave_trat, nota_comp = None, None, None, None


                                        if data_match != 0 :
                                            try:
                                                for x in range(1, 5):
                                                    data_ajust = data_match.group(x)
                                                    if data_ajust is not None:
                                                        break
                                                data_emissao = data_ajust.replace('.', '/') if data_match else '0'
                                            except IndexError:
                                                data_emissao = 0
                                            print(data_emissao)
                                        replica_data.append(data_emissao)

                                        if serie_match != 0:
                                            try:
                                                for x in range(1, 6):
                                                    serie = serie_match.group(x)
                                                    if serie is not None:
                                                        break
                                            except IndexError:
                                                serie = 0
                                            print(serie)
                                        replica_serie.append(serie)

                                        if chave !=0:
                                            try:
                                                for x in range(0, 2):
                                                    ch = chave.group(x)
                                                    if ch is not None:
                                                        break
                                                ch_trat = ''.join(ch.split())
                                                chave_trat = re.sub(r'[\D]', '', ch_trat)
                                            except IndexError:
                                                chave_trat = 0
                                            print(f' chave - {chave_trat}')
                                        replica_chave_comp.append(chave_trat)

                                        if nfe_comp != 0:
                                            try:

                                                for x in range(1, 10):
                                                    if nfe_comp and isinstance(nfe_comp, re.Match):
                                                        comp_nota = nfe_comp.group(x)
                                                        if comp_nota is not None:
                                                            break
                                                    else:
                                                        comp_nota = None
                                                if comp_nota:
                                                    nota_comp = comp_nota.replace('.', '')
                                                else:
                                                    nota_comp = 0
                                            except IndexError:
                                                nota_comp = 0

                                        replica_nfe.append(nota_comp)



                            for nota, serie_match, data_match, chave, nfe_comp in zip(replica_nota, replica_serie, replica_data,
                                                                                    replica_chave_comp, replica_nfe):
                                print(f'teste {nfe_comp}')
                                nf_pdf_map[nota] = {
                                    'nota_fiscal': nota,
                                    'data_email': message.received,
                                    'chave_acesso': '0',
                                    'email_vinculado': message.subject,
                                    'serie_nf': serie_match,
                                    'data_emissao': data_match,
                                    'cnpj': cnpj_adm,
                                    'nfe': nfe_comp.lstrip('0'),
                                    'chave_comp': chave,
                                    'transportadora': 'ADM',
                                    'peso_comp': '0',
                                    'serie_comp': '0'
                                }
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
                    os.remove(temp_pdf.name)
