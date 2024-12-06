import os, re, base64, tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text


def process_cofco(message, save_folder, nf_pdf_map):
    """
    Essa função é responsavel por ler e tratar os dados vindos do Cliente COFCO, extraído de um arquivo PDF, podendo ser
    alterado os padrões de captura da Nota Fiscal para se adequar à esse projeto, os padrões atuais de captura são:

    Padrão 1.0:

    notas_fiscais.extend(re.finditer(
    r'NF:\s*(\d{4,8})(?:\s*,\s*(\d{4,8}))*|'
    #r'nota\(s\)\s*fiscais:\s*(?:\d+\s*,\s*)?(\d{4,9})|'
    r'nota\(s\)\s*fiscais:\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
    r'PESO\s+REF\.\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
    r'PESO\s+REF\..A\s*NF\s*(\d{4,8})|'
    r'REF.\s+NF\s*(\d{4,8})',
    pdf_reader, re.IGNORECASE))

    Caso seja realizado alguma alteração, favor documentar.
    :param message: Variável pertecente a lista Messages.
    :param save_folder: Local onde vai ser salvo o arquivo.
    :param nf_pdf_map: Dicionário responsável por salvar os dados referente aos PDF's.
    """
    if re.search(r'@cofcointernational\.com', message.body):
        print('Tem COFCO')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":

                    decoded_content = base64.b64decode(attachment.content)
                    #print(decoded_content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name

                        pdf_reader = extract_text(temp_pdf_path)

                        cnpj_cofco = '06315338000542'

                        notas_fiscais = []
                        notas_fiscais.extend(re.finditer(
                            r'NOTAS\s+FISCAIS:\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'(?:#NF:|Nota\s+Fiscal:|fiscais:|NF:|'
                            r'Ref\s+NF|'
                            r'REF.\s+NF|'
                            r'NF\s+n|,\s*NF|'
                            r'Nfe\s+de\s+n\s+:|'
                            r'Referente\s+NF|'
                            r'Referente\s+a\s+NF|'
                            r'nota\(s\)\s+fiscal\(is\):|'
                            r'nota\(s\)\s+fiscais|'
                            r'Nota\(s\)\s+de\s+Origem:\s*|'
                            r'REF\s+A\s+NOTA|'
                            r'Referente\s+a\s+NFE|REF.\s+NFE:|Emissao\s+Original\s+NF-e:\s+\d+|'
                            r'REF.\s+NOTA\s+FISCAL|A\s+NOTA\s+FISCAL\s+N\s+:|'
                            r'REF\s+NFE|REFE\s+NFE|ICMS\s+NF\(s\):)'
                            r'\s*(?:\d+\s*,\s*)?(\d{4,9})|'
                            r'NF:\s*(\d{4,8})(?:\s*,\s*(\d{4,8}))*|'
                            r'REF.\s+NOTAS\s+FISCAIS\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            # r'nota\(s\)\s*fiscais:\s*(?:\d+\s*,\s*)?(\d{4,9})|'
                            r'nota\(s\)\s*fiscais:\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'PESO\s+REF\.\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'PESO\s+REF\..A\s*NF\s*(\d{4,8})|'
                            r'REF.\s+NF\s*(\d{4,8})|'
                            r'REF.\s+NF\s+nº\s*([\d\s\-]+)|'
                            r'(\b(\d{9}-\d{1,3})\b|\b(\d{5}\n\d{4}-\d{3}))',
                            #r'COMPLEMENTO\s*REFERENTE\s*NF(\d+)',
                            pdf_reader, re.IGNORECASE | re.DOTALL))

                        if not notas_fiscais:
                            notas_fiscais.append(0)
                        replica_nota = []

                        for match in notas_fiscais:

                            if match == 0:
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
                                    'transportadora': 'COFCO',
                                    'peso_comp': 'SEM LEITURA',
                                    'serie_comp': 'SEM LEITURA',
                                }
                                print('zerado')
                            else:
                                try:
                                    for x in range(1, 20):
                                        notas = match.group(x)
                                        if notas is not None:
                                            break

                                    print(notas)

                                    nota_trata = notas.replace('\n', '')

                                except AttributeError as e:
                                    print(f"Erro de atributo: {e}")
                                    nota_trata = 0
                                except IndexError as e:
                                    print(f"Erro de índice: {e}")
                                    nota_trata = 0
                                except Exception as e:
                                    print(f"Erro inesperado: {e}")
                                    nota_trata = 0

                                if nota_trata != 0:
                                    numeros_notas_fiscais = re.findall(r'\d{4,9}', nota_trata)

                                    for i in numeros_notas_fiscais:

                                        replica_nota.append(i)


                                else:
                                    pass

                        serie_matches = []
                        serie_matches.extend(
                            re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,3})|'
                                        r'(?:FOLHA:\s*\d{1,3}.\d{1,3}.\d{1,3}\s*|FOLHA\s*\d{1,3}.\d{1,3}.\d{1,3}\s*)(\d{1,3})|'
                                        r'(?:SÉRIE:\s*\d{2,6}\s*)(\d+)',
                                        pdf_reader, re.IGNORECASE))

                        chaves_comp = (re.findall(
                            r'\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}|\d{44}',
                            pdf_reader, re.IGNORECASE))

                        if not chaves_comp:
                            chaves_comp.append(0)

                        nfe_match = []
                        nfe_match.extend(
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s*N.|NF-e\s+Nº|VALOR:\s+)\s+(\d+(?:\.\d+)*)|'
                                        r'SÉRIE\s*(\d{4,8})|(?:FOLHA:|FOLHA)\s*(\d{1,3}.\d{1,3}.\d{1,3})|'
                                        r'Nº\s*(\d{1,3}.\d{1,3}.\d{1,3})', pdf_reader,

                                        re.IGNORECASE))

                        pdf_reader2 = PdfReader(temp_pdf)
                        pdf_text = ''
                        for page in pdf_reader2.pages:
                            pdf_text += page.extract_text()
                            pesocomp = []
                            pesocomp.extend(
                                re.finditer(r'PESO\s+LÍQUIDO\s*(\d+(?:\.\d+)*(?:,\d{1,3})?)', pdf_text, re.IGNORECASE)
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
                                r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))|'
                                r'(?:NFe)(\d{44})',
                                pdf_reader, re.IGNORECASE))

                        data_emissao = []
                        data_emissao.extend(
                            re.finditer(r'(?:DATA\s+DA\s+EMISSÃO\s+|DATA\s+DE\s+EMISSÃO\s+)(\d{2}/\d{2}/\d{4})|'
                                        r'(?:DATA\s+DA\s+EMISSÃO\s+)(\d{2}.\d{2}.\d{4})|'
                                        r'(?:\d{2}.\d{3}.\d{3}/\d{4}-\d{2}\s+)(\d{2}/\d{2}/\d{2,4})|'
                                        r'(?:\s+EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,4})', pdf_reader, re.IGNORECASE))

                        vazio = [replica_nota, chaves_comp, nfe_match, serie_matches, data_emissao]

                        if None in vazio:
                            print(f'Encontrado vazio em um dos valores da COFCO\n'
                                  f'Nota: {replica_nota}\n'
                                  f'NFE: {nfe_match}\n'
                                  f'Data: {data_emissao}\n'
                                  f'Chave: {chaves_comp}\n'
                                  f'Serie: {serie_matches}')
                            break

                        else:
                            for nf, comp, nfe, serie, dta_comp in zip(replica_nota, chaves_comp,
                                                                             nfe_match,
                                                                             serie_matches, data_emissao):

                                try:
                                    if comp != 0:
                                        ch_comp = ''.join(comp.split())
                                    else:
                                        ch_comp = 0
                                except IndexError:
                                    ch_comp = 0
                                dta_emissao = 0
                                try:
                                    if dta_comp != 0:
                                        for x in range(1, 9):
                                            data = dta_comp.group(x)
                                            if data is not None:
                                                break
                                        dta_emissao = re.sub(r'[.]', '/', data)
                                except IndexError:
                                    dta_emissao = 0

                                try:
                                    for i in range(0, 6):
                                        notas = nfe.group(i)
                                        if notas is not None:
                                            break
                                    nota_trata = re.sub(r'[\D]', '', notas)

                                except IndexError:
                                    nota_trata = 0
                                serie_trata = 0
                                try:
                                    if serie != 0:
                                        for i in range(1, 8):
                                            series = serie.group(i)
                                            if series is not None:
                                                break
                                        serie_trata = re.sub(r'[\D]', '', series)
                                except IndexError:
                                    serie_trata = 0

                                for nota in replica_nota:

                                    nf_pdf_map[nota.lstrip('0')] = {
                                        'nota_fiscal': nota,
                                        'chave_acesso': '0',
                                        'data_email': message.received,
                                        'email_vinculado': message.subject,
                                        'serie_nf': serie_trata,
                                        'serie_comp': '0',
                                        'data_emissao': dta_emissao,
                                        'cnpj': cnpj_cofco,
                                        'nfe': nota_trata.lstrip('0'),
                                        'chave_comp': ch_comp,
                                        'transportadora': 'COFCO',
                                        'peso_comp': '0'

                                    }


                    os.remove(temp_pdf.name)
