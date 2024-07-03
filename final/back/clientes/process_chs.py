import os,re,base64,tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_chs(message, save_folder, nf_pdf_map):
    corpo_email = message.body
    if re.search(r'@chsinc\.com',corpo_email):
        print('')
        print('Tem CHS')
        if message.attachments:
            #print(corpo_email)
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name
                        pdf_reader = extract_text(temp_pdf_path)
                        notas_fiscais = []
                        replica_chave = []
                        notas_fiscais.extend(re.finditer(
                            r'(?:#NF:|Nota\s+Fiscal:|fiscais:|NF:|'  # padrao
                            r'Ref\s+NF|'  # adicionado royal
                            r'NF\s+n|'  # adicionado royal
                            r'Nfe\s+de\s+n\s+:|'
                            r'Referente\s+NF|'
                            r'nota\(s\)\s+fiscal\(is\):|'
                           # r'REFERENTE\s+A\s+NF|'
                            r'Nota\(s\)\s+de\s+Origem:\s*|'
                            r'REF\s+A\s+NOTA)'  # adicionado para usimat destilaria
                            r'\s*(?:\d+\s*,\s*)?(\d{4,8})|'  # padrao 
                            r'ORIGEM\s+NR\.: (\d+(\.\d+)?)|'  # adicionado para agricola gemelli
                            r'(:?REF\s+NFS\s+)(\d+/\d+/\d+)\s+\((.*?)\)',
                            pdf_reader, re.IGNORECASE))
                        if not notas_fiscais:
                            notas_fiscais.append(0)

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
                                r'(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b)',
                                pdf_reader, re.IGNORECASE))
                        if not chave_comp:
                            chave_acesso_match.append(0)

                        serie_matches = []
                        serie_matches.extend(
                            re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})',
                                        pdf_reader, re.IGNORECASE))
                        if not serie_matches:
                            serie_matches.append(0)
                        data_matches = []
                        data_matches.extend(re.finditer(
                            r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                            r'EMISSAO\s+(\d{2}/\d{2}/\d{2,})|'
                            r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                            , pdf_reader, re.IGNORECASE))
                        if not data_matches:
                            data_matches.append(0)

                        nfe_match = []
                        nfe_match.extend(
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|Nf-e\s+Nº.)\s+(\d+(?:\.\d+)*)', pdf_reader,
                                        re.IGNORECASE))
                        segundo_cnpj = []
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

                        for nf, chave, serie, cnpj, nfe in zip(notas_fiscais, chave_comp, serie_matches,segundo_cnpj, nfe_match):

                            for i in range(0, 2):
                                chave_comp = chave.group(i)
                                if chave_comp is not None:
                                    break
                            chave_trata = re.sub(r'[\D]', '', chave_comp)
                            for i in range(0,2):
                                serie_comp = serie.group(i)
                                if serie_comp is not None:
                                    break
                            serie_trata = re.sub(r'[\D]','',serie_comp)

                            cn = cnpj.group(1)
                            cnpj_trata = re.sub(r'[./-]', '', cn)

                            nfe_formatado = nfe.group(1)
                            nfe_semponto = nfe_formatado.replace('.', '')
                            nfe_semzero = nfe_semponto if nfe_semponto[
                                                          0:3] != '000' else nfe_semponto[3:]

                            nota = []
                            for i in notas_fiscais:
                                if i != 0:
                                    if i != 'e':
                                        notas = None
                                        for x in range(0, 10):
                                            notas = i.group(x)
                                            if notas is not None:
                                                break
                                        nota_trata = re.sub(r'[\D]', '', notas)
                                        nota.append(nota_trata)

                            if nf == 0:
                                nf_pdf_map[nfe_semzero] = {
                                    'nota_fiscal': nf,
                                    'data_email': message.received,
                                    'chave_acesso': chave_trata,
                                    'email_vinculado': message.subject,
                                    'serie_nf': serie_trata,
                                    'data_emissao': '0',
                                    'cnpj': cnpj_trata,
                                    'nfe': nfe_semzero,
                                    'chave_comp': chave_trata,
                                    'transportadora': 'CHS'
                                }
                            else:
                                nf_pdf_map[nota_trata] = {
                                    'nota_fiscal': nota_trata,
                                    'data_email': message.received,
                                    'chave_acesso': chave_trata,
                                    'email_vinculado': message.subject,
                                    'serie_nf': serie_trata,
                                    'data_emissao': '0',
                                    'cnpj': cnpj_trata,
                                    'nfe': nfe_semzero,
                                    'chave_comp': chave_trata,
                                    'transportadora': 'CHS'
                                }

                    os.remove(temp_pdf.name)