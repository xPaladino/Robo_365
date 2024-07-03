import os,re,base64,tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_royal(message, save_folder, nf_pdf_map):
    corpo_email = message.body
    if re.search(r'@royalagro\.com',corpo_email):
        print('Tem Royal')
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

                        # print(chave_acesso_match)
                        chave_comp = []
                        chave_comp.extend(
                            re.finditer(
                                r'(?:CHAVE\s+DE\s+ACESSO\s+P/\s+CONSULTA\s+DE\s+AUTENTICIDADE\s*(\d{44}))|'
                                r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))',
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
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº)\s+(\d+(?:\.\d+)*)', pdf_reader,
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
                        #cnpj_royal = ''
                        for nf, chaves, serie_match, data_match, cnpj, nfe, chv_comp in zip(notas_fiscais, chave_acesso_match,
                                                                                  serie_matches, data_matches,
                                                                                  segundo_cnpj, nfe_match,chave_comp):


                            for i in range(0,2):
                                chave_comp = chv_comp.group(i)
                                if chave_comp is not None:
                                    break
                            chave_trata = re.sub(r'[\D]','',chave_comp)

                            if nf != 0:
                                if nf.group(1) == '73077':
                                    pass
                                else:
                                    if nf == 0:
                                        """if chaves != 0:
                                            chave = ''.join(chaves.split())
                                            # print(chave)
                                        else:
                                            chave = '0'
                                        if serie_match != 0:
                                            serie = serie_match.group(3)
                                        else:
                                            serie = 'NÃO ENCONTREI SERIE'
                                        if cnpj.group() != 0:
                                            cnpj_segundo = cnpj.group(0)
                                        else:
                                            cnpj_segundo = 'NÃO ENCONTREI CNPJ'
                                        if nfe.group() != 0:
                                            nfe_formatado = nfe.group(1)
                                        else:
                                            nfe_formatado = 'NÃO ENCONTREI NOTA COMP'

                                        data_emissao = 'NÃO ENCONTREI DATA'

                                        nf_formatado = nf
                                        nfe_semponto = nfe_formatado.replace('.', '')
                                        nfe_semzero = nfe_semponto if nfe_semponto[
                                                                      0:4] != '0000' else nfe_semponto[4:]
                                        cn_tratada = re.sub(r'[./-]', '', cnpj_segundo)
                                        nf_pdf_map[message.subject] = {
                                            'nota_fiscal': nf,
                                            'data_email': message.received,
                                            'chave_acesso': chave,
                                            'email_vinculado': message.subject,
                                            'serie_nf': serie,
                                            'data_emissao': data_emissao,
                                            'cnpj': cn_tratada,
                                            'nfe': nfe_semzero,
                                            'chave_comp': chave_trata,
                                            'transportadora': 'ROYAL'
                                       
                                        }"""
                                        pass
                                    elif chaves == 0:
                                        pass
                                    elif re.match(r'^\d{1,3}$', str(serie_match.group(2))):
                                        chave = ''.join(chaves.split())
                                        serie = serie_match.group(2)
                                        nf_formatado = nf.group(1)
                                        cnpj_segundo = cnpj.group(0)
                                        cn_tratada = re.sub(r'[./-]', '', cnpj_segundo)
                                        nfe_formatado = nfe.group(1)
                                        nfe_semponto = nfe_formatado.replace('.', '')
                                        nfe_semzero = nfe_semponto if nfe_semponto[
                                                                      0:4] != '0000' else nfe_semponto[4:]
                                        data_pdf = data_match.group(2)
                                        data_emissao = data_pdf.replace('.', '/') if data_match else '0'

                                        nf_pdf_map[nf_formatado] = {
                                            'nota_fiscal': nf_formatado,
                                            'data_email': message.received,
                                            'chave_acesso': chave if chave else None,
                                            'email_vinculado': message.subject,
                                            'serie_nf': serie,
                                            'data_emissao': data_emissao,
                                            'cnpj': cn_tratada,
                                            'nfe': nfe_semzero,
                                            'chave_comp': chave_trata,
                                            'transportadora': 'ROYAL'
                                        }
                                    elif data_match != 0 and data_match.group(1) is not None:
                                        # feito para tratar a royal
                                        chave = ''.join(chaves.split())
                                        serie = serie_match.group(3)
                                        nf_formatado = nf.group(1)
                                        cnpj_segundo = cnpj.group(0)
                                        nfe_formatado = nfe.group(1)
                                        nfe_semponto = nfe_formatado.replace('.', '')
                                        nfe_semzero = nfe_semponto if nfe_semponto[0:4] != '0000' else nfe_semponto[4:]
                                        data_pdf = data_match.group(1)
                                        data_emissao = data_pdf.replace('.', '/') if data_match else '0'
                                        nf_pdf_map[nf_formatado] = {
                                            'nota_fiscal': nf_formatado,
                                            'data_email': message.received,
                                            'chave_acesso': chave if chave else None,
                                            'email_vinculado': message.subject,
                                            'serie_nf': serie,
                                            'data_emissao': data_emissao,
                                            'cnpj': cn_tratada,
                                            'nfe': nfe_semzero,
                                            'chave_comp': chave_trata,
                                            'transportadora': 'ROYAL'

                                        }
                                    else:
                                        # feito para tratar a royal
                                        chave = ''.join(chaves.split())
                                        serie = serie_match.group(3)
                                        nf_formatado = nf.group(1)
                                        cnpj_segundo = cnpj.group(0)
                                        cn_tratada = re.sub(r'[./-]', '', cnpj_segundo)
                                        nfe_formatado = nfe.group(1)
                                        nfe_semponto = nfe_formatado.replace('.', '')
                                        nfe_semzero = nfe_semponto if nfe_semponto[0:4] != '0000' else nfe_semponto[4:]
                                        data_pdf = data_match.group(3)
                                        if nf_formatado is None:
                                            nf_formatado = nf.group(2).replace('.', '')  # tratamento para agricola gemelli
                                        data_emissao = data_pdf.replace('.', '/') if data_match else '0'
                                        nf_pdf_map[nf_formatado] = {
                                            'nota_fiscal': nf_formatado,
                                            'data_email': message.received,
                                            'chave_acesso': chave if chave else None,
                                            'email_vinculado': message.subject,
                                            'serie_nf': serie,
                                            'data_emissao': data_emissao,
                                            'cnpj': cn_tratada,
                                            'nfe': nfe_semzero,
                                            'chave_comp': chave_trata,
                                            'transportadora': 'ROYAL'
                                        }

                    os.remove(temp_pdf.name)