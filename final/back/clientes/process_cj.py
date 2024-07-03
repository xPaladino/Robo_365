import os, re, base64, tempfile
from PyPDF2 import PdfReader
import pandas as pd
from pdfminer.high_level import extract_text


def process_cj(message, save_folder, nf_pdf_map, nf_excel_map):
    corpo_email = message.body
    if re.search(r'@cjtrade\.net', corpo_email):
        print('Tem CJ')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            a = {}
            excel = 0
            pdf_nfe, pdf_chave, pdf_chave_comp, pdf_serie = [], [], [], []
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension in a:
                    a[file_extension] += 1
                else:
                    a[file_extension] = 1
            if '.xlsx' in a:
                excel += 1
            else:
                pass
            if excel == 1:
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    if file_extension == ".pdf":
                        decoded_content = base64.b64decode(attachment.content)
                        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                            temp_pdf.write(decoded_content)
                            pdf_reader = extract_text(temp_pdf.name)
                            chave_comp, serie_matches, nfe_match = [], [], []

                            chave_acesso = re.findall(
                                r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}',
                                pdf_reader, re.IGNORECASE)
                            if not chave_acesso:
                                chave_acesso.append(0)

                            chave_comp.extend(
                                re.finditer(
                                    r'NFe Ref\.:\((\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b)',
                                    pdf_reader, re.IGNORECASE))
                            if not chave_comp:
                                chave_acesso.append(0)

                            serie_matches.extend(
                                re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|'
                                            r'(?:FOLHA\s*)(\d{1,})',
                                            pdf_reader, re.IGNORECASE))

                            for match in chave_comp:
                                chave = match.group(1)
                                chaves = ''.join(chave.split())
                                pdf_chave_comp.append(chaves)

                            if not serie_matches:
                                serie_matches.append(0)
                            nfe_match.extend(
                                re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|NFe)\s+(\d+(?:\.\d+)*)',
                                            pdf_reader, re.IGNORECASE))
                            if not nfe_match:
                                nfe_match.append(0)

                            for i in nfe_match:
                                nfe_formatado = i.group(1)
                                tiraponto = nfe_formatado.replace('.', '')
                                tirazero = tiraponto if tiraponto[0:4] != '0000' else tiraponto[
                                                                                      4:]
                                pdf_nfe.append(tirazero)
                                for match in chave_acesso:
                                    pass
                                    if match == 0:
                                        pass
                                    else:
                                        pdf_chave.append(match)

                            for match in serie_matches:
                                serie = match.group(3)
                                if serie != None:
                                    pdf_serie.append(serie)
                                if match.group(4) != None:
                                    pdf_serie.append(match.group(4))

                        os.remove(temp_pdf.name)
                for attachment2 in message.attachments:
                    file_extension2 = os.path.splitext(attachment2.name)[1].lower()
                    if file_extension2 == '.xlsx':
                        decoded_content = base64.b64decode(attachment2.content)
                        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_xlsx:
                            temp_xlsx.write(decoded_content)
                            pd.set_option('display.max_columns', None)
                            pd.set_option('display.max_rows', None)
                            df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=1)
                            df.columns = df.columns.str.strip()
                            df2 = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=0)
                            df2.columns = df2.columns.str.strip()
                            cnpj_cj = '21294708000334'
                            #print(df2.columns)
                            if 'Unnamed: 0' in df2.columns:
                                pass
                                if 'NF sobra' in df.columns:
                                    nota_cleaned = df['NF'].dropna().astype(str)
                                    nota_cleaned = [nf.split('.')[0] for nf in nota_cleaned]
                                    nf_ref_cleaned = df['NF sobra'].dropna().astype(str)
                                    nf_ref_cleaned = [comp.split('.')[0] for comp in
                                                      nf_ref_cleaned]
                                    emissao_data = df['DATA DESCARGA'].dropna().astype(str)
                                    emissao_data = [dta.split('.')[0] for dta in emissao_data]
                                    peso = df['DIFERENÇA'].dropna().astype(str)
                                    peso = [pes.split('.')[0] for pes in peso]
                                    replica_chave, replica_nfe, replica_serie, pdf_ajeita = [], [], [], []
                                    ajeita = sorted(pdf_nfe)
                                    ajeita_nfe = sorted(nf_ref_cleaned)
                                    #print(df.columns)
                                    for i in ajeita:
                                        pdf_ajeita.append(i)

                                    for i, nf_ref in enumerate(nf_ref_cleaned):

                                        for nfe, serie, chave in zip(pdf_nfe, pdf_serie, pdf_chave):

                                            if nfe == nf_ref:
                                                replica_nfe.append(nfe)
                                                replica_serie.append(serie)
                                                replica_chave.append(chave)
                                            else:
                                                pass
                                    for nota, nf, nfe, chave_comp, \
                                            peso_nfe, data, serie in zip(
                                        nota_cleaned, nf_ref_cleaned, replica_nfe,
                                        replica_chave,
                                        peso, emissao_data, replica_serie):
                                        chaves = ''.join(chave_comp.split())

                                        nf_excel_map.append({'nota_fiscal': nota,
                                                             'data_email': message.received,
                                                             'chave_acesso': '0',
                                                             'serie_nf': serie,
                                                             'data_emissao': '0',
                                                             'nfe': nfe,
                                                             'chave_comp': chaves,
                                                             'cnpj': cnpj_cj,
                                                             'peso_nfe': peso_nfe,
                                                             'email_vinculado': message.subject,
                                                             'transportadora': 'CJ'})

                            else:
                                if 'COMPLEMENTO' in df2.columns and 'NOTA' in df2.columns:
                                    nota_cleaned = df2['NOTA'].dropna().astype(str)
                                    nota_cleaned = [nf.split('.')[0] for nf in nota_cleaned]
                                    nf_ref_cleaned = df2['COMPLEMENTO'].dropna().astype(str)
                                    nf_ref_cleaned = [comp.split('.')[0] for comp in
                                                      nf_ref_cleaned]
                                    emissao_data = df2['DATA EMISSÃO'].dropna().astype(str)
                                    emissao_data = [dta.split('.')[0] for dta in emissao_data]
                                    cnpj_nfe = df2['CNPJ'].dropna().astype(str)
                                    cnpj_nfe = [re.sub(r'[/\-.]', '', cnpj) for cnpj in
                                                cnpj_nfe]
                                    peso = df2['DIF'].dropna().astype(str)
                                    peso = [pes.split('.')[0] for pes in peso]
                                    replica_chave, replica_nfe, replica_chave_comp, replica_serie = [], [], [], []
                                    for i, nf_ref in enumerate(nf_ref_cleaned):
                                        for nfe, chave, chave_comp, series in zip(pdf_nfe,
                                                                                  pdf_chave,
                                                                                  pdf_chave_comp,
                                                                                  pdf_serie):
                                            if nfe == nf_ref:
                                                replica_nfe.append(nfe)
                                                replica_chave.append(chave)
                                                replica_chave_comp.append(chave_comp)
                                                replica_serie.append(series)
                                            else:
                                                pass

                                    for nota, nf, nfe, chave, chaves_comp, serie, dta_emiss, \
                                            cnpj, peso_nfe in zip(
                                        nota_cleaned,
                                        nf_ref_cleaned,
                                        replica_nfe,
                                        replica_chave,
                                        replica_chave_comp,
                                        replica_serie,
                                        emissao_data,
                                        cnpj_nfe,
                                        peso):
                                        chaves = ''.join(chaves)

                                        nf_excel_map.append({'nota_fiscal': nota,
                                                             'data_email': message.received,
                                                             'chave_acesso': chaves_comp,
                                                             'serie_nf': serie,
                                                             'data_emissao': '0',
                                                             'nfe': nf,
                                                             'chave_comp': chaves,
                                                             'cnpj': cnpj_cj,
                                                             'peso_nfe': peso_nfe,
                                                             'email_vinculado': message.subject,
                                                             'transportadora': 'CJ'})
                        os.remove(temp_xlsx.name)
            else:

                print(excel)
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    if file_extension == ".pdf":
                        print("CAMINHO PDF")
                        decoded_content = base64.b64decode(attachment.content)
                        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                            temp_pdf.write(decoded_content)
                            pdf_reader = extract_text(temp_pdf.name)
                            chave_acesso = (re.findall(
                                r'(?<!NFe Ref\.:\()\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b',
                                pdf_reader, re.IGNORECASE))
                            if not chave_acesso:
                                chave_acesso.append(0)

                            notas_fiscais, segundo_cnpj, cnpj_match, nfe_match, data_matches, serie_matches, replica_nota = [], [], [], [], [], [], []
                            replica_cnpj = []
                            cnpj_cj = '21294708000334'

                            notas_fiscais.extend(
                                re.finditer(r'(?:número:|REF\s+A\s+NOTA'
                                            r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
                                            r'|referente\s+a\s+NF|'
                                            r'NF:)'
                                            r'\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                            r'nota\s+complementar\s+ref\s+a\s+nfe\s*(?:\d+\s*,\s*)?(\d{4,})'
                                            r'|nota\s+complementar\s+ref.\s+NF\s+n\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                            r'\b(\d+)\s+(?:de\s+\d{2}/\d{2}/\d{4})|'
                                            r'NF\(\s*s\)\s*\(([\d,\s*]+)\)|'
                                            r'NFES\s+((?:\d+\s*,\s*)*\d+\s*E\s*\d+)',
                                            # r'NFs\s+de\s+(\d+/\d+/\d+)\s+\((.*?)\)',

                                            pdf_reader,
                                            re.IGNORECASE))
                            if not notas_fiscais:
                                notas_fiscais.append(0)

                            for i in notas_fiscais:
                                if (len(notas_fiscais)) > 1:
                                    nota = i.group(4)

                            cnpj_match.extend(
                                re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_reader,
                                            re.IGNORECASE))
                            if len(cnpj_match) >= 2:
                                segundo_cnpj.append(cnpj_match[1])
                            if not segundo_cnpj:
                                segundo_cnpj.append(0)
                            nfe_match.extend(
                                re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|NFe)\s+(\d+(?:\.\d+)*)',
                                            pdf_reader, re.IGNORECASE))
                            if not nfe_match:
                                nfe_match.append(0)

                            data_matches.extend(re.finditer(
                                r'EMISSÃO\s+(\d{2}\.\d{2}\.\d{2,}|\d{2}/\d{2}/\d{2,})|'
                                r'EMISSAO\s*(\d{2}/\d{2}/\d{2,})|'
                                r'(?:EMISSÃO:\s+)(\d{2}/\d{2}/\d{2,})'
                                , pdf_reader, re.IGNORECASE))
                            if not data_matches:
                                data_matches.append(0)

                            serie_matches.extend(
                                re.finditer(
                                    r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|FOLHA\s+(\d{1,})',
                                    pdf_reader, re.IGNORECASE))
                            if not serie_matches:
                                serie_matches.append(0)

                            for notas in notas_fiscais:
                                if notas != 0:
                                    nota = None
                                    for x in range(0,15):
                                        nota = notas.group(x)
                                        if nota is not None:
                                            break
                                    #print(nota)
                            for nf, chave in zip(notas_fiscais, chave_acesso):
                                nota = []
                                for i in notas_fiscais:
                                    if i != 0:
                                        if i != 'e':
                                            notas = None
                                            for x in range(0, 12):
                                                notas = i.group(x)
                                                if notas is not None:
                                                    break
                                            #nota_trata = re.sub(r'[\D]', '', notas)
                                            nota.append(notas)
                                            for i in nota:
                                                nota_split = re.split(r',',i)
                                                for x in nota_split:
                                                    print(x)

                                for i in nota:
                                    procura = re.findall(r'\d+',i)
                                    filtra = [num for num in procura if len(num) > 1]

                                    replica_nota.append(filtra)
                                    replica_cnpj.append(cnpj_cj)

                            #for nota,cnpj,chave,serie,nfe in zip(replica_nota,replica_cnpj,chave_acesso,serie_matches,nfe_match):
                            for nota,cnpj,serie in zip(replica_nota,replica_cnpj,serie_matches):
                                """chaves = ''.join(chave.split())

                                
                                for i in range(0,2):
                                    nfe_comp = nfe.group(i)
                                    if nfe_comp is not None:
                                        break

                                nfe_split = re.sub(r'[\D]','',nfe_comp)
                                """
                                for i in range(0,3):
                                    series = serie.group(i)
                                    if series is not None:
                                        break
                                serie_split = re.sub(r'[\D]','',series)
                                for match in nota:
                                    print(match)
                                    nf_pdf_map[match] = {
                                                'nota_fiscal': match,
                                                'data_email': message.received,
                                                'chave_acesso': '0',
                                                'email_vinculado': message.subject,
                                                'serie_nf': serie_split,
                                                'data_emissao': '0',
                                                'cnpj': cnpj,
                                                'nfe': '0',
                                                'chave_comp': '0',
                                                'transportadora': 'CJ'
                                    }

                            """for chave, cnpj, nfe_comp, emissao, serie, nota in zip(
                                    chave_acesso,
                                    cnpj_match,
                                    nfe_match,
                                    data_matches,
                                    serie_matches,
                                    notas_fiscais):
                                chaves = ''.join(chave.split())
                                cn = cnpj.group(1)
                                cn_tratada = re.sub(r'[./-]', '',
                                                    cn)  # remoção de ponto, barra e traço
                                cn_cj = '21294708000334'
                                nfe = nfe_comp.group(1)
                                nfe_ponto = re.sub(r'[.]', '', nfe)  # remoção do ponto
                                nfe_tratada = nfe_ponto if nfe_ponto[
                                                           0:4] != '0000' else nfe_ponto[4:]
                                dta_emissao = emissao.group(3)
                                #print(notas_fiscais)
                                for i in range(0, 3):
                                    series = serie.group(i)
                                    if series is not None:
                                        break
                                #print(series)
                                #print(nfe_match)
                                series = serie.group(3)
                                if series == None:
                                    series = serie.group(2)
                                if series == None:
                                    series = serie.group(1)
                                if series == None:
                                    series = serie.group(4)
                                if nota == 0:
                                    print("nota nao encontrada")
                                    nf_pdf_map[message.subject] = {
                                        'nota_fiscal': '0',
                                        'data_email': message.received,
                                        'chave_acesso': '0',
                                        'email_vinculado': message.subject,
                                        'serie_nf': series,
                                        'data_emissao': dta_emissao,
                                        'cnpj': cn_cj,
                                        'nfe': nfe_tratada,
                                        'chave_comp': chaves,
                                        'transportadora': 'CJ'
                                    }
                                else:
                                    notas = nota.group(2)

                                    if notas == None:
                                        notas = nota.group(3)
                                    if notas == None:
                                        notas = nota.group(2)
                                    if notas == None:
                                        notas = nota.group(1)
                                    if notas == None:
                                        print('não há nota')
                                    else:
                                        nf_pdf_map[notas] = {
                                            'nota_fiscal': notas,
                                            'data_email': message.received,  # data_email,
                                            'chave_acesso': chaves,
                                            'email_vinculado': message.subject,
                                            'serie_nf': series,
                                            'data_emissao': dta_emissao,
                                            'cnpj': cn_cj,
                                            'nfe': nfe_tratada,
                                            'chave_comp': '0',
                                            'transportadora': 'CJ'
                                        }
                                    if notas == None:
                                        notas = nota.group(4)
                                        for i in notas_fiscais:
                                            if (len(notas_fiscais)) > 1:
                                                nota = i.group(4)
                                                # notaseparada = re.split(r'de', nota)
                                                nfe_ajust = nfe_tratada if nfe_tratada[
                                                                           0:3] != '000' else nfe_tratada[
                                                                                              3:]
                                                nf_pdf_map[nota] = {
                                                    'nota_fiscal': nota,
                                                    'data_email': message.received,  # data_email,
                                                    'chave_acesso': chaves,
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': series,
                                                    'data_emissao': dta_emissao,
                                                    'cnpj': cn_cj,
                                                    'nfe': nfe_ajust,
                                                    'chave_comp': '0',
                                                    'transportadora': 'CJ'
                                                }

                                    if notas is None:
                                        #print(notas)
                                        notas = nota.group(5)
                                        if notas is not None:

                                            notaseparada = re.split(r',', notas)
                                            #print(notaseparada)
                                            if dta_emissao == None:
                                                dta_emissao = emissao.group(2)
                                            for notaajustada in notaseparada:
                                                chave_ajustada = ''.join(notaajustada.split())
                                                nf_pdf_map[chave_ajustada] = {
                                                    'nota_fiscal': chave_ajustada,
                                                    'data_email': message.received,  # data_email,
                                                    'chave_acesso': chaves,
                                                    'email_vinculado': message.subject,
                                                    'serie_nf': series,
                                                    'data_emissao': dta_emissao,
                                                    'cnpj': cn_cj,
                                                    'nfe': nfe_tratada,
                                                    'chave_comp': '0',
                                                    'transportadora': 'CJ'
                                                }
                                    if notas == None:
                                        notas = nota.group(6)
                                        notaseparada = re.split(r'\s*,\s*|\s+E\s+', notas)

                                        for ajuste in notaseparada:
                                            chave_ajustada = ''.join(ajuste.split())
                                            nf_pdf_map[chave_ajustada] = {
                                                'nota_fiscal': chave_ajustada,
                                                'data_email': message.received,  # data_email,
                                                'chave_acesso': chaves,
                                                'email_vinculado': message.subject,
                                                'serie_nf': series,
                                                'data_emissao': dta_emissao,
                                                'cnpj': cn_cj,
                                                'nfe': nfe_tratada,
                                                'chave_comp': '0',
                                                'transportadora': 'CJ'
                                            }
"""
                        os.remove(temp_pdf.name)
