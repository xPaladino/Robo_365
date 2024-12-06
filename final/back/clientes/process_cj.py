import os, re, base64, tempfile
from PyPDF2 import PdfReader
import pandas as pd
from pdfminer.high_level import extract_text
from dateutil.parser import parse


def convert_to_dia_mes_ano(date_str):
    try:
        # Tenta analisar a data em diferentes formatos
        date_obj = parse(date_str, dayfirst=False)  # Primeiro tenta MES/DIA/ANO
        if date_obj.month > 12:
            date_obj = parse(date_str, dayfirst=True)  # Se falhar, tenta DIA/MES/ANO
        return date_obj.strftime('%d/%m/%Y')
    except ValueError:
        # Retorna a string original se a conversão falhar
        return date_str


def process_cj(message, save_folder, nf_pdf_map, nf_excel_map):
    """
    Essa função é responsável por ler e tratar os dados vindos do Cliente CJ, extraído de um arquivo EXCEL e/ou PDF,
    nesse caso em específico, se houver ao menos um arquivo Excel sendo encaminhado pelo cliente, é feito um tratamento
    em baseado na NFE do PDF combinando com a mesma informação vinda do EXCEL.
    Caso encontre somente o PDF, o padrão é esse:

    Padrão PDF 1.0:

    notas_fiscais.extend(
    re.finditer(r'(?:número:|REF\s+A\s+NOTA'
    r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
    r'|referente\s+a\s+NF|a\s+NFE|'
    r'COMPLEMENTO\s+DA\s+NF|'
    r'NF:|REFERENTE\s+NFO)'
    r'\s*(?:\d+\s*,\s*)?(\d{4,})|'
    r'nota\s+complementar\s+ref\s+a\s+nfe\s*(?:\d+\s*,\s*)?(\d{4,})'
    r'|nota\s+complementar\s+ref.\s+NF\s+n\s*(?:\d+\s*,\s*)?(\d{4,})|'
    r'NF\(\s*s\)\s*\(([\d,\s*]+)\)|'
    r'NFES\s+((?:\d+\s*,\s*)*\d+\s*E\s*\d+)|'
    r'FISCAIS\s+No\s+\d+(?:/\d+)?\s+E\s+\d+|'
    r'NOTAS\s+REFERENCIA:\s+\d+(?:/\d+)?\s+E\s+\d+|'
    r'FISCAIS\s+No\s*((?:\d+(?:,\s*)?)+)|'
    r'NOTAS\s+REFERENCIA:\s+((?:\d+(?:,\s*)?)+(?:\s+E\s+\d+)?)',
    pdf_reader,
    re.IGNORECASE))

    Caso seja realizado alguma alteração, favor documentar.
    :param message: Variável pertencente a lista Messages.
    :param save_folder: Local onde vai ser salvo o arquivo.
    :param nf_pdf_map: Dicionário responsável por salvar os dados referente aos PDF's.
    :param nf_excel_map: Lista responsável por salvar os dados referente aos EXCEL's.
    """
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
                                tirazero1 = tiraponto if tiraponto[0:4] != '0000' else tiraponto[
                                                                                       4:]
                                tirazero2 = tirazero1.lstrip('0')
                                pdf_nfe.append(tirazero2)
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

                                        nf_excel_map[nota] = {'nota_fiscal': nota,
                                                              'data_email': message.received,
                                                              'chave_acesso': '0',
                                                              'serie_nf': serie,
                                                              'data_emissao': '0',
                                                              'nfe': nfe,
                                                              'chave_comp': chaves,
                                                              'cnpj': cnpj_cj,
                                                              'peso_nfe': peso_nfe,
                                                              'email_vinculado': message.subject,
                                                              'transportadora': 'CJ',
                                                              'serie_comp': '0',
                                                              'peso_comp': '0'
                                                              }

                            else:
                                if 'NF' in df2.columns and 'COMPLEMENTO' in df2.columns:
                                    nota_cleaned = df2['NF'].dropna().astype(str)
                                    nota_cleaned = [nf.split('.')[0] for nf in nota_cleaned]
                                    nf_ref_cleaned = df2['COMPLEMENTO'].dropna().astype(str)
                                    nf_ref_cleaned = [comp.split('.')[0] for comp in
                                                      nf_ref_cleaned]
                                    emissao_data = df2['DATA DESCARGA'].dropna().astype(str)
                                    emissao_data = [dta.split('.')[0] for dta in emissao_data]
                                    data_ajust = []
                                    for x in emissao_data:
                                        dta = x.split('.')[0]
                                        try:
                                            date_obj = parse(dta, dayfirst=False)
                                            if date_obj.month > 12:
                                                date_obj = parse(dta, dayfirst=True)
                                            data_formatada = date_obj.strftime('%d/%m/%Y')
                                        except ValueError:
                                            data_formatada = dta
                                        data_ajust.append(data_formatada)
                                    cnpj_nfe = df2['CNPJ'].dropna().astype(str)
                                    cnpj_nfe = [re.sub(r'[/\-.]', '', cnpj) for cnpj in
                                                cnpj_nfe]
                                    peso = df2['DIF'].dropna().astype(str)
                                    peso = [pes.split('.')[0] for pes in peso]
                                    replica_nf, replica_chaves, replica_series = [], [], []
                                    for i, nf_ref in enumerate(nf_ref_cleaned):
                                        for x, ch, se in zip(pdf_nfe, pdf_chave, pdf_serie):
                                            if x == nf_ref:
                                                replica_nf.append(x)
                                                replica_chaves.append(ch)
                                                replica_series.append(se)

                                    for nota, nfe, dta, peso_nota, ch_comp, se_comp in zip(nota_cleaned, nf_ref_cleaned,
                                                                                           data_ajust, peso,
                                                                                           replica_chaves,
                                                                                           replica_series):
                                        chs_comp = ''.join(ch_comp.split())

                                        nf_excel_map[nota] = {'nota_fiscal': nota,
                                                              'data_email': message.received,
                                                              'chave_acesso': '0',
                                                              'serie_nf': se_comp,
                                                              'data_emissao': dta,
                                                              'nfe': nfe,
                                                              'chave_comp': chs_comp,
                                                              'cnpj': cnpj_cj,
                                                              'peso_nfe': peso_nota,
                                                              'email_vinculado': message.subject,
                                                              'transportadora': 'CJ',
                                                              'serie_comp': '0',
                                                              'peso_comp': '0'
                                                              }

                                elif 'COMPLEMENTO' in df2.columns and 'NOTA' in df2.columns:
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

                                        nf_excel_map[nota] = {'nota_fiscal': nota,
                                                              'data_email': message.received,
                                                              'chave_acesso': chaves_comp,
                                                              'serie_nf': serie,
                                                              'data_emissao': dta_emiss,
                                                              'nfe': nf,
                                                              'chave_comp': chaves,
                                                              'cnpj': cnpj_cj,
                                                              'peso_nfe': peso_nfe,
                                                              'email_vinculado': message.subject,
                                                              'transportadora': 'CJ',
                                                              'serie_comp': '0',
                                                              'peso_comp': '0'
                                                              }
                        os.remove(temp_xlsx.name)
            else:
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
                            replica_data = []
                            replica_serie = []
                            replica_chave = []
                            replica_nfe = []
                            cnpj_cj = '21294708000334'
                            # print(chave_acesso)
                            notas_fiscais.extend(
                                re.finditer(r'REF.\s+Compl.\s+NF:\s*([\d\s,]+)|'
                                            r'(?:número:|REF\s+A\s+NOTA|REFERENTE\s+A\s+NOTA'
                                            r'|NF\s+Nº:|ORIGEM\s+NR.:|referente\s+NF'
                                            r'|referente\s+a\s+NF|a\s+NFE|'
                                            r'COMPLEMENTO\s+DA\s+NF|'
                                            r'NF:|REFERENTE\s+NFO|NOTA\s+FISCAL|REF.\s+A\s+NF.)'
                                            r'\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                            r'NFS\s+(\d+\s*)+|'
                                            r'nota\s+complementar\s+ref\s+a\s+nfe\s*(?:\d+\s*,\s*)?(\d{4,})'
                                            r'|nota\s+complementar\s+ref.\s+NF\s+n\s*(?:\d+\s*,\s*)?(\d{4,})|'
                                            r'NF\(\s*s\)\s*\(([\d,\s*]+)\)|'
                                            r'NFES\s+((?:\d+\s*,\s*)*\d+\s*E\s*\d+)|'
                                            r'FISCAIS\s+No\s+((?:\d+/\d+)?\s+E\s+\d+)|'
                                            r'NOTAS\s+REFERENCIA:\s+\d+(?:/\d+)?\s+E\s+\d+|'
                                            r'FISCAIS\s+No\s*((?:\d+(?:,\s*)?)+)|'
                                            r'NOTAS\s+REFERENCIA:\s+((?:\d+(?:,\s*)?)+(?:\s+E\s+\d+)?)|'
                                            r'REFERENTE\s+AS\s+NOTAS\s*(\d+\s+E\s+\d+|\d.\s/]+)|'
                                            r'complementar\s+da\(s\)\s+nota\(s\):\s*([\d\s,]+)|'
                                            r'REFERENTE\s+AS\s+NFS\s*([\d\s,]+)|'
                                            r'NOTAS\s+FISCAIS\s+N\s*([\d\s,]+?\s+E\s+\d+)|'
                                            r'REF\s+COMPL.\s+NOTA\s+((?:\d+\s*,\s*)*\d+\s*E\s*\d+)|'
                                            r'NF\s+Nr\s+(\d+)',
                                            pdf_reader,
                                            re.IGNORECASE))
                            if not notas_fiscais:
                                notas_fiscais.append(0)
                            cnpj_match.extend(
                                re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_reader,
                                            re.IGNORECASE))
                            if len(cnpj_match) >= 2:
                                segundo_cnpj.append(cnpj_match[1])
                            if not segundo_cnpj:
                                segundo_cnpj.append(0)
                            nfe_match.extend(
                                re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº'
                                            r'|RECEBEDOR\s*NFe)'
                                            # r'|NFe)'
                                            r'\s+(\d+(?:\.\d+)*)',
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
                                        'transportadora': 'CJ',
                                        'peso_comp': 'SEM LEITURA',
                                        'serie_comp': 'SEM LEITURA'
                                    }
                                    print('zerado')

                            serie_matches.extend(
                                re.finditer(
                                    r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})|FOLHA\s+(\d{1,})',
                                    pdf_reader, re.IGNORECASE))
                            if not serie_matches:
                                serie_matches.append(0)

                            for nf, chave in zip(notas_fiscais, chave_acesso):
                                nota = []
                                for i in notas_fiscais:

                                    if i != 0 and i is not None:
                                        print(i)
                                        if i != 'e':
                                            print('aqui')
                                            notas = None
                                            for x in range(1, 20):
                                                notas = i.group(x)
                                                if notas is not None:
                                                    break

                                            if notas and notas != 0:
                                                print(notas)
                                                notas_split = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|-|\s+', notas)
                                                for valor in notas_split:
                                                    print(valor)
                                                    if valor != 'E':
                                                        if valor != '':
                                                            nota.append(valor)
                                                            print(valor)
                                            else:
                                                print('zero')

                                            # nota_trata = re.sub(r'[\D]', '', notas)
                                            #print(notas)

                                            #nota.append(notas)

                                for i in nota:
                                    print(f' a {i}')
                                    if i is not None:
                                        procura = re.findall(r'\d+', i)

                                        filtra = [num for num in procura if len(num) > 1]

                                        replica_nota.append(filtra)
                                        replica_cnpj.append(cnpj_cj)
                                        for dta, chavea, nfe, serie in zip(data_matches,chave_acesso,nfe_match, serie_matches):
                                            try:
                                                chavesa = ''.join(chavea.split())
                                            except IndexError:
                                                chavesa = 0

                                            try:
                                                for i in range(1, 4):
                                                    nfe_comp = nfe.group(i)
                                                    if nfe_comp is not None:
                                                        break

                                                nfe_split = re.sub(r'[\D]', '', nfe_comp)
                                            except IndexError:
                                                nfe_split = 0

                                            try:
                                                for x in range(1, 4):
                                                    data_emissao = dta.group(x)
                                                    if data_emissao is not None:
                                                        break
                                            except IndexError:
                                                data_emissao = 0

                                            try:
                                                for i in range(0, 3):
                                                    series = serie.group(i)
                                                    if series is not None:
                                                        break
                                                serie_split = re.sub(r'[\D]', '', series)
                                            except IndexError:
                                                serie_split = 0

                                            replica_data.append(data_emissao)
                                            replica_chave.append(chavesa)
                                            replica_nfe.append(nfe_split)
                                            replica_serie.append(serie_split)

                            for nota, cnpj, serie, dta, chave, nfe in zip(replica_nota, replica_cnpj, replica_serie,
                                                                          replica_data, replica_chave, replica_nfe):



                                for match in nota:
                                    ajusta = match.replace('.', '')
                                    nfe_ajusta = nfe_split.lstrip('0')
                                    nf_pdf_map[ajusta] = {
                                        'nota_fiscal': ajusta,
                                        'data_email': message.received,
                                        'chave_acesso': '0',
                                        'email_vinculado': message.subject,
                                        'serie_nf': serie,
                                        'data_emissao': dta,
                                        'cnpj': cnpj,
                                        'nfe': nfe_ajusta,
                                        'chave_comp': chave,
                                        'transportadora': 'CJ',
                                        'peso_comp': '0',
                                        'serie_comp': '0'
                                    }

                        os.remove(temp_pdf.name)
