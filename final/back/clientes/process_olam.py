import os
import re
import base64
import tempfile
import zipfile

import pandas as pd
from PyPDF2 import PdfReader

def process_olam(message, salve_folder, nf_excel_map):
    """
    Essa função é responsavel por ler e tratar os dados vindos do Cliente OLAM, extraído de um arquivo Excel,
    podendo ser alterado o estilo de captura da Nota Fiscal para se adequar ao projeto. A captura de um Excel é
    realizada pela coluna que é lida, diferente do PDF/ZIP que precisa achar um padrão.
    :param message: Variável pertencente a lista Messages.
    :param salve_folder: Local onde vai ser salvo o arquivo.
    :param nf_excel_map: Lista responsável por salvar os dados referente ao EXCEL.
    """
    corpo_email = message.body
    print(f"{message.received} - {message.subject}")
    if re.search(r'@olam\.com|@olamagri\.com', corpo_email):
        print('Tem Olam')
        teste = []
        zip_nfe = []
        zip_nota = []
        zip_cnpj = []
        if message.attachments:
                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    decoded_content = base64.b64decode(attachment.content)
                    if file_extension == ".zip":
                        print("teste zip")
                        with tempfile.TemporaryDirectory() as tmp_dir:
                            zip_file_path = os.path.join(tmp_dir, 'attachment.zip')
                            with open(zip_file_path, 'wb') as f:
                                f.write(decoded_content)
                            with zipfile.ZipFile(zip_file_path,'r') as zip_ref:
                                zip_ref.extractall(tmp_dir)
                            for file_name in os.listdir(tmp_dir):
                                file_path = os.path.join(tmp_dir, file_name)
                                if file_path.endswith('.pdf'):
                                    with open(file_path,'rb') as pdf_file:
                                        pdf_reader = PdfReader(pdf_file)
                                        pdf_text = ''
                                        for page in pdf_reader.pages:
                                            pdf_text += page.extract_text()

                                        notas_fiscais = []
                                        nota_comp = []
                                        cnpj_match = []
                                        primeiro_cnpj = []
                                        replica_nota = []
                                        replica_comp = []
                                        notas_fiscais.extend(re.finditer(
                                            r'REF.\s+NFe:\s+(\d{4,})|'
                                            r'nfe\s+referenciadas:\s+(\d+(?:\s*(?:,\s*|\s+E\s+|\s*e\s*|\s*/\s*)?\d+)*)|'
                                            r'Ref.\s+NF\s+n\s+(\d{4,})|'
                                            r'Eletronica\s+nr\s+(\d{4,})|'
                                            r'REF\s+NF\s+(\d{4,})|REF.\s+NF\s+(\d{4,})|'
                                            r'REFERENTE\s+A\s+(?:NOTA\s+FISCAL\s+|NF\s+)(\d{4,})', pdf_text,
                                            re.IGNORECASE
                                        ))
                                        if not notas_fiscais:
                                            notas_fiscais.append(0)
                                        nota_comp.extend(re.finditer(
                                            r'SÉRIE:\s*\d+\s*(\d+)', pdf_text, re.IGNORECASE
                                        ))
                                        if not nota_comp:
                                            nota_comp.append(0)

                                        cnpj_match.extend(
                                            re.finditer(r'(\d{2}.\d{3}.\d{3}/\d{4}-\d{2})', pdf_text,
                                                        re.IGNORECASE))

                                        if len(cnpj_match) >= 2:
                                            primeiro_cnpj.append(cnpj_match[0])
                                        if not primeiro_cnpj:
                                            primeiro_cnpj.append(0)



                                        for nt in nota_comp:
                                            replica_comp.append(nt.group(1))
                                            for valor in replica_comp:
                                                zip_nfe.append(valor)
                                                zip_cnpj.append(primeiro_cnpj)
                                                for i in notas_fiscais:
                                                    if i != 0:
                                                        if i != 'e':
                                                            notas = None
                                                            for x in range(1, 9):
                                                                notas = i.group(x)
                                                                if notas is not None:
                                                                    break
                                                            notax = re.split(r'\s*/\s*|\s*,\s*|\s*e\s*|\s+', notas)
                                                            if len(notas_fiscais) >= 2:
                                                                for i in notax:
                                                                    if i != 0:
                                                                        replica_nota.append(i)
                                                            else:
                                                                for i in notax:
                                                                    replica_nota.append(i)
                                                zip_nota.append(replica_nota)


                                    teste.append('ultimo')

                for attachment in message.attachments:
                    file_extension = os.path.splitext(attachment.name)[1].lower()
                    decoded_content = base64.b64decode(attachment.content)
                    if file_extension == '.xlsx':
                        print('teste xlsx')
                        cnpj_olam = '03902252001680'
                        with tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False) as temp_xlsx:
                            temp_xlsx.write(decoded_content)
                        try:
                            with pd.ExcelFile(temp_xlsx.name, engine='openpyxl') as excel_file:
                                sheet_names = excel_file.sheet_names

                                if "NOTAS ORIGEM" in sheet_names:
                                    df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=3,
                                                       sheet_name="NOTAS ORIGEM")  # COMPLEMENTO

                                    nf_ref_cleaned = df['Nº Nota Fiscal']
                                    chavenfe_cleaned = df['Chave de acesso de compra']  # CHAVE DE ACESSO COMPLEMENTO
                                    df2 = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=1,
                                                        sheet_name="COMPLEMENTOS")
                                    nota_comp = df2['Nº Nota Fiscal']
                                    chavecomp_cleaned = df2['Chave de acesso de compra']
                                    nota_ori = df2['NF ORIGEM'].dropna().astype(str)
                                    nota_ori = [nf.split('/')[0] for nf in nota_ori]

                                    for nf, chave, chavecomp, notacomp in zip(nf_ref_cleaned, chavenfe_cleaned,
                                                                              chavecomp_cleaned, nota_comp):

                                        chaves = ''.join(chave)
                                        nf_excel_map.append({'nota_fiscal': nf,
                                                             'data_email': message.received,
                                                             'chave_acesso': chaves,
                                                             'email_vinculado': message.subject,
                                                             'serie_nf': '0',
                                                             'nfe': notacomp,
                                                             'cnpj': cnpj_olam,
                                                             'chave_comp': chavecomp,
                                                             'data_emissao': '0',
                                                             'transportadora': 'OLAM',
                                                             'serie_comp': '0',
                                                             'peso_comp': '0'
                                                             })
                                elif "COMPLEMENTOS" in sheet_names:
                                    df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=1,
                                                       sheet_name="COMPLEMENTOS")
                                    nf_origem = df['NF ORIGEM'].astype(str).str.lstrip('0')
                                    chave_origem = df['CHAVE DE ACESSO NF ORIGEM']
                                    chave_comp = df['Chave de acesso COMPLEMENTO']
                                    cnpj_origem = df['CNPJ NF ORIGEM'].astype(str)
                                    nf_comp = df['Nº Nota Fiscal']
                                    serie = df['Série']
                                    peso = df['Peso Nota Fiscal (kg)']

                                    for nfo, chave, cnpj, nfe, series, pesos, chavecomp in zip(nf_origem, chave_origem,
                                                                                               cnpj_origem, nf_comp,
                                                                                               serie, peso, chave_comp
                                                                                                     ):
                                        nf_excel_map.append({
                                            'nota_fiscal': nfo,
                                            'data_email': message.received,
                                            'chave_acesso': chave,
                                            'chave_comp': chavecomp,
                                            'email_vinculado': message.subject,
                                            'serie_nf': series,
                                            'nfe': nfe,
                                            'cnpj': cnpj_olam,  # cnpj,
                                            'data_emissao': '0',
                                            'peso_nfe': pesos,
                                            'transportadora': 'OLAM',
                                            'serie_comp': '0',
                                            'peso_comp': '0'

                                        })
                                else:
                                    nf_excel_map.append({'nota_fiscal': '0',
                                                         'data_email': message.received,
                                                         'chave_acesso': 'SEM LEITURA',
                                                         'email_vinculado': message.subject,
                                                         'serie_nf': 'SEM LEITURA',
                                                         'data_emissao': 'SEM LEITURA',
                                                         'nfe': attachment.name,
                                                         'cnpj': 'SEM LEITURA',
                                                         'chave_comp': 'SEM LEITURA',
                                                         'peso_nfe': '0',
                                                         'serie_comp': 'SEM LEITURA',
                                                         'peso_comp': 'SEM LEITURA',
                                                         'transportadora': 'OLAM'})

                        except Exception as e:
                            print(f"Erro ao ler arquivo {e}")
                        finally:
                            os.remove(temp_xlsx.name)

                    if file_extension == '.xlsm':
                        with tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False) as temp_xlsx:
                            temp_xlsx.write(decoded_content)
                            pd.set_option('display.max_columns', None)
                            pd.set_option('display.max_rows', None)
                            print("Arquivo temporario gerado", temp_xlsx.name)
                            df = pd.read_excel(temp_xlsx.name, engine='openpyxl', header=0)
                            df.columns = df.columns.str.strip()
                            if 'NF' in df.columns and 'Chave de acesso' in df.columns:
                                nf_ref_cleaned = df['NF'].dropna().astype(str)
                                nf_ref_cleaned = [nf.split('.')[0] for nf in nf_ref_cleaned]
                                chave_ref_clean = df['Chave de acesso'].dropna().astype(
                                    str).str.extract(
                                    r'(\d{44})')
                                chave_ref_clean = chave_ref_clean.dropna().values.tolist()

                                serie_nfe = df['Série'].dropna().astype(str)
                                serie_nfe = [serie.split('.')[0] for serie in serie_nfe]

                                peso_olarim = df['Peso NF'].dropna().astype(str)
                                peso_olarim = [peso.split('.')[0] for peso in peso_olarim]

                                for nf, chave, nfe_serie, peso_nfe, nota, nota_comp, cnpj in zip(nf_ref_cleaned,
                                                                                           chave_ref_clean, serie_nfe,
                                                                                           peso_olarim, zip_nota,
                                                                                           zip_nfe,zip_cnpj):

                                    for i in cnpj:
                                        cn = i.group(1)
                                        cn_tratada = re.sub(r'[./-]', '', cn)
                                    for match in nota:
                                        chaves = ''.join(chave)
                                        nf_excel_map.append({'nota_fiscal': match,
                                                             'data_email': message.received,
                                                             'chave_acesso': '0',
                                                             'email_vinculado': message.subject,
                                                             'serie_nf': nfe_serie,
                                                             'data_emissao': '0',
                                                             'nfe': nf,
                                                             'cnpj': cn_tratada,
                                                             'chave_comp': chaves,
                                                             'peso_nfe': peso_nfe,
                                                             'transportadora': 'OLAM',
                                                             'serie_comp': '0',
                                                             'peso_comp': '0'
                                                             })
                        os.remove(temp_xlsx.name)
