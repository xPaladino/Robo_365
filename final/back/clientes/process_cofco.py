import os, re, base64, tempfile
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text

def process_cofco(message, save_folder, nf_pdf_map):
    if re.search(r'@cofcointernational\.com',message.body):
        print('Tem COFCO')
        if message.attachments:
            print(f'{message.received} - {message.subject}')
            for attachment in message.attachments:
                file_extension = os.path.splitext(attachment.name)[1].lower()
                if file_extension == ".pdf":
                    decoded_content = base64.b64decode(attachment.content)
                    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                        temp_pdf.write(decoded_content)
                        temp_pdf_path = temp_pdf.name
                        pdf_reader = extract_text(temp_pdf_path)
                        cnpj_cofco = '06315338000542'
                        notas_fiscais = []
                        notas_fiscais.extend(re.finditer(
                            r'NF:\s*(\d{4,8})(?:\s*,\s*(\d{4,8}))*|'
                            r'nota\(s\)\s*fiscais:\s+(\d{4,9})|'
                            r'PESO\s+REF\.\s+AS\s+NFS\s+(\d+(?:\s*(?:,\s*|\s+e\s+|\s*/\s*)?\d+)*)|'
                            r'REF\.A\s+NF\s+(\d{4,8})',
                            pdf_reader, re.IGNORECASE))
                        if not notas_fiscais:
                            notas_fiscais.append(0)
                        """COMPLEMENTO
                        DE
                        PESO
                        REF.AS
                        NFS
                        31618, 31619, 31620, 31626, 31629, 31630, 31631, 31632, 31636, 31637,
                        31639, 31640, 31641, 31642, 31644, 31647, 31673
                        E
                        32437"""

                        replica_nota = []

                        serie_matches = []
                        serie_matches.extend(
                            re.finditer(r'(?:VALOR:\s+)(\d{9}\s+(\d+))|Série\s+(\d{1,})',
                                        pdf_reader, re.IGNORECASE))

                        nfe_match = []
                        nfe_match.extend(
                            re.finditer(r'(?:NF-e\s*No.|NF-e\s+Nº|VALOR:\s+)\s+(\d+(?:\.\d+)*)', pdf_reader,
                                        re.IGNORECASE))
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
                                r'(?:CHAVE\s+DE\s+ACESSO\s*(\b\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\b))',
                                pdf_reader, re.IGNORECASE))
                        for i in (notas_fiscais):
                            if i != 0:
                                numeros_notas_fiscais = re.findall(r'\d{4,9}', i.group())
                                for i in (numeros_notas_fiscais):
                                    replica_nota.append(i)
                            else:
                                pass

                        for nf, chave in zip(replica_nota,chave_acesso_match):
                            chaves = ''.join(chave.split())
                            for nf in replica_nota:

                                nf_pdf_map[nf] = {
                                    'nota_fiscal': nf,
                                    'chave_acesso': chaves,
                                    'data_email': message.received,
                                    'email_vinculado': message.subject,
                                    'serie_nf': '0',
                                    'data_emissao': '0',
                                    'cnpj': cnpj_cofco,
                                    'nfe': '0',
                                    'chave_comp': '0',
                                    'transportadora': 'COFCO'

                                }

                    os.remove(temp_pdf.name)