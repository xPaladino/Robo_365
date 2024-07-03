from openpyxl import Workbook
from openpyxl.styles import PatternFill
from datetime import datetime
import pyodbc
from dotenv import load_dotenv
import os


def parse_datetime(date_str):
    formats = ['%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S']
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Data'{date_str}")


def trocar_cores(ws, row_index):
    # largura
    for column in ['A', 'E', 'F','J']:
        ws.column_dimensions[column].width = 10
    for column in ['H', 'I']:
        ws.column_dimensions[column].width = 45
    for column in ['A', 'D']:
        ws.column_dimensions[column].width = 6
    for column in ['G','B']:
        ws.column_dimensions[column].width = 18
    for column in ['M']:
        ws.column_dimensions[column].width = 16
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['L'].width = 27
    ws.column_dimensions['N'].width = 59

    # cores
    for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K','L','M','N']:
        fundo = PatternFill(start_color="a6e1fa", end_color="a6e1fa", fill_type="solid")
        ws[f"{column}{row_index}"].fill = fundo
    for column in ['F', 'I']:
        amarelo = PatternFill(start_color="ffea00", end_color="ffea00", fill_type="solid")
        ws[f"{column}{row_index}"].fill = amarelo
    for column in ['D','E', 'H','J','K']:
        marrom = PatternFill(start_color="ffe6a7", end_color="ffe6a7", fill_type="solid")
        ws[f"{column}{row_index}"].fill = marrom


def save_to_excel(nf_pdf_map, nf_zip_map, nf_excel_map, file_name):

    wb = Workbook()
    ws = wb.active
    ws.title = 'Notas Fiscais'
    ws['A1'] = 'Tipo'
    ws['B1'] = 'Data Email'
    ws['C1'] = 'Transportadora'
    ws['D1'] = 'Série'
    ws['E1'] = 'Nota Fiscal'
    ws['F1'] = 'NF Comp'
    ws['G1'] = 'Data Descarga'
    ws['H1'] = 'Chave de Acesso'
    ws['I1'] = 'Chave NFE Comp'
    ws['J1'] = 'Peso'
    ws['K1'] = 'Lote'
    ws['L1'] = 'Produto'
    ws['M1'] = 'Cnpj Destinatario'
    ws['N1'] = 'Origem'

    row_index = 2
    resultados_pdf, resultados_zip, resultados_excel = conecta_sql(nf_zip_map,nf_pdf_map,nf_excel_map)

    try:
        #for nf_formatado, info in nf_pdf_map.items():
        for (nf_formatado, info), resultado in zip(nf_pdf_map.items(),resultados_pdf):
            data_descarga = str(resultado[8])
            formatob = '%d/%m/%Y %H:%M:%S'

            data_formatada = parse_datetime(data_descarga)
            data_ajustada = data_formatada.strftime(formatob)
            ws.cell(row=row_index, column=1, value="PDF")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2, value=str(info['data_email']))  # Converte para string se não for datetime
            ws.cell(row=row_index, column=3, value=str(resultado[2]))
            ws.cell(row=row_index, column=4, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))
            #ws.cell(row=row_index, column=5, value=str(info.get('nota_fiscal', '')))
            ws.cell(row=row_index, column=5, value=str(resultado[0]))
            ws.cell(row=row_index, column=6, value=str(info['nfe'] if info['nfe'] != '0' else '0'))
            #valor = ws.cell(row=row_index, column=5).value
            #if valor == '0':
            ws.cell(row=row_index, column=7, value=data_ajustada)
            ws.cell(row=row_index, column=8, value=str(resultado[1]))
            ws.cell(row=row_index, column=9, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            ws.cell(row=row_index, column=10, value=str(resultado[4]).replace('.', ','))  # PESO
            ws.cell(row=row_index, column=11, value=str(resultado[7]))  # LOTE
            ws.cell(row=row_index, column=12, value=str(resultado[9]))  # PRODUTO
            ws.cell(row=row_index, column=13, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info.get('email_vinculado', ''))
            trocar_cores(ws, row_index)
            row_index += 1
    except Exception as e:
        print(f"Erro ao salvar a planilha vindo do PDF {e}")

    """try:
        for nf_formatado, info in nf_pdf_map.items():
        #for (nf_formatado, info), resultado in zip(nf_pdf_map.items(),resultados_pdf):
            #data_descarga = str(resultado[8])
            #formatob = '%d/%m/%Y %H:%M:%S'

            #data_formatada = parse_datetime(data_descarga)
            #data_ajustada = data_formatada.strftime(formatob)
            ws.cell(row=row_index, column=1, value="PDF")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2, value=str(info['data_email']))  # Converte para string se não for datetime
            ws.cell(row=row_index, column=3, value=str(info['transportadora']))
            ws.cell(row=row_index, column=4, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))
            ws.cell(row=row_index, column=5, value=str(info.get('nota_fiscal', '')))
            ws.cell(row=row_index, column=6, value=str(info['nfe'] if info['nfe'] != '0' else '0'))
            #ws.cell(row=row_index, column=7, value=data_ajustada)
            #ws.cell(row=row_index, column=8, value=str(resultado[1]))
            ws.cell(row=row_index, column=9, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            #ws.cell(row=row_index, column=10, value=str(resultado[4]).replace('.', ','))  # PESO
            #ws.cell(row=row_index, column=11, value=str(resultado[7]))  # LOTE
            #ws.cell(row=row_index, column=12, value=str(resultado[9]))  # PRODUTO
            ws.cell(row=row_index, column=13, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info.get('email_vinculado', ''))
            trocar_cores(ws, row_index)
            row_index += 1
    except Exception as e:
        print(f"Erro ao salvar a planilha vindo do PDF {e}")"""

    try:
        for (valor, info), resultado in zip(nf_zip_map.items(), resultados_zip):
            data_descarga = str(resultado[8])
            formatob = '%d/%m/%Y %H:%M:%S'

            data_formatada = parse_datetime(data_descarga)
            data_ajustada = data_formatada.strftime(formatob)
            ws.cell(row=row_index, column=1, value="ZIP")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2, value=str(info['data_email']))
            ws.cell(row=row_index, column=3, value=str(resultado[2]))
            ws.cell(row=row_index, column=4, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))
            #ws.cell(row=row_index, column=5, value=str(info.get('nota_fiscal', 'não encontrei NF')))
            ws.cell(row=row_index, column=5, value=str(resultado[0]))
            ws.cell(row=row_index, column=6, value=str(info['nfe'] if info['nfe'] != '0' else '0'))
            #ws.cell(row=row_index, column=6, value=info['data_emissao'] if info['data_emissao'] != '0' else '0')
            ws.cell(row=row_index, column=8, value=resultado[1])
            ws.cell(row=row_index, column=7, value=data_ajustada) #DATA DESCARGA
            #ws.cell(row=row_index, column=8, value=info['chave_acesso'] if info['chave_acesso'] != '0' else str(resultado[1]))

            ws.cell(row=row_index, column=9, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            ws.cell(row=row_index, column=10, value=str(resultado[4]).replace('.',','))  # PESO
            ws.cell(row=row_index, column=11, value=str(resultado[7]))  # LOTE
            ws.cell(row=row_index, column=12, value=str(resultado[9]))  # PRODUTO
            ws.cell(row=row_index, column=13, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info.get('email_vinculado', ''))

            trocar_cores(ws, row_index)
            row_index += 1
    except Exception as e:
        print(f"Erro ao salvar a planilha vinda do ZIP {e}")

    try:
        if not nf_excel_map:
            print('nenhum arquivo')
        else:
            for info, resultado in zip(nf_excel_map, resultados_excel):
                # for info in nf_excel_map:
                data_descarga = str(resultado[8])
                formatob = '%d/%m/%Y %H:%M:%S'

                data_formatada = parse_datetime(data_descarga)
                data_ajustada = data_formatada.strftime(formatob)

                ws.cell(row=row_index, column=1, value="EXCEL")
                if isinstance(info['data_email'], datetime):
                    ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
                else:
                    ws.cell(row=row_index, column=2, value=str(info['data_email']))
                ws.cell(row=row_index, column=3, value=str(resultado[2]))
                ws.cell(row=row_index, column=4, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))
               # ws.cell(row=row_index, column=5, value=str(info.get('nota_fiscal', 'não encontrei NF')))
                ws.cell(row=row_index,column=5,value=str(resultado[0]))
                # ws.cell(row=row_index, column=4, value=info.get('nota_fiscal', ''))
                ws.cell(row=row_index, column=6, value=str(info['nfe'] if info['nfe'] != '0' else '0'))
                ws.cell(row=row_index, column=7, value=data_ajustada)  # DATA DESCARGA
                # ws.cell(row=row_index, column=6, value=info['data_emissao'] if info['data_emissao'] != '0' else '0')
                # ws.cell(row=row_index, column=7, value=info.get('chave_acesso', '0'))
                ws.cell(row=row_index, column=8, value=resultado[1])  # CHAVE DE ACESSO
                ws.cell(row=row_index, column=9, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
                ws.cell(row=row_index, column=10, value=str(resultado[4]).replace('.', ','))  # PESO
                ws.cell(row=row_index, column=11, value=str(resultado[7]))  # LOTE
                ws.cell(row=row_index, column=12, value=str(resultado[9]))  # PRODUTO
                ws.cell(row=row_index, column=13, value=info['cnpj'] if info['cnpj'] != '0' else '0')
                ws.cell(row=row_index, column=14, value=info.get('email_vinculado', ''))
                trocar_cores(ws, row_index)
                row_index += 1

    except Exception as e:
        print(f"Erro ao salvar a planilha vinda do EXCEL {e}")

    wb.save(file_name)
    print(f"Planilha '{file_name}' salva com sucesso.")

def conecta_sql(nf_zip_map, nf_pdf_map, nf_excel_map):#, file_name):

    load_dotenv()
    server = os.getenv('DB_SERVER')
    database = os.getenv('DB_DATABASE')
    user = os.getenv('DB_USER')
    senha = os.getenv('DB_PASSWORD')


    consulta_nfe_zip = ', '.join([f"'{info.get('nfe', '0')}'" for valor, info in nf_zip_map.items()])
    #consulta_nfe_pdf = ', '.join([f"'{info.get('nfe', '0')}'" for valor, info in nf_pdf_map.items()])
    consulta_nfe_excel=', '.join([f"'{info.get('nfe', '0')}'" for info in nf_excel_map])

    chaves_consulta_zip = ', '.join([f"'{info.get('nota_fiscal', '0')}'" for valor, info in nf_zip_map.items()])
    chaves_consulta_pdf = [f"'{info.get('nota_fiscal', '0')}'" for valor, info in nf_pdf_map.items()]
    chaves_consulta_excel = ', '.join([f"'{info.get('nota_fiscal', '0')}'" for info in nf_excel_map])

    cnpj_consulta_zip = ', '.join([f"'{info.get('cnpj', '0')}'" for valor, info in nf_zip_map.items()])
    cnpj_consulta_pdf = ', '.join([f"'{info.get('cnpj', '0')}'" for valor, info in nf_pdf_map.items()])
    cnpj_consulta_excel = ', '.join([f"'{info.get('cnpj', '0')}'" for info in nf_excel_map])


    conn_str = f'DRIVER={{SQL SERVER}};SERVER={server};DATABASE={database};UID={user};' \
               f'PWD={senha}'

    try:
        conn = pyodbc.connect(conn_str)
        print('Conectou')
        cursor_pdf = conn.cursor()
        cursor_zip = conn.cursor()
        cursor_excel = conn.cursor()


        #consulta_pdf = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
        #               f" from eis_v_mapa_estoque_python where (Nota in ({chaves_consulta_pdf}) and cnpj in ({cnpj_consulta_pdf}))"
        consulta_pdf = f"""
        WITH ChavesNumeradas AS (
            SELECT
                Nota,
                ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Posicao
            FROM (VALUES {', '.join([f"({nota})" for nota in chaves_consulta_pdf])}) AS V(Nota)
        ),
            UltNota AS (SELECT
                m.Nota,
                m.Chave_nfe,
                m.RazaoSocial,
                m.cnpj,
                m.diferenca,
                m.Dt_Emissao,
                m.codigo,
                m.lote,
                m.dtMovimento,
                m.Produto,
                o.Posicao,
                ROW_NUMBER() OVER (PARTITION BY m.Nota ORDER BY m.Dt_Emissao DESC) AS rn
            FROM
                eis_v_mapa_estoque_python m
            JOIN
                ChavesNumeradas o ON m.Nota = o.Nota
            WHERE
                m.cnpj IN ({cnpj_consulta_pdf})
            )
        SELECT
            Nota,
            Chave_nfe,
            RazaoSocial,
            cnpj,
            diferenca,
            Dt_Emissao,
            codigo,
            lote,
            dtMovimento,
            Produto,
            Posicao
        From
            UltNota
        where
            rn = 1
        ORDER BY
            Posicao;
        """
        #print(f'consulta pdf : {consulta_pdf}')
        consulta_zip = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
                       f" from eis_v_mapa_estoque_python where (Nota in ({chaves_consulta_zip}) and cnpj in ({cnpj_consulta_zip}) and FlagCompl = 'N') OR" \
                       f" (Nota in ({consulta_nfe_zip}) and cnpj in ({cnpj_consulta_zip}))"
        consulta_excel = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
                       f" from eis_v_mapa_estoque_python where (Nota in ({chaves_consulta_excel}) and cnpj in ({cnpj_consulta_excel}) and FlagCompl = 'N') OR" \
                       f" (Nota in ({consulta_nfe_excel}) and cnpj in ({cnpj_consulta_excel}))"
        # consulta_sql = f"Select diferenca from eis_v_mapa_estoque_python where Chave_Nfe on ({chaves_consulta})"
        resultado_pdf, resultado_zip, resultado_excel = [], [], []

        try:
            cursor_pdf.execute(consulta_pdf)

            resultado_pdf = cursor_pdf.fetchall()
        except Exception as e:
            print(f"Erro ao executar a consulta PDF: {e}")
        try:
            cursor_zip.execute(consulta_zip)
            resultado_zip = cursor_zip.fetchall()
        except Exception as e:
            print(f"Erro ao executar a consulta ZIP: {e}")
        try:
            cursor_excel.execute(consulta_excel)
            resultado_excel = cursor_excel.fetchall()
        except Exception as e:
            print(f"Erro ao executar a consulta EXCEL: {e}")

        cursor_pdf.close()
        cursor_zip.close()
        cursor_excel.close()
        conn.close()

        if not resultado_pdf:
            print("Nenhum dado encontrado para PDF.")
        if not resultado_zip:
            print("Nenhum dado encontrado para ZIP.")
        if not resultado_excel:
            print("Nenhum dado encontrado para EXCEL.")

        return resultado_pdf, resultado_zip, resultado_excel

    except Exception as e:
        print(f"Erro ao preparar ou fechar a conexão: {e}")
        return None, None, None
        #return resultados

