from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
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
    for column in ['A', 'E', 'H', 'L', 'M', 'H', 'O', 'Q']:
        ws.column_dimensions[column].width = 12
    for column in ['J', 'N']:
        ws.column_dimensions[column].width = 45
    for column in ['A', 'I']:
        ws.column_dimensions[column].width = 6
    for column in ['B', 'G', 'F', 'K', 'P']:
        ws.column_dimensions[column].width = 18
    for column in ['D']:
        ws.column_dimensions[column].width = 16
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['R'].width = 106
    ws.column_dimensions['L'].width = 8
    ws.column_dimensions['K'].width = 10

    # cores
    for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N','O','P','Q','R']:
        fundo = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
        ws[f"{column}{row_index}"].fill = fundo
    for column in ['H', 'I', 'J', 'K', 'L']:
        cor_comp = PatternFill(start_color="C7F9CC", end_color="C7F9CC", fill_type="solid")
        ws[f"{column}{row_index}"].fill = cor_comp
    for column in ['D', 'E', 'F', 'G', 'K', 'M', 'N', 'O']:
        cor_nf = PatternFill(start_color="FCBC5D", end_color="FCBC5D", fill_type="solid")
        ws[f"{column}{row_index}"].fill = cor_nf


def save_to_excel(nf_pdf_map, nf_zip_map, nf_excel_map, file_name):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Notas Fiscais'
    headers = ['Tipo', 'Data Email', 'Transportadora', 'Cnpj Destinatario',
               'Cod Dest', 'Produto', 'Data Descarga', 'Nota Fiscal', 'Serie',
               'Chave de Acesso','Peso Nota', 'Lote',
               'NF Comp', 'Chave NFE Comp', 'Serie Comp', 'Data Emissao Comp', 'Peso Comp',
               'Origem']

    cor = PatternFill(start_color="f0812c", end_color="f0812c", fill_type="solid")
    fonte_cor = Font(color="FFFFFF")

    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_index, value=header)
        cell.fill = cor
        cell.font = fonte_cor

    row_index = 2
    resultados_pdf, resultados_zip, resultados_excel = conecta_sql(nf_zip_map, nf_pdf_map, nf_excel_map)
    print(f"nf pdf {len(nf_pdf_map.items())}")
    print(f"nf pdf teste{len(resultados_pdf)}")

    #print(resultados_dict)
    try:
        row_index = 2
        resultados_dict = {str(r[0]): r for r in resultados_pdf}

        for nf_formatado, info in nf_pdf_map.items():

            resultado = resultados_dict.get(nf_formatado.lstrip('0'), None)
            ws.cell(row=row_index, column=1, value="PDF")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2,
                        value=str(info['data_email']))
            ws.cell(row=row_index, column=3, value=info.get('transportadora', ''))
            ws.cell(row=row_index, column=4, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=5, value=str(resultado[6]) if resultado else 'SEM LEITURA')  # CODDEST
            ws.cell(row=row_index, column=6, value=str(resultado[9]) if resultado else 'SEM LEITURA')  # PRODUTONOTA
            ws.cell(row=row_index, column=7,
                    value=resultado[8].strftime('%d/%m/%Y %H:%M:%S') if resultado else 'SEM LEITURA')  # DATADESCARGA
            ws.cell(row=row_index, column=8, value=nf_formatado)
            ws.cell(row=row_index, column=9, value=str(resultado[11]) if resultado else 'SEM LEITURA')
            ws.cell(row=row_index, column=10, value=str(resultado[1]) if resultado else 'SEM LEITURA')  # CHAVEACESSO
            ws.cell(row=row_index, column=11,
                    value=str(resultado[4]).replace('.', ',') if resultado else 'SEM LEITURA')  # PESONOTA
            ws.cell(row=row_index, column=12, value=str(resultado[7]) if resultado else 'SEM LEITURA')  # LOTENOTA
            ws.cell(row=row_index, column=13, value=info['nfe'] if info['nfe'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            ws.cell(row=row_index, column=15, value=info['serie_nf'] if info['serie_nf'] != '0' else '0')
            if isinstance(info['data_emissao'], datetime):
                ws.cell(row=row_index, column=16, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=16,
                        value=str(info['data_emissao']))
            ws.cell(row=row_index, column=17, value=info.get('peso_comp', ''))
            ws.cell(row=row_index, column=18, value=info.get('email_vinculado', ''))
            trocar_cores(ws, row_index)
            row_index += 1

    except Exception as e:
        print(f"Erro ao salvar a planilha vindo do PDF {e}")

    try:

        resultados_dict = {str(r[0]): r for r in resultados_zip}

        for valor, info in nf_zip_map.items():
            resultado = resultados_dict.get(valor.lstrip('0'), None)

            ws.cell(row=row_index, column=1, value="ZIP")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2,
                        value=str(info['data_email']))
            ws.cell(row=row_index, column=3, value=info.get('transportadora', ''))
            ws.cell(row=row_index, column=4, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=5, value=str(resultado[6]) if resultado else 'SEM LEITURA')  # CODDEST
            ws.cell(row=row_index, column=6, value=str(resultado[9]) if resultado else 'SEM LEITURA')  # PRODUTONOTA
            ws.cell(row=row_index, column=7,
                    value=resultado[8].strftime('%d/%m/%Y %H:%M:%S') if resultado else 'SEM LEITURA')  # DATADESCARGA
            ws.cell(row=row_index, column=8, value=valor)
            ws.cell(row=row_index, column=9, value=str(resultado[11]) if resultado else 'SEM LEITURA')
            ws.cell(row=row_index, column=10, value=str(resultado[1]) if resultado else 'SEM LEITURA')  # CHAVEACESSO
            ws.cell(row=row_index, column=11,
                    value=str(resultado[4]).replace('.', ',') if resultado else 'SEM LEITURA')  # PESONOTA
            ws.cell(row=row_index, column=12, value=str(resultado[7]) if resultado else 'SEM LEITURA')  # LOTENOTA
            ws.cell(row=row_index, column=13, value=info['nfe'] if info['nfe'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            ws.cell(row=row_index, column=15, value=info['serie_nf'] if info['serie_nf'] != '0' else '0')
            if isinstance(info['data_emissao'], datetime):
                ws.cell(row=row_index, column=16, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=16,
                        value=str(info['data_emissao']))
            ws.cell(row=row_index, column=17, value=info.get('peso_comp', ''))
            ws.cell(row=row_index, column=18, value=info.get('email_vinculado', ''))
            trocar_cores(ws, row_index)
            row_index += 1
    except Exception as e:
        print(f"Erro ao salvar a planilha vinda do ZIP {e}")

    """try:
        for valor, info in nf_zip_map.items():
            ws.cell(row=row_index, column=1, value="ZIP")
            if isinstance(info['data_email'], datetime):
                ws.cell(row=row_index, column=2, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
            else:
                ws.cell(row=row_index, column=2, value=str(info['data_email']))
            #ws.cell(row=row_index, column=3, value=str(resultado[2]))
            ws.cell(row=row_index, column=4, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))
            ws.cell(row=row_index, column=5, value=str(info.get('nota_fiscal', 'não encontrei NF')))
            #ws.cell(row=row_index, column=5, value=str(resultado[0]))
            ws.cell(row=row_index, column=6, value=str(info['nfe'] if info['nfe'] != '0' else '0'))
            # ws.cell(row=row_index, column=6, value=info['data_emissao'] if info['data_emissao'] != '0' else '0')
            #ws.cell(row=row_index, column=8, value=resultado[1])
            #ws.cell(row=row_index, column=7, value=data_ajustada)  # DATA DESCARGA
            # ws.cell(row=row_index, column=8, value=info['chave_acesso'] if info['chave_acesso'] != '0' else str(resultado[1]))

            ws.cell(row=row_index, column=9, value=info['chave_comp'] if info['chave_comp'] != '0' else '0')
            #ws.cell(row=row_index, column=10, value=str(resultado[4]).replace('.', ','))  # PESO
            #ws.cell(row=row_index, column=11, value=str(resultado[7]))  # LOTE
            #ws.cell(row=row_index, column=12, value=str(resultado[9]))  # PRODUTO
            ws.cell(row=row_index, column=13, value=info['cnpj'] if info['cnpj'] != '0' else '0')
            ws.cell(row=row_index, column=14, value=info.get('email_vinculado', ''))

            trocar_cores(ws, row_index)
            row_index += 1
    except Exception as e:
        print(f"Erro ao salvar a planilha vinda do ZIP {e}")"""

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
                ws.cell(row=row_index, column=3, value=str(info.get('transportadora', '')))  # str(resultado[2]))
                ws.cell(row=row_index, column=4, value=info['cnpj'] if info['cnpj'] != '0' else '0')  # CNPJDEST
                ws.cell(row=row_index, column=5, value=str(resultado[6]))  # CODDEST
                ws.cell(row=row_index, column=6, value=str(resultado[9]))  # PRODUTONOTA
                ws.cell(row=row_index, column=7, value=data_ajustada)  # DATADESCARGA
                ws.cell(row=row_index, column=8, value=str(resultado[0]))  # NOTA FISCAL
                ws.cell(row=row_index, column=9, value=str(resultado[11]))

                ws.cell(row=row_index, column=10, value=str(resultado[1]))  # CHAVEACESSO
                ws.cell(row=row_index, column=11, value=str(resultado[4]).replace('.', ','))  # PESONOTA
                ws.cell(row=row_index, column=12, value=str(resultado[7]))  # LOTENOTA
                ws.cell(row=row_index, column=13, value=str(info['nfe'] if info['nfe'] != '0' else '0'))  # NOTACOMP
                ws.cell(row=row_index, column=14,
                        value=info['chave_comp'] if info['chave_comp'] != '0' else '0')  # CHAVECOMP
                ws.cell(row=row_index, column=15, value=str(info['serie_nf'] if info['serie_nf'] != '0' else '0'))  # SERIECOMP
                if isinstance(info['data_emissao'], datetime):
                    ws.cell(row=row_index, column=16, value=info['data_email'].strftime('%Y-%m-%d %H:%M:%S'))
                else:
                    ws.cell(row=row_index, column=16,
                            value=str(info['data_emissao']))

                ws.cell(row=row_index, column=17, value=str(info.get('peso_comp', '')))
                ws.cell(row=row_index, column=18, value=info.get('email_vinculado', ''))
                trocar_cores(ws, row_index)
                row_index += 1

    except Exception as e:
        print(f"Erro ao salvar a planilha vinda do EXCEL {e}")

    wb.save(file_name)
    print(f"Planilha '{file_name}' salva com sucesso.")
    return len(nf_pdf_map.items()), len(resultados_pdf), len(nf_zip_map.items()), len(resultados_zip),\
        len(nf_excel_map), len(resultados_excel)


def conecta_sql(nf_zip_map, nf_pdf_map, nf_excel_map):  # , file_name):

    load_dotenv()
    server = os.getenv('DB_SERVER')
    database = os.getenv('DB_DATABASE')
    user = os.getenv('DB_USER')
    senha = os.getenv('DB_PASSWORD')

    # consulta_nfe_zip = ', '.join([f"'{info.get('nfe', '0')}'" for valor, info in nf_zip_map.items()])
    # consulta_nfe_pdf = ', '.join([f"'{info.get('nfe', '0')}'" for valor, info in nf_pdf_map.items()])
    # consulta_nfe_excel=', '.join([f"'{info.get('nfe', '0')}'" for info in nf_excel_map])

    chaves_consulta_zip = [f"'{info.get('nota_fiscal', '0')}'" for valor, info in nf_zip_map.items()]
    chaves_consulta_pdf = [f"'{info.get('nota_fiscal', '0')}'" for valor, info in nf_pdf_map.items()]
    chaves_consulta_excel = [f"'{info.get('nota_fiscal', '0')}'" for info in nf_excel_map]

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
        # consulta_pdf = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
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
                m.SerieNF,
                ROW_NUMBER() OVER (PARTITION BY m.Nota ORDER BY m.Dt_Emissao DESC) AS rn
            FROM
                eis_v_mapa_estoque_python m
            JOIN
                ChavesNumeradas o ON m.Nota = o.Nota
            WHERE
                m.cnpj IN ({cnpj_consulta_pdf})
                and m.TipoNota = 'NOTA_FISCAL'

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
            Posicao,
            SerieNF
        From
            UltNota
        where
            rn = 1
        ORDER BY
            Posicao;
        """

        # print(f'consulta pdf : {consulta_pdf}')
        # consulta_zip = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
        #               f" from eis_v_mapa_estoque_python where (Nota in ({chaves_consulta_zip}) and cnpj in ({cnpj_consulta_zip}) and FlagCompl = 'N')"
        consulta_zip = f"""
                WITH ChavesNumeradas AS (
                    SELECT
                        Nota,
                        ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Posicao
                    FROM (VALUES {', '.join([f"({nota})" for nota in chaves_consulta_zip])}) AS V(Nota)
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
                        m.SerieNF,
                        ROW_NUMBER() OVER (PARTITION BY m.Nota ORDER BY m.Dt_Emissao DESC) AS rn
                    FROM
                        eis_v_mapa_estoque_python m
                    JOIN
                        ChavesNumeradas o ON m.Nota = o.Nota
                    WHERE
                        m.cnpj IN ({cnpj_consulta_zip})
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
                    Posicao,
                    SerieNF
                From
                    UltNota
                where
                    rn = 1
                ORDER BY
                    Posicao;
                """
        # consulta_excel = f"Select Nota, Chave_nfe, RazaoSocial, cnpj, diferenca, Dt_Emissao, codigo, lote, dtMovimento, Produto" \
        #               f" from eis_v_mapa_estoque_python where (Nota in ({chaves_consulta_excel}) and cnpj in ({cnpj_consulta_excel}) and FlagCompl = 'N') OR" \
        #               f" (Nota in ({consulta_nfe_excel}) and cnpj in ({cnpj_consulta_excel}))"
        consulta_excel = f"""
                        WITH ChavesNumeradas AS (
                            SELECT
                                Nota,
                                ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Posicao
                            FROM (VALUES {', '.join([f"({nota})" for nota in chaves_consulta_excel])}) AS V(Nota)
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
                                m.SerieNF,
                                ROW_NUMBER() OVER (PARTITION BY m.Nota ORDER BY m.Dt_Emissao DESC) AS rn
                            FROM
                                eis_v_mapa_estoque_python m
                            JOIN
                                ChavesNumeradas o ON m.Nota = o.Nota
                            WHERE
                                m.cnpj IN ({cnpj_consulta_excel})
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
                            Posicao,
                            SerieNF
                        From
                            UltNota
                        where
                            rn = 1
                        ORDER BY
                            Posicao;
                        """
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
        # return resultados
