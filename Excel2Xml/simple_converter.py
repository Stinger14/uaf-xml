"""
Traductor de Excel a XML
========================

Este programa traduce un archivo Excel generado por el módulo RTE a XML
requerido por la Unidad de Análisis Financiera(UAF).
"""

import PySimpleGUI as sg
import sys
from openpyxl import load_workbook
from resources import get_resources_path
from yattag import Doc, indent
from datetime import datetime
from xml.etree.ElementTree import parse, ParseError
import timeit
import time

SUCURSALES = {
    "0": 'PRINCIPAL',
    "1": 'SANTO DOMINGO',
    "2": 'HIGUEY',
    "3": 'SAN CRISTOBAL',
    "4": 'BARAHONA',
    "5": 'SAN JUAN DE LA MAGUANA',
    "6": 'SAN FRANCISCO MACORIS',
    "7": 'COMENDADOR',
    "8": 'COTUI',
    "9": 'LA VEGA',
    "10": 'SANTIAGO RODRIGUEZ',
    "11": 'MONTE CRISTI',
    "12": 'PUERTO PLATA',
    "13": 'NAGUA',
    "15": 'EL SEYBO',
    "16": 'SANTIAGO',
    "17": 'SAN JOSE DE OCOA',
    "18": 'AZUA',
    "19": 'BANI',
    "20": 'VALVERDE MAO',
    "21": 'ARENOSO',
    "22": 'HATO MAYOR',
    "23": 'MOCA',
    "24": 'SAMANA',
    "25": 'BONAO',
    "26": 'NEYBA',
    "27": 'DAJABON',
    "28": 'SAN JOSE DE LAS MATAS',
    "29": 'RIO SAN JUAN',
    "30": 'VILLA RIVA',
    "31": 'SALCEDO',
    "32": 'MONTE PLATA',
    "33": 'CONSTANZA',
}

DIRECCIONES = {
    "PRINCIPAL": "Ave.George Washington NO. 601",
    "SANTO DOMINGO": "Ave.George Washington NO. 601",
    "SANTIAGO RODRIGUEZ": "C/ Dr. Darío Gómez No. 64",
    "MONTE CRISTI": "C/ BENITO MONCION, ESQ. SANCHEZ NO.60",
    "PUERTO PLATA": "C/ PRINCIPAL PROF. JUAN BOSCH NO.4",
    "CONSTANZA": "AV. ANTONIO ABUD ISSAC",
    "NAGUA": "C/ 27 DE FEBRERO, ESQ. MERCEDES BELLO NO.24",
    "EL SEYBO": "C/ MANUELA DIEZ JIMENEZ No.10",
    "SANTIAGO": "AV. JUAN PABLO DUARTE, ESQ. ESTADO DE ISRAEL",
    "SAN JOSE DE OCOA": "C/ DUARTE ESQ. ALTAGRACIA No.40",
    "AZUA": "C/ EMILIO PRUDHOMME No.35",
    "BANI": "C/ NUESTRA SRA. DE REGLA No.19 ESQ. SANCHEZ",
    "HIGUEY": "C/ ALTAGRACIA ESQ. LAS CARRERAS",
    "VALVERDE MAO": "AV. BENITO MONCION",
    "ARENOSO": "AV. DUARTE No.61",
    "HATO MAYOR": "C/ MELCHOR CONTIN ALPHAU No.39, MILLON",
    "MOCA": "CARR. MOCA LA VEGA, AL LADO DEL HOSPITAL",
    "SAMANA": "C/ SANTA BARBARA, EDIF. OFICINAS PUBLICAS",
    "BONAO": "C/ DUARTE No.279, SECTOR EL 90",
    "NEYBA": "C/ SAN BARTOLOME No.41",
    "DAJABON": "AV. PABLO REYES NO.29",
    "SAN JOSE DE LAS MATAS": "C/ 16 DE AGOSTO No.18, CENTRO DEL PUEBLO",
    "RIO SAN JUAN": "AUTOPISTA DR. DOMINGO ANTIONIO GONZALEZ",
    "SAN CRISTOBAL": "C/ GENERAL CABRAL No. 34",
    "VILLA RIVA": "C/ 27 DE FEBRERO ESQ. COLON NO.57",
    "MONTE PLATA": "C/ GENERAL MATIAS MORENO ESQ. ALTAGRACIA No. 7",
    "CONSTANZA": "AV. ANTONIO ABUD ISSAC",
    "BARAHONA": "AV. LUIS E. DEL MONTE No.59, CENTRO DE LA CIU",
    "SAN JUAN DE LA MAGUANA": "C/ SANCHEZ No.67",
    "SAN FRANCISCO MACORIS": "Av. Frank Grullón, Salida hacia Nagua",
    "SALCEDO": "C/ DUARTE, ESQ. HERMANAS MIRABAL",
    "COMENDADOR": "C/ 27 DE FEB. ESQ. LUZ CELESTE LARA No.20",
    "COTUI": "C/ PADRE FANTINO NO.5",
    "LA VEGA": "PROFESOR JUAN BOSCH, ESQ. COMANDANTE JIMÉNEZ",
}


# RTEMAP = get_resources_path("data/mapped_elements.xlsx")


class Converter:
    """This class converts an RTE module excel file to a valid
        XML to meet goAML platform requirements.
    """
    # time program execution
    start = time.time()

    def __init__(self):
        self.k = []
        self.ele = []
        self.wb = None

    def run(self):
        self.wb = sys.argv[1] if len(sys.argv) > 1 else sg.popup_get_file('Seleccionar archivo')
        if not self.wb:
            sg.popup("Cancelando", "No se seleccionó ningún archivo.")
            raise SystemExit("Cancelado.")
        else:
            gen_xml(self.wb)
            sg.popup("Archivo XML generado exitosamente.", self.wb)

    time.sleep(1)
    stop = time.time()
    print(f"Program execution time: {stop - start}")


def gen_xml(workbook):
    """
    :param workbook: str - Archivo Excel a convertir.
    :return: void
    """
    # Load our Excel File
    wb = load_workbook(workbook)
    # Getting an object of active sheet 1
    ws = wb.worksheets[0]
    # Returning returns a triplet
    doc, tag, text = Doc().tagtext()

    xml_header = '<?xml version="1.0" encoding="iso-8859-3"?>'
    # xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema-instance"></xs:schema>'

    # ? Agrega el header al documento.
    doc.asis(xml_header)

    # ? Se crean los elementos o nodos.
    # ? El nodo report(root) se debe generar con el esquema especificado para
    # ? que el reporte siga las reglas definidas en la plataforma goAML.
    with tag('report', ('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance"),
             ('xsi:noNamespaceSchemaLocation', "goAMLSchema.xsd")):
        # ? Se recorre solo la primera fila para crear el esqueleto
        # ? Algunos datos están hardcode ya que el RTE no genera esta data o
        # ? estos datos serán siempre fijos.
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=8, min_col=1, max_col=ws.max_column)):
            row = [cell.value for cell in row if row is not None]
            # ? La indentación define la relación padre-hijo entre los nodos.
            with tag('rentity_id'):
                text(int('1059'))  # int
            with tag('rentity_branch'):
                text("Principal")  # str
            with tag('submission_code'):
                if row[36] == "EFECTIVO" or row[36] == "CHEQUE":
                    text("E")
                else:
                    text("E")  # Si existe otro método cambiar letra según documento.
            with tag('report_code'):
                text("CTR")
            with tag('submission_date'):
                tmp = datetime.strptime(str(row[4]), "%Y-%m-%d %H:%M:%S")
                text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))  # Formato válido "%Y-%m-%dT%H:%M:%S"
            with tag("currency_code_local"):
                text("DOP")
            with tag("reporting_person"):
                with tag("gender"):
                    text("F")
                with tag("title"):
                    text("Oficial de Cumplimiento")
                with tag("first_name"):
                    text("Naty")
                with tag("last_name"):
                    text("Abreu")
                with tag("birthdate"):
                    tmp = datetime.strptime("1972-10-03 00:00:00", "%Y-%m-%d %H:%M:%S")
                    text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                with tag("id_number"):
                    text("00109551382")
                with tag("nationality1"):
                    text("DO")
                with tag("phones"):
                    with tag("phone"):
                        with tag("tph_contact_type"):
                            text(2)
                        with tag("tph_communication_type"):
                            text("L")
                        with tag("tph_country_prefix"):
                            text("809")
                        with tag("tph_number"):
                            text("5358088")
                        with tag("tph_extension"):
                            text("3212")
                with tag("addresses"):
                    with tag("address"):
                        with tag("address_type"):
                            text(int("2"))
                        with tag("address"):
                            text("Ave. George Washington NO. 601")
                        with tag("town"):
                            text("Santo Domingo")
                        with tag("city"):
                            text("Santo Domingo")
                        with tag("zip"):
                            text("10103")
                        with tag("country_code"):
                            text("DO")
                with tag("email"):
                    text("N.abreu@bagricola.gob.do")
                with tag("occupation"):
                    text("Ing. de Sistemas")

            with tag("location"):
                with tag("address_type"):
                    text(int("2"))
                with tag("address"):
                    text("Calle Ave. George Washington No. 601")
                with tag("city"):
                    text("SANTO DOMINGO,D.N.")
                with tag("country_code"):
                    text("DO")
                with tag("state"):
                    text("SANTO DOMINGO")

            with tag("reason"):
                if row[74] is None:
                    row[74] = "Transacciones que sobrepasaron los USD$15,000.00"
                else:
                    text(row[74])

        # -------------------------------- TRANSACCION ---------------------------------#

        # ? Una vez que el esqueleto XML está hecho se recorren todas las filas
        # ? y se añade un nodo transacción por fila del RTE(Archivo Excel)
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
            row = [cell.value for cell in row if row is not None]
            # ? Si tiene número de reporte entonces la fila es válida
            if row[0]:
                with tag("transaction"):
                    with tag("transactionnumber"):
                        text(row[77])
                    with tag("transaction_location"):
                        if row[79] is None or row[79] == '':
                            suc = SUCURSALES.get(str(row[3])).upper()
                            text(DIRECCIONES.get(suc))
                        else:
                            text(row[79])
                    with tag("transaction_description"):
                        if row[41] is None:
                            row[41] = ""
                            text(row[41])
                        else:
                            text(row[41])
                    with tag("date_transaction"):
                        if row[37] is None:
                            row[37] = datetime.strptime("1900-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
                        tmp = datetime.strptime(str(row[37]), "%Y-%m-%d %H:%M:%S")
                        text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                    with tag("teller"):
                        text(SUCURSALES.get(str(row[3])).upper())
                    with tag("authorized"):
                        text(SUCURSALES.get(str(row[3])).upper())
                    with tag("transmode_code"):
                        text("6")
                    with tag("transmode_comment"):
                        if row[31] is None:
                            row[31] = ""
                            text(row[31])
                        else:
                            text(row[31])
                    with tag("amount_local"):
                        if row[34] is None or row[34] == '':
                            row[34] = 0.0
                            text(row[34])
                        else:
                            text(row[34])
                    with tag("t_from_my_client"):
                        with tag("from_funds_code"):
                            text("K")

                        with tag("from_person"):
                            with tag("gender"):
                                if row[48] == 'SI':
                                    pass
                                elif row[9] is None or row[9] == '':
                                    row[9] = "M"
                                    text(row[9])
                                else:
                                    text(row[9])
                            with tag("title"):
                                if row[9] == "M":
                                    text("Sr.")
                                elif row[9] == "F":
                                    text("Sra.")
                                else:
                                    text("n/a")
                            with tag("first_name"):
                                text(row[10])
                            with tag("middle_name"):
                                if len(row[10].split(' ')) > 1:
                                    text(row[10].split(' ')[1])
                                else:
                                    text(row[10])
                            with tag("last_name"):
                                if row[11] is None or row[11] == '':
                                    row[11] = "."
                                    text(row[11])
                                else:
                                    text(row[11])
                            with tag("birthdate"):
                                text("1988-06-17T00:00:00")
                            with tag("id_number"):
                                text(''.join(row[15].split('-')).strip())
                            with tag("nationality1"):
                                text("DO")
                            with tag("residence"):
                                text("DO")
                            with tag("phones"):
                                with tag("phone"):
                                    with tag("tph_contact_type"):
                                        text(1)
                                    with tag("tph_communication_type"):
                                        if row[28] is None or row[28] == '':
                                            text("L")
                                        else:
                                            text("M")
                                    with tag("tph_country_prefix"):
                                        if row[27] is None and row[28] is None and row[26] is None:
                                            text("n/a")
                                        elif row[27]:
                                            text(''.join(row[27][:3].split('-')))
                                        elif row[28]:
                                            text(''.join(row[28][:3].split('-')))
                                        elif row[26]:
                                            text(''.join(row[26][:3].split('-')))
                                    with tag("tph_number"):
                                        if row[27] is None and row[28] is None and row[26] is None:
                                            text("n/a")
                                        elif row[27]:
                                            text(''.join(row[27][4:12].split('-')))
                                        elif row[26]:
                                            text(''.join(row[26][4:12].split('-')))
                                        elif row[28]:
                                            text(''.join(row[28][4:12].split('-')))
                            with tag("addresses"):
                                with tag("address"):
                                    with tag("address_type"):
                                        text(int("1"))
                                    with tag("address"):
                                        if row[25] is None or row[25] == '':
                                            text("n/a")
                                        else:
                                            text(row[25])
                                    with tag("city"):
                                        if row[23] is None or row[23] == '':
                                            text("n/a")
                                        else:
                                            text(row[23])
                                    with tag("country_code"):
                                        text("DO")
                                    with tag("state"):
                                        if row[22] is None or row[22] == '':
                                            text("n/a")
                                        else:
                                            text(row[22])
                            with tag("occupation"):
                                if row[17] is None or row[17] == '':
                                    text("n/a")
                                else:
                                    text(row[17])
                            with tag("identification"):
                                with tag("type"):
                                    text("1")
                                with tag("number"):
                                    if row[15] is None or row[15] == '':
                                        text("n/a")
                                    else:
                                        text(''.join(row[15].split('-')).strip())
                                with tag("issue_country"):
                                    text("DO")
                        with tag("from_country"):
                            text("DO")
                    with tag("t_to_my_client"):
                        with tag("to_funds_code"):
                            text("K")
                        with tag("to_account"):
                            with tag("institution_name"):
                                text("Banco Agrícola de la República Dominicana")
                            with tag("swift"):
                                text("401007665")
                            with tag("branch"):
                                text(SUCURSALES.get(str(row[3])))
                            with tag("account"):
                                text(row[19])
                            with tag("currency_code"):
                                text("DOP")
                            with tag("account_name"):
                                text(row[10] + ' ' + row[11])
                            with tag("client_number"):
                                text(row[78])
                            with tag("personal_account_type"):
                                text("C")
                            with tag("signatory"):
                                with tag("t_person"):
                                    with tag("gender"):
                                        text(row[9])
                                    with tag("title"):
                                        if row[9] == "M":
                                            text("Sr.")
                                        elif row[9] == "F":
                                            text("Sra.")
                                        else:
                                            text(row[9])
                                    with tag("first_name"):
                                        text(row[10])           # FIXME: row[50]
                                    with tag("middle_name"):
                                        if len(row[10].split(' ')) > 1:         # FIXME: row[50]
                                            text(row[10].split(' ')[1])
                                        else:
                                            text(row[10])
                                    with tag("last_name"):
                                        if row[11] is None or row[11] == '':        # FIXME: row[51]
                                            row[11] = "."
                                            text(row[11])
                                        else:
                                            text(row[11])
                                    with tag("birthdate"):
                                        text("1988-06-17T00:00:00")
                                    with tag("id_number"):
                                        text(''.join(row[15].split('-')).strip())
                                    with tag("nationality1"):
                                        text("DO")
                                    with tag("residence"):
                                        text("DO")
                                    with tag("phones"):
                                        with tag("phone"):
                                            with tag("tph_contact_type"):
                                                text(1)
                                            with tag("tph_communication_type"):
                                                if row[28] is None or row[28] == '':
                                                    text("L")
                                                else:
                                                    text("M")
                                            with tag("tph_country_prefix"):
                                                if row[27] is None and row[28] is None and row[26] is None:
                                                    text("n/a")
                                                elif row[27] == '' and row[28] == '' and row[26] == '':
                                                    text("n/a")
                                                elif row[27]:
                                                    text(''.join(row[27][:3].split('-')))
                                                elif row[28]:
                                                    text(''.join(row[28][:3].split('-')))
                                                elif row[26]:
                                                    text(''.join(row[26][:3].split('-')))
                                            with tag("tph_number"):
                                                if row[27] is None and row[28] is None and row[26] is None:
                                                    text("n/a")
                                                elif row[27] == '' and row[28] == '' and row[26] == '':
                                                    text("n/a")
                                                elif row[27]:
                                                    text(''.join(row[27][4:12].split('-')))
                                                elif row[28]:
                                                    text(''.join(row[28][4:12].split('-')))
                                                elif row[26]:
                                                    text(''.join(row[26][4:12].split('-')))
                                    with tag("addresses"):
                                        with tag("address"):
                                            with tag("address_type"):
                                                text(int("1"))
                                            with tag("address"):
                                                if row[79] is None or row[79] == '':
                                                    text("n/a")
                                                else:
                                                    text(row[79])
                                            with tag("city"):
                                                if row[23] is None or row[23] == '':
                                                    text("n/a")
                                                else:
                                                    text(row[23])
                                            with tag("country_code"):
                                                text("DO")
                                            with tag("state"):
                                                if row[22] is None or row[22] == '':
                                                    text("n/a")
                                                else:
                                                    text(row[22])
                                    with tag("occupation"):
                                        if row[17] is None or row[17] == '':
                                            text("n/a")
                                        else:
                                            text(row[17])
                                    with tag("identification"):
                                        with tag("type"):
                                            text("1")
                                        with tag("number"):
                                            if row[15] is None or row[15] == '':
                                                text("n/a")
                                            else:
                                                text(''.join(row[15].split('-')).strip())
                                        with tag("issue_country"):
                                            text("DO")
                                with tag("role"):
                                    text("A")
                            with tag("opened"):  # FIXME: Fecha apertura cta
                                if row[37] is None or row[37] == '':
                                    row[37] = datetime.strptime("2000-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
                                tmp = datetime.strptime(str(row[37]), "%Y-%m-%d %H:%M:%S")
                                text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                            with tag("balance"):  # FIXME: Balance luego de realizar trx
                                if row[34] is None or row[34] == '':
                                    row[34] = 0.0
                                    text(row[34])
                                else:
                                    text(row[34])
                            with tag("status_code"):  # FIXME: Fetch estado cta al inicio de trx
                                text("A")
                            with tag("beneficiary"):  # FIXME: Beneficiaro final
                                text(' '.join(row[10].split()[:5]))
                        with tag("to_country"):
                            text("DO")

    # ? doc.getvalue() contiene el archivo XML en formato string
    result = indent(
        doc.getvalue(),
        indentation='   ',
        indent_text=False
    )

    # ? Se agrega la fecha del día en que se genera el XML
    date = ''.join(str(datetime.now()).replace('-', '_')[:10])
    filename = "_UAF_Web_Report_" + date + ".xml"

    # ? Guardar contenido en un archivo.
    with open(filename, "w") as f:
        f.write(result)

    # ? Once generated, edit file to remove root node attribs.
    try:
        tree = parse(filename)
        if tree:
            tree.getroot().attrib.popitem()
            tree.write(filename)
    except ParseError as e:
        print('>>Exception<<: File not well-formed')


if __name__ == '__main__':
    app = Converter()
    app.run()
