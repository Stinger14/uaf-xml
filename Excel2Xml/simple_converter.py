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

# RTEMAP = get_resources_path("data/mapped_elements.xlsx")

k = []
ele = []


def gen_xml(workbook):
    """
    :param name: workbook - Archivo Excel a convertir.
    :param type: str
    :return: void
    """
    # Load our Excel File
    wb = load_workbook(workbook)
    # Getting an object of active sheet 1
    ws = wb.worksheets[0]
    # Returning returns a triplet
    doc, tag, text = Doc().tagtext()

    xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
    # xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema-instance"></xs:schema>'

    # ? Se agrega el header al documento.
    doc.asis(xml_header)

    # ? Se crean los elementos o nodos.
    # ? El nodo report(root) se debe generar con los atributos especificados para
    # ? que el reporte sigan las reglas definidas en el esquema de la UAF.
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
                text("1")  # str
            with tag('submission_code'):
                if row[36] == "EFECTIVO" or row[36] == "CHEQUE":
                    text("M")
                else:
                    text("M")
            with tag('report_code'):
                text("CTR")
            with tag('entity_reference'):
                text("Bagricola")
            with tag('fiu_ref_number'):
                text("UAF Santo Domingo")
            with tag('submission_date'):
                tmp = datetime.strptime(str(row[4]), "%Y-%m-%d %H:%M:%S")
                text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
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
                    text("001-0955138-2")
                with tag("nationality1"):
                    text("DO")
                with tag("phones"):
                    with tag("phone"):
                        with tag("tph_contact_type"):
                            text(1)
                        with tag("tph_communication_type"):
                            text("L")
                        with tag("tph_country_prefix"):
                            text("809")
                        with tag("tph_number"):
                            text("535-8088")
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
                    text(int("1"))
                with tag("address"):
                    text("Calle Ave. George Washington No. 601")
                with tag("city"):
                    text("Santo Domingo")
                with tag("country_code"):
                    text("DO")

            with tag("reason"):
                if row[105] is None:
                    row[105] = "Transaccion sospechosa"
                else:
                    text(row[105])
            with tag("action"):
                if row[80] is None:
                    row[80] = "Acciones a tomar"
                else:
                    text(row[80])

        # ? Una vez que el esqueleto XML está hecho se recorren todas las filas
        # ? y se añade un elemento transacción por file del RTE
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
            row = [cell.value for cell in row if row is not None]

            with tag("transaction"):
                with tag("transactionnumber"):
                    if row[97] is None:
                        row[97] = ""
                    else:
                        text(row[97])
                with tag("internal_ref_number"):
                    if row[19] is None:
                        row[19] = ""
                    else:
                        text(row[19])
                with tag("transaction_location"):
                    if row[3] is None:
                        row[3] = ""
                    else:
                        text(row[3])
                with tag("transaction_description"):
                    if row[31] is None:
                        row[31] = ""
                    else:
                        text(row[31])
                with tag("date_transaction"):
                    if row[37] is None:
                        row[37] = datetime.strptime("1900-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
                        text(str(row[37]))
                    else:
                        tmp = datetime.strptime(str(row[37]), "%Y-%m-%d %H:%M:%S")
                        text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))

                with tag("transmode_code"):
                    text("A")

                with tag("transmode_comment"):
                    if row[105] is None:
                        row[105] = ""
                    else:
                        text(row[105])
                with tag("amount_local"):
                    if row[34] is None:
                        row[34] = 0.0
                    else:
                        text(row[34])
                with tag("involved_parties"):
                    with tag("party"):
                        with tag("role"):
                            text("B")
                        with tag("person_my_client"):
                            with tag("first_name"):
                                text(row[10])
                            with tag("middle_name"):
                                text(row[10])
                            with tag("last_name"):
                                if row[11] is None or row[11] == "":
                                    row[11] = "n/a"
                                else:
                                    text(row[11])
                            with tag("birthdate"):
                                tmp = datetime.strptime("1900-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
                                text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                            with tag("ssn"):
                                if row[15] is None:
                                    row[15] = "n/a"
                                else:
                                    text(row[15])
                            with tag("id_number"):
                                if row[15] is None:
                                    row[15] = "n/a"
                                else:
                                    text(row[15])

                            with tag("phones"):
                                with tag("phone"):
                                    with tag("tph_contact_type"):
                                        text(2)
                                    with tag("tph_communication_type"):
                                        text("L")
                                    with tag("tph_number"):
                                        text("809-222-2222")
                                        # if row[28] is None or row[28] == "":
                                        #     row[28] = "8098881722"
                                        # else:
                                        #     text(row[28])

                            with tag("addresses"):
                                with tag("address"):
                                    with tag("address_type"):
                                        text(int("1"))
                                    with tag("address"):
                                        if row[25] is None or row[25] == "":
                                            row[25] = "prueba"
                                        else:
                                            text(row[25])
                                    with tag("city"):
                                        if row[22] is None:
                                            row[22] = "prueba"
                                        else:
                                            text(row[22])
                                    with tag("country_code"):
                                        text("DO")

                            with tag("email"):
                                text("prueba@prueba.com")

                            with tag("employer_address_id"):
                                with tag("address_type"):
                                    text(int("2"))
                                with tag("address"):
                                    text("Calle prueba no. 2")
                                with tag("city"):
                                    text(row[22])
                                with tag("country_code"):
                                    text("DO")

                            with tag("employer_phone_id"):
                                with tag("tph_contact_type"):
                                    text(1)
                                with tag("tph_communication_type"):
                                    text("M")
                                with tag("tph_number"):
                                    text("829-000-0000")

                            with tag("identification"):
                                with tag("type"):
                                    text(int("1"))
                                with tag("number"):
                                    if row[55] is None:
                                        row[55] = "n/a"
                                    else:
                                        text(row[55])
                                with tag("issue_date"):
                                    tmp = datetime.strptime("2020-03-03 00:00:00", "%Y-%m-%d %H:%M:%S")
                                    text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                                with tag("expiry_date"):
                                    tmp = datetime.strptime("2020-03-03 00:00:00", "%Y-%m-%d %H:%M:%S")
                                    text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                                with tag("issue_country"):
                                    text("DO")
                        with tag("funds_code"):
                            text("K")
                        with tag("funds_comment"):
                            text(row[31])
                        with tag("country"):
                            text("DO")
                        with tag("significance"):
                            text(int("6"))

                with tag("comments"):
                    text("Prueba")

    # ? doc.getvalue() contiene el archivo XML en formato string
    result = indent(
        doc.getvalue(),
        indentation='   ',
        indent_text=False
    )

    # ? Se agrega la fecha del día en que se genera el XML
    date = ''.join(str(datetime.now()).replace('-', '_')[:10])
    filename = "_UAF_Web_Report_" + date + ".xml"

    with open(filename, "w") as f:
        f.write(result)


# ? Esta variable guarda el archivo seleccionado.
fname = sys.argv[1] if len(sys.argv) > 1 else sg.popup_get_file('Seleccionar archivo')

if not fname:
    sg.popup("Cancel", "No se seleccionó ningún archivo")
    raise SystemExit("Cancelando: no filename supplied")
else:
    gen_xml(fname)
    sg.popup('Archivo XML generado exitosamente', fname)
