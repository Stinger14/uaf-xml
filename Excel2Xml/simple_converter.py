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

# RTEMAP = get_resources_path("data/mapped_elements.xlsx")

k = []
ele = []


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

    xml_header = '<?xml version="1.0" encoding="iso-8859-1"?>'
    # xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema-instance"></xs:schema>'

    # ? Se agrega el header al documento.
    doc.asis(xml_header)

    # ? Se crean los elementos o nodos.
    # ? El nodo report(root) se debe generar con los atributos especificados para
    # ? que el reporte siga las reglas definidas en el esquema de la UAF.
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
                    text("E")
            with tag('report_code'):
                text("CTR")
            # with tag('entity_reference'):
            #     text("Bagricola")
            # with tag('fiu_ref_number'):
            #     text("UAF Santo Domingo")
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
            # with tag("action"):
            #     if row[80] is None:
            #         row[80] = "Acciones a tomar"
            #     else:
            #         text(row[80])

        # ? Una vez que el esqueleto XML está hecho se recorren todas las filas
        # ? y se añade un nodo transacción por fila del RTE
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=ws.max_row, min_col=1, max_col=ws.max_column)):
            row = [cell.value for cell in row if row is not None]
            # ? Si tiene número de reporte entonces la fila es válida
            if row[0]:
                with tag("transaction"):
                    with tag("transactionnumber"):  # ! CHECK col in excel file
                        text(row[77])
                    # with tag("internal_ref_number"):
                    #     if row[19] is None or row[19] == '':
                    #         row[19] = "test1234"
                    #     else:
                    #         text(row[19])
                    with tag("transaction_location"):       # FIXME: retornar nombre de sucursal, no el id
                        if row[3] is None:
                            row[3] = "n/a"
                            text(row[3])
                        else:
                            text(row[3])
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
                        text("SANTO DOMINGO")
                    with tag("authorized"):
                        text("SANTO DOMINGO")
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
                                if row[9] is None or row[9] == '':
                                    row[9] = "M"
                                    text(row[9])
                                else:
                                    text(row[9])
                            with tag("title"):
                                if row[9] is None or row[9] == '':
                                    if row[9] == "M":
                                        text("Sr.")
                                    elif row[9] == "F":
                                        text("Sra.")
                                else:
                                    text(row[9])
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
                                text(row[15])
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
                                            text("M")
                                        else:
                                            text("L")
                                    with tag("tph_country_prefix"):
                                        if row[27]:
                                            text(row[27][:3])
                                        elif row[28]:
                                            text(row[28][:3])
                                        elif row[26]:
                                            text(row[26][:3])
                                    with tag("tph_number"):
                                        if row[27]:
                                            text(row[27][4:12])     #! BUG
                                        elif row[26]:
                                            text(row[26][4:12])
                                        elif row[28]:
                                            text(row[28][4:12])
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
                                        text(row[15])
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
                                text(row[3])
                            with tag("account"):
                                text(row[19])
                            with tag("currency_code"):
                                text("DOP")
                            with tag("account_name"):
                                text(row[10] + ' ' + row[11])
                            with tag("client_number"):      # FIXME: Fetch client acc no.
                                text("5596595")
                            with tag("personal_account_type"):
                                text("C")
                            with tag("signatory"):
                                with tag("t_person"):
                                    with tag("gender"):
                                        text(row[9])
                                    with tag("title"):
                                        if row[9] is None or row[9] == '':
                                            if row[9] == "M":
                                                text("Sr.")
                                            elif row[9] == "F":
                                                text("Sra.")
                                        else:
                                            text(row[9])
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
                                        text(row[15])
                                    with tag("nationality1"):
                                        text("DO")
                                    with tag("residence"):
                                        text("DO")
                                    with tag("phones"):
                                        with tag("phone"):
                                            with tag("tph_contact_type"):
                                                text(1)
                                            with tag("tph_communication_type"):
                                                if row[27] is None or row[27] == '':
                                                    text("M")
                                                else:
                                                    text("L")
                                            with tag("tph_country_prefix"):
                                                if row[27]:
                                                    text(row[27][:3])
                                                elif row[28]:
                                                    text(row[28][:3])
                                                elif row[26]:
                                                    text(row[26][:3])
                                            with tag("tph_number"):
                                                if row[27]:
                                                    text(row[27][4:12])
                                                elif row[26]:
                                                    text(row[26][4:12])
                                                elif row[28]:
                                                    text(row[28][4:12])
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
                                                text(row[15])
                                        with tag("issue_country"):
                                            text("DO")
                                with tag("role"):
                                    text("A")
                            with tag("opened"):         # FIXME: Fecha apertura cta
                                if row[37] is None or row[37] == '':
                                    if row[37] is None:
                                        row[37] = datetime.strptime("2000-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
                                    tmp = datetime.strptime(str(row[37]), "%Y-%m-%d %H:%M:%S")
                                    text(tmp.strftime("%Y-%m-%dT%H:%M:%S"))
                            with tag("balance"):        # FIXME: Balance luego de realizar trx
                                text("1000000")
                            with tag("status_code"):    # FIXME: Fetch estado cta al inicio de trx
                                text("A")
                            with tag("beneficiary"):    # FIXME: Beneficiaro final
                                text(row[10] + ' ' + row[11])
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


# ? Esta variable guarda el archivo seleccionado.
fname = sys.argv[1] if len(sys.argv) > 1 else sg.popup_get_file('Seleccionar archivo')

if not fname:
    sg.popup("Cancel", "No se seleccionó ningún archivo")
    raise SystemExit("Cancelando: no filename supplied")
else:
    gen_xml(fname)
    sg.popup('Archivo XML generado exitosamente', fname)
