import PySimpleGUI as sg
import sys
from openpyxl import load_workbook
from resources import get_resources_path
import os.path
from yattag import Doc, indent
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET

RTEMAP = get_resources_path("data/mapped_elements.xlsx")

k = []
ele = []


# TODO pass abs path to workbook file
def get_row_headers(workbook):
    # Loading the Excel file
    tmp_wb = load_workbook(workbook)
    # creating the sheet 1 object
    ws = tmp_wb.worksheets[0]
    # Iterating rows for getting the values of each row
    for row in ws.iter_rows(min_row=7, max_row=7, min_col=1, max_col=127):
        header = [cell.value for cell in row]
    return header


# TODO pass abs path to rtemap file
def get_rte_keys(keys):
    """
        Return key values from xml mapped elemnents of rte's excel file.
    """
    try:
        df = pd.read_excel(RTEMAP)
        rte_map = df[df['RTE Banke'].notna()]
        for i, j in rte_map.iterrows():
            keys.append(j.values[1])
    except FileNotFoundError as e:
        print(f'The file {e} does not exist.')


def get_xml_elements(elements):
    """
        Return key values from xml mapped elemnents of rte's excel file.
    """
    try:
        df = pd.read_excel(RTEMAP)
        elements_map = df[df['Elementos UAF'].notna()]
        for i, j in elements_map.iterrows():
            elements.append(j.values[0])
    except FileNotFoundError as e:
        print(f'The file {e} does not exist.')


def gen_keymap(elements, keys):
    """Translate excel columns into XML elements names"""
    try:
        df = pd.DataFrame([elements, keys])
        keymap = dict.fromkeys(elements, '')
        for i in range(len(df.columns)):
            for k in elements:
                if keymap[k] == '' and k == df.iloc[0][i]:
                    if df.iloc[1][i] in keys and df.iloc[1][i] != 'parent node':
                        keymap[k] = df.iloc[1][i]
                        break
                    elif df.iloc[1][i] == 'parent node':
                        break
        return keymap
    except KeyError as e:
        print(e)


def gen_xml(workbook):
    # Load our Excel File
    wb = load_workbook(workbook)
    # Getting an object of active sheet 1
    ws = wb.worksheets[0]
    # Returning returns a triplet
    doc, tag, text = Doc().tagtext()
    headers = get_row_headers(workbook)
    get_xml_elements(ele)
    get_rte_keys(k)
    keymap = gen_keymap(k, ele)

    xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
    # xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema-instance"></xs:schema>'

    doc.asis(xml_header)
    # doc.asis(xml_schema)

    with tag('report', ('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance"), \
             ('xsi:noNamespaceSchemaLocation', "_UAF_Web_Report_.xsd")):
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=186, min_col=1, max_col=127)):
            row = [cell.value for cell in row if row is not None]
            with tag('rentity_id'):
                text(int('0590'))  # int
            with tag('rentity_branch'):
                text(row[3])  # str
            with tag('submission_code'):
                if row[36] is None:
                    row[36] = "E"
                else:
                    text(row[36])
            with tag('report_code'):
                if row[42] is None:
                    row[42] = 'CTR'
                else:
                    text(row[42])
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
                            if row[89] is None:
                                row[89] = 1
                            else:
                                text(row[89])
                        with tag("tph_communication_type"):
                            if row[90] is None:
                                row[90] = 1
                            else:
                                text(row[90])
                        with tag("tph_country_prefix"):
                            text("809")
                        with tag("tph_number"):
                            text("535-8088")
                        with tag("tph_extension"):
                            text("3212")
                with tag("addresses"):
                    with tag("address"):
                        with tag("address_type"):
                            if row[94] is None:
                                row[94] = 1
                            else:
                                text(row[94])
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
                    if row[94] is None:
                        row[94] = 1
                    else:
                        text(row[94])
                with tag("address"):
                    text("Calle Ave. George Washington No. 601")
                with tag("city"):
                    text("Santo Domingo")
                with tag("country_code"):
                    text("DO")

            with tag("reason"):
                if row[104] is None:
                    row[104] = "Transaccion sospechosa"
                else:
                    row[104]
            with tag("action"):
                if row[80] is None:
                    row[80] = "Acciones a tomar"
                else:
                    text(row[80])

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
                    if row[36] is None:
                        row[36] = ""
                    else:
                        text(row[36])
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
                                        text(int("1"))
                                    with tag("tph_communication_type"):
                                        text("C")
                                    with tag("tph_number"):
                                        if row[26] is None:
                                            row[26] = "n/a"
                                        else:
                                            text(row[26])

                            with tag("addresses"):
                                with tag("address"):
                                    with tag("address_type"):
                                        text(int("1"))
                                    with tag("address"):
                                        text(row[25])
                                    with tag("city"):
                                        if row[22] is None:
                                            row[22] = ""
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
                                    text(int("2"))
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
                            text(row[30])
                        with tag("funds_comment"):
                            text(row[31])
                        with tag("country"):
                            text("DO")
                        with tag("significance"):
                            text(int("6"))

                with tag("comments"):
                    text("Prueba")

    result = indent(
        doc.getvalue(),
        indentation='   ',
        indent_text=False
    )

    date = ''.join(str(datetime.now()).replace('-', '_')[:10])
    filename = "_UAF_Web_Report_" + date + ".xml"

    with open(filename, "w") as f:
        f.write(result)


fname = sys.argv[1] if len(sys.argv) > 1 else sg.popup_get_file('Seleccionar archivo')

if not fname:
    sg.popup("Cancel", "No se seleccionó ningún archivo")
    raise SystemExit("Cancelando: no filename supplied")
else:
    # sg.popup('El archivo que seleccionó es: ', fname)
    gen_xml(fname)
    sg.popup('Archivo XML generado exitosamente', fname)
