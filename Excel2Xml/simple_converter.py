import PySimpleGUI as sg
from openpyxl import load_workbook
from resources import get_resources_path
import os.path
from yattag import Doc, indent
import pandas as pd

# Approach
#
# Import modules
# Load Excel file
# Create sheet object
# (Check) Iterate through rows
# Read data
# Create XML format page
# Append to file
# Save file

WORKBOOK = get_resources_path("data/transacciones_efectivo_2019_alt.xlsx")
TEMPLATE = get_resources_path("data/_Web_Report_ReportID_3234-0-0.xml")
RTEMAP = get_resources_path("data/mapped_elements.xlsx")
k = []
ele = []


def get_row_headers():
    # Loading the Excel file
    tmp_wb = load_workbook(WORKBOOK)
    # creating the sheet 1 object
    ws = tmp_wb.worksheets[0]
    # Iterating rows for getting the values of each row
    for row in ws.iter_rows(min_row=7, max_row=7, min_col=1, max_col=127):
        header = [cell.value for cell in row]
    return header


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
    headers = get_row_headers()
    get_xml_elements(ele)
    get_rte_keys(k)
    keymap = gen_keymap(k, ele)

    xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
    xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"></xs:schema>'

    with tag('report'):
        for idx, row in enumerate(ws.iter_rows(min_row=8, max_row=186, min_col=1, max_col=127)):
            row = [cell.value for cell in row if row is not None]
            with tag('rentity_id'):
                doc.asis(row[1])
            with tag('rentity_branch'):
                text(row[3])
            with tag('submission_code'):
                if row[77] is None:
                    row[77] = ""
                else:
                    text(row[77])
            with tag('report_code'):
                if row[78] is None:
                    row[78] = ""
                else:
                    text(row[78])
            with tag('entity_reference'):
                doc.asis("Bagricola")
            with tag('fiu_ref_number'):
                text("UAF Santo Domingo")
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
                    text("1972-10-03T00:00:00")
                with tag("id_number"):
                    text("001-0955138-2")
                with tag("nationality1"):
                    text("DO")
                with tag("phones"):
                    with tag("phone"):
                        with tag("tph_contact_type"):
                            if row[88] is None:
                                row[88] = ""
                            else:
                                text(row[88])
                        with tag("tph_communication_type"):
                            if row[89] is None:
                                row[89] = ""
                            else:
                                text(row[89])
                        with tag("tph_number"):
                            text("535-8088")
                        with tag("tph_country_prefix"):
                            text("809")
                        with tag("tph_extension"):
                            text("3212")
                with tag("addresses"):
                    with tag("address"):
                        with tag("address_type"):
                            if row[94] is None:
                                row[94] = ""
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

            with tag("location"):
                with tag("address_type"):
                    if row[93] is None:
                        row[93] = ""
                    else:
                        text(row[93])
                with tag("address"):
                    text("Calle Ave. George Washington No. 601")
                with tag("city"):
                    text("Santo Domingo")
                with tag("country_code"):
                    text("DO")

            with tag("reason"):
                if row[104] is None:
                    row[104] = ""
                else:
                    text(row[104])
            # with tag("action"):
            #     text(row[80])

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
                # with tag("date_transaction"):
                #     text(row[37])
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
                        row[34] = ""
                    else:
                        text(row[34])
                with tag("involved_parties"):
                    with tag("party"):
                        with tag("role"):
                            pass

    result = indent(
        doc.getvalue(),
        indentation='   ',
        indent_text=False
    )

    with open("output.xml", "w") as f:
        f.write(result)


sg.theme('TanBlue')  # No gray windows please!

# STEP 1 define the layout
layout = [
    [sg.Text('Convertidor simple de formato Excel a XML')],
    [
        sg.Text("Archivo Excel:"),
        sg.In(size=(28, 10), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(),
    ],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 10), key="-FILE LIST-"
        )
    ],
    [sg.Button('Convertir'), sg.Button('Salir')]
]

# STEP 2 - create the window
window = sg.Window('Excel2XML', layout, grab_anywhere=True)

# STEP3 - the event loop
while True:
    event, values = window.read()  # Read the event that happened and the values dictionary
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Salir':  # If user closed window with X or if user clicked "Exit" button
        # then exit
        break
    elif event == 'Convertir':
        # get_resources_path("data/transacciones_efectivo_2019_alt.xlsx")
        wb = get_resources_path("data/" + values["-FILE LIST-"][0])
        # arg = values["-FILE LIST-"][0]
        gen_xml(wb)
    elif event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
               and f.lower().endswith((".xlsx", ".xls"))
        ]
        window["-FILE LIST-"].update(fnames)
    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
        except:
            pass

window.close()
