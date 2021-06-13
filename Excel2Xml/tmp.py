import xml.etree.ElementTree as ET
import openpyxl as xl
import pathlib
from resources import get_resources_path
from xml.dom import minidom
import os

XML_TEMPLATE = get_resources_path("../data/_Web_Report_ReportID_19307-0-1.xml")

if __name__ == '__main__':
    def get_tree():
        return ET.parse("../data/" + XML_TEMPLATE.name)

    elms = []
    root = get_tree().getroot()

    # for elm in root.findall('./'):
    #     elms.append(elm.tag)

    for elm in root.iter():
        elms.append(elm)

    print(elms)
