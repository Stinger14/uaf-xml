# Excel2Xml
---

Convert an Excel file generated by the RTE module from Banke to XML format.

This project uses the python libraries *yattag* and *xml.etree.ElementTree* 
to generate and update the XML tree from values obtained from the Excel file fetched
with *pandas* and *openpyxl*.

This is a gui desktop app built with *PySimpleGUI*.

---

## Pack all components into a windows executable program with *pyinstaller*:

```pyinstaller Excel2Xml\simple_converter.py -F --name "Excel2Xml" --add-data ".\data\*;data" --clean```