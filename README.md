# Excel To Ubl Xml 
A python script to convert Excel File to XML UBL Portuguese Invoiced Format

## Universal Business Language
Universal Business Language is an open library of standard electronic XML business documents for procurement and transportation such as purchase orders, invoices, transport logistics and waybills. 

This script translate a Excel file ampping to a XMl UBL 2.1 Portuguese eletric invoice format.
For that we give the desired Excel file, and with python we map the cells to xml estruture. It utilizes the openpyxl library for Excel file handling and xml.etree.ElementTree for XML generation.

## Requirementst 
. Python 3.6 or later
. Required Python packages can be installed using the following command:

```
pip install openyxl
```

## Usage 

```
import openpyxl
from datetime import datetime
from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace, QName
from xml.dom import minidom
from openpyxl import load_workbook

```

## Adicional Notes 

- Make sure your Excel file follows the expected structure with necessary columns.
- Ensure that the Excel file is not open while running the script.
- Check the XML output and customize the create_ubl_xml function if needed.


