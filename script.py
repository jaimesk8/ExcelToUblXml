import openpyxl
from datetime import datetime
from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace, QName
from xml.dom import minidom
from openpyxl import load_workbook

# Create UBL XML document
def create_ubl_xml(data):
    common_namespace = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
    basic_namespace = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
    
    ubl_doc = Element(f'{{{basic_namespace}}}Invoice')

    # Register the 'cbc' and 'cac' namespace prefix
    register_namespace('cbc', basic_namespace) #cbc 
    register_namespace('cac', common_namespace) #cac

    #incremental id variable
    invoice_id = 1

    # Create elements and add data from Excel
    for row in data:
        
        invoice_line = SubElement(ubl_doc, f'{{{basic_namespace}}}InvoiceLine')

        invoice_id_element = SubElement(invoice_line, QName(basic_namespace, 'ID'))
        invoice_id_element.text = str(invoice_id)
        
        invoice_id += 1
        
        note_element = SubElement(invoice_line, 'cbc:Note')
        note_element.text = "STATIC_VALUE"

        # InvoicedQuantity
        invoiced_quantity = SubElement(invoice_line, f'{{{basic_namespace}}}InvoicedQuantity', unitCode="KWT")
        invoiced_quantity.text = str(row['InvoicedQuantity'])

        # LineExtensionAmount
        line_extension_amount = SubElement(invoice_line, f'{{{basic_namespace}}}LineExtensionAmount', currencyID="EUR")
        line_extension_amount.text = str(row['LineExtensionAmount'])

        # InvoicePeriod
        additional_item_InvoicePeriod_element = SubElement(invoice_line, QName(common_namespace, 'InvoicePeriod'))
        item_InvoicePeriod = SubElement(additional_item_InvoicePeriod_element, f'{{{basic_namespace}}}StartDate')
        item_InvoicePeriod.text = str(row['StartDate'])
        item_InvoicePeriod = SubElement(additional_item_InvoicePeriod_element, f'{{{basic_namespace}}}EndDate')
        item_InvoicePeriod.text = str(row['EndDate'])
        
        item_element = SubElement(invoice_line, QName(common_namespace, 'Item'))

        name_element = SubElement(item_element, QName(basic_namespace, 'Name'))
        name_element.text = "STATIC_VALUE"

        classified_tax_category_element = SubElement(item_element, QName(common_namespace, 'ClassifiedTaxCategory'))

        id_element = SubElement(classified_tax_category_element, QName(basic_namespace, 'ID'))
        id_element.text = "STATIC_VALUE"

        percent_element = SubElement(classified_tax_category_element, QName(basic_namespace, 'Percent'))
        percent_element.text = "STATIC_VALUE"

        tax_scheme_element = SubElement(classified_tax_category_element, QName(common_namespace, 'TaxScheme'))

        id_tax_scheme_element = SubElement(tax_scheme_element, QName(basic_namespace, 'ID'))
        id_tax_scheme_element.text = "STATIC_VALUE"

        additional_item_property_element = SubElement(item_element, QName(common_namespace, 'AdditionalItemProperty'))

        name_property_element = SubElement(additional_item_property_element, QName(basic_namespace, 'Name'))
        name_property_element.text = "STATIC_VALUE"

        #Value
        item_value = SubElement(additional_item_property_element, f'{{{basic_namespace}}}Value')
        item_value.text = str(row['Value'])
        
        #PriceAmount
        additional_item_price_element = SubElement(invoice_line, QName(common_namespace, 'Price'))
        item_PriceAmount = SubElement(additional_item_price_element, f'{{{basic_namespace}}}PriceAmount', currencyID="EUR")
        item_PriceAmount.text = str(row['PriceAmount'])
    return ubl_doc


def save_xml_to_file(xml_data, filename):
    # Save the XML to a file
    with open(filename, 'w') as xml_file:
        xml_file.write(prettify(xml_data))

def prettify(elem):
    # Prettify the XML output
    rough_string = tostring(elem, 'utf-8').decode('utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

#map the values from excel cell to xml estruture
def excel_to_ubl(input_excel, output_xml):
    # Load Excel data
    workbook = load_workbook(input_excel, data_only=True)
    sheet = workbook.active
    
    # Extract data from Excel
    excel_data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):

         # Format the date to "yyyy-mm-dd"
        dateStart_value = row[0]
        formatted_date = datetime.strftime(dateStart_value, "%Y-%m-%d") if dateStart_value is not None else None

        dateEnd_value = row[1]
        formatted_date1 = datetime.strftime(dateEnd_value, "%Y-%m-%d") if dateEnd_value is not None else None
        
        # Mapping the Excel File
        excel_data.append({
            'InvoicedQuantity': row [13],
            'LineExtensionAmount': row[2],
            'Value': row[3],
            'PriceAmount': row[14],
            'StartDate': formatted_date,
            'EndDate': formatted_date1
        })

    # Create UBL XML document
    ubl_xml = create_ubl_xml(excel_data)
    
    # Save UBL XML to file
    save_xml_to_file(ubl_xml, output_xml)

if __name__ == "__main__":
    # Specify the input Excel file and output UBL XML file
    input_excel_file = "YOUR_PATH_TO_INPUTFILE"
    output_ubl_xml_file = "YOUR_PATH_TO_OUTPUFILE.xml"

    # Convert Excel to UBL XML
    excel_to_ubl(input_excel_file, output_ubl_xml_file)
