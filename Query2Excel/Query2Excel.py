import xml.etree.ElementTree as ET
import pyodbc
import openpyxl
import datetime
import argparse
import os

def generate_excel(sheet_name, cnn_str, query, wb):
    print(f'Preparing the sheet {sheet_name}')

    ws = wb.create_sheet(title=sheet_name)

    cnn = pyodbc.connect(cnn_str)
    cur = cnn.cursor()
    cur.execute(query)
    rows = cur.fetchall()

    if not rows:
        return

    row_no = 1

    # Add Header
    for i in range(len(rows[0].cursor_description)):
        cell = ws.cell(row=row_no, column=i + 1)
        # First item in cursor_description is the field name
        cell.value = rows[0].cursor_description[i][0]

        if isinstance(rows[0][i], datetime.datetime):
            cell.number_format = 'dd-MMM-yy'
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')

    row_no += 1

    for i in range(len(rows)):
        row = rows[i]

        for j in range(len(row)):
            cell = ws.cell(row=row_no, column=j + 1)
            cell.value = row[j]
            if isinstance(row[j], datetime.datetime):
                cell.number_format = 'dd-MMM-yy'
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')

        row_no += 1


def main_xml():
    aparser = argparse.ArgumentParser(
        description='Mitesh: Query 2 Excel.  Give a XML file with query and have output in Excel .',
        epilog='You are seeing EPILOG'
    )
    aparser.add_argument('xml', type=str, help='xml file having queries, connections, etc...')
    args = aparser.parse_args()
    file_path_xml = args.xml

    tree = ET.parse(file_path_xml)
    root = tree.getroot()

    # Connection string given in queries root node
    global_cnn_str = ''
    if 'connection' in root.attrib:
        global_cnn_str = root.attrib['connection']


    output_excel = os.path.join(root.attrib['filepath'], root.attrib['filename'])
    if 'appenddate' in root.attrib:
        if root.attrib['appenddate'] == 'yes':
            today = datetime.date.today()
            output_excel = f"{output_excel}_{today.strftime('%d-%b-%Y')}"
    output_excel = output_excel + '.xlsx'

    wb = openpyxl.Workbook()
    for query in root:
        sheet_name = query.attrib['name']

        # If connection attribute provided at query, pick connection string from the query node.
        if 'connection' in query.attrib:
            cnn_str = query.attrib['connection']
        else:
            cnn_str = global_cnn_str

        sql = query.find('sql')
        # If file attribute given..
        if 'file' in sql.attrib:
            # Pick query from the sql file
            with open(sql.attrib['file'], 'r') as sqlfile:
                sql_text = sqlfile.read()
        else:
            sql_text = sql.text

        generate_excel(sheet_name, cnn_str, sql_text, wb)

    wb.save(output_excel)
    print(f'File {output_excel} created successfully.')

if __name__ == '__main__':
    #main_json()
    main_xml()
