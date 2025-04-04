"""
pseg_parse.py
"""

import re
from csv import reader
from argparse import ArgumentParser
from datetime import datetime
import xlsxwriter

def add_title_cells(book, sheet) -> None:
    """
    Adds titles for various TOU plans
    """
    merge_format = book.add_format({
        "bold": 0,
        "border": 0,
        "align": "center",
        "valign": "vcenter"
    })

    sheet.merge_range('B1:C1', 'Non-TOU Billing', merge_format)
    sheet.merge_range('E1:H1', 'Off-Peak Billing', merge_format)
    sheet.merge_range('J1:O1', 'Super Off-Peak Billing', merge_format)
    sheet.merge_range('Q1:R1', 'Standard Billing', merge_format)

    sheet.merge_range('E2:F2', 'Peak', merge_format)
    sheet.merge_range('G2:H2', 'Off-Peak', merge_format)

    sheet.merge_range('J2:K2', 'Peak', merge_format)
    sheet.merge_range('L2:M2', 'Off-Peak', merge_format)
    sheet.merge_range('N2:O2', 'Super Off-Peak', merge_format)

    sheet.write_row('A3', ['Time', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated'])


def pseg_parse(csv_file_name: str, xslx_file_name: str) -> None:
    """
    Parse PSEG downloaded CSV file and create Excel xslx file.
    """
    workbook: Workbook = xlsxwriter.Workbook(xslx_file_name)
    worksheet: Unknown | Worksheet = workbook.add_worksheet('PSEG TOU Usage')
    if worksheet is None:
        print('Error creating worksheet!')
        return

    add_title_cells(workbook, worksheet)

    worksheet.autofit()
    workbook.close()

    with open(csv_file_name) as csv_file:
        csv_reader = reader(csv_file)

        try:
            kwh_col: int = next(csv_reader).index('kWh')
        except ValueError:
            print('Downloaded usage data must include kWh!')
            return
        
        try:
            line: int = 0
            meter: list[str]
            for line, meter in enumerate(csv_reader):
                gen_meter = next(csv_reader)
                try:
                    meter_num = int(re.split(' #| - ', meter[1])[1])
                    gen_meter_num = int(re.split(' #|g - ', gen_meter[1])[1])
                except ValueError:
                    print(f'Error parsing usage data at lines [{line} - {line + 1}].  Check that both meter and generated meter data are included.')
        except StopIteration:
            print(f'Read beyond end of file at line {line}.  Check that both meter and generated meter data are included.')
            
                
                meter = line[1].split(' '






def main():
    """
    Main function for pseg_parse.
    """
    parser = ArgumentParser()
    parser.add_argument('--bill', '-b', type=str, required=True, help='Path to downloaded bill as csv.')
    parser.add_argument('--excel', '-e', type=str, required=True, help='Path to Excel file that will be generated.')
    args = parser.parse_args()

    pseg_parse (args.bill, args.excel)


if __name__ == "__main__":
    main()
