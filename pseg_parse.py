"""
pseg_parse.py
"""

from typing import Union
import re
from csv import reader
from argparse import ArgumentParser
from datetime import datetime, time, timedelta
import xlsxwriter
from set_outer_border_for_range_xlsx import apply_outer_border_to_range

FIRST_HEADER_ROW: int = 3

# Peak and super off peak time periods.
# Peak time is from 3:00 PM up to 7:00 PM except weekends.  3:00 PM reading is last meter reading for off-peak.  3:15 is first peam reading.
# Similarly, the last peak reading is at 7:00 PM, and off-peak starts with the 7:15 PM reading.
PEAK_START: time = time(15, 0, 1)
PEAK_END: time = time(19, 0, 0)

# Super Off-Peak is from 10:00 PM up to 6:00:00 AM including weekends.
# Similar to peak start and end, The last off-peak reading ends at 10:00 PM and 10:15 is the first meter reading for super off-peak.
# The last meter reading for super off-peak occurs at 6:00 AM the next day.
# Shift by three hours to make comparisons easier.  This gets the time periods in the same day (1 AM to 9 AM).
TIME_OFFSET: timedelta = timedelta (hours = 2)
SUPER_OFF_PEAK_START: time = (datetime(1, 1, 1, 22, 0, 1) + TIME_OFFSET).time()
SUPER_OFF_PEAK_END: time = (datetime(1, 1, 1, 6, 00, 00) + TIME_OFFSET).time()

def get_day_of_week(date: datetime) -> int:
    """
    Return day of week (0 = Sunday, 1 = Monday, ..., 6 = Saturday) given a datetime object.
    """
    return int(date.strftime('%w'))

def is_peak_time(date_time_val: datetime) -> bool:
    """
    Returns True if date_time_val is in peak period.

    Peak begins at PEAK_START and ends at PEAK_END, except Saturady and Sunday.
    """

    # Weekends aren't peak.
    if get_day_of_week(date_time_val) in [0, 6]:
        return False
    
    return (date_time_val.time() >= PEAK_START) and (date_time_val.time() <= PEAK_END)

def is_super_off_peak_time(date_time_val: datetime) -> bool:
    """
    Returns True if date_time_val is in super_off_peak_period.

    Super off peak is between 10:00 PM and 5:59:59 AM inclusive.  Adding two hours to super off peak start, super off peak end,
    and the candidate date_time_val makes the comparisons easier, since then they all fall within the same day.
    """
    time_val: time = (date_time_val + TIME_OFFSET).time()

    return (time_val >= SUPER_OFF_PEAK_START) and (time_val <= SUPER_OFF_PEAK_END)

def add_title_cells(book, sheet) -> int:
    """
    Adds titles for various TOU plans
    """
    merge_center = book.add_format({
        "bold": 0,
        "border": 0,
        "align": "center",
        "valign": "vcenter"
    })
    merge_left = book.add_format({
        "bold": 0,
        "border": 0,
        "align": "left",
        "valign": "vcenter"
    })

    row_num: int = FIRST_HEADER_ROW
    sheet.write(f'B{row_num}', 'Non-TOU Consumed:', merge_left)
    sheet.write(f'E{row_num}', 'Off-Peak Consumed:', merge_left)
    sheet.write(f'H{row_num}', 'Super Off-Peak Consumed:', merge_left)
    row_num += 1

    sheet.write(f'B{row_num}', 'Non-TOU Generated:', merge_left)
    sheet.write(f'E{row_num}', 'Off-Peak Generated:', merge_left)
    sheet.write(f'H{row_num}', 'Super Off-Peak Generated:', merge_left)
    row_num += 1

    sheet.write(f'E{row_num}', 'Peak Consumed:', merge_left)
    sheet.write(f'H{row_num}', 'Off-Peak Consumed:', merge_left)
    row_num += 1

    sheet.write(f'E{row_num}', 'Peak Generated:', merge_left)
    sheet.write(f'H{row_num}', 'Off-Peak Generated:', merge_left)
    row_num += 1

    sheet.write(f'H{row_num}', 'Peak Consumed:', merge_left)
    row_num += 1

    sheet.write(f'H{row_num}', 'Peak Generated:', merge_left)
    row_num += 2

    sheet.write(f'B{row_num}', 'Non-TOU Net:', merge_left)
    sheet.write(f'E{row_num}', 'Off-Peak Net:', merge_left)
    sheet.write(f'H{row_num}', 'Super Off-Peak Net:', merge_left)
    row_num += 1

    sheet.write(f'E{row_num}', 'Peak Net:', merge_left)
    sheet.write(f'H{row_num}', 'Off-Peak Net:', merge_left)
    row_num += 1

    sheet.write(f'H{row_num}', 'Peak Net:', merge_left)
    row_num += 2

    sheet.write(f'B{row_num}', 'Bill:', merge_left)
    sheet.write(f'E{row_num}', 'Bill:', merge_left)
    sheet.write(f'H{row_num}', 'Bill:', merge_left)
    row_num += 2

    sheet.merge_range(f'B{row_num}:C{row_num}', 'Non-TOU Billing', merge_center)
    sheet.merge_range(f'E{row_num}:H{row_num}', 'Off-Peak Billing', merge_center)
    sheet.merge_range(f'J{row_num}:O{row_num}', 'Super Off-Peak Billing', merge_center)
    row_num += 1

    sheet.merge_range(f'E{row_num}:F{row_num}', 'Peak', merge_center)
    sheet.merge_range(f'G{row_num}:H{row_num}', 'Off-Peak', merge_center)

    sheet.merge_range(f'J{row_num}:K{row_num}', 'Peak', merge_center)
    sheet.merge_range(f'L{row_num}:M{row_num}', 'Off-Peak', merge_center)
    sheet.merge_range(f'N{row_num}:O{row_num}', 'Super Off-Peak', merge_center)
    row_num += 1

    sheet.write_row(f'A{row_num}', ['Time', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', 'Consumed', 'Generated'])
    row_num += 1

    sheet.freeze_panes(f'B{row_num}')

    # Return first row number after titles
    return row_num

def add_formulas(sheet, first_header_row: int, first_data_row: int, last_data_row: int) -> None:
    """
    Populates formulas in header of worksheet.
    """

    # Non-TOU billing formulas
    sheet.write_formula(f'C{first_header_row}', f'=SUM(B{first_data_row}:B{last_data_row})')                # Consumed
    sheet.write_formula(f'C{first_header_row + 1}', f'=SUM(C{first_data_row}:C{last_data_row})')            # Generated
    sheet.write_formula(f'C{first_header_row + 7}', f'=C{first_header_row} + C{first_header_row + 1}')      # Net

    # Off-Peak billing formulas
    sheet.write_formula(f'F{first_header_row}', f'=SUM(G{first_data_row}:G{last_data_row})')                # Off-Peak consumed
    sheet.write_formula(f'F{first_header_row + 1}', f'=SUM(H{first_data_row}:H{last_data_row})')            # Off-Peak generated
    sheet.write_formula(f'F{first_header_row + 7}', f'=F{first_header_row} + F{first_header_row + 1}')      # Off-Peak net
    sheet.write_formula(f'F{first_header_row + 2}', f'=SUM(E{first_data_row}:E{last_data_row})')            # Peak consumed
    sheet.write_formula(f'F{first_header_row + 3}', f'=SUM(F{first_data_row}:F{last_data_row})')            # Peak generated
    sheet.write_formula(f'F{first_header_row + 8}', f'=F{first_header_row + 2} + F{first_header_row + 3}')  # Peak net

    # Super Off-Peak billing formulas
    sheet.write_formula(f'I{first_header_row}', f'=SUM(N{first_data_row}:N{last_data_row})')                # Super Off-Peak consumed
    sheet.write_formula(f'I{first_header_row + 1}', f'=SUM(O{first_data_row}:O{last_data_row})')            # Super Off-Peak generated
    sheet.write_formula(f'I{first_header_row + 7}', f'=I{first_header_row} + I{first_header_row + 1}')      # Super Off-Peak net
    sheet.write_formula(f'I{first_header_row + 2}', f'=SUM(L{first_data_row}:L{last_data_row})')            # Off-Peak consumed
    sheet.write_formula(f'I{first_header_row + 3}', f'=SUM(M{first_data_row}:M{last_data_row})')            # Off-Peak generated
    sheet.write_formula(f'I{first_header_row + 8}', f'=I{first_header_row + 2} + I{first_header_row + 3}')  # Off-Peak net
    sheet.write_formula(f'I{first_header_row + 4}', f'=SUM(J{first_data_row}:J{last_data_row})')            # Peak consumed
    sheet.write_formula(f'I{first_header_row + 5}', f'=SUM(K{first_data_row}:K{last_data_row})')            # Peak generated
    sheet.write_formula(f'I{first_header_row + 9}', f'=I{first_header_row + 4} + I{first_header_row + 5}')  # Peak net


def format_cells(book, sheet, first_data_row: int, last_data_row: int) -> None:
    """
    Add borders to data and header columns.
    """
    # Non-TOU Billing Header
    apply_outer_border_to_range(book, sheet, {"range_string": "B18:B18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "C18:C18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "B16:C18", "border_style": 5})

    # Non-TOU Billing data
    apply_outer_border_to_range(book, sheet, {"range_string": f"B{first_data_row}:B{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"B{first_data_row}:C{last_data_row}", "border_style": 5})

    # Off-Peak Billing Header
    apply_outer_border_to_range(book, sheet, {"range_string": "E17:F17", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "G17:H17", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "E18:E18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "F18:F18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "G18:G18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "H18:H18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "E16:H18", "border_style": 5})

    # Off-Peak Billing data
    apply_outer_border_to_range(book, sheet, {"range_string": f"E{first_data_row}:E{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"F{first_data_row}:F{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"G{first_data_row}:G{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"E{first_data_row}:H{last_data_row}", "border_style": 5})

    # Super Off-Peak Billing Header
    apply_outer_border_to_range(book, sheet, {"range_string": "J17:K17", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "L17:M17", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "N17:O17", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "J18:J18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "K18:K18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "L18:L18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "M18:M18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "N18:N18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "O18:O18", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": "J16:O18", "border_style": 5})

    # Super Off-Peak Billing data
    apply_outer_border_to_range(book, sheet, {"range_string": f"J{first_data_row}:J{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"K{first_data_row}:K{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"L{first_data_row}:L{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"M{first_data_row}:M{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"N{first_data_row}:N{last_data_row}", "border_style": 1})
    apply_outer_border_to_range(book, sheet, {"range_string": f"J{first_data_row}:O{last_data_row}", "border_style": 5})

    # Center all cells
    centered = book.add_format({'align': 'center', 'valign': 'vcenter'})
    sheet.conditional_format(f'B16:O{last_data_row}', {'type': 'no_errors', 'format' : centered})


def pseg_parse(csv_file_name: str, xslx_file_name: str) -> None:
    """
    Parse PSEG downloaded CSV file and create Excel xslx file.
    """
    workbook: Workbook = xlsxwriter.Workbook(xslx_file_name, {'default_date_format': 'mmm dd yyyy hh:mm'})
    worksheet: Unknown | Worksheet = workbook.add_worksheet('PSEG TOU Usage')
    if worksheet is None:
        print('Error creating worksheet!')
        return

    # first_row is cell addressing (one-based).
    # sheet_row is row addressing (zero-based).
    # add_title_cells returns row in cell addressing mode, which is one-based.
    first_row: int = add_title_cells(workbook, worksheet)
    sheet_row: int = first_row - 1

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
                    return
                
                if meter_num != gen_meter_num:
                    print(f'Mismatched meters at lines [{line} - {line + 1}].')
                    return
                
                meter_time: datetime = datetime.strptime(meter[0], '%m/%d/%Y %I:%M:%S %p')
                gen_meter_time: datetime = datetime.strptime(gen_meter[0], '%m/%d/%Y %I:%M:%S %p')

                if meter_time != gen_meter_time:
                    print(f'Mismatched meter times at lines [{line} - {line + 1}].')
                    return
                
                # Column 0: Time
                worksheet.write_datetime(sheet_row, 0, meter_time)

                consumed: float = float(meter[kwh_col])
                generated: float = float(gen_meter[kwh_col])
                # Columns B-D - Non-TOU Billing: Consumed, Gen'd, and empty cell.
                kwh_data: list[Union[float, str]] = [consumed, generated, '']

                # Rate 195 Super off-peak is off-peak for rate 194
                if is_super_off_peak_time(meter_time):
                    # Columns E-I = Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, and empty cell.
                    kwh_data.extend([0, 0, consumed, generated, ''])
                    # Columns J-P - Super Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, Super Off-Peak Consumed and Gen'd, and empty cell.
                    kwh_data.extend([0, 0, 0, 0, consumed, generated, ''])
                else:
                    if is_peak_time(meter_time):
                        # Columns E-I = Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, and empty cell.
                        kwh_data.extend([consumed, generated, 0, 0, ''])
                        # Columns J-P - Super Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, Super Off-Peak Consumed and Gen'd, and empty cell.
                        kwh_data.extend([consumed, generated, 0, 0, 0, 0, ''])
                    else:
                        # If it's neither super off-peak nor peak, then it's off-peak.
                        # Columns E-I = Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, and empty cell.
                        kwh_data.extend([0, 0, consumed, generated, ''])
                        # Columns J-P - Super Off-Peak Billing - Peak Consumed and Gen'd, Off-Peak Consumed and Gen'd, Super Off-Peak Consumed and Gen'd, and empty cell.
                        kwh_data.extend([0, 0, consumed, generated, 0, 0, ''])
                worksheet.write_row(sheet_row, 1, kwh_data)
                sheet_row += 1

        except StopIteration:
            print('Read beyond end of file.  Check that both meter and generated meter data are included.')

    worksheet.autofit()
    add_formulas(worksheet, FIRST_HEADER_ROW, first_row, sheet_row)
    format_cells(workbook, worksheet, first_row, sheet_row)
    workbook.close()

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
