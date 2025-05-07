"""
pseg_parse.py
"""

from typing import Union
import re
from csv import reader
from argparse import ArgumentParser
from datetime import datetime, time, timedelta
import xlsxwriter

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

    sheet.merge_range('E2:F2', 'Peak', merge_format)
    sheet.merge_range('G2:H2', 'Off-Peak', merge_format)

    sheet.merge_range('J2:K2', 'Peak', merge_format)
    sheet.merge_range('L2:M2', 'Off-Peak', merge_format)
    sheet.merge_range('N2:O2', 'Super Off-Peak', merge_format)

    sheet.write_row('A3', ['Time', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', '',
                           'Consumed', 'Generated', 'Consumed', 'Generated', 'Consumed', 'Generated'])

def pseg_parse(csv_file_name: str, xslx_file_name: str) -> None:
    """
    Parse PSEG downloaded CSV file and create Excel xslx file.
    """
    workbook: Workbook = xlsxwriter.Workbook(xslx_file_name, {'default_date_format': 'mmm dd yyyy hh:mm'})
    worksheet: Unknown | Worksheet = workbook.add_worksheet('PSEG TOU Usage')
    if worksheet is None:
        print('Error creating worksheet!')
        return

    add_title_cells(workbook, worksheet)

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
            sheet_row: int = 3  # Last title row is 3 and worksheet cell indexing is zero based.
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
