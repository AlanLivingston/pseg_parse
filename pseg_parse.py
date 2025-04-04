"""
pseg_parse.py
"""

from csv import reader
import xslxwriter

def pseg_parse(csv_file_name: str, xslx_file_name: str) -> None:
    """
    Parse PSEG downloaded CSV file and create Excel xslx file.
    """
    with open(csv_file_name) as csv_file:
        csv_reader = reader(csv_file)26



def main():
    """
    Main function for pseg_parse.
    """
    pseg_parse ('2025-01-16 to 2025-02-13--Usage.csv', 'usage.xslx')


if __name__ == "__main__":
    main()
