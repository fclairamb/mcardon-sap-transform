import re
from pathlib import Path
from typing import Optional

import openpyxl
import logging

logging.basicConfig(
    format='%(asctime)s.%(msecs)03d | %(levelname)-4s | %(lineno)4d | %(message)s',
    datefmt='%H:%M:%S',
    level=logging.INFO,
)

CELL_POS_DATE = 2
CELL_POS_ACCOUNT = 7

DATE_PATTERN = re.compile(r'[0-9]{2}\.[0-9]{2}\.[0-9]{4}')

input_dir = Path('input')

output_wb = openpyxl.Workbook()
output_ws = output_wb.active

nb_rows_copied: int = 0

for input_xlsx in input_dir.glob('**/*.xlsx'):
    logging.info("Opening \"%s\" ...", input_xlsx)
    input_wb: openpyxl.Workbook = openpyxl.open(input_xlsx)
    logging.info("Opened !")
    account: Optional[str] = None

    for input_ws in input_wb.worksheets:
        for row in input_ws.iter_rows():
            values = [cell.value for cell in row]
            date_cell = values[2]
            account_cell = values[7]
            if account_cell:
                account = account_cell
            if date_cell and DATE_PATTERN.match(values[2]):
                values[7] = account
                output_ws.append(values)
                nb_rows_copied += 1
                if nb_rows_copied % 1000 == 0:
                    logging.info("Copied %d rows ...", nb_rows_copied)
        input_wb.close()

# Save the file
output_wb.save("output.xlsx")

