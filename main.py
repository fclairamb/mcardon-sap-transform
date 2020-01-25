import os
import re
from pathlib import Path
from typing import Optional, List

import openpyxl
import logging

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

logging.basicConfig(
    format='%(asctime)s.%(msecs)03d | %(levelname).4s | %(lineno)4d | %(message)s',
    datefmt='%H:%M:%S',
    level=logging.INFO,
)

CELL_MARKER = 'BP04'
CELL_POS_MARKER = 3
CELL_POS_MARKER_FALLBACK = 2

CELL_POS_DATE = 2
CELL_POS_ACCOUNT_SRC = 7
CELL_POS_ACCOUNT_SRC_FALLBACK = 6
CELL_POS_ACCOUNT_DST = 0

CELL_POS_AMOUNT = 17

DATE_PATTERN = re.compile(r'[0-9]{2}\.[0-9]{2}\.[0-9]{4}')


def copy_rows(input_ws: Worksheet, output_ws: Worksheet):
    nb_rows_copied: int = 0
    nb_rows_read: int = 0
    account: Optional[str] = None
    cell_account_src_pos: int = 0
    cell_marker_pos: int = 0

    if input_ws['D6'].value == 'BP04':
        logging.info("  BP04 at D6 !")
        cell_account_src_pos = CELL_POS_ACCOUNT_SRC
        cell_marker_pos = CELL_POS_MARKER
    elif input_ws['C6'].value == 'BP04':
        logging.info("  BP04 at C6 !")
        cell_account_src_pos = CELL_POS_ACCOUNT_SRC_FALLBACK
        cell_marker_pos = CELL_POS_MARKER_FALLBACK
    else:
        raise ValueError('Could not find BP04 marker !')

    for row in input_ws.iter_rows():

        nb_rows_read += 1

        values: List[str] = list([cell.value for cell in row])
        cell_date: str = values[CELL_POS_DATE]

        if values[cell_marker_pos] == CELL_MARKER:
            account = values[cell_account_src_pos]

        # If we have a cell date
        if cell_date and DATE_PATTERN.match(cell_date):
            # Rewriting the date
            values[CELL_POS_DATE] = cell_date.replace('.', '/')
            values[CELL_POS_ACCOUNT_DST] = account

            # Rewriting the amount
            amount = values[CELL_POS_AMOUNT]
            if type(amount) is str:
                values[CELL_POS_AMOUNT] = float(amount.replace(' ', '').replace(',', ''))

            output_ws.append(values)
            nb_rows_copied += 1
            if nb_rows_copied % 1000 == 0:
                logging.info("  Copied %d rows on %d input rows ...", nb_rows_copied, nb_rows_read)

    logging.info("  Finished (with %d / %d rows) !", nb_rows_copied, nb_rows_read)


def process_files(input_dir: Path, output_ws: Worksheet):
    # For each file, we will
    for input_xlsx in input_dir.glob('**/*.xlsx'):
        # Open it
        logging.info("Opening \"%s\" ...", input_xlsx)

        # For test
        # if input_xlsx.stat().st_size > 100000:
        #    logging.warning("Skipping this big file for now")
        #    continue

        input_wb: openpyxl.Workbook = openpyxl.open(input_xlsx)
        logging.info("  Opened !")

        # Open each sheet
        for input_ws in input_wb.worksheets:
            # Copy it
            copy_rows(input_ws, output_ws)

        # And then close it
        input_wb.close()


def main():
    # Input directory
    input_dir = Path(os.path.join(os.path.dirname(__file__), 'input'))

    # Output workbook
    output_wb = Workbook()
    output_ws = output_wb.active

    output_ws.append([
        'No Compte',
        None,
        'Date Comptable',
        None,
        'Code Journal',
        None,
        'No Piece',
        None,
        'Date piece',
        'CC',
        'CN',
        'Descriptif',
        None,
        None,
        None,
        'Devise',
        'Montant (Devise)',
        'Montant (Euro)',
        'Libellé écriture'
    ])

    # Core of the processing
    process_files(input_dir, output_ws)

    logging.info("Saving output file ...")
    # Save the output workbook
    output_wb.save("output.xlsx")
    logging.info("  Done !")


main()
