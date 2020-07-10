from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Alignment, colors, PatternFill, Font, Fill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
import pandas as pd
from filepaths import *


class InitialiseDb:
    def __init__(self):
        """
        Gets the names of the input stocks, initialise formatting and other
        structures.
        """
        self.input_df = pd.read_excel(input_path)
        self.input_names = self.input_df['Stocks']
        wb = Workbook()
        self.ws = wb.active
        self.border = Border(left=Side(border_style='thin', color='000000'),
                             right=Side(border_style='thin', color='000000'),
                             top=Side(border_style='thin', color='000000'),
                             bottom=Side(border_style='thin', color='000000'))
        self.ft = Font(color='FFFFFF', bold=True, name='Times New Roman')
        self.allign_style = 'center'
        self.cell_range = 'A1:A{}'.format(1+len(self.input_names))
        self.ws['A1'] = 'Stocks'
        self.store_names()
        self.stylise_cells()
        wb.save(stored_path)

    def store_names(self):
        """
        Adds names of initial stocks to the excel sheet.
        """
        st_row, st_col = 2, 'A'
        for name in self.input_names:
            curr_cell = f'{st_col}{st_row}'
            self.ws[curr_cell] = name
            st_row += 1

    def stylise_cells(self):
        """
        Called upon by init after getting  to stylise the rows
        """
        rows = self.ws[self.cell_range]
        for row in rows:
            for cell in row:
                cell.border = self.border
                cell.alignment = Alignment(horizontal=self.allign_style,
                                           vertical=self.allign_style)


def main():
    obj = InitialiseDb()


if __name__ == '__main__':
    main()
