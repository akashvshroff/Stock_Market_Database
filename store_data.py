from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Alignment, colors, PatternFill, Font, Fill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
import pandas as pd
from filepaths import *
import requests
import win32com.client
from datetime import date


class StoreData()
   def __init__(self):
        """
        Initialises the various dataframes, stylistic variables and data
        structures for the program! Accepts the parameters that are to be
        stored.
        As reports are released for the prev day then the current day is
        considered to be yesterday.
        """
        date_today = date.today() - timedelta(days=1)
        self.d1 = date_today.strftime("%d-%m-%Y")
        self.url_date = date_today.strftime("%d%m%Y")
        self.base_url = 'https://archives.nseindia.com/products/content/'
        self.border = Border(left=Side(border_style='thin', color='000000'),
                             right=Side(border_style='thin', color='000000'),
                             top=Side(border_style='thin', color='000000'),
                             bottom=Side(border_style='thin', color='000000'))
        self.ft = Font(color='FFFFFF', bold=True, name='Times New Roman')
        self.allign_style = 'center'
        self.parameters = ['CLOSE_PRICE', 'DELIV_PER']
        self.start_row = 1
        self.stored_names, self.input_data, self.input_names, self.cell_range, self.url = [], [], [], '', ''
        self.cells_ref = ['' for i in range(len(self.parameters))]
        self.get_file()
        wb = load_workbook(stored_path)
        self.ws = wb.active
        self.date_column()
        self.enter_data()
        os.remove(self.file_path)

    def get_file(self):
        """
        Retrieves the url for the file that is to be scraped.
        """
        self.url = self.base_url + 'sec_bhavdata_full_{}.csv'.format(self.url_date)
        # print(self.url)
        values = requests.get(self.url)
        self.file_path = 'sec_bhavdata_full_{}.csv'.format(self.url_date)
        fhand = open(self.file_path, 'wb')
        fhand.write(values.content)
        fhand.close()
        self.retrieve_data()

    def retrieve_data(self):
        """
        Scrapes the data retrieved, adds it to the data frames and parses the
        data.
        """
        raw_input_df = pd.read_csv(self.file_path, sep=r'\s*,\s*', engine='python')
        input_df = raw_input_df[raw_input_df["SERIES"] == 'EQ']
        self.input_names = input_df['SYMBOL']
        self.input_data = input_df[self.parameters]
        stored_data = pd.read_excel(stored_path)
        self.stored_names = stored_data['STOCKS']


    def date_column(self):
        """
        Calculates which column is to be used for the day. Fills in the date
        that is to be used for the present day. Finds the next columns as well,
        depending on the lenght of the parameters.
        """
        rows = self.ws.iter_rows(min_row=self.start_row,
                            max_row=self.start_row)  # the row where your headings are
        row = next(rows)
        headings = [c.value for c in row]
        col_letter = ''
        for col, heading in enumerate(headings):
            if not col:
                continue  # skip the first col
            if heading == self.d1 or heading is None:
                for i in range(len(self.parameters)):
                    self.cells_ref[i] = get_column_letter(col+i+1)
        if not self.cells_ref:
            for i in range(len(self.parameters)):
                self.cells_ref[i] = get_column_letter(len(headings)+i+1)
        self.enter_initial()

    def enter_initial(self):
        """
        Enters the date, as well as the parameters list into the sheet.
        """
        start_cell,end_cell = f'{self.cells_ref[0]}{self.start_row}',f'{self.cells_ref[-1]}{self.start_row}'
        merge = '{}:{}'.format(start_cell,end_cell)
        self.ws.merge_cells(merge)
        self.ws[start_cell] = self.d1
        cell.border = self.border
        cell.alignment = Alignment(horizontal=self.allign_style, vertical=self.allign_style)
        curr_row = self.start_row + 1
        for i,para in enumerate(self.parameters):
            curr_cell = f'{self.cells_ref[i]}{curr_row}'
            self.ws[curr_cell] = para

    def stylise_cells(self):
        """
        Stylises cells within a range with the border, allignment etc.
        """
        rows = self.ws[self.cell_range]
        for row in rows:
            for cell in row:
                cell.border = self.border
                cell.alignment = Alignment(horizontal=self.allign_style,
                                           vertical=self.allign_style)

    def enter_data(self):
        """
        Enters the data in the excel sheet based on the names that are already
        stored and those that are new are apended to the end of the STOCKS list
        on the first column.
        """
        curr_row = self.start_row + 2
        for stock in self.stored_names:
            if stock in self.input_names: #there is data for it
                stock_index = self.input_names.index(stock)
                for num,parameter in enumerate(self.parameters):
                    curr_cell = '{}{}'.format(self.cells_ref[i],self.curr_row)
                    value = self.input_data[parameter][stock_index]
                    self.ws[curr_cell] = value
                self.input_names.pop(stock_index)




    def arrange_rows(self):
        """
        Sorts all the names within the different rows.
        """
        pass
