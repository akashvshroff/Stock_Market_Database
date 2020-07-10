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
        self.url_date = date_today.strftime("%d%m%Y")
        self.base_url = 'https://archives.nseindia.com/products/content/'
        self.border = Border(left=Side(border_style='thin', color='000000'),
                             right=Side(border_style='thin', color='000000'),
                             top=Side(border_style='thin', color='000000'),
                             bottom=Side(border_style='thin', color='000000'))
        self.ft = Font(color='FFFFFF', bold=True, name='Times New Roman')
        self.allign_style = 'center'
        self.parameters = ['CLOSE_PRICE', 'DELIV_PER']
        self.stored_names, self.input_names, self.cell_range, self.url = [], [], '', ''
        self.get_file()
        wb = load_workbook(stored_path)
        self.ws = wb.active
        self.date_column()
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
        stored_data = pd.read_excel(stored_path)
        self.stored_names = stored_data['STOCKS']

    def date_column(self):
        """
        Calculates which column is to be used for the day. Fills in the date
        that is to be used for the present day. Merges the cell, stylises it,

        """
        pass

    def stylise_cells(self):
        """
        Stylises cells within a range with the border, allignment etc.
        """
        rows = self.ws[self.cell_range]
        for row in rows:
            for cell in row:
                cell.border = self.border
                cell.alignment = Alignment(horizontal=self.allign_style, vertical=self.allign_style)

    def enter_data(self):
        """
        Enters the data in the excel sheet based on the names that are already
        stored and those that are new
        """
        pass

    def arrange_rows(self):
        """
        Sorts all the names within the different rows.
        """
        pass
