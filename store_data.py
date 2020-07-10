from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Alignment, colors, PatternFill, Font, Fill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
import pandas as pd
from filepaths import *
import requests
import win32com.client


class StoreData()
   def __init__(self, parameters):
        """
        Initialises the various dataframes, stylistic variables and data
        structures for the program! Accepts the parameters that are to be
        stored.
        """
        pass

    def get_url(self):
        """
        Retrieves the url for the file that is to be scraped.
        """
        pass

    def scrape_data(self):
        """
        Scrapes the data retrieved, adds it to the data frames and parses the
        data.
        """
        pass

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
        pass

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
