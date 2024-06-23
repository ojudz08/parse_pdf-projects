"""
    Author: Ojelle Rogero
    Created on: June 23, 2024
    About:
        Parses the transactions of a bank statement
"""

import pymupdf, tabula
import os, sys, time
from config import config_cred
import pandas as pd
from openpyxl import load_workbook



class bank_estatement():

    def __init__(self):
        self.init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))


    def bound_box(self, doc_page):
        try:
            top = doc_page.search_for("Sale Date")[0].irect[3]
            left = doc_page.search_for("Sale Date")[0].irect[0]
            bot = doc_page.search_for("For more details")[0].irect[3] - 4
            right = doc_page.search_for("Amount")[0].irect[2]
        
            rect_area = [top, left, bot, right]
            return rect_area
        except:
            return None



if __name__ == '__main__':
    pass
    