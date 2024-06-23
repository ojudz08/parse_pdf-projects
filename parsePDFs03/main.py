"""
    Author: Ojelle Rogero
    Created on: June 23, 2024
    About:
        Parses the transactions of a bank statement
"""

import pymupdf, tabula
import os, sys, time
from config import config01
import pandas as pd
import numpy as np
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
    

    def transform(self, df):
        idx = []
        for i in range(0, len(df) - 3):
            i_n = i + 1
            if isinstance(df["Sale Date"][i_n], str): 
                pass
            elif np.isnan(df["Sale Date"][i_n]) and np.isnan(df["Post Date"][i_n]):
                temp = df["Transaction"][i] + " (" + df["Transaction"][i_n] + ")"
                df.loc[i, "Transaction"] = temp
                idx.append(i_n)

        result = df.drop(idx).reset_index()
        result = result.drop('index', axis=1)
        return result


    def bdo_es(self, file_path):
        doc = pymupdf.open(file_path)
        pg_cnt = doc.page_count
        
        if doc.authenticate(config01()[2]) == 6:
            result = pd.DataFrame()

            for pg in range(0, pg_cnt):
                rect_area = self.bound_box(doc.load_page(pg))
            
                if rect_area == None: continue

                df = tabula.read_pdf(file_path, password=config01()[2], pages=pg + 1, area=rect_area, pandas_options={'header': None}, output_format='dataframe', stream = True)[0]
                result = pd.concat([result, df], axis=0, ignore_index=True)
        
        result = result.rename(columns={0: "Sale Date", 1: "Post Date", 2: "Transaction", 3: "Amount"}).drop(index=[0, 1]).reset_index()
        result = result.drop('index', axis=1)
        result = self.transform(result)
        return result


    def transactions(self, bank):
        if bank == config01()[1]:
            es_dir = os.path.join(self.init_dir, config01()[0], config01()[1])

            for filename in os.listdir(es_dir):
                file_path = os.path.join(es_dir, filename)
                output = self.bdo_es(file_path)
            
        return output



if __name__ == '__main__':
    bnk = "BDO"

    test = bank_estatement().transactions(bnk)    
    print(test)
    