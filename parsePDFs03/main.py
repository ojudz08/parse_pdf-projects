"""
    Author: Ojelle Rogero
    Created on: June 23, 2024
    About:
        Parses the transactions of a bank statement
"""

from tkinter import filedialog
import pymupdf, tabula
import os, sys
from config import config01
import pandas as pd
import numpy as np



class bank_estatement():

    def __init__(self):
        self.init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))


    def open(self):
        open_file = filedialog.askdirectory(initialdir=self.init_dir)
        return open_file


    def save(self, out_path, df):
        temp = pd.read_excel(out_path, sheet_name='Summary', engine='openpyxl')
        last_row = len(temp) + 1
        with pd.ExcelWriter(out_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name='Summary', startrow=last_row, index=False, header=False)


    def check(self, file_path, init_date):
        doc = pymupdf.open(file_path)
        if doc.authenticate(config01()[1]) == 6:
            doc_page = doc.load_page(0)
            rect_area = self.bound_box(doc_page)
            temp = self.period(doc_page, "Statement Date", rect_area[3])
            if temp in init_date: 
                return True
            else: 
                return False


    def bound_box(self, doc_page):
        try:
            top = doc_page.search_for("Sale Date")[0].irect.y1
            left = doc_page.search_for("Sale Date")[0].irect.x0
            right = doc_page.search_for("Amount")[0].irect.x1
            bot1 = doc_page.search_for("For more details")[0].irect.y1 - 4
            bot2 = doc_page.search_for("BDO Unibank, Inc. is regulated by")[0].irect.y1 - 4

            if bot1 < bot2: bot = bot1
            else: bot = bot2
            rect_area = [top, left, bot, right]

            return rect_area
        except:
            return None
    

    def period(self, doc_page, word, rect_r):
        x0 = doc_page.search_for("(PHP)")[0].irect.x1
        rect = doc_page.search_for(word)[0].irect
        new_rect = pymupdf.Rect(x0, rect.y0, rect_r + 20, rect.y1)
        dt = doc_page.get_textbox(new_rect)
        return dt


    def transform(self, df):
        temp_df = df.drop(index=[0, 1]).reset_index().drop('index', axis=1)
        idx = []

        for i in range(0, len(temp_df)):
            i_n = i + 1
            try:
                if (np.isnan(temp_df["Sale Date"][i_n]) and np.isnan(temp_df["Post Date"][i_n])) and temp_df["Transaction"][i_n] not in ["SUBTOTAL", "TOTAL"]:
                    temp = temp_df["Transaction"][i] + " (" + temp_df["Transaction"][i_n] + ")"
                    temp_df.loc[i, "Transaction"] = temp
                    idx.append(i_n)
            except:
                pass
        
        idx.append(len(temp_df) - 1)
        temp_df = temp_df.drop(idx).reset_index().drop('index', axis=1)                                         # drops na rows and row -2
        total_df = pd.DataFrame({"Total": temp_df.tail(1)['Amount']}).reset_index().drop('index', axis=1)       # concat Total
        result = pd.concat([temp_df, total_df], axis=1).drop(index=[temp_df.tail(1).index[0]])                  # drop last row -- Total

        return result


    def bdo_es(self, file_path):
        doc = pymupdf.open(file_path)
        pg_cnt = doc.page_count
        
        if doc.authenticate(config01()[1]) == 6:
            temp = pd.DataFrame()

            for pg in range(0, pg_cnt):
                doc_page = doc.load_page(pg)
                rect_area = self.bound_box(doc_page)

                if pg == 0:
                    stmnt_dt, due_dt = self.period(doc_page, "Statement Date", rect_area[3]), self.period(doc_page, "Payment Due Date", rect_area[3])
                    dt_df = pd.DataFrame({"Statement Date": [stmnt_dt], "Due Date": due_dt})

                if rect_area == None: continue

                df = tabula.read_pdf(file_path, password=config01()[1], pages=pg + 1, area=rect_area, pandas_options={'header': None}, output_format='dataframe', stream = True)[0]
                temp = pd.concat([temp, df], axis=0, ignore_index=True)
        
        temp = temp.rename(columns={0: "Sale Date", 1: "Post Date", 2: "Transaction", 3: "Amount"})
        temp = self.transform(temp)
        result = pd.concat([temp, dt_df], axis=1)
        return result


    def transactions(self, bank):
        output_file = os.path.join(self.init_dir, "Transactions.xlsx")
        temp1 = pd.read_excel(output_file, sheet_name='Summary', engine='openpyxl')
        init_dt = temp1["Statement Date"].dropna().values

        if bank == config01()[0]:
            es_dir = os.path.realpath(self.open())

            for filename in os.listdir(es_dir):
                if filename[-4:] == ".pdf":
                    file_path = os.path.join(es_dir, filename)
                    
                    if self.check(file_path, init_dt) == True: pass
                    else:
                        output = self.bdo_es(file_path)
                        self.save(output_file, output)

        temp2 = pd.read_excel(output_file, sheet_name='Summary', engine='openpyxl')
        fin_dt = temp2["Statement Date"].dropna().values

        if len(init_dt) < len(fin_dt) and init_dt in fin_dt:
            return "New period added"
        elif len(init_dt) == len(fin_dt) and len(set(init_dt).intersection(fin_dt)) == len(fin_dt):
            return "No changes"


if __name__ == '__main__':
    bnk = "BDO"

    result = bank_estatement().transactions(bnk)    
    print(result) 