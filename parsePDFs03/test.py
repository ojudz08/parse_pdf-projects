import pymupdf, tabula
import os, sys, time
from config import config01
import pandas as pd
import numpy as np
from openpyxl import load_workbook


# initial parameters
init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))


def save_wb():


def bound_box(doc_page):
    try:
        top = doc_page.search_for("Sale Date")[0].irect.y1
        left = doc_page.search_for("Sale Date")[0].irect.x0
        bot = doc_page.search_for("For more details")[0].irect.y1 - 4
        right = doc_page.search_for("Amount")[0].irect.x1
    
        rect_area = [top, left, bot, right]
        return rect_area
    except:
        return None


def period(doc_page, word, rect_r):
    x0 = doc_page.search_for("(PHP)")[0].irect.x1
    rect = doc_page.search_for(word)[0].irect
    new_rect = pymupdf.Rect(x0, rect.y0, rect_r + 20, rect.y1)
    dt = doc_page.get_textbox(new_rect)
    return dt


def transform(df):
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


def bdo_es(file_path):
    doc = pymupdf.open(file_path)
    pg_cnt = doc.page_count
    
    if doc.authenticate(config01()[2]) == 6:
        result = pd.DataFrame()

        for pg in range(0, pg_cnt):
            doc_page = doc.load_page(pg)
            rect_area = bound_box(doc_page)

            if pg == 0:
                stmnt_dt = period(doc_page, "Statement Date", rect_area[3])
                due_dt = period(doc_page, "Payment Due Date", rect_area[3])
                
        
            if rect_area == None: continue

            df = tabula.read_pdf(file_path, password=config01()[2], pages=pg + 1, area=rect_area, pandas_options={'header': None}, output_format='dataframe', stream = True)[0]
            result = pd.concat([result, df], axis=0, ignore_index=True)
    
    result = result.rename(columns={0: "Sale Date", 1: "Post Date", 2: "Transaction", 3: "Amount"}).drop(index=[0, 1])   # drops first 2 rows
    result = result.reset_index().drop('index', axis=1)                                                                  # reset index and drop old index col
    result = transform(result)                                                                                           # remove nan values and transform
    temp_df = pd.DataFrame({"Statement Date": [stmnt_dt] * len(result), "Due Date": [due_dt] * len(result)})             
    
    result = pd.concat([result, temp_df], axis=1)
    return result


def transactions(bank):
    if bank == config01()[1]:
        es_dir = os.path.join(init_dir, config01()[0], config01()[1])

        for filename in os.listdir(es_dir):
            file_path = os.path.join(es_dir, filename)
            output = bdo_es(file_path)
        
    return output




if __name__ == '__main__':    
    bnk = "BDO"

    test = transactions(bnk)
    
    print(test)
    
    
    #mod_time = time.ctime(os.path.getmtime(file_path)) file_size = os.path.getsize(file_path)
    