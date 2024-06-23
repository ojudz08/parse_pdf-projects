import pymupdf, tabula
import os, sys, time
from config import config_cred
import pandas as pd
from openpyxl import load_workbook


# initial parameters
init_dir = os.path.abspath(os.path.dirname(sys.argv[0]))


def save_as():
    with pd.ExcelWriter(output_file) as writer:
        idxRet.to_excel(writer, sheet_name="index_returns", index=False)
        comm.to_excel(writer, sheet_name="commodities", index=False)
        curr.to_excel(writer, sheet_name="currencies", index=False)
        rtSp.to_excel(writer, sheet_name="rates_spreads", index=False)
    pass

def bound_box(doc_page):
    try:
        top = doc_page.search_for("Sale Date")[0].irect[3]
        left = doc_page.search_for("Sale Date")[0].irect[0]
        bot = doc_page.search_for("For more details")[0].irect[3] - 4
        right = doc_page.search_for("Amount")[0].irect[2]
    
        rect_area = [top, left, bot, right]
        return rect_area
    except:
        return None


def bdo_es(es_dir):

    for filename in os.listdir(es_dir):
        file_path = os.path.join(es_dir, filename)
        doc = pymupdf.open(file_path)
        pg_cnt = doc.page_count
        
        if doc.authenticate(config_cred()[1]) != 6: break
        
        result = pd.DataFrame()
        for pg in range(0, pg_cnt):
            rect_area = bound_box(doc.load_page(pg))
            
            if rect_area == None: continue

            df = tabula.read_pdf(file_path,
                                 password = config_cred()[1],
                                 pages = pg + 1,
                                 area = rect_area,
                                 pandas_options = {'header': None},
                                 output_format = 'dataframe',
                                 stream = True)[0]
            result = pd.concat([result, df], axis=0, ignore_index=True)
        
        result = result.rename(columns={0: "Sale Date", 1: "Post Date", 2: "Transaction Details", 3: "Amount"}).drop(index=[0, 1]).reset_index()
        result = result.drop('index', axis=1)
        return result


def transactions(bank):
    if bank == config_cred()[0]:
        es_dir = os.path.join(init_dir, "e-statements\BDO")
        output = bdo_es(es_dir)

    for filename in os.listdir(es_dir):
        summary_path = os.path.join(init_dir, "Transactions.xlsx")
        #wb = load_workbook(summary_path)
        #writer = pd.ExcelWriter(summary_path, engine='openpyxl')
        #reader = pd.read_excel(summary_path, sheet_name="BDO", header=True)
        

    return output




if __name__ == '__main__':    
    bnk = "BDO"

    test = transactions(bnk)
    
    print(test)
    
    #mod_time = time.ctime(os.path.getmtime(file_path)) file_size = os.path.getsize(file_path)