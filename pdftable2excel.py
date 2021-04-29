import pdfplumber
from pdfplumber import utils
import pandas as pd
import numpy as np
import shutil
import copy
import glob
import os
file_type = 'xlsx'
def Save_table(table, table_path, tag=file_type):
    """
    save the table to csv / excel
    :param table:
    :param table_path:
    :return:
    """
    df = pd.DataFrame(table, columns=None)
    # path = os.path.join(save_dir, 'page-{}_table-{}.xlsx'.format(i, j))
    # path = r'D:\work\浦发理财产品说明\page_{}table_{}.xlsx'.format(i,j)
    # j += 1
    drop_item = []
    for index, row in df.iteritems():
        vals = df[index].values.tolist()
        flag = 0
        for val in vals:
            if val:
                flag = 1
                break
        if flag == 0:
            drop_item.append(index)
    if drop_item:
        df = df.drop(drop_item, axis=1)
    if not df.empty:
        if tag == 'xlsx':
            df.to_excel(table_path, index=False, header=False)
        if tag == 'csv':
            df.to_csv(table_path, index=False, header=False)
pdf_name = "table.pdf"
with pdfplumber.open(pdf_name) as pdf:
    page_text = []
    for index, e_page in enumerate(pdf.pages):
        # e_page=pdf.pages[index]
        table = e_page.extract_tables()[0]
        table_path = "save_table.xlsx"
        df = pd.DataFrame(table, columns=None)
        df.to_excel(table_path, index=False, header=False)
        # Save_table(table, table_path, tag=file_type)
        print(table)