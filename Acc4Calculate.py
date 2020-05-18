import numpy as np
import pandas as pd
import os

def readDataSet(field=True):
    if field == True:
        cur_path = os.getcwd()
        dataset_path = cur_path + '\现场数据\超短期单点数据集'
        files = os.listdir(dataset_path)
        print(files)
        if not files:
            print('There is no Dataset in ...\现场数据\超短期单点数据集, please check it!')
            return
        df_all = pd.DataFrame()
        for file in files:
            df = pd.read_excel(dataset_path + '\\' + file)
            df_all = pd.concat([df_all, df])
    else:



def calc():
    pass

def write2Excel():
    pass

if __name__ == "__main__":
    readDataSet()
    calc()
    write2Excel()
