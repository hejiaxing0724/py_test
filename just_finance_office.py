# -*- coding: utf-8 -*-
import pandas as pd
from openpyxl import load_workbook

def getExcel(filefullpath):
    df = pd.read_excel(filefullpath)
    return df

def getExcelWriter(filefullpath):
    excelWriter = pd.ExcelWriter(filefullpath, engine='openpyxl')
    return excelWriter

def addSheet(df,excelWriter,sheetName):
    # book=load_workbook(excelWriter.path)
    # excelWriter.book = book
    df.to_excel(excel_writer=excelWriter,sheet_name=sheetName,index=None)
    excelWriter.close()

def write2Excel(df,excelWriter):
    df = df.sort_values(by=['项目代码'],ascending=True)
    project_code_list = df['项目代码'].unique()
    project_code_list = sorted(project_code_list)
    addSheet(df, excelWriter, '源数据')
    df.to_excel(excel_writer=excelWriter, sheet_name='源数据', index=None)
    for project_code in project_code_list:
        new_df = df.loc[df['项目代码'].isin([project_code])]
        new_df=  new_df.sort_values(by=['编号','日','序号'],ascending=True)
        addSheet(new_df,excelWriter,project_code)
        new_df.to_excel(excel_writer=excelWriter,sheet_name=project_code,index=None)
        excelWriter.save()
    excelWriter.close()


if __name__ == '__main__':
    filefullpath = input("请输入要处理的Excel的所在目录：")
    df=getExcel(filefullpath)
    excelWriter=getExcelWriter(filefullpath)
    write2Excel(df,excelWriter)








