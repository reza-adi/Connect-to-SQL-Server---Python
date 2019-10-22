# -*- coding: utf-8 -*-
"""
Created on Tue Oct 15 10:45:59 2019

@author: rezaa
"""

import pandas as pd
import pyodbc
import numpy as np
import xlsxwriter
from openpyxl import load_workbook

#Install SQL Server Express Edition first
server = 'YourPCName\SQLEXPRESS'
database = 'Northwind'
username = 'USERNAME'
password = 'PASSWORD'
conn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
writer = pd.ExcelWriter('test.xlsx',engine='xlsxwriter')

'*****************Customers Table******************'
query = "SELECT * from Customers"
first = pd.read_sql(query,conn)
df = pd.DataFrame(first)

df.to_excel(writer,sheet_name='Customers',startrow=0 , startcol=0) 
print('Customers Downloaded.')

writer.save()