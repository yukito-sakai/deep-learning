import numpy as np, pandas as pd
from pulp import *
from ortoolpy import addvars, addbinvars
import random
import datetime
import openpyxl

# ここでExcelのデータを取得
sheet1 = pd.read_excel('/Users/sakaiyukito/Downloads/python_study/まっちゃん面談.xlsx',index_col=0)
# Sheet1シートのカラムとインデックス
date = sheet1.index
person = sheet1.columns
# 確認用にコピー
sheet1_show = sheet1

print(sheet1)