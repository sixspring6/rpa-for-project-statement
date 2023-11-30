#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlwings as xw
import pandas as pd


app = xw.App(visible=False)  # 엑셀 프로그램을 화면에 표시하지 않음
wb = xw.Book('c:/칼럼실습.xlsx')
sht = wb.sheets[0]
df = sht.range('a1').options(pd.DataFrame,expand='table',Header=1,index=False).value
df


# In[2]:


df = df.drop('B', axis=1)
df


# In[3]:


대상 = ['A','I','J']
df = df[대상]
df


# In[4]:


df.loc[:, 'new'] = '임의'
df


# In[5]:


df.to_excel('c:/result.xlsx', index=False)
df


# In[ ]:




