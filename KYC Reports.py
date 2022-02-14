#!/usr/bin/env python
# coding: utf-8

# In[16]:


import pandas as pd
import numpy as np


# In[17]:


rw = pd.read_excel(r"C:\Users\3363\Desktop\KYC.xlsx")

rw.head()


# In[18]:


rw['Date'] = pd.to_datetime(rw['Timestamp'], errors='coerce')

rw['Date'] = rw['Date'].dt.strftime('%d-%b')

rw.head()


# In[19]:


Result =pd.pivot_table(data= rw, index = ['EMP ID'],columns= ['Date'],values = 'PROCESS', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result.head()


# In[20]:


Result1 =pd.pivot_table(data= rw, index = ['EMP ID'],columns= ['PROCESS'],values = 'Date', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result1.head()


# In[ ]:





# In[21]:




writer = pd.ExcelWriter(r"C:\Users\3363\Desktop\Akshay\KYCReport.xlsx", engine='xlsxwriter')


# In[22]:




Result.to_excel(writer, sheet_name='Agent wise')
Result1.to_excel(writer, sheet_name='Process Wise')


writer.save()

