#!/usr/bin/env python
# coding: utf-8

# In[110]:



import pandas as pd
import numpy as np


# In[111]:


rw = pd.read_excel(r"C:\Users\3363\Desktop\Reports\KYC\Att.xlsx")
rw.head()


# In[112]:


rw['Date'] = pd.to_datetime(rw['Datetime']).dt.date
rw['Time'] = pd.to_datetime(rw['Datetime']).dt.time
rw.head()


# In[113]:


df = rw.sort_values(["Date", "ID"], ascending = (False, True))
df = df.drop(columns=['Datetime'])

df = df[['Date', 'ID', 'Name', 'Count','Time']]
df


# In[114]:


Result = df.pivot(index=['ID', 'Name'], columns='Count', values='Time')


# In[115]:


Result.to_excel(r"C:\Users\3363\Desktop\Akshay\KYC Attendence.xlsx")


# In[ ]:




