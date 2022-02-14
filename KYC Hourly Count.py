#!/usr/bin/env python
# coding: utf-8

# In[127]:



import pandas as pd
import numpy as np


# In[128]:


rw = pd.read_excel(r"C:\Users\3363\Desktop\Reports\kycdata.xlsx")
rw.head()


# In[129]:


rw['Time'] = pd.to_datetime(rw['Timestamp']).dt.time
rw['Hour'] = rw['Timestamp'].dt.hour


# In[130]:


rw.head()


# In[131]:


import datetime


# In[132]:


def test(x):
    if x == 7:
        return "7 AM - 8 AM"
    if x == 8:
        return "8 AM - 9 AM"
    if x == 9:
        return "9 AM - 10 AM"
    if x == 10:
        return "10 AM - 11 AM"
    if x == 11:
        return "11 AM - 12 PM"
    if x == 12:
        return "12 PM - 1 PM"
    if x == 13:
        return "1 PM - 2 PM"
    if x == 14:
        return "2 PM - 3 PM"
    if x == 15:
        return "3 PM - 4 PM"
    if x == 16:
        return "4 PM - 5 PM"
    if x == 17:
        return "5 PM - 6 PM"
    if x == 18:
        return "6 PM - 7 PM"
    if x == 19:
        return "7 PM - 8 PM"
    if x == 20:
        return "8 PM - 9 PM"
    else:
        return "Time exceeds"
    


# In[133]:


rw['Timing'] = rw['Hour'].apply(lambda x: test(x))
rw.head()


# In[134]:


Sonal =  rw.loc[rw['Shift Timing'] == "07:00 to 04:00 pm"]
shivani =  rw.loc[rw['Shift Timing'] == "09:00 to 06:00 pm"]
hema =  rw.loc[rw['Shift Timing'] == "12:00 to 09:00 pm"]


# In[135]:


Result1 =pd.pivot_table(data= Sonal, index = ['Agent Name'],columns= ['Hour'], values = 'URL', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result1 = Result1.rename(columns={6: '6:00 to 7:00', 7: '7:00 to 8:00', 8: '8:00 to 9:00', 9: '9:00 to 10:00', 10: '10:00 to 11:00', 11: '11:00 to 12:00', 12: '12:00 to 1:00', 13: '1:00 to 2:00', 14: '2:00 to 3:00', 15: '3:00 to 4:00', 16: '4:00 to 5:00', 17: '5:00 to 6:00',18:'6:00 to 7:00' })

Result1.head()


# In[136]:


Result2 =pd.pivot_table(data= shivani, index = ['Agent Name'],columns= ['Hour'], values = 'URL', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result2 = Result2.rename(columns={6: '6:00 to 7:00', 7: '7:00 to 8:00', 8: '8:00 to 9:00', 9: '9:00 to 10:00', 10: '10:00 to 11:00', 11: '11:00 to 12:00', 12: '12:00 to 1:00', 13: '1:00 to 2:00', 14: '2:00 to 3:00', 15: '3:00 to 4:00', 16: '4:00 to 5:00', 17: '5:00 to 6:00',18:'6:00 to 7:00' })

Result2.head()


# In[137]:


Result3 =pd.pivot_table(data= hema, index = ['Agent Name'],columns= ['Hour'], values = 'URL', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result3 = Result3.rename(columns={6: '6:00 to 7:00', 7: '7:00 to 8:00', 8: '8:00 to 9:00', 9: '9:00 to 10:00', 10: '10:00 to 11:00', 11: '11:00 to 12:00', 12: '12:00 to 1:00', 13: '1:00 to 2:00', 14: '2:00 to 3:00', 15: '3:00 to 4:00', 16: '4:00 to 5:00', 17: '5:00 to 6:00',18:'6:00 to 7:00' })

Result3.head()


# In[138]:




writer = pd.ExcelWriter(r"C:\Users\3363\Desktop\Akshay\KYC Hourlycount.xlsx", engine='xlsxwriter')


# In[139]:


Result1.to_excel(writer, sheet_name='Sonal')
Result2.to_excel(writer, sheet_name='Shivani')
Result3.to_excel(writer, sheet_name='Hema')




writer.save()


# In[140]:


rw['Hour'].value_counts()


# In[ ]:




