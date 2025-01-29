# -*- coding: utf-8 -*-
"""
Created on Tue Jan 28 18:17:42 2025

@author: harsh
"""
import pandas as pd
import os

#reading excel file from local
data = pd.read_excel(r'C:\Users\harsh\OneDrive\Desktop\orders_data.xlsx', sheet_name='Sheet1')
#remove null values
data.isnull().sum() 
print(data.head(10))

data['order_date_clean'] = data['order_date'].str.replace(r'^\w{3},\s', '', regex=True) #remove Weekday 
data['order_date_clean'] = data['order_date_clean'].str.replace(r'\sIST$', '', regex=True)  # Remove " IST"

#Standardize month abbreviations (Sept → Sep) as i was not able to convert into datetime with fixed format
data['order_date_clean'] = data['order_date_clean'].str.replace(r'Sept', 'Sep', regex=True)
print(data['order_date_clean'])

data['order_date'] = pd.to_datetime(data['order_date_clean'], format='%d %b, %Y, %I:%M %p')

print(data['order_date'])


data['day_of_week'] = data['order_date'].dt.strftime('%A')   # Full weekday name
data['time'] = data['order_date'].dt.strftime('%I:%M %p')    # Time in 12-hour format
data['month'] = data['order_date'].dt.strftime('%B')         # Full month name
data['year'] = data['order_date'].dt.year 

print(data[['order_date', 'day_of_week', 'time', 'month', 'year']])

#remove rupees from field
data['item_total'] = data['item_total'].str.replace(r'₹','',regex=True)
print(data['item_total'])
data['shipping_fee'] = data['shipping_fee'].str.replace(r'₹','',regex=True)

#only first letter to be capitalised and removed trailling comma : CHANDIGARH, -> Chandigarh
data['ship_city'] = data['ship_city'].str.capitalize().str.rstrip(',')
data['ship_state'] = data['ship_state'].str.capitalize()

data.to_excel('cleaned_orders.xlsx', index=False)

# Define the filename
filename = "cleaned_orders.xlsx"
#if want to save at specific location
#filename = "C:/Users/YourUsername/Documents/cleaned_orders.xlsx"

# Save the DataFrame to Excel
data.to_excel(filename, index=False)

# Check if the file exists
if os.path.exists(filename):
    print(f"✅ File '{filename}' has been successfully created!")
else:
    print(f"❌ File '{filename}' was NOT created. Check file path or permissions.")

#location of the file
print("File saved at:", os.getcwd())

