# This file contains the code for the job application (Naukri.com) of Data Analyst at Cointab

# Importing Packages
import numpy as np
import pandas as pd
import xlsxwriter as xw


# Reading the data from given data files
XOR = pd.read_excel('Assignment details/Company X - Order Report.xlsx')                          # Company X - order report workbook
X_SKU = pd.read_excel('Assignment details/Company X - SKU Master.xlsx').to_dict('list')          # Company X - SKU Master Workbook
PinZones = pd.read_excel('Assignment details/Company X - Pincode Zones.xlsx').to_dict('list')    # Company X - Pincode Zones
CCI = pd.read_excel('Assignment details/Courier Company - Invoice.xlsx')                         # Courier Company - Invoice Workbook
CCR = pd.read_excel('Assignment details/Courier Company - Rates.xlsx')                           # Courier Company - Rates Workbook


# Creating two xlsx workbooks for two outputs
output1 = xw.Workbook('Order_info.xlsx')
output2 = xw.Workbook('Summary_Report.xlsx')


# Adding the columns to the output1
worksheet1 = output1.add_worksheet()
col_names1 = ['Order Id', 'AWB Number', 'Total weight as per X (KG)', 'Weight slab as per X (KG)',\
    'Total weight as per Courier Company (KG)', 'Weight slab charged by Courier Company (KG)', 'Delivery Zone as per X',\
        'Delivery Zone charged by Courier Company', 'Expected Charge as per X (Rs.)', 'Charges Billed by Courier Company (Rs.)',\
            'Difference Between Expected Charges and Billed Charges (Rs.)']

for i in range(0, len(col_names1)):
    worksheet1.write(0, i, col_names1[i])                           # Writing column names in the worksheet 1
    if i == 1 or i == 0:
        worksheet1.set_column(i, i, len(col_names1[i])+5)           # Scaling of first and second column (Order Id and AWB number)
    else:
        worksheet1.set_column(i, i, len(col_names1[i]))             # Scaling of column width according to size of column names

# Similarly for output2
worksheet2 = output2.add_worksheet()
row_names2 = ['Total orders where X has been correctly charged', 'Total orders where X has been overcharged',\
     'Total orders where X has been undercharged']
col_names2 = ['Count', 'Amount (Rs.)']
for i in range(0, len(row_names2)):
    if i == 0 or i == 1:
        worksheet2.write(0, i+1, col_names2[i])                     # Writing column names in worksheet 2
        worksheet2.set_column(i+1, i+1, len(col_names2[i])+5)       # Setting column width
        worksheet2.write(i+1, 0, row_names2[i])                     # Writing row names in worksheet 2
    else:
        worksheet2.write(i+1, 0, row_names2[i])                     
    worksheet2.set_column(0, 0, len(row_names2[0]))


# Order level calculations
order_no = list(set(XOR['ExternOrderNo']))                          # Extracting the order ids
XOR_new = XOR.groupby('ExternOrderNo')                              # Grouping by order ids
for i in range(0, len(order_no)):
    order_group = XOR_new.get_group(order_no[i]).reset_index()                                  # Extracting all entries of particular order Id

    # Extracting all the necessary information from the provided data
    AWB_number = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['AWB Code'][0]           
    cust_pin = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['Customer Pincode'][0]
    zone_CC = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['Zone'][0]
    zone_X = PinZones['Zone'][PinZones['Customer Pincode'].index(cust_pin)]
    weight_slabs_X = CCR.loc[CCR['Zone'] == zone_X.upper()].reset_index()['Weight Slabs'][0]
    fwd_fixed_charge = CCR.loc[CCR['Weight Slabs'] == weight_slabs_X].reset_index()['Forward Fixed Charge'][0]
    RTO_fixed_charge = CCR.loc[CCR['Weight Slabs'] == weight_slabs_X].reset_index()['RTO Fixed Charge'][0]
    fwd_add_charge = CCR.loc[CCR['Weight Slabs'] == weight_slabs_X].reset_index()['Forward Additional Weight Slab Charge'][0]
    RTO_add_charge = CCR.loc[CCR['Weight Slabs'] == weight_slabs_X].reset_index()['RTO Additional Weight Slab Charge'][0]
    shipment_type = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['Type of Shipment'][0]
    charges_CC = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['Billing Amount (Rs.)'][0]
    total_weight_CC = CCI.loc[CCI['Order ID'] == order_no[i]].reset_index()['Charged Weight'][0]
    weight_slabs_CC = CCR.loc[CCR['Zone'] == zone_CC.upper()].reset_index()['Weight Slabs'][0]

    payment_mode = order_group['Payment Mode'][0]

    # Finding total weight and total amount of the order
    total_amount_order = 0
    total_weight_order_X = 0
    for j in range(0, len(order_group)):
        SKU = order_group['SKU'][j]
        WeightCorrSKU = X_SKU['Weight (g)'][X_SKU['SKU'].index(SKU)]
        total_weight_order_X += order_group['Order Qty'][j] * WeightCorrSKU
        total_amount_order += order_group['Order Qty'][j] * order_group['Item Price(Per Qty.)'][j]

    # Finding the applicable weight and calculating charges (fwd + rto)
    app_weight_X = 0
    total_charge = 0
    k = 0
    while total_weight_order_X/1000 > app_weight_X:
        app_weight_X += weight_slabs_X
        if shipment_type == 'Forward charges':
            if k == 0:
                total_charge += fwd_fixed_charge
            else:
                total_charge += fwd_add_charge
        else:
            if k == 0:
                total_charge += fwd_fixed_charge + RTO_fixed_charge
            else:
                total_charge += fwd_add_charge + RTO_add_charge
        k += 1
    
    # Calculating COD charges
    if payment_mode == 'COD' and total_amount_order <= 300:
        total_charge += 15

    elif payment_mode == 'COD' and total_amount_order > 300:
        total_charge += (5/100)*total_amount_order

    else:
        total_charge += 0

    # Calculating the difference between the bill charged by Courier Company and Expected charges
    difference = total_charge - charges_CC
    
    # Writing all the calculated values in output sheet (Order_info.xlsx)
    cal_values = [order_no[i], AWB_number, total_weight_order_X/1000, weight_slabs_X, total_weight_CC, weight_slabs_CC,\
         zone_X.upper(), zone_CC.upper(), total_charge, charges_CC, difference]
    for l in range(0, len(cal_values)):
        worksheet1.write(i+1, l, cal_values[l])

output1.close()


# Preparing Summary Report
order_info = pd.read_excel('Order_info.xlsx')

# Calculating overcharging counts and amount
overcharging_counts = len(order_info.loc[order_info['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0])
print('Overcharging count:', overcharging_counts)
overcharging_amt = sum(order_info.loc[order_info['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0]['Difference Between Expected Charges and Billed Charges (Rs.)'])
print('Overcharging amount:', overcharging_amt)

# Calculating Undercharging counts and amounts
undercharging_counts = len(order_info.loc[order_info['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0])
print('Undercharging count:', undercharging_counts)
undercharging_amt = sum(order_info.loc[order_info['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0]['Difference Between Expected Charges and Billed Charges (Rs.)'])
print('Undercharging amount:', undercharging_amt)

# Calculating the counts of orders charged correctly and total invoice by Courier Company
correctly_charged = len(order_info.loc[order_info['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0])
print('Correctly charged:', correctly_charged)

# This is the total of invoice amount billed by Courier Company
total_invoice_amt = sum(order_info['Charges Billed by Courier Company (Rs.)'])
print('Total of charges billed by Courier Company', total_invoice_amt)
total_expected_amt = sum(order_info['Expected Charge as per X (Rs.)'])
print('Total expected charges:', total_expected_amt)

# Writing all the summary in output file (Summary_Report.xlsx)
summary = [correctly_charged, total_invoice_amt, overcharging_counts, overcharging_amt, undercharging_counts, undercharging_amt]
d = 0
for p in range(1, 4):
    for o in range(1, 3):
        worksheet2.write(p, o, summary[d])
        d += 1

output2.close()