from flask import Flask, request, url_for, render_template, redirect, send_file
import pandas as pd
import numpy as np
from openpyxl import Workbook
import os
import tempfile
from datetime import datetime
import re
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def process_step1(file_path):
    xls = pd.ExcelFile(file_path)
    
    df1 = pd.read_excel(xls, sheet_name="Outward supply")
    df2 = pd.read_excel(xls, sheet_name="Doc. Series")
    df3 = pd.read_excel(xls, sheet_name="Amendments(Invoices)")
    df4 = pd.read_excel(xls, sheet_name="Debit&CreditNotes")
    df5 = pd.read_excel(xls, sheet_name="Amendments (CDN)")
    df6 = pd.read_excel(xls, sheet_name="Advances")
    df7 = pd.read_excel(xls, sheet_name="Amendment(Advances)")

    # -------------------------------------------* Processing of Doc. Series Sheet *----------------------------------------------------
    
    # It will keep only those rows where not all elements in the row are either NaN or empty strings
    filtered_df = df2.apply(lambda row: not all(pd.isna(row) | (row == "")), axis=1)
    df2 = df2[filtered_df]
    
    # Delete first column
    df2 = df2.drop(df1.columns[0], axis=1)
    
    newvals = pd.to_datetime(df2.iloc[:, 1], errors="coerce").dt.date
    df2[df2.columns[1]] = newvals
    
    new_header = df2.iloc[5]  # Get the 6th row as the new header
    df2 = df2[6:]  # Remove the first 5 rows
    df2.columns = new_header  # Set the new header
    df2.reset_index(drop=True, inplace=True)
    
    # Remove all columns after the 2nd column
    df2 = df2.iloc[:, :2]
    # Remove all rows after the 2nd row
    df2 = df2.iloc[:2, :]
    
    df2.insert(2, 'Column3', df2.iloc[:, 1])
    #  Transposes the DataFrame
    df2 = df2.T
    
    # Promote the second row to be the column headers
    df2.columns = df2.iloc[0]

    # Remove the original first row (which is now redundant)
    df2 = df2.iloc[1:]
    
    
    # -------------------------------------------* Processing of Amendment(Advances) Sheet *----------------------------------------------------
    
    df7 = df7.iloc[5:]
    # Step 3: Set the 6th row as the header.
    header_row = df7.iloc[0]
    df7 = df7[1:]
    df7.columns = header_row
    
    # to select data from 1st column to 29th column 
    df7 = df7.iloc[:, 1:27]
    
    # To check if duplicate entries of any Origional document number is present
    df7['Count'] = df7.groupby('Original document number')['Original document number'].transform('count')
    df7['Duplicate Amendment(Advances) for original'] = df7['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df7.drop(columns=['Count'], inplace=True) # Remove Count column
    
    # To check if duplicate entries of any Revised receipt voucher number is present
    df7['Count'] = df7.groupby('Revised receipt voucher number')['Revised receipt voucher number'].transform('count')
    df7['Duplicate Amendment(Advances) for revised voucher'] = df7['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df7.drop(columns=['Count'], inplace=True) # Remove Count column
    
    # Function to check status of Recipient
    def comp_status_of_recipient_check(row):
        if row['Original status of recipient'] == row['Revised status of recipient']:
            return 'Match'
        elif pd.isna(row['Revised status of recipient']):
            return 'It should not be blank'
        elif row['Revised status of recipient'] == 'Unregistered':
            return 'Changes made - Supply made to unregistered GSTN should be blank'
        else:
            return 'Changes made - Supply made to registered GSTN should not be blank'

    df7['Comp. Status Of Recipient Check'] = df7.apply(comp_status_of_recipient_check, axis=1)
    
    # Function to check Type of Supply
    def comp_type_of_supply_check(row):
        if row['Original type of supply'] == row['Revised type of supply']:
            return 'Match'
        elif pd.isna(row['Revised type of supply']):
            return 'It should not be blank'
        elif 'SEZ supplies without' in str(row['Revised type of supply']) or 'SEZ without' in str(row['Revised type of supply']):
            return 'Changes made - It is a Zero rated supply without payment all tax columns should be blank'
        elif 'SEZ supplies with' in str(row['Revised type of supply']) or 'SEZ with' in str(row['Revised type of supply']):
            return 'Changes made - It is an Zero rated supply CGST+SGST should be blank'
        elif 'Export with' in str(row['Revised type of supply']) or 'Export supplies with' in str(row['Revised type of supply']):
            return 'Changes made - It is an Zero rated supply CGST+SGST should be blank'
        elif 'Exempt' in str(row['Revised type of supply']):
            return 'Changes made - It is an exempt supply tax value should be blank'
        elif 'Regular' in str(row['Revised type of supply']):
            return 'Correct'
        elif 'Export without' in str(row['Revised type of supply']) or 'Export supplies without' in str(row['Revised type of supply']):
            return 'Changes made - It is a Zero rated supply without payment all tax columns should be blank'
        else:
            return 'Prima facie it observes that it is other than Regular supply'

    df7['Comp. Type OF Supply Check'] = df7.apply(comp_type_of_supply_check, axis=1)
    
    # Function to check Taxability
    def comp_taxability_check(row):
        if row['Original taxability'] == row['Revised taxability']:
            return 'Match'
        elif row['Revised taxability'] == 'Exempt':
            return 'Changes made - Exempt supply made by the Company which attract reversal under rule 42 & 43 also tax amount should be zero'
        elif row['Revised taxability'] == 'non GST':
            return 'Changes made - It is a No GST supply hence tax amount should be zero'
        elif pd.isna(row['Revised taxability']):
            return 'It should not be blank'
        elif row['Revised taxability'] == 'Taxable':
            return 'Changes made - Normal taxable supply'
        else:
            return "Didn't match"

    df7['Comp. Taxability Check'] = df7.apply(comp_taxability_check, axis=1)
    
    # Function to check Origional Document number and Revised receipt voucher number
    def comp_document_number_check(row):
        if row['Original document number'] == '0':
            return 'It should not be blank'
        elif row['Original document number'] == row['Revised receipt voucher number']:
            return 'Match'
        else:
            return "Didn't match / Need to check the Invoice copy"

    df7['Comp. Document number Check'] = df7.apply(comp_document_number_check, axis=1)
    
    # Function to check Invoice Number
    def invoice_number_check(row):
        if len(str(row['Revised receipt voucher number'])) <= 16:
            return 'Correct'
        else:
            return 'Need to check the Invoice copy'

    df7['Invoice no. Check'] = df7.apply(invoice_number_check, axis=1)
    
    # Function to check Document Date
    def document_date_check(row):
        original_date = str(row['Original document date'])
        revised_date = str(row['Revised receipt voucher date'])
        
        if original_date == revised_date:
            return 'Match'
        elif original_date == '0':
            return 'It should not be blank'
        else:
            return "Didn't match"

    df7['Comp. Document Date Check'] = df7.apply(document_date_check, axis=1)
    
    # To extract first two Numbers from String
    df7['Revised Recipients GSTIN - Copy'] = df7.iloc[:, 12].str.extract(r'(\d{2})')
    df7['Revised place Of Supply - Copy'] = df7['Revised place Of Supply'].str.extract(r'(\d{2})')
    
    # Function to check Place of Supply and Recipient GSTIN 
    def pos_recipient_check(row):
        pos_copy = row['Revised place Of Supply - Copy']
        gstin_copy = row['Revised Recipients GSTIN - Copy']
        
        if pos_copy == gstin_copy:
            return "Match"
        elif pd.isna(pos_copy):
            if pd.isna(gstin_copy):
                return "Place of Supply and Recipients GSTIN should not be blank"
        elif pd.isna(gstin_copy):
                return "Recipients GSTIN should not be blank"
        else:
            return "Place of Supply and Recipients GSTIN need to check"

    df7['POS & Recipient check'] = df7.apply(pos_recipient_check, axis=1)
    
    df7 = df7.drop(columns=["Revised place Of Supply - Copy", "Revised Recipients GSTIN - Copy"]) #Remove specified columnns
    
    # Function to check Revised HSN
    def hsn_check(row):
        revised_hsn = row['Revised HSN']
        
        # Check if the value is NaN
        if pd.notna(revised_hsn):
            revised_hsn = int(revised_hsn)
            
            if revised_hsn == 99999999:
                return 'Need to mention correct HSN'
            elif revised_hsn == 999:
                return 'Correct'
            elif revised_hsn == 0:
                return 'HSN should not be blank'
            else:
                return 'Need to mention correct HSN'
        else:
            return 'Revised HSN is NaN'  # Handle NaN values

    df7['HSN check'] = df7.apply(hsn_check, axis=1)
    
    # Function to check and convert Revised GST Rate
    def gst_rate_check(row):
        revised_rate = row['Revised GST rate(%)']
        
        if revised_rate == 0.28:
            return 28
        elif revised_rate == 28:
            return 28
        elif revised_rate == 0.18:
            return 18
        elif revised_rate == 18:
            return 18
        elif revised_rate == 0.12:
            return 12
        elif revised_rate == 12:
            return 12
        elif revised_rate == 0.05:
            return 5
        elif revised_rate == 5:
            return 5
        elif revised_rate == 0.025:
            return 2.5
        elif revised_rate == 2.5:
            return 2.5
        elif revised_rate == 0.01:
            return 0.1
        elif revised_rate == 0.03:
            return 3
        elif revised_rate == 3:
            return 3
        elif pd.isna(revised_rate):
            return "Should not be blank"
        else:
            return "Need to mention correct GST Rate"

    df7['GST Rate check'] = df7.apply(gst_rate_check, axis=1)
    
    # Function to Find GST Difference
    def compute_gst_difference(row):
        revised_taxable_value = row.iloc[20]
        revised_igst = row.iloc[21]
        revised_cgst = row.iloc[22]
        revised_sgst_utgst = row.iloc[23]
        gst_rate_check = row['GST Rate check']
        
        # To check for NaN values and convert columns to numeric types
        if pd.notna(revised_taxable_value) and pd.notna(gst_rate_check):
            # Handle the case where the conversion fails using try and except
            try:
                taxable_value = float(revised_taxable_value)
                igst = float(revised_igst)
                cgst = float(revised_cgst)
                sgst_utgst = float(revised_sgst_utgst)
                rate_check = float(gst_rate_check)
                
                return round((taxable_value * rate_check / 100) - (igst + cgst + sgst_utgst), 0)
            except ValueError:
                return "Revised taxable Value (Rs.)/ IGST/ Revised IGST/ Revised CGST/ Revised SGST/UTGST Should not be blank"  
        else:
            return None

    df7['GST Difference'] = df7.apply(compute_gst_difference, axis=1)
    
    # Function to check Origional and Revised GSTIN
    def gstin_of_recipient_check(row):
        original_gstin = row['Original GSTIN of recipient']
        revised_gstin = row['Revised Recipients GSTIN (Billing party GSTIN)']
        
        if original_gstin == "0":
            return "It should not be blank"
        elif original_gstin == revised_gstin:
            return "Match"
        else:
            return "Didn't match"
        
    df7['GSTIN of recipient check'] = df7.apply(gstin_of_recipient_check, axis=1)
    df7['GSTIN length check'] = df7.apply(lambda row: len(str(row.iloc[12])), axis=1)
    
    # Function to check Unusual transation by Revised HSN
    def identify_unusual_transaction(row):
        revised_hsn = row['Revised HSN']
        if pd.isna(revised_hsn):
            return "HSN should not be blank"
        elif revised_hsn == "9997":
            return "This sort of recovery made need to check the transaction"
        elif revised_hsn == "9965":
            return "This sort of GTA Supply made need to check the transaction"
        elif revised_hsn == "996601":
            return "Motor vehicle provided on rent along with operator and cost of fuel is recovered in rent or Motor vehicle provided on rent along with operator but cost of fuel is not recovered in rent(Need to verify the GST Rate)"
        elif revised_hsn == "9973":
            return "Motor vehicle provided on rent without operator whether or not fuel cost is recovered in rent"
        elif revised_hsn == "8703":
            return "Prima facie it is sale of used car (Need to check the transaction)"
        elif revised_hsn == "9972":
            return "Prima facie it is renting of immovable property (Need to check the transaction)"
        elif revised_hsn == "4902":
            return "Prima facie it is supply of MEIS scripts (Need to check the transaction)"
        elif revised_hsn == "8471":
            return "Prima facie it is sale of used Laptops/Desktops (Need to check the transaction)"
        elif revised_hsn == "997331":
            return "Prima facie it is supply of Licensing services for the right to use computer software and databases (Need to check the transaction)"
        elif revised_hsn == "9954":
            return "Prima facie it is supply of works contract service (Need to check the transaction)"
        else:
            return "-"
        
    df7['Identification of unusual transaction by HSN'] = df7.apply(identify_unusual_transaction, axis=1)
    
    # Function to check Unusual transation by Revised description
    def identify_unusual_transaction_by_description(row):
        revised_description = row['Revised description']
        if pd.isna(revised_description) or revised_description == "-":
            return "Description should not be blank"
        elif 'recovery' in revised_description:
            return "Prima facie it is observed that some recovery made by the Company"
        elif 'reimb' in revised_description:
            return "Prima facie it is observed that some reimbursement made by the Company"
        elif 'works contract' in revised_description:
            return "Prima facie it is observed that works contract service provided by the Company"
        elif 'rent' in revised_description:
            return "Prima facie it is observed that renting service provided by the Company (need to check the transaction)"
        elif 'scrap' in revised_description:
            return "Prima facie it is observed that scrap sale is made by the Company"
        elif 'gift' in revised_description:
            return "Prima facie it is observed that gift provided by the Company"
        elif 'dest' in revised_description:
            return "Prima facie it is observed that material destroyed and sale made by the Company"
        elif 'stolen' in revised_description:
            return "Prima facie it is observed that the material is stolen in the Company"
        elif 'lost' in revised_description:
            return "Prima facie it is observed that some material is lost in the Company"
        elif 'disposed' in revised_description:
            return "Prima facie it is observed that inputs/Capital goods disposed by the Company"
        elif 'free sample' in revised_description:
            return "Prima facie it is observed that free sample supply made by the Company(Need to check whether ITC on the same is reversed)"
        elif 'written off' in revised_description:
            return "Prima facie it is observed that made by the Company"
        elif 'cheque bounce' in revised_description:
            return "Prima facie it is observed that recovery made by the Company"
        elif 'damage' in revised_description:
            return "Prima facie it is observed that damage material sold by the Company"
        elif 'penalty' in revised_description:
            return "Prima facie it is observed that penalty recovered by the Company"
        elif 'interest' in revised_description:
            return "Prima facie it is observed that recovery made by the Company"
        elif 'delay' in revised_description:
            return "Prima facie it is observed that recovery made by the Company"
        elif revised_description in ['Interest', 'Works', 'Recovery', 'Reimb', 'Rent', 'Scrap', 'Gift', 'Dest', 'Stolen', 'Lost', 'Disposed', 'Free sample', 'Written off', 'Cheque', 'Damage', 'Penalty', 'Delay']:
            return f"Prima facie it is observed that {revised_description}"
        else:
            return "Description need to check"
        
    df7['Identification of unusual transaction by Description'] = df7.apply(identify_unusual_transaction_by_description, axis=1)
    
    df7.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Amendment(Advances) sheet and Save in Amendment(Advances) sheet
    df7 = pd.concat([df2, df7], axis=1)
    
    df7_columns = df7.columns.tolist()
    df7_new_columns = df7_columns[2:] + df7_columns[:2] # Move the first and second columns to the end
    df7 = df7[df7_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df7_column_to_extend = 'Start date'
    df7_value_to_extend = df7.at[1, df7_column_to_extend]
    df7[df7_column_to_extend] = df7[df7_column_to_extend].fillna(df7_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df7_column_to_extend1 = 'End date'
    df7_value_to_extend1 = df7.at[1, df7_column_to_extend1]
    df7[df7_column_to_extend1] = df7[df7_column_to_extend1].fillna(df7_value_to_extend1)
    
    df7['Original document date'] = pd.to_datetime(df7['Original document date'])
    df7['Revised receipt voucher date'] = pd.to_datetime(df7['Revised receipt voucher date'])
    df7['Start date'] = pd.to_datetime(df7['Start date'])
    df7['End date'] = pd.to_datetime(df7['End date'])

    # Function to check Origional Document Date as it should be between start and end date
    def origional_document_date_check(row):
        if pd.isna('Original document date'):
            return "Original document date is Blank"
        elif pd.notna('Original document date'):
            if row['Original document date'] > row['End date']:
                return "Original document date is not pertaining to this FY"
            elif row['Original document date'] < row['Start date']:
                return "Original document date is not pertaining to this FY"
            else:
                return "Correct"
        
    df7['Original document date check'] = df7.apply(origional_document_date_check, axis=1)
    
    # Function to check Reviseed receipt voucher date as it should be between start and end date
    def revised_document_date_check(row):
        if pd.isna('Revised receipt voucher date'):
            return "Revised receipt voucher date is Blank"
        elif pd.notna('Revised receipt voucher date'):
            if row['Revised receipt voucher date'] > row['End date']:
                return "Revised receipt voucher date is not pertaining to this FY"
            elif row['Revised receipt voucher date'] < row['Start date']:
                return "Revised receipt voucher date is not pertaining to this FY"
            else:
                return "Correct"
        
    df7['Revised receipt voucher date check'] = df7.apply(revised_document_date_check, axis=1)
    
    df7 = df7.drop(columns=["Start date", "End date"]) # Remove the specified columns
    df7 = df7.dropna(subset=df7.columns[0:25], how='all') # Remove unnecessarily created checks even if rows not contains Data
    

    # -------------------------------------------* Processing of Advances Sheet *----------------------------------------------------
    
    df6 = df6.iloc[5:] # Set the 6th row as the header.
    header_row = df6.iloc[0]
    df6 = df6[1:]
    df6.columns = header_row
    
    # To select data from 1st column to 29th column 
    df6 = df6.iloc[:, 1:30]
    
    # To check if duplicate entries of any Receipt voucher number is present
    df6['Count'] = df6.groupby('Receipt voucher number')['Receipt voucher number'].transform('count')
    df6['Duplicate receipt voucher no check'] = df6['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df6.drop(columns=['Count'], inplace=True)
    
    # Function to create a new column 'Status of Recipient Check 1'
    def check_status_of_recipient(row):
        status_of_recipient = row['Status of recipient']
        
        if pd.notna(status_of_recipient) and "Unregistered" in status_of_recipient:
            return row[6]
        else:
            return "Correct / Supply made to registered GSTN"

    df6['Status of Recipient Check 1'] = df6.apply(check_status_of_recipient, axis=1)
    
    # Function to create a new column 'Status of Recipient Check 2'
    def check_status_of_recipient(row):
        status_of_recipient = row['Status of recipient']
        
        if pd.notna(status_of_recipient) and "Registered" in status_of_recipient:
            return row[6]
        else:
            return "GST number should be blank"

    df6['Status of Recipient Check 2'] = df6.apply(check_status_of_recipient, axis=1)
    
    # Function to check status of Recipient
    def check_status_of_recipient(row):
        status_check_1 = row['Status of Recipient Check 1']
        status_check_2 = row['Status of Recipient Check 2']
        gstin = row[6]

        if status_check_1 == "-":
            return "Correct"
        elif status_check_2 == "-":
            return "Incorrect / Supply made to registered GSTN should not be blank"
        elif gstin == status_check_2:
            return "Correct"
        elif gstin == status_check_1:
            return "Incorrect / GSTIN should be blank"
        else:
            return "-"

    df6['Status of Recipient check'] = df6.apply(check_status_of_recipient, axis=1)

    df6 = df6.drop(columns=["Status of Recipient Check 1", "Status of Recipient Check 2"]) # Remove specified columnns
    
    # Function to check Type of supply
    def type_of_supply_check(row):
        if isinstance(row, str):
            if "Regular" in row:
                return "Correct"
            elif "Export with" in row:
                return "It is a Zero-rated supply CGST+SGST should be blank"
            elif "SEZ without" in row:
                return "It is a Zero-rated supply without payment; all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%."
            elif "Export without" in row:
                return "It is a Zero-rated supply without payment; all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%."
            elif "SEZ with" in row:
                return "It is a Zero-rated supply CGST+SGST should be blank"
            elif "Exempt" in row:
                return "It is an exempt supply; tax value should be blank"
            elif row == "-":
                return "It should not be blank."
        return "Prima facie it observes that it is other than Regular supply"

    df6["Type of Supply Check"] = df6["Type of supply"].apply(type_of_supply_check)
    
    # Creating column to check taxability
    df6['Taxability Check'] = df6.apply(lambda row: 
        "Exempt supply made by the Company which attaract reversal under rule 42 & 43"
        if "Exempt" in str(row['Taxability']) else
        "It is an No GST supply hence tax amount should be zero"
        if str(row['Taxability']) == "non GST" else
        "It should not be blank"
        if str(row['Taxability']) == "-" else
        "Normal taxable supply", axis=1)
    
    # Creating column to check Length of Voucher number
    df6['Voucher No Length Check'] = df6.apply(lambda row: 
        "Correct" if len(str(row['Receipt voucher number'])) <= 16 else "Need to check the Invoice copy",
        axis=1)
    
    # Creating column to get Length of GST
    df6['GST Length Check'] = df6.iloc[:, 6].apply(lambda x: len(x) if pd.notna(x) else None)
    
    # Function to check Length of GST Number(Should be Equal to 15)
    def calculate_gstin_check(row):
        gst_length = row['GST Length Check']
        if pd.isnull(gst_length):
            return "GSTIN Cannot be blank in case of registered supply"
        elif gst_length == 15:
            return "Correct"
        else:
            return "Need to mention the correct GST Number"

    df6['GSTIN check'] = df6.apply(calculate_gstin_check, axis=1)
    
    # To extract first two Numbers from String
    df6['Recipients GSTIN - Copy'] = df6.iloc[:, 6].str.extract(r'(\d{2})')
    df6['Place Of Supply - Copy'] = df6['Place Of Supply'].str.extract(r'(\d{2})')
    
    # Function to check Place of Supply and Recipient GSTIN 
    def pos_recipient_check(row):
        pos_copy = row['Place Of Supply - Copy']
        gstin_copy = row['Recipients GSTIN - Copy']
        
        if pos_copy == gstin_copy:
            return "Match"
        elif pd.isna(pos_copy):
            if pd.isna(gstin_copy):
                return "Place of Supply and Recipients GSTIN should not be blank"
        elif pd.isna(gstin_copy):
                return "Recipients GSTIN should not be blank"
        else:
            return "Place of Supply and Recipients GSTIN need to check"

    df6['POS & Recipient check'] = df6.apply(pos_recipient_check, axis=1)
    
    df6 = df6.drop(columns=["Place Of Supply - Copy", "Recipients GSTIN - Copy"]) # Remove specified columnns
    
    # Function to check HSN
    def hsn_check(row):
        hsn = row['HSN']
        
        if hsn > 99999999:
            return "Need to mention correct HSN"
        elif hsn > 999:
            return "Correct"
        elif hsn == 0:
            return "HSN should not be blank"
        else:
            return "Need to mention correct HSN"

    df6['HSN check'] = df6.apply(hsn_check, axis=1)
    
    # Function to check Unusual transation by Description
    def unusual_transaction_by_description(row):
        description = row['Description']
        
        if pd.notna(description):
            if "recovery" in description:
                return "Prima facie it is observed that some recovery made by the Company"
            elif "reimb" in description:
                return "Prima facie it is observed that some reimbursement made by the Company"
            elif "works contract" in description:
                return "Prima facie it is observed that works contract service provided by the Company"
            elif "rent" in description:
                return "Prima facie it is observed that renting service provided by the Company (need to check the transaction)"
            elif "scrap" in description:
                return "Prima facie it is observed that scrap sale is made by the Company"
            elif "gift" in description:
                return "Prima facie it is observed that gift provided by the Company"
            elif "dest" in description:
                return "Prima facie it is observed that material destroyed and sale made by the Company"
            elif "stolen" in description:
                return "Prima facie it is observed that the material is stolen in the Company"
            elif "lost" in description:
                return "Prima facie it is observed that some material is lost in the Company"
            elif "disposed" in description:
                return "Prima facie it is observed that inputs/Capital goods disposed by the Company"
            elif "free sample" in description:
                return "Prima facie it is observed that free sample supply made by the Company(Need to check whether ITC on the same is reversed)"
            elif "written off" in description:
                return "Prima facie it is observed that made by the Company"
            elif "cheque bounce" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "damage" in description:
                return "Prima facie it is observed that damage material sold by the Company"
            elif "penalty" in description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "interest" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "delay" in description:
                return "Prima facie it is observed that recovery made by the Company"
            else:
                return "Normal Description"
        else:
            return "Description should not be blank"

    df6['Unusual Transaction by Description'] = df6.apply(unusual_transaction_by_description, axis=1)
    
    # Function to check and convert Revised GST Rate
    def gst_rate_check(row):
        gst_rate = row['GST Rate(%)']
        
        if pd.notna(gst_rate):
            if gst_rate == 0.28:
                return 28
            elif gst_rate == 28:
                return 28
            elif gst_rate == 0.18:
                return 18
            elif gst_rate == 18:
                return 18
            elif gst_rate == 0.12:
                return 12
            elif gst_rate == 12:
                return 12
            elif gst_rate == 0.05:
                return 5
            elif gst_rate == 5:
                return 5
            elif gst_rate == 0.025:
                return 2.5
            elif gst_rate == 2.5:
                return 2.5
            elif gst_rate == 0.01:
                return 0.1
            elif gst_rate == 0.03:
                return 3
            elif gst_rate == 3:
                return 3
            else:
                return "Need to mention correct GST Rate"
        else:
            return "Should not be blank"

    df6['GST Rate Check'] = df6.apply(gst_rate_check, axis=1)
    
    # Function to Find GST Difference
    def compute_gst_difference(row):
        taxable_value = row.iloc[14]
        igst = row.iloc[15]
        cgst = row.iloc[16]
        sgst_utgst = row.iloc[17]
        gst_rate_check = row['GST Rate Check']
        
        # Check for NaN values and convert columns to numeric types
        if pd.notna(taxable_value) and pd.notna(gst_rate_check):
            # Handle the case where the conversion fails using try and except
            try:
                taxable_value = float(taxable_value)
                igst = float(igst)
                cgst = float(cgst)
                sgst_utgst = float(sgst_utgst)
                rate_check = float(gst_rate_check)
                
                return round((taxable_value * rate_check / 100) - (igst + cgst + sgst_utgst), 0)
            except ValueError:
                return "taxable Value (Rs.)/ IGST/ Revised IGST/ Revised CGST/ Revised SGST/UTGST Should not be blank"
        else:
            return None

    df6['GST Difference'] = df6.apply(compute_gst_difference, axis=1)
    
    # Function to Check Document Type
    def document_type_check(row):
        document_type = row['Type of document']
        
        if pd.notna(document_type):
            if "Invoice" in document_type:
                return "Invoice date is after advance received"
            elif "Invoice-cum-bill of supply" in document_type:
                return "GST charge on taxable portion"
            elif "Bill of supply" in document_type:
                return "GST should not be charged"
            elif "Refund voucher" in document_type:
                return "Refund voucher should be after than advance received"
            else:
                return "Invalid Type of Document"
        else:
            return "It should not be blank"

    df6['Type of Document check'] = df6.apply(document_type_check, axis=1)
    
    # Function to check Length of Document number(should be less than equal to 16)
    def document_number_length_check(row):
        document_number = str(row['Document number'])
        if len(document_number) <= 16:
            return "Correct"
        else:
            return "Need to check the Invoice copy"

    df6['Document No Length Check'] = df6.apply(document_number_length_check, axis=1)
    
    df6.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Amendment(Advances) sheet and Save in Amendment(Advances) sheet
    df6 = pd.concat([df2, df6], axis=1)
    
    df6_columns = df6.columns.tolist()
    df6_new_columns = df6_columns[2:] + df6_columns[:2] # Move the first and second columns to the end
    df6 = df6[df6_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df6_column_to_extend = 'Start date'
    df6_value_to_extend = df6.at[1, df6_column_to_extend]
    df6[df6_column_to_extend] = df6[df6_column_to_extend].fillna(df6_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df6_column_to_extend1 = 'End date'
    df6_value_to_extend1 = df6.at[1, df6_column_to_extend1]
    df6[df6_column_to_extend1] = df6[df6_column_to_extend1].fillna(df6_value_to_extend1)
    
    df6['Document date'] = pd.to_datetime(df6['Document date'])
    df6['Start date'] = pd.to_datetime(df6['Start date'])
    df6['End date'] = pd.to_datetime(df6['End date'])

    # Function to check Document Date as it should be between start and end date
    def document_date_check(row):
        if pd.isna('Document date'):
            return "Document Date is Blank"
        elif pd.notna('Document date'):
            if row['Document date'] > row['End date']:
                return "Document date is not pertaining to this FY"
            elif row['Document date'] < row['Start date']:
                return "Document date is not pertaining to this FY"
            else:
                return "Correct"
        
    df6['Document Date check'] = df6.apply(document_date_check, axis=1)
    
    df6 = df6.drop(columns=["Start date", "End date"]) # Remove specified columnns
    df6 = df6.dropna(subset=df6.columns[0:28], how='all') # Remove unnecessarily created checks even if rows not contains Data
            
    
    # -------------------------------------------* Processing of Amendments (CDN) Sheet *----------------------------------------------------
    
    df5 = df5.iloc[5:] # Step 3: Set the 6th row as the header.
    header_row = df5.iloc[0]
    df5 = df5[1:]
    df5.columns = header_row
    
    # to select data from 1st column to 29th column 
    df5 = df5.iloc[:, 1:30]
    
    # Function to check Origional Document number and Revised Document number
    def document_number_check(row):
        original_doc_num = row['Original document number']
        revised_doc_num = row['Revised document number']
        
        if pd.isna(revised_doc_num):
            if pd.isna(original_doc_num):
                return "Original document number and Revised document number should not be blank"
            else:
                return "Revised document number should not be blank"
        elif pd.isna(original_doc_num):
            return "Original document number should not be blank"
        elif original_doc_num == revised_doc_num:
                return "Match"
        else :
            return "Didn't Match"

    df5['Comp. Document number Check'] = df5.apply(document_number_check, axis=1)
    
    # Function to check Invoice number(should be less than or equal to 16)
    def invoice_number_check(row):
        revised_doc_num = row['Revised document number']
        
        if pd.notna(revised_doc_num):
            revised_doc_num_str = str(revised_doc_num)
            if len(revised_doc_num_str) <= 16:
                return "Correct"
            else:
                return "Need to mention Invoice copy"
        else:
            return "Revised document number should not be blank"

    df5['Invoice no Check'] = df5.apply(invoice_number_check, axis=1)
    
    # Function to check Origional Document Date and Revised Document date
    def document_date_check(row):
        original_date = row['Original document date']
        revised_date = row['Revised document date']
        
        if pd.isna(revised_date):
            if pd.isna(original_date):
                return "Original document date and Revised document date should not be blank"
            else:
                return "Revised document date should not be blank"
        elif pd.isna(original_date):
            return "Original document date should not be blank"
        elif original_date == revised_date:
                return "Match"
        else :
            return "Didn't Match"

    df5['Comp. Document date Check'] = df5.apply(document_date_check, axis=1)
    df5['GSTN Length check'] = df5.iloc[:, 13].str.len()
    
    # Function to check Origional GSTIN of recipient and Revised GSTIN of recipient
    def gstn_recipient_check(row):
        original_gstin = row['Original GSTIN of recipient']
        revised_gstin = row['Revised GSTIN of recipient']
        gstn_length_check = row['GSTN Length check']
        
        if pd.isna(revised_gstin):
            if pd.isna(original_gstin):
                return "Origional and Revised GSTIN Cannot be blank in case of registered supply"
            else:
                return "Revised GSTIN Cannot be blank in case of registered supply"
        elif pd.isna(original_gstin):
                return "Origional GSTIN Cannot be blank in case of registered supply"
        elif gstn_length_check != 15:
            return "Need to mention the correct GST Number"
        elif original_gstin == revised_gstin:
            return "Match"
        elif gstn_length_check == 15:
            return "Incorrect length of Revised GSTIN of recipient / Didn't match"
        else:
            return "Didn't match / Need to mention the correct GST Number. GSTIN is a 15-digit alphanumeric code. The first two digits represent the state code, the next 10 digits represent the PAN (Permanent Account Number) of the taxpayer, the 13th digit represents the number of registrations the entity has within a state, the 14th digit is the default 'Z', and the last digit is a checksum digit calculated using the Modulus 10 algorithm"

    df5['Comp. GSTN of Recipient Check'] = df5.apply(gstn_recipient_check, axis=1)
    
    # Function to check Origional note type and Revised note type
    def note_type_check(row):
        original_note_type = row['Original note type']
        revised_note_type = row['Revised note type']
    
        if pd.isna(revised_note_type):
            if pd.isna(original_note_type):
                return "Original note type and Revised note type should not be blank"
            else:
                return "Revised note type should not be blank"
        elif pd.isna(original_note_type):
            return "Original note type should not be blank"
        elif original_note_type == revised_note_type:
                return "Match"
        else :
            return "Didn't Match"

    df5['Comp. Note type Check'] = df5.apply(note_type_check, axis=1)
    
    # Function to check Origional note Number and Revised note Number
    def note_number_check(row):
        original_note_number = row['Original note number']
        revised_note_number = row['Revised note number']
 
        if pd.isna(revised_note_number):
            if pd.isna(original_note_number):
                return "Original note number and Revised note number should not be blank"
            else:
                return "Revised note number should not be blank"
        elif pd.isna(original_note_number):
            return "Original note number should not be blank"
        elif original_note_number == revised_note_number:
                return "Match"
        else :
            return "Didn't Match"

    df5['Comp. Note number Check'] = df5.apply(note_number_check, axis=1)
    
    # Function to check Origional note Date and Revised note Date    
    def note_date_check(row):
        original_note_date = row['Original note date']
        revised_note_date = row['Revised note date']
 
        if pd.isna(revised_note_date):
            if pd.isna(original_note_date):
                return "Original note date and Revised note date should not be blank"
            else:
                return "Revised note date should not be blank"
        elif pd.isna(original_note_date):
            return "Original note date should not be blank"
        elif original_note_date == revised_note_date:
                return "Match"
        else :
            return "Didn't Match"

    df5['Comp. Note date Check'] = df5.apply(note_date_check, axis=1)
    
    # Function to check Revised HSN
    def hsn_check(row):
        revised_hsn = row['Revised HSN']
        
        if pd.notna(revised_hsn):
            if revised_hsn > 99999999:
                return "Need to mention correct HSN"
            elif revised_hsn > 999:
                return "Correct"
            elif revised_hsn == 0:
                return "HSN should not be blank"
        
        return "Need to mention correct HSN."

    df5['HSN Check'] = df5.apply(hsn_check, axis=1)
    
    # Function to check Revised GST Rate
    def gst_rate_check(row):
        revised_rate = row['Revised rate (%)']
        
        if pd.notna(revised_rate):
            if revised_rate == 0.28:
                return 28
            elif revised_rate == 28:
                return 28
            elif revised_rate == 0.18:
                return 18
            elif revised_rate == 18:
                return 18
            elif revised_rate == 0.12:
                return 12
            elif revised_rate == 12:
                return 12
            elif revised_rate == 0.05:
                return 5
            elif revised_rate == 5:
                return 5
            elif revised_rate == 0.025:
                return 2.5
            elif revised_rate == 2.5:
                return 2.5
            elif revised_rate == 0.01:
                return 0.1
            elif revised_rate == 0.03:
                return 3
            elif revised_rate == 3:
                return 3
        
        return "Need to mention correct GST Rate"

    df5['GST Rate Check'] = df5.apply(gst_rate_check, axis=1)

    # Function to Find GST Difference
    def calculate_gst_difference(row):
        taxable_value = row[21]
        gst_rate_check = row['GST Rate Check']
        igst = row[22]
        cgst = row[23]
        sgst = row[24]
        
        # Check if any of the relevant columns are NaN
        if pd.notna(taxable_value) and pd.notna(gst_rate_check) and pd.notna(igst) and pd.notna(cgst) and pd.notna(sgst):
            gst_difference = (taxable_value * gst_rate_check / 100) - (igst + cgst + sgst)
            return round(gst_difference, 0)
        
        return None

    df5['GST Diffrence'] = df5.apply(calculate_gst_difference, axis=1)
    
    mask1 = df5.iloc[:, 7].notna() # Create a boolean mask for non-null values
    df5.loc[mask1, df5.columns[7]] = df5.loc[mask1, df5.columns[7]].astype(str) # Convert non-null values to strings
    mask2 = df5.iloc[:, 13].notna() # Create a boolean mask for non-null values
    df5.loc[mask2, df5.columns[13]] = df5.loc[mask2, df5.columns[13]].astype(str) # Convert non-null values to strings
    
    # Apply the .str.extract() method to the modified column
    df5['Original GSTIN of recipient - Copy'] = df5.loc[mask1, df5.columns[7]].str.extract(r'(\d{2})')
    df5['Revised GSTIN of recipient - Copy'] = df5.loc[mask2, df5.columns[13]].str.extract(r'(\d{2})')
    
    # Function to Check Origional Recipient GSTIN and Revised Recipient GSTIN
    def compute_pos_recipient_gstin_check(row):
        revised_recipient_gstin = row['Revised GSTIN of recipient - Copy']
        origional_recipient_gstin = row['Original GSTIN of recipient - Copy']
        
        if pd.isna(revised_recipient_gstin):
            if pd.isna(origional_recipient_gstin):
                return "Revised and Origional GSTIN of recipients should not be blank"
            else:
                return "Revised GSTIN of recipients should not be blank"
        elif pd.isna(origional_recipient_gstin):
                return "Origional GSTIN of recipients should not be blank"
        elif revised_recipient_gstin == origional_recipient_gstin:
            return "Match"
        else:
            return "Incorrect POS need to check"

    df5['POS & recipient GSTIN Check'] = df5.apply(compute_pos_recipient_gstin_check, axis=1)
    df5 = df5.drop(columns=["Original GSTIN of recipient - Copy", "Revised GSTIN of recipient - Copy"]) #Remove specified columnns
    
    # Function to check Unusual transation by Revised HSN
    def identify_unusual_transaction(row):
        revised_hsn = row['Revised HSN']
        
        if revised_hsn == "9997":
            return "This sort of recovery made need to check the transaction"
        elif revised_hsn == "9965":
            return "This sort of GTA Supply made need to check the transaction"
        elif revised_hsn == "996601":
            return "Motor vehicle provided on rent along with operator and cost of fuel is recovered in rent or Motor vehicle provided on rent along with operator but cost of fuel is not recovered in rent(Need to verify the GST Rate)"
        elif revised_hsn == "9973":
            return "Motor vehicle provided on rent without operator whether or not fuel cost is recovered in rent"
        elif revised_hsn == "8703":
            return "Prima facie it is sale of used car (Need to check the transaction)"
        elif revised_hsn == "9972":
            return "Prima facie it is renting of immovable property (Need to check the transaction)"
        elif revised_hsn == "4902":
            return "Prima facie it is supply of MEIS scripts (Need to check the transaction)"
        elif revised_hsn == "8471":
            return "Prima facie it is sale of used Laptops/Desktops (Need to check the transaction)"
        elif revised_hsn == "997331":
            return "Prima facie it is supply of Licensing services for the right to use computer software and databases (Need to check the transaction)"
        elif revised_hsn == "9954":
            return "Prima facie it is supply of works contract service (Need to check the transaction)"
        elif revised_hsn == "0":
            return "HSN should not be blank"
        else:
            return "-"

    df5['Identification of unusal transaction by HSN'] = df5.apply(identify_unusual_transaction, axis=1)
    
    # To check if duplicate entries of any Origional document number is present
    df5['Count'] = df5.groupby('Original document number')['Original document number'].transform('count')
    df5['Origional Invoice Duplicates'] = df5['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df5.drop(columns=['Count'], inplace=True)
    
    # To check if duplicate entries of any Revised document number is present
    df5['Count'] = df5.groupby('Revised document number')['Revised document number'].transform('count')
    df5['Revised Invoice Duplicates'] = df5['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df5.drop(columns=['Count'], inplace=True)
    
    df5.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Amendments (CDN) sheet and Save in Amendment(CDN) sheet
    df5 = pd.concat([df2, df5], axis=1)
    
    df5_columns = df5.columns.tolist()
    df5_new_columns = df5_columns[2:] + df5_columns[:2] # Move the first and second columns to the end
    df5 = df5[df5_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df5_column_to_extend = 'Start date'
    df5_value_to_extend = df5.at[1, df5_column_to_extend]
    df5[df5_column_to_extend] = df5[df5_column_to_extend].fillna(df5_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df5_column_to_extend1 = 'End date'
    df5_value_to_extend1 = df5.at[1, df5_column_to_extend1]
    df5[df5_column_to_extend1] = df5[df5_column_to_extend1].fillna(df5_value_to_extend1)
    
    df5['Original document date'] = pd.to_datetime(df5['Original document date'])
    df5['Revised document date'] = pd.to_datetime(df5['Revised document date'])
    df5['Start date'] = pd.to_datetime(df5['Start date'])
    df5['End date'] = pd.to_datetime(df5['End date'])

    # Function to check Origional Document Date as it should be between start and end date
    def origional_document_date_check(row):
        if pd.isna('Original document date'):
            return "Original document date is Blank"
        elif pd.notna('Original document date'):
            if row['Original document date'] > row['End date']:
                return "Origional Document Date is not pertaining to this FY"
            elif row['Original document date'] < row['Start date']:
                return "Origional Document Date is not pertaining to this FY"
            else:
                return "Correct"
        
    df5['Original document date check'] = df5.apply(origional_document_date_check, axis=1)
    
    # Function to check Revised Document Date as it should be between start and end date
    def revised_document_date_check(row):
        if pd.isna('Revised document date'):
            return "Revised document date is Blank" 
        elif pd.notna('Revised document date'):
            if row['Revised document date'] > row['End date']:
                return "Revised document date is not pertaining to this FY"
            elif row['Revised document date'] < row['Start date']:
                return "Revised document date is not pertaining to this FY"
            else:
                return "Correct"
        
    df5['Revised document date check'] = df5.apply(revised_document_date_check, axis=1)
    
    df5 = df5.drop(columns=["Start date", "End date"]) # Remove specified columnns
    df5 = df5.dropna(subset=df5.columns[0:28], how='all') # Remove unnecessarily created checks even if rows not contains Data



    # -------------------------------------------* Processing of Debit&CreditNotes Sheet *----------------------------------------------------
    df4 = df4.iloc[4:]

    # Step 3: Set the 6th row as the header.
    header_row = df4.iloc[0]
    df4 = df4[1:]
    df4.columns = header_row
    
    # to select data from 1st column to 24th column 
    df4 = df4.iloc[:, 1:25]
    
    # To check if duplicate entries of any Receipt voucher number is present
    cndn_df4 = df4.groupby('Document number').size().reset_index(name='Count')
    cndn_df4['Invoice Duplicates'] = cndn_df4['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique')
    df4 = df4.merge(cndn_df4[['Document number', 'Invoice Duplicates']], on='Document number', how='left')
    print(df4)
    
    # Function to check the length of GSTIN Number(Should be equal to 15)
    def gstn_check(row):
        if pd.notna(row[7]):
            gst_length = len(str(row[7]))
            if pd.isna(gst_length):
                return "GSTIN Cannot be blank in case of registered supply"
            elif gst_length == 15:
                return "Correct"
        return ""

    df4['GSTN Check'] = df4.apply(gstn_check, axis=1)
    
    # Function to check Type of supply
    def type_of_supply_check(row):
        type_of_supply = row['Type of supply']

        if pd.notna(type_of_supply):
            if "Regular" in type_of_supply:
                return "Correct"
            elif "Export with" in type_of_supply:
                return "It is a Zero-rated supply CGST+SGST should be blank"
            elif "SEZ without" in type_of_supply:
                return "It is a Zero-rated supply without payment all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%."
            elif "Export without" in type_of_supply:
                return "It is a Zero-rated supply without payment all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%."
            elif "SEZ with" in type_of_supply:
                return "It is a Zero-rated supply CGST+SGST should be blank"
            elif "Exempt" in type_of_supply:
                return "It is an exempt supply tax value should be blank"
            else:
                return "Prima facie it observes that it is other than Regular supply"
        else :
            return "It should not be blank"

    df4['Type of Supply check'] = df4.apply(type_of_supply_check, axis=1)
    
    # Function to check Document Number(Should be less than or equal to 16)
    def invoice_no_check(row):
        document_number = row['Document number']
        
        if pd.notna(document_number) and len(str(document_number)) <= 16:
            return "Correct"
        else:
            return "Need to check the Invoice copy"

    df4['Invoice No Check'] = df4.apply(invoice_no_check, axis=1)
    
    # Function to check HSN
    def hsn_check(row):
        hsn_value = row['HSN ']
        
        if pd.notna(hsn_value):
            if hsn_value > 99999999:
                return "Need to mention correct HSN"
            elif hsn_value > 999:
                return "Correct"
            elif hsn_value == 0:
                return "HSN should not be blank"
        
        return "Need to mention correct HSN"

    df4['HSN check'] = df4.apply(hsn_check, axis=1)
    
    # Function which converts any type of gst Rate into Number
    def gst_rate_check(row):
        rate_percent = row[15]
        
        if pd.notna(rate_percent):
            if rate_percent == 0.28:
                return 28
            elif rate_percent == 28:
                return 28
            elif rate_percent == 0.18:
                return 18
            elif rate_percent == 18:
                return 18
            elif rate_percent == 0.12:
                return 12
            elif rate_percent == 12:
                return 12
            elif rate_percent == 0.05:
                return 5
            elif rate_percent == 5:
                return 5
            elif rate_percent == 0.025:
                return 2.5
            elif rate_percent == 2.5:
                return 2.5
            elif rate_percent == 0.01:
                return 0.1
            elif rate_percent == 0.03:
                return 3
            elif rate_percent == 3:
                return 3
        
        return "Need to mention correct GST Rate"

    df4['GST Rate Check'] = df4.apply(gst_rate_check, axis=1)
    
    # Function to Find GST Difference
    def compute_gst_difference(row):
        taxable_value = row[16]
        igst = row[17]
        cgst = row[18]
        sgst_utgst = row[19]
        gst_rate_check = row['GST Rate Check']
        
        # Check for NaN values and convert columns to numeric types
        if pd.notna(taxable_value) and pd.notna(gst_rate_check):
            return round((float(taxable_value) * float(gst_rate_check) / 100) - (float(igst) + float(cgst) + float(sgst_utgst)), 0)
        else:
            return None

    df4['GST Difference'] = df4.apply(compute_gst_difference, axis=1)
    
    # Function to check Place of Supply and Recipients GSTIN
    def compute_pos_recipient_gstin_check(row):
        place_of_supply = row.iloc[8]
        recipient_gstin = row.iloc[7]
        
        if pd.isna(place_of_supply):
            if pd.isna(recipient_gstin):
                return "Place of supply and GSTIN of recipients should not be blank"
            else:
                return "Place of supply should not be blank"
        elif pd.isna(recipient_gstin):
                return "GSTIN of recipients should not be blank"
        elif place_of_supply == recipient_gstin:
            return "Match"
        else:
            return "Incorrect POS need to check"

    df4['POS & Recipient GSTIN Check'] = df4.apply(compute_pos_recipient_gstin_check, axis=1)
    
    # Function to check Unusual transation by HSN
    def unusual_transaction_by_hsn(row):
        hsn_value = row['HSN ']
        
        if hsn_value == 9997:
            return "This sort of recovery made need to check the transaction"
        elif hsn_value == 9965:
            return "This sort of GTA Supply made need to check the transaction"
        elif hsn_value == 996601:
            return "Motor vehicle provided on rent along with operator and cost of fuel is recovered in rent or Motor vehicle provided on rent along with operator but cost of fuel is not recovered in rent (Need to verify the GST Rate)"
        elif hsn_value == 9973:
            return "Motor vehicle provided on rent without operator whether or not fuel cost is recovered in rent"
        elif hsn_value == 8703:
            return "Prima facie it is sale of used car (Need to check the transaction)"
        elif hsn_value == 9972:
            return "Prima facie it is renting of immovable property (Need to check the transaction)"
        elif hsn_value == 4902:
            return "Prima facie it is supply of MEIS scripts (Need to check the transaction)"
        elif hsn_value == 8471:
            return "Prima facie it is sale of used Laptops/Desktops (Need to check the transaction)"
        elif hsn_value == 997331:
            return "Prima facie it is supply of Licensing services for the right to use computer software and databases (Need to check the transaction)"
        elif hsn_value == 9954:
            return "Prima facie it is supply of works contract service (Need to check the transaction)"
        else:
            return "-"

    df4['Identification of Unusual Transaction by HSN'] = df4.apply(unusual_transaction_by_hsn, axis=1)
    
    # Function to check Taxability
    def taxability_check(row):
        taxability_value = row['Taxability']
        
        if pd.notna(taxability_value):
            if "Exempt" in taxability_value:
                return "Exempt supply made by the Company which attracts reversal under rule 42 & 43"
            elif taxability_value == "non GST":
                return "It is a No GST supply hence tax amount should be zero"
            elif taxability_value == "-":
                return "It should not be blank"
        
        return "Normal taxable supply"

    df4['Taxability Check'] = df4.apply(taxability_check, axis=1)
    
    df4.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Debit&CreditNotes sheet and Save in Debit&CreditNotes sheet
    df4 = pd.concat([df2, df4], axis=1)
    
    df4_columns = df4.columns.tolist()
    df4_new_columns = df4_columns[2:] + df4_columns[:2] # Move the first and second columns to the end
    df4 = df4[df4_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df4_column_to_extend = 'Start date'
    df4_value_to_extend = df4.at[1, df4_column_to_extend]
    df4[df4_column_to_extend] = df4[df4_column_to_extend].fillna(df4_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df4_column_to_extend1 = 'End date'
    df4_value_to_extend1 = df4.at[1, df4_column_to_extend1]
    df4[df4_column_to_extend1] = df4[df4_column_to_extend1].fillna(df4_value_to_extend1)
    
    df4['Document date'] = pd.to_datetime(df4['Document date'])
    df4['Start date'] = pd.to_datetime(df4['Start date'])
    df4['End date'] = pd.to_datetime(df4['End date'])

    # Function to check Document Date as it should be between start and end date
    def document_date_check(row):
        if pd.isna('Document date'):
            return "Document Date is Blank"
        elif pd.notna('Document date'):
            if row['Document date'] > row['End date']:
                return "Document date is not pertaining to this FY"
            elif row['Document date'] < row['Start date']:
                return "Document date is not pertaining to this FY"
            else:
                return "Correct"
        
    df4['Document Date check'] = df4.apply(document_date_check, axis=1)
    df4 = df4.drop(columns=["Start date", "End date"]) # Remove specified columnns
    
    # function to check Reasons for issue
    def reasons_for_issue_check(row):
        reasons_value = row['Reasons for issue of credit/debit note']
        
        if pd.isna(reasons_value):
            return "-"
        else:
            return "correct"

    # Apply the function to create a new column 'Reasons for issue check'
    df4['Reasons for issue check'] = df4.apply(reasons_for_issue_check, axis=1)
    
    df4 = df4.dropna(subset=df4.columns[0:23], how='all') # Remove unnecessarily created checks even if rows not contains Data
    
        
    # -------------------------------------------* Processing of Amendments(Invoices) Sheet *----------------------------------------------------
    
    df3 = df3.iloc[5:]
    # Step 3: Set the 6th row as the header.
    header_row = df3.iloc[0]
    df3 = df3[1:]
    df3.columns = header_row
    
    # Remove the 1st to 37th column from the DataFrame
    df3 = df3.iloc[:, 1:38]
    
    
    # To check if duplicate entries of any Origional document number is present
    dn_df3 = df3.groupby('Original document number').size().reset_index(name='Count')
    dn_df3['Original Invoice Duplicates'] = dn_df3['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique')
    df3 = df3.merge(dn_df3[['Original document number', 'Original Invoice Duplicates']], on='Original document number', how='left')
    print(df3)
    
    # To check if duplicate entries of any Revised document number is present
    rdn_df3 = df3.groupby('Revised document number').size().reset_index(name='Count')
    rdn_df3['Revised Invoice Duplicates'] = rdn_df3['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique')
    df3 = df3.merge(rdn_df3[['Revised document number', 'Revised Invoice Duplicates']], on='Revised document number', how='left')
    print(df3)
    
    # Function to check the Origional and Revised Status of recipient
    def compute_comp_status(row):
        if row['Original status of recipient'] == row['Revised status of recipient']:
            return "Match"
        elif pd.isnull(row['Revised status of recipient']):
            return "It should not be blank"
        elif row['Revised status of recipient'] == "Unregister":
            return "Changes made - Supply made to unregistered GSTN should be blank"
        else:
            return "Changes made - Supply made to registered GSTN should not be blank"

    df3['Comp. Status Of Recipient Check'] = df3.apply(compute_comp_status, axis=1)
    
    # Function to check the Origional and Revised type of Supply
    def compute_comp_type_of_supply(row):
        revised_type_of_supply = row['Revised type of supply']
        
        if row['Original type of supply'] == revised_type_of_supply:
            return "Match"
        elif pd.isnull(revised_type_of_supply):
            return "It should not be blank"
        elif "SEZ supplies without" in revised_type_of_supply or "SEZ without" in revised_type_of_supply:
            return "Changes made - It is a Zero-rated supply without payment all tax columns should be blank"
        elif "SEZ supplies with" in revised_type_of_supply or "SEZ with" in revised_type_of_supply or "Export with" in revised_type_of_supply:
            return "Changes made - It is a Zero-rated supply CGST+SGST should be blank"
        elif "Exempt" in revised_type_of_supply:
            return "Changes made - It is an exempt supply tax value should be blank"
        elif "Regular" in revised_type_of_supply:
            return "Correct"
        elif "Export without" in revised_type_of_supply or "Export supplies without" in revised_type_of_supply:
            return "Changes made - It is a Zero-rated supply without payment all tax columns should be blank"
        else:
            return "Prima facie it observes that it is other than Regular supply"

    df3['Comp. Type Of Supply Check'] = df3.apply(compute_comp_type_of_supply, axis=1)
    
    # Function to check Taxibility
    def compute_comp_taxability_check(row):
        revised_taxability = row['Revised Taxability']
        
        if row['Original taxability'] == revised_taxability:
            return "Match"
        elif pd.isna(revised_taxability):
            return "Should not be blank"
        elif pd.notna(revised_taxability):
            if "Exempt" in revised_taxability:
                return "Changes made - Exempt supply made by the Company which attracts reversal under rule 42 & 43 also tax amount should be zero"
            elif "non GST" in revised_taxability:
                return "Changes made - It is a No GST supply hence tax amount should be zero"
            elif revised_taxability == "-":
                return "It should not be blank"
            elif "Taxable" in revised_taxability:
                return "Changes made - Normal taxable supply"
        return "Didn't match"

    df3['Comp. Taxability Check'] = df3.apply(compute_comp_taxability_check, axis=1)
    
    # Function to check the Origional and Revised type of Document
    def compute_comp_type_of_document_check(row):
        revised_type_of_document = row['Revised type of document']
        
        if row['Original type of documents'] == revised_type_of_document:
            return "Match"
        elif pd.isna(revised_type_of_document):
            return "Should not be blank"
        else:
            return "Didn't match"

    df3['Comp. Type of Document Check'] = df3.apply(compute_comp_type_of_document_check, axis=1)
    
    # Function to check the Origional and Revised Document Number
    def compute_comp_document_number_check(row):
        revised_document_number = row['Revised document number']
        
        if row['Original document number'] == revised_document_number:
            return "Match"
        elif pd.isna(revised_document_number):
            return "It should not be blank"
        else:
            return "Didn't match / Need to check the Invoice copy"

    df3['Comp. Document number Check'] = df3.apply(compute_comp_document_number_check, axis=1)
    
    # Function to check the length of Revised document Number
    def compute_revised_doc_no_check(row):
        revised_document_number = row['Revised document number']
        
        if len(str(revised_document_number)) <= 16:
            return "Correct"
        elif pd.isna(revised_document_number):
            return "It should not be blank"
        else:
            return "Need to check the Invoice copy"

    df3['Invoice no. Check'] = df3.apply(compute_revised_doc_no_check, axis=1)
    
    # Function to check the Origional and Revised Document Date
    def compute_comp_document_date_check(row):
        original_document_date = row['Original document date']
        revised_document_date = row['Revised document date']
        
        if original_document_date == revised_document_date:
            return "Match"
        elif pd.isna(original_document_date):
            return "It should not be blank"
        else:
            return "Didn't match"

    df3['Comp. Document Date Check'] = df3.apply(compute_comp_document_date_check, axis=1)
    
    # Check the length of GST Number
    df3['GSTN Length check'] = df3.iloc[14].apply(lambda x: len(str(x)))
    
    # Function to check the Origional GSTIN and Revised GSTIN
    def compute_comp_gstn_of_recipient_check(row):
        gstn_length_check = row['GSTN Length check']
        original_gstin = row.iloc[7]
        revised_gstin = row.iloc[14]
        
        if pd.isna(gstn_length_check):
            return "GSTIN Cannot be blank in case of registered supply"
        elif gstn_length_check != 15:
            return "Need to mention the correct GST Number"
        elif original_gstin == revised_gstin:
            return "Match"
        elif gstn_length_check == 15:
            return "Revised Length Correct / Didn't match"
        else:
            return (
                "Didn't match / Need to mention the correct GST Number, GSTIN is a 15-digit alphanumeric code. The first two digits represent the state code, the next 10 digits represent the PAN (Permanent Account Number) of the taxpayer, the 13th digit represents the number of registrations the entity has within a state, the 14th digit is the default 'Z', and the last digit is a checksum digit calculated using the Modulus 10 algorithm")

    df3['Comp. GSTN of Recipient Check'] = df3.apply(compute_comp_gstn_of_recipient_check, axis=1)
    
    # Function to Check Revised HSN
    def compute_hsn_check(row):
        revised_hsn = row['Revised HSN']
        
        if revised_hsn > 99999999:
            return "Need to mention correct HSN"
        elif revised_hsn > 999:
            return "Correct"
        elif pd.isna(revised_hsn):
            return "HSN should not be blank"
        else:
            return "Need to mention correct HSN"

    df3['HSN Check'] = df3.apply(compute_hsn_check, axis=1)

    # Function which converts any type of gst Rate into Number
    def gst_rate_check(row):
        rate_percent = row['Revised rate (%)']
        
        if pd.notna(rate_percent):
            if rate_percent == 0.28:
                return 28
            elif rate_percent == 28:
                return 28
            elif rate_percent == 0.18:
                return 18
            elif rate_percent == 18:
                return 18
            elif rate_percent == 0.12:
                return 12
            elif rate_percent == 12:
                return 12
            elif rate_percent == 0.05:
                return 5
            elif rate_percent == 5:
                return 5
            elif rate_percent == 0.025:
                return 2.5
            elif rate_percent == 2.5:
                return 2.5
            elif rate_percent == 0.01:
                return 0.1
            elif rate_percent == 0.03:
                return 3
            elif rate_percent == 3:
                return 3
        
        return "Need to mention correct GST Rate"

    df3['GST Rate Check'] = df3.apply(gst_rate_check, axis=1)
    
    # Function to check the GST Difference
    def compute_gst_difference(row):
        revised_taxable_value = row['Revised taxable Value (Rs.)']
        revised_igst = row.iloc[27]
        revised_cgst = row.iloc[28]
        revised_sgst_utgst = row.iloc[29]
        gst_rate_check = row['GST Rate Check']
        
        # Check for NaN values and convert columns to numeric types
        if pd.notna(revised_taxable_value) and pd.notna(gst_rate_check):
            # Handle the case where the conversion fails using try and except
            try:
                taxable_value = float(revised_taxable_value)
                igst = float(revised_igst)
                cgst = float(revised_cgst)
                sgst_utgst = float(revised_sgst_utgst)
                rate_check = float(gst_rate_check)
                
                return round((taxable_value * rate_check / 100) - (igst + cgst + sgst_utgst), 0)
            except ValueError:
                return "Revised taxable Value (Rs.)/ IGST/ Revised IGST/ Revised CGST/ Revised SGST/UTGST Should not be blank"  
        else:
            return None

    df3['GST Difference'] = df3.apply(compute_gst_difference, axis=1)
    
    # Function to check Revised Place of supply and Revised Recipients GSTIN
    def compute_pos_recipient_gstin_check(row):
        revised_place_of_supply = row.iloc[17]
        revised_recipient_gstin = row.iloc[14]
        
        if pd.isna(revised_place_of_supply):
            if pd.isna(revised_recipient_gstin):
                return "Revised Place of supply and Revised GSTIN of recipients should not be blank"
            else:
                return "Revised Place of supply should not be blank"
        elif pd.isna(revised_recipient_gstin):
                return "Revised GSTIN of recipients should not be blank"
        elif revised_place_of_supply == revised_recipient_gstin:
            return "Match"
        else:
            return "Incorrect POS need to check"

    df3['POS & Recipient GSTIN Check'] = df3.apply(compute_pos_recipient_gstin_check, axis=1)
    
    # Function to check Unusual transation by Revised HSN
    def identify_unusual_transaction_by_hsn(row):
        revised_hsn = row['Revised HSN']
        
        if revised_hsn == "9997":
            return "This sort of recovery made need to check the transaction"
        elif revised_hsn == "9965":
            return "This sort of GTA Supply made need to check the transaction"
        elif revised_hsn == "996601":
            return "Motor vehicle provided on rent along with operator and cost of fuel is recovered in rent or Motor vehicle provided on rent along with operator but cost of fuel is not recovered in rent(Need to verify the GST Rate)"
        elif revised_hsn == "9973":
            return "Motor vehicle provided on rent without operator whether or not fuel cost is recovered in rent"
        elif revised_hsn == "8703":
            return "Prima facie it is sale of used car (Need to check the transaction)"
        elif revised_hsn == "9972":
            return "Prima facie it is renting of immovable property (Need to check the transaction)"
        elif revised_hsn == "4902":
            return "Prima facie it is supply of MEIS scripts (Need to check the transaction)"
        elif revised_hsn == "8471":
            return "Prima facie it is sale of used Laptops/Desktops (Need to check the transaction)"
        elif revised_hsn == "997331":
            return "Prima facie it is supply of Licensing services for the right to use computer software and databases (Need to check the transaction)"
        elif revised_hsn == "9954":
            return "Prima facie it is supply of works contract service (Need to check the transaction)"
        elif pd.isna(revised_hsn):
            return "HSN should not be blank"
        else:
            return "-"

    df3['Identification of unusal transaction by HSN'] = df3.apply(identify_unusual_transaction_by_hsn, axis=1)
    
    # Function to check Unusual transation by Description
    def identify_unusual_transaction_by_description(row):
        revised_description = row['Revised description']
        
        if pd.notna(revised_description):
            if "recovery" in revised_description:
                return "Prima facie it is observed that some recovery made by the Company"
            elif "reimb" in revised_description:
                return "Prima facie it is observed that some reimbursement made by the Company"
            elif "works contract" in revised_description:
                return "Prima facie it is observed that works contract service provided by the Company"
            elif "rent" in revised_description:
                return "Prima facie it is observed that renting service provided by the Company (need to check the transaction)"
            elif "scrap" in revised_description:
                return "Prima facie it is observed that scrap sale is made by the Company"
            elif "gift" in revised_description:
                return "Prima facie it is observed that gift provided by the Company"
            elif "dest" in revised_description:
                return "Prima facie it is observed that material destroyed and sale made by the Company"
            elif "stolen" in revised_description:
                return "Prima facie it is observed that the material is stolen in the Company"
            elif "lost" in revised_description:
                return "Prima facie it is observed that some material is lost in the Company"
            elif "disposed" in revised_description:
                return "Prima facie it is observed that inputs/Capital goods disposed by the Company"
            elif "free sample" in revised_description:
                return "Prima facie it is observed that free sample supply made by the Company (Need to check whether ITC on the same is reversed)"
            elif "written off" in revised_description:
                return "Prima facie it is observed that made by the Company"
            elif "cheque bounce" in revised_description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "damage" in revised_description:
                return "Prima facie it is observed that damage material sold by the Company"
            elif "penalty" in revised_description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "interest" in revised_description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "delay" in revised_description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "Interest" in revised_description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "Works" in revised_description:
                return "Prima facie it is observed that works contract service provided by the Company"
            elif "Recovery" in revised_description:
                return "Prima facie it is observed that some recovery made by the Company"
        
        return "Description should not be blank"

    df3['Identification of unusual transaction by Description'] = df3.apply(identify_unusual_transaction_by_description, axis=1)
    
    # To check if duplicate entries of any Origional document number is present
    df3['Count'] = df3.groupby('Original document number')['Original document number'].transform('count')
    df3['Origional Invoice Duplicates'] = df3['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df3.drop(columns=['Count'], inplace=True)
    
    # To check if duplicate entries of any Revised document number is present
    df3['Count'] = df3.groupby('Revised document number')['Revised document number'].transform('count')
    df3['Revised Invoice Duplicates'] = df3['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique' if x > 1 else '-')
    df3.drop(columns=['Count'], inplace=True)
    
    # Function to check Revised Reverse Charge
    def reverse_charge_check(row):
        revised_reverse_charge = row['Revised applicability of Reverse Charge']
        
        if revised_reverse_charge in ["Yes", "Y", "y", "yes"]:
            return "Prima facie it is observed that this transaction covered under reverse charge (Need to check)"
        elif pd.isna(revised_reverse_charge):
            return "It should not be blank"
        else:
            return "Forward Charge Supply"

    df3['Reverse Charge Check'] = df3.apply(reverse_charge_check, axis=1)
    
    # Function to check Bill to Ship party GSTIN
    def bill_to_ship_party_gstin_check(row):
        bill_to_gstin = row.iloc[14]
        ship_to_gstin = row.iloc[15]
        
        if bill_to_gstin == ship_to_gstin:
            return "Same"
        elif pd.isna(bill_to_gstin):
            return "It shouldn't be blank"
        else:
            return "It is observed that bill to GST number is different than Ship to GST number"

    df3['Bill to Ship party GSTIN Check'] = df3.apply(bill_to_ship_party_gstin_check, axis=1)
    
    # Function to check Revised shipping bill number
    def shipping_bill_check(row):
        shipping_bill_number = row['Revised shipping bill number']
        
        if pd.notna(shipping_bill_number):
            if len(str(shipping_bill_number)) == 7:
                return "Correct"
            else:
                return "Incorrect shipping bill details mentioned, need to correct the same"
        else:
            return "It should not be blank"

    df3['Shipping bill Check'] = df3.apply(shipping_bill_check, axis=1)
    
    # Function to check Revised port code
    def port_code_check(row):
        port_code = row['Revised port code']
        
        if pd.notna(port_code):
            if len(str(port_code)) == 6:
                return "Correct"
            else:
                return "Need to mention correct port code"
        else:
            return "It should not be blank"

    df3['Port Code Check'] = df3.apply(port_code_check, axis=1)
    
    df3.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Amendments(Invoices) sheet and Save in Amendments(Invoices) sheet
    df3 = pd.concat([df2, df3], axis=1)
    
    df3_columns = df3.columns.tolist()
    df3_new_columns = df3_columns[2:] + df3_columns[:2] # Move the first and second columns to the end
    df3 = df3[df3_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df3_column_to_extend = 'Start date'
    df3_value_to_extend = df3.at[1, df3_column_to_extend]
    df3[df3_column_to_extend] = df3[df3_column_to_extend].fillna(df3_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df3_column_to_extend1 = 'End date'
    df3_value_to_extend1 = df3.at[1, df3_column_to_extend1]
    df3[df3_column_to_extend1] = df3[df3_column_to_extend1].fillna(df3_value_to_extend1)
    
    df3['Original document date'] = pd.to_datetime(df3['Original document date'])
    df3['Revised document date'] = pd.to_datetime(df3['Revised document date'])
    df3['Start date'] = pd.to_datetime(df3['Start date'])
    df3['End date'] = pd.to_datetime(df3['End date'])

    # Function to check Origional Document Date as it should be between start and end date
    def origional_document_date_check(row):
        if pd.isna('Original document date'):
            return "Original document date is Blank"
        elif pd.notna('Original document date'):
            if row['Original document date'] > row['End date']:
                return "Origional Document Date is not pertaining to this FY"
            elif row['Original document date'] < row['Start date']:
                return "Origional Document Date is not pertaining to this FY"
            else:
                return "Correct"
        
    df3['Original document date check'] = df3.apply(origional_document_date_check, axis=1)
    
    # Function to check Revised Document Date as it should be between start and end date
    def revised_document_date_check(row):
        if pd.isna('Revised document date'):
            return "Revised document date is Blank"
        elif pd.notna('Revised document date'):
            if row['Revised document date'] > row['End date']:
                return "Revised document date is not pertaining to this FY"
            elif row['Revised document date'] < row['Start date']:
                return "Revised document date is not pertaining to this FY"
            else:
                return "Correct"
         
    df3['Revised document date check'] = df3.apply(revised_document_date_check, axis=1)
    
    df3 = df3.drop(columns=["Start date", "End date"]) # Remove specified columnns
    df3 = df3.dropna(subset=df3.columns[0:38], how='all') # Remove unnecessarily created checks even if rows not contains Data
    
    
    # -------------------------------------------* Processing of Outward supply Sheet *----------------------------------------------------
    
    df1 = df1.iloc[5:]
    # Step 3: Set the 6th row as the header.
    header_row = df1.iloc[0]
    df1 = df1[1:]
    df1.columns = header_row
    
    # to select data from 1st column to 29th column
    df1 = df1.iloc[:, 1:40]
    
    # Convert the "Document date" column to datetime, if it's not already
    df1['Document date'] = pd.to_datetime(df1['Document date'])

    # Extract only the date portion and overwrite the column
    df1['Document date'] = df1['Document date'].dt.date
    
    # To check if duplicate entries of any Document number is present
    result_df = df1.groupby('Document number').size().reset_index(name='Count')
    result_df['Invoice duplicates check'] = result_df['Count'].apply(lambda x: 'Unique' if x == 1 else 'Not Unique')
    df1 = df1.merge(result_df[['Document number', 'Invoice duplicates check']], on='Document number', how='left')
    print(df1)
    
    # Checking GSTN is less than 16 only if Status of recipient is Registered
    def check_gstn(status, gstin):
        if status == 'Registered':
            if len(str(gstin)) <= 15:
                return 'Correct'
            else:
                return 'Incorrect / GSTIN should be blank'
        else:
            return ''

    df1.loc['GSTN Check'] = df1.apply(lambda row: check_gstn(row.iloc[1], row.iloc[7]), axis=1)

    # Function to check type of supply
    def type_of_supply_check(type_of_supply):
        if pd.notna(type_of_supply):
            if "Regular" in type_of_supply:
                return "Correct"
            elif "Export with" in type_of_supply:
                return "It is an Zero rated supply CGST+SGST should be blank."
            elif "SEZ without" in type_of_supply:
                return "It is a Zero rated supply without payment all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%"
            elif "Export without" in type_of_supply:
                return "It is a Zero rated supply without payment all tax columns should be blank. Zero-rated supply under Goods and Services Tax (GST) refers to the supplies of goods or services that are taxable but have a GST rate of 0%"
            elif "SEZ with" in type_of_supply:
                return "It is an Zero rated supply CGST+SGST should be blank"
            elif "Exempt" in type_of_supply:
                return "It is an exempt supply tax value should be blank"
            else:
                return "Prima facie it observes that it is other than Regular supply"
        else:
            return "It should not be blank."

    df1['Type of Supply check'] = df1['Type of supply'].apply(type_of_supply_check)
    
    # Adding the "Invoice No Length Check" column
    df1['Invoice Check'] = df1['Document number'].apply(lambda x: 'It should not be blank' if pd.notna(x) and len(str(x)) > 0 else ('Correct' if len(str(x)) < 16 else 'Need to check the Invoice copy'))

    
    # Function to check HSN
    def hsn_check(hsn):
        if hsn > 99999999:
            return "Need to mention correct HSN"
        elif hsn > 999:
            return "Correct"
        elif hsn == 0:
            return "HSN should not be blank"
        else:
            return "Need to mention correct HSN"

    df1['HSN check'] = df1['HSN'].apply(hsn_check)
    
    # Function to check and convert Revised GST Rate
    def gst_rate_check(row):
        revised_rate = row['GST Rate (%)']
        
        if pd.notna(revised_rate):
            if revised_rate == 0.28:
                return 28
            elif revised_rate == 28:
                return 28
            elif revised_rate == 0.18:
                return 18
            elif revised_rate == 18:
                return 18
            elif revised_rate == 0.12:
                return 12
            elif revised_rate == 12:
                return 12
            elif revised_rate == 0.05:
                return 5
            elif revised_rate == 5:
                return 5
            elif revised_rate == 0.025:
                return 2.5
            elif revised_rate == 2.5:
                return 2.5
            elif revised_rate == 0.01:
                return 0.1
            elif revised_rate == 0.03:
                return 3
            elif revised_rate == 3:
                return 3
        
        return "Need to mention correct GST Rate"

    df1['GST Rate Check'] = df1.apply(gst_rate_check, axis=1)
    
    # Function to calculate GST Difference
    def compute_gst_difference(row):
        taxable_value = row.iloc[19]
        igst = row.iloc[20]
        cgst = row.iloc[21]
        sgst_utgst = row.iloc[22]
        gst_rate_check = row['GST Rate Check']
        
        # Check for NaN values and convert columns to numeric types
        if pd.notna(taxable_value) and pd.notna(gst_rate_check):
            # Handle the case where the conversion fails using try and except
            try:
                taxable_value_f = float(taxable_value)
                igst_f = float(igst)
                cgst_f = float(cgst)
                sgst_utgst_f = float(sgst_utgst)
                rate_check = float(gst_rate_check)
                
                return round((taxable_value_f * rate_check / 100) - (igst_f + cgst_f + sgst_utgst_f), 0)
            except ValueError:
                return "Revised taxable Value (Rs.)/ IGST/ Revised IGST/ Revised CGST/ Revised SGST/UTGST Should not be blank"  
        else:
            return None

    df1['GST Difference'] = df1.apply(compute_gst_difference, axis=1)
    
    mask = df1.iloc[:, 7].notna() # Create a boolean mask for non-null values
    df1.loc[mask, df1.columns[7]] = df1.loc[mask, df1.columns[7]].astype(str) # Convert non-null values to strings
    
    # Apply the .str.extract() method to the modified column
    df1['Recipients GSTIN - Copy'] = df1.loc[mask, df1.columns[7]].str.extract(r'(\d{2})')
    df1['Place Of Supply - Copy'] = df1['Place Of Supply'].str.extract(r'(\d{2})')
    
    # Function to check Revised Place of supply and Revised Recipients GSTIN
    def pos_recipient_check(row):
        pos_copy = row['Place Of Supply - Copy']
        gstin_copy = row['Recipients GSTIN - Copy']
        
        if pos_copy == gstin_copy:
            return "Match"
        elif pd.isna(pos_copy):
            if pd.isna(gstin_copy):
                return "Place of Supply and Recipients GSTIN need to check"
            else:
                return "Place of Supply need to check"
        elif pd.isna(gstin_copy):
                return "Recipients GSTIN need to check"

    df1['POS & Recipient check'] = df1.apply(pos_recipient_check, axis=1)
    
    # Function to check Unusual transaction by HSN
    def identify_unusual_transaction_by_hsn(hsn):
        if pd.isna(hsn):
            return "HSN should not be blank"
        elif hsn == 9997:
            return "This sort of recovery made need to check the transaction"
        elif hsn == 9965:
            return "This sort of GTA Supply made need to check the transaction"
        elif hsn == 996601:
            return "Motor vehicle provided on rent along with operator and cost of fuel is recovered in rent or Motor vehicle provided on rent along with operator but cost of fuel is not recovered in rent(Need to verify the GST Rate)"
        elif hsn == 9973:
            return "Motor vehicle provided on rent without operator whether or not fuel cost is recovered in rent"
        elif hsn == 8703:
            return "Prima facie it is sale of used car (Need to check the transaction)"
        elif hsn == 9972:
            return "Prima facie it is renting of immovable property (Need to check the transaction)"
        elif hsn == 4902:
            return "Prima facie it is supply of MEIS scripts (Need to check the transaction)"
        elif hsn == 8471:
            return "Prima facie it is sale of used Laptops/Desktops (Need to check the transaction)"
        elif hsn == 997331:
            return "Prima facie it is supply of Licensing services for the right to use computer software and databases (Need to check the transaction)"
        elif hsn == 9954:
            return "Prima facie it is supply of works contract service (Need to check the transaction)"
        else:
            return "-"  # Default case, modify this according to your requirement

    df1['Identification of Unusal Transaction by HSN'] = df1['HSN'].apply(identify_unusual_transaction_by_hsn)
    
    
    # Function to check Unusual transaction by Description
    def identify_unusual_description(description):
        if isinstance(description, str):
            if "recovery" in description:
                return "Prima facie it is observed that some recovery made by the Company"
            elif "reimb" in description:
                return "Prima facie it is observed that some reimbursement made by the Company"
            elif "works contract" in description:
                return "Prima facie it is observed that works contract service provided by the Company"
            elif "rent" in description:
                return "Prima facie it is observed that renting service provided by the Company (need to check the transaction)"
            elif "scrap" in description:
                return "Prima facie it is observed that scrap sale is made by the Company"
            elif "gift" in description:
                return "Prima facie it is observed that gift provided by the Company"
            elif "dest" in description:
                return "Prima facie it is observed that material destroyed and sale made by the Company"
            elif "stolen" in description:
                return "Prima facie it is observed that the material is stolen in the Company"
            elif "lost" in description:
                return "Prima facie it is observed that some material is lost in the Company"
            elif "disposed" in description:
                return "Prima facie it is observed that inputs/Capital goods disposed by the Company"
            elif "free sample" in description:
                return "Prima facie it is observed that free sample supply made by the Company (Need to check whether ITC on the same is reversed)"
            elif "written off" in description:
                return "Prima facie it is observed that made by the Company"
            elif "cheque bounce" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "damage" in description:
                return "Prima facie it is observed that damage material sold by the Company"
            elif "penalty" in description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "interest" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "delay" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "Interest" in description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "Works" in description:
                return "Prima facie it is observed that works contract service provided by the Company"
            elif "Recovery" in description:
                return "Prima facie it is observed that some recovery made by the Company"
            elif "Reimb" in description:
                return "Prima facie it is observed that some reimbursement made by the Company"
            elif "Rent" in description:
                return "Prima facie it is observed that renting service provided by the Company (need to check the transaction)"
            elif "Scrap" in description:
                return "Prima facie it is observed that scrap sale is made by the Company"
            elif "Gift" in description:
                return "Prima facie it is observed that gift provided by the Company"
            elif "Dest" in description:
                return "Prima facie it is observed that material destroyed and sale made by the Company"
            elif "Stolen" in description:
                return "Prima facie it is observed that the material is stolen in the Company"
            elif "Lost" in description:
                return "Prima facie it is observed that some material is lost in the Company"
            elif "Disposed" in description:
                return "Prima facie it is observed that inputs/Capital goods disposed by the Company"
            elif "Free sample" in description:
                return "Prima facie it is observed that free sample supply made by the Company (Need to check whether ITC on the same is reversed)"
            elif "Written off" in description:
                return "Prima facie it is observed that made by the Company"
            elif "Cheque" in description:
                return "Prima facie it is observed that recovery made by the Company"
            elif "Damage" in description:
                return "Prima facie it is observed that damage material sold by the Company"
            elif "Penalty" in description:
                return "Prima facie it is observed that penalty recovered by the Company"
            elif "Delay" in description:
                return "Prima facie it is observed that recovery made by the Company"
        return "-"

    df1['Identification of Unusual Transaction by Description'] = df1['Description'].apply(identify_unusual_description)
    
    # Function to check Bill to ship party GSTIN
    def bill_to_ship_party_gstin_check(row):
        if pd.isna(row.iloc[7]) or pd.isna(row.iloc[8]):
            return "-"
        elif row.iloc[7] == row.iloc[8]:
            return "Same"
        else:
            return "It is observed that bill to GST number is different than Ship to GST number"

    df1['Bill to Ship Party GSTIN Check'] = df1.apply(bill_to_ship_party_gstin_check, axis=1)
    
    # Function to check Taxability
    def taxability_check(row):
        # Convert 'Taxability' column to string data type within the function
        taxability = str(row["Taxability"])

        if "Exempt" in taxability:
            return "Exempt supply made by the Company which attracts reversal under rule 42 & 43 also tax amount should be zero. Exempt supply under Goods and Services Tax (GST) refers to the supply of goods or services that are not taxable under GST."
        elif taxability == "non GST":
            return "It is a No GST supply hence tax amount should be zero"
        else:
            return "Normal taxable supply"

    df1['Taxability Check'] = df1.apply(taxability_check, axis=1)

    # Function to check Reverse Charge
    def reverse_charge_check(row):
        value = row["Applicability of Reverse Charge"]
        if pd.isna(value):
            return "It should not be blank"
        elif value.lower() in ["yes", "y"]:
            return "Prima facie it is observed that this transaction is covered under reverse charge (Need to check)"
        else:
            return "Forward Charge Supply"

    df1['Reverse Charge Check'] = df1.apply(reverse_charge_check, axis=1)
    
    # Calculating length of Shipping bill number
    df1['SB Length'] = df1['Shipping bill number'].apply(lambda x: len(str(x)) if pd.notna(x) and str(x) != '-' else None)

    # Function to check Shipping bill number
    def shipping_bill_check(row):
        value = row["SB Length"]
        if pd.isna(value):
            return "-"
        elif value == 7:
            return "Correct"
        else:
            return "Incorrect shipping bill details mentioned, need to correct the same"

    df1['Shipping bill check'] = df1.apply(shipping_bill_check, axis=1)

    df1['Port Code Length'] = df1['Port code'].str.len()
    
    # Function to check port code length(Should be equa to 6)
    def port_code_check(row):
        value = row["Port Code Length"]
        if pd.isna(value):
            return "It should not be blank"
        elif value == 6:
            return "Correct"
        else:
            return "Need to mention correct port code"

    df1['Port code check'] = df1.apply(port_code_check, axis=1)
    df1 = df1.drop(columns=["SB Length", "Port Code Length"])
    
    print(df1.index.duplicated().any())
    
    df1.reset_index(inplace=True, drop=True)
    df2.reset_index(inplace=True, drop=True)
    # Merge Doc. Series and Outward supply sheet and Save in Outward supply sheet
    df1 = pd.concat([df2, df1], axis=1)
    
    df1_columns = df1.columns.tolist()
    df1_new_columns = df1_columns[2:] + df1_columns[:2] # Move the first and second columns to the end
    df1 = df1[df1_new_columns] # Reorganize the DataFrame columns
    
    # Extend Values for below cells  in Start date column
    df1_column_to_extend = 'Start date'
    df1_value_to_extend = df1.at[1, df1_column_to_extend]
    df1[df1_column_to_extend] = df1[df1_column_to_extend].fillna(df1_value_to_extend)
    
    # Extend Values for below cells  in End date column
    df1_column_to_extend1 = 'End date'
    df1_value_to_extend1 = df1.at[1, df1_column_to_extend1]
    df1[df1_column_to_extend1] = df1[df1_column_to_extend1].fillna(df1_value_to_extend1)
    
    df1['Document date'] = pd.to_datetime(df1['Document date'])
    df1['Start date'] = pd.to_datetime(df1['Start date'])
    df1['End date'] = pd.to_datetime(df1['End date'])

    # Function to check Document Date as it should be between start and end date
    def document_date_check(row):
        if pd.isna('Document date'):
            return "Document Date is Blank"
        elif pd.notna('Document date'):
            if row['Document date'] > row['End date']:
                return "Document date is not pertaining to this FY"
            elif row['Document date'] < row['Start date']:
                return "Document date is not pertaining to this FY"
            else:
                return "Correct"
        
    df1['Document Date check'] = df1.apply(document_date_check, axis=1)
    
    df1 = df1.drop(columns=["Start date", "End date"]) # Remove specified columnns
    df1 = df1.dropna(subset=df1.columns[0:38], how='all') # Remove unnecessarily created checks even if rows not contains Data
    # columns_to_remove_df1 = [20, 22, 26,28, 29, 37, 39, 41, 43, 45]
    # df1 = df1.drop(df1.columns[columns_to_remove_df1], axis=1)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Outward supply', index=False)
        df3.to_excel(writer, sheet_name='Amendments(Invoices)', index=False)
        df4.to_excel(writer, sheet_name='Debit&CreditNotes', index=False)
        df5.to_excel(writer, sheet_name='Amendments (CDN)', index=False)
        df6.to_excel(writer, sheet_name='Advances', index=False)
        df7.to_excel(writer, sheet_name='Amendment(Advances)', index=False)
    return file_path

def process_step2(file_path):
    xls = pd.ExcelFile(file_path)
    
    df1 = pd.read_excel(xls, sheet_name="Outward supply")
    df3 = pd.read_excel(xls, sheet_name="Amendments(Invoices)")
    df4 = pd.read_excel(xls, sheet_name="Debit&CreditNotes")
    df5 = pd.read_excel(xls, sheet_name="Amendments (CDN)")
    df6 = pd.read_excel(xls, sheet_name="Advances")
    df7 = pd.read_excel(xls, sheet_name="Amendment(Advances)")
    

    b2b = df1.loc[df1['Taxability'] == 'Taxable']

    # Dataframe will contain Selected columns only
    selected_columns_indices = [1, 2, 5, 6, 7, 9, 10, 12, 13, 18, 19, 21, 23, 25, 27, 29]  # Indices of the columns you want to select (0-based index)
    b2b = b2b.iloc[:, selected_columns_indices]

    b2b["Applicable % of Tax Rate"] = None

    # Selected entries from particular column will be deleted
    values_to_remove = ['Employee recoveries', 'Export with payment', 'Export without payment']
    b2b = b2b[~b2b['Type of supply'].isin(values_to_remove)]

    # Selected entries from particular column will be remained in the dataframe
    b2b = b2b.loc[b2b['Status of recipient'] == 'Registered']

    # Dataframe will contain Selected columns with the given order
    selected_columns_reordered = [4,5,2,3,15,6,7,16,1,8,9,10,14,11,12,13,0]  # Indices of the columns you want to select (0-based index)
    b2b = b2b.iloc[:, selected_columns_reordered]


    b2cs = df1.loc[df1['Taxability'] == 'Taxable']

    # Selected entries from particular column will be remained in the dataframe
    b2cs = b2cs[(b2cs['Type of supply'].isin(['Employee recoveries', 'Regular', 'Regular B2B'])) & (b2cs['Status of recipient'] == 'Unregistered')]

    b2cs = b2cs[b2cs.iloc[:, 29] <= 250000]

    b2cs["Applicable % of Tax Rate"] = None
    b2cs['blank column'] = None

    selected_columns_reordered1 = [67,10,18,66,19,27,13,21,23,25,2,1]  # Indices of the columns you want to select (0-based index)
    b2cs = b2cs.iloc[:, selected_columns_reordered1]

    # Selected entries from particular column will be remained in the dataframe
    b2cl = df1.loc[df1['Taxability'] == 'Taxable']
    # Selected entries from particular column will be remained in the dataframe
    b2cl = b2cl[(b2cl['Type of supply'] == 'Regular') | (b2cl['Type of supply'] == 'Regular B2B')]
    b2cl = b2cl.loc[b2cl['Status of recipient'] == 'Unregistered']

    b2cl["Applicable % of Tax Rate"] = None

    def categorize_type(row):
        if "Regular" in row['Type of supply']:
            return "Unreg/Regular"
        elif "Regular B2B" in row['Type of supply']:
            return "Unreg/Regular"
        elif "Export with payment" in row['Type of supply']:
            return "EWP/EWOP"
        elif "Export without payment" in row['Type of supply']:
            return "EWP/EWOP"
        else:
            return None

    b2cl['Type'] = b2cl.apply(categorize_type, axis=1)

    def categorize_custom(row):
        if row['Type'] == 'EWP/EWOP':
            return row['Type']
        elif row['Invoice value (Rs.)'] >= 250000:
            return row['Type']
        else:
            return None

    b2cl['Custom'] = b2cl.apply(categorize_custom, axis=1)

    b2cl = b2cl[b2cl['Custom'].notnull()]

    selected_columns_reordered2 = [5,6,29,10,66,18,19,27,13,21,23,25,2,1]  # Indices of the columns you want to select (0-based index)
    b2cl = b2cl.iloc[:, selected_columns_reordered2]

    b2ba = df3.copy()
    b2ba["Applicable % of Tax Rate"] = None

    b2ba = b2ba[
    (b2ba['Revised type of supply'].notnull()) &
    (b2ba['Revised type of supply'] != 'Export with payment') &
    (b2ba['Revised type of supply'] != 'Export without payment') &
    (b2ba['Revised type of supply'] != 'Highseas sale') &
    (b2ba['Revised status of recipient'] == 'Registered')
    ]
    
    selected_columns_reordered2 = [14,16,5,6,12,13,31,17,19,61,9,20,25,26,30,27,28,29,1,2,8]  # Indices of the columns you want to select (0-based index)
    b2ba = b2ba.iloc[:, selected_columns_reordered2]

    b2cla = df3.copy()
    b2cla["Applicable % of Tax Rate"] = None

    b2cla = b2cla[
    (b2cla['Revised type of supply'].notnull()) |
    (b2cla['Revised type of supply'] == 'Export with payment') |
    (b2cla['Revised type of supply'] == 'Export without payment') |
    (b2cla['Revised type of supply'] == 'Regular') |
    (b2cla['Revised type of supply'] == 'Regular B2B')
    ]

    b2cla = b2cla[(b2cla['Revised status of recipient'] == 'Unregistered')]

    def categorize_type(row):
        if row['Revised type of supply'] == 'Regular':
            return 'Unregistered/Regular'
        elif row['Revised type of supply'] == 'Regular B2B':
            return 'Unregistered/Regular'
        elif row['Revised type of supply'] == 'Export with payment':
            return 'EWP/EWOP'
        elif row['Revised type of supply'] == 'Export without payment':
            return 'EWP/EWOP'
        else:
            return None

    b2cla['Type'] = b2cla.apply(categorize_type, axis=1)

    def categorize_custom(row):
        if row['Type'] == 'EWP/EWOP':
            return row['Type']
        elif row['Revised Invoice value (Rs.)'] >= 250000:
            return row['Type']
        else:
            return None

    b2cla['Custom'] = b2cla.apply(categorize_custom, axis=1)

    b2cla = b2cla[b2cla['Custom'].notnull()]

    selected_columns_reordered3 = [5,6,17,12,13,31,25,26,27,28,29,30,8,9,61]  # Indices of the columns you want to select (0-based index)
    b2cla = b2cla.iloc[:, selected_columns_reordered3]

    exp=df1.copy()

    exp = exp[(exp['Type of supply'] == 'Export with payment') | (exp['Type of supply'] == 'Export without payment')]

    selected_columns_reordered3 = [2,5,6,29,32,30,31,18,19,27,21,23,25,1]  # Indices of the columns you want to select (0-based index)
    exp = exp.iloc[:, selected_columns_reordered3]

    exemp = df1.loc[df1['Taxability'] == 'Taxable']
    selected_columns_reordered4 = [1,2,3,19]  # Indices of the columns you want to select (0-based index)
    exemp = exemp.iloc[:, selected_columns_reordered4]

    expa = df3.copy()

    expa = expa[(expa['Revised type of supply'] == 'Export with payment') | (expa['Revised type of supply'] == 'Export without payment') | (expa['Revised type of supply'] == 'WOPAY') | (expa['Revised type of supply'] == 'WPAY')]

    selected_columns_reordered4 = [11,5,6,12,13,31,34,32,25,26,27,30,33]  # Indices of the columns you want to select (0-based index)
    expa = expa.iloc[:, selected_columns_reordered4]

    hsn_OUTWARD = df1.copy()

    hsn_OUTWARD.at[0, "Type"] = 'HSN'

    sum_tv = pd.to_numeric(hsn_OUTWARD.iloc[:, 19]).sum()
    hsn_OUTWARD.at[0, 'Taxable_count'] = sum_tv

    sum_igst = pd.to_numeric(hsn_OUTWARD.iloc[:, 21]).sum()
    hsn_OUTWARD.at[0, 'IGST_count'] = sum_igst

    sum_cgst = pd.to_numeric(hsn_OUTWARD.iloc[:, 23]).sum()
    hsn_OUTWARD.at[0, 'CGST_count'] = sum_cgst

    sum_sgst = pd.to_numeric(hsn_OUTWARD.iloc[:, 25]).sum()
    hsn_OUTWARD.at[0, 'SGST_count'] = sum_sgst

    hsn_OUTWARD['Total duty'] = hsn_OUTWARD['IGST_count'] + hsn_OUTWARD['CGST_count'] + hsn_OUTWARD['SGST_count']

    hsn_OUTWARD['Total value'] = hsn_OUTWARD['Taxable_count'] + hsn_OUTWARD['IGST_count'] + hsn_OUTWARD['CGST_count'] + hsn_OUTWARD['SGST_count']

    selected_columns_reordered5 = [66,67,68,69,70,71,72]  # Indices of the columns you want to select (0-based index)
    hsn_OUTWARD = hsn_OUTWARD.iloc[:, selected_columns_reordered5]

    hsn_CDNR = df4.copy()

    selected_columns_reordered6 = [9,16,17,18,19]  # Indices of the columns you want to select (0-based index)
    hsn_CDNR = hsn_CDNR.iloc[:, selected_columns_reordered6]

    numeric_columns = hsn_CDNR.columns[1:5]
    hsn_CDNR[numeric_columns] = hsn_CDNR[numeric_columns].apply(pd.to_numeric)

    # Group by 'Note type' and calculate the sums for each column
    hsn_CDNR = hsn_CDNR.groupby('Note type')[numeric_columns].sum().reset_index()

    new_column_names = ['Type', 'Taxable_count', 'IGST_count', 'CGST_count', 'SGST_count']
    # Assign new column names to the DataFrame
    hsn_CDNR.columns = new_column_names


    # Add a new column 'Total duty' by summing up 'IGST_count', 'CGST_count', and 'SGST_count'
    hsn_CDNR['Total duty'] = hsn_CDNR['IGST_count'] + hsn_CDNR['CGST_count'] + hsn_CDNR['SGST_count']

    # Add a new column 'Total value' by summing up 'Taxable_count', 'IGST_count', 'CGST_count', and 'SGST_count'
    hsn_CDNR['Total value'] = hsn_CDNR['Taxable_count'] + hsn_CDNR['IGST_count'] + hsn_CDNR['CGST_count'] + hsn_CDNR['SGST_count']


    docs_1 = df1['Document number']

    # Sort the Series using a lambda function to handle mixed data types
    docs_1 = sorted(docs_1, key=lambda x: (int(x) if str(x).isdigit() else float('inf'), x))
    # Convert the sorted list back to a pandas Series
    docs_1 = pd.Series(docs_1)
    # Convert the sorted Series back to a DataFrame, reset index, and rename the columns
    docs_1 = docs_1.to_frame().reset_index(drop=True)
    docs_1.columns = ['Document number']  # Replace column name

    docs_1 = docs_1.drop_duplicates()

    # Replace "-" with an empty string in the specified column
    docs_1['Document number'] = docs_1['Document number'].replace("-", "", regex=True)
    docs_1['Document number'] = docs_1['Document number'].replace("/", "", regex=True)

    # Convert the column to strings
    docs_1['Document number'] = docs_1['Document number'].astype(str)

    # Split the column into two columns at positions 0 and 4
    docs_1['Invoice Number - Copy.1'] = docs_1['Document number'].astype(str).str[:-4]  # Extract without last four digits
    docs_1['Invoice Number - Copy.2'] = docs_1['Document number'].astype(str).str[-4:]

    # Drop the original column
    docs_1.drop(columns=['Document number'], inplace=True)

    docs_1 = (
    docs_1.groupby('Invoice Number - Copy.1')
    .agg(start=('Invoice Number - Copy.2', 'min'),
         end=('Invoice Number - Copy.2', 'max'),
         count=('Invoice Number - Copy.2', 'size'))
    .reset_index()
    )

    docs_1.rename(columns={'Invoice Number - Copy.1': 'Invoice Number',}, inplace=True)

    # Assuming 'grouped_df' is your DataFrame
    docs_1['start'] = docs_1['start'].astype(float)  # Convert 'Start' column to numeric
    docs_1['end'] = docs_1['end'].astype(float)  # Convert 'End' column to numeric

    # Assuming 'grouped_df' is your DataFrame and 'Start', 'End', and 'Count' are columns in it
    docs_1['Cancelled'] = docs_1['end'] - docs_1['start'] - docs_1['count']

    # Assuming 'grouped_df' is your DataFrame
    docs_1['Cancelled'] = docs_1['Cancelled'].replace(-1, 0)


    docs_2 = df4['Document number']
    docs_2 = docs_2.drop_duplicates()

    docs_2 = docs_2.to_frame().reset_index(drop=True)
    docs_2.columns = ['Document number']  # Replace column name

    # Replace "-" with an empty string in the specified column
    docs_2['Document number'] = docs_2['Document number'].replace("-", "", regex=True)
    docs_2['Document number'] = docs_2['Document number'].replace("/", "", regex=True)

    docs_2['Document number'] = docs_2['Document number'].apply(lambda x: str(x) if pd.notnull(x) else x)
    docs_2.sort_values(by='Document number', ascending=True, inplace=True, na_position='last')

    # Split the column into two columns at positions 0 and 4
    docs_2['Invoice Number - Copy.1'] = docs_2['Document number'].astype(str).str[:-4]  # Extract without last four digits
    docs_2['Invoice Number - Copy.2'] = docs_2['Document number'].astype(str).str[-4:]

    docs_2 = (
    docs_2.groupby('Invoice Number - Copy.1')
    .agg(start=('Invoice Number - Copy.2', 'min'),
         end=('Invoice Number - Copy.2', 'max'),
         count=('Invoice Number - Copy.2', 'size'))
    .reset_index()
    )

    docs_2.rename(columns={'Invoice Number - Copy.1': 'Invoice Number',}, inplace=True)

    # Assuming 'grouped_df' is your DataFrame
    docs_2['start'] = docs_2['start'].astype(float)  # Convert 'Start' column to numeric
    docs_2['end'] = docs_2['end'].astype(float)  # Convert 'End' column to numeric

    # Assuming 'grouped_df' is your DataFrame and 'Start', 'End', and 'Count' are columns in it
    docs_2['Cancelled'] = docs_2['end'] - docs_2['start'] - docs_2['count']

    # Assuming 'grouped_df' is your DataFrame
    docs_2['Cancelled'] = docs_2['Cancelled'].replace(-1, 0)

    docs_2 = docs_2.dropna(subset=['Invoice Number'])


    docs_3 = df6['Document number']
    docs_3 = docs_3.drop_duplicates()

    docs_3 = docs_3.to_frame().reset_index(drop=True)
    docs_3.columns = ['Document number']  # Replace column name

    # Replace "-" with an empty string in the specified column
    docs_3['Document number'] = docs_3['Document number'].replace("-", "", regex=True)
    docs_3['Document number'] = docs_3['Document number'].replace("/", "", regex=True)

    # Split the column into two columns at positions 0 and 4
    docs_3['Invoice Number - Copy.1'] = docs_3['Document number'].astype(str).str[:-4]  # Extract without last four digits
    docs_3['Invoice Number - Copy.2'] = docs_3['Document number'].astype(str).str[-4:]

    docs_3 = (
    docs_3.groupby('Invoice Number - Copy.1')
    .agg(start=('Invoice Number - Copy.2', 'min'),
         end=('Invoice Number - Copy.2', 'max'),
         count=('Invoice Number - Copy.2', 'size'))
    .reset_index()
    )

    docs_3.rename(columns={'Invoice Number - Copy.1': 'Invoice Number',}, inplace=True)

    # Assuming 'grouped_df' is your DataFrame
    docs_3['start'] = docs_3['start'].astype(float)  # Convert 'Start' column to numeric
    docs_3['end'] = docs_3['end'].astype(float)  # Convert 'End' column to numeric

    # Assuming 'grouped_df' is your DataFrame and 'Start', 'End', and 'Count' are columns in it
    docs_3['Cancelled'] = docs_3['end'] - docs_3['start'] - docs_3['count']

    # Assuming 'grouped_df' is your DataFrame
    docs_3['Cancelled'] = docs_3['Cancelled'].replace(-1, 0)

    docs_3.dropna(subset=['Invoice Number'], inplace=True)


    cdnr = df4.copy()
    cdnr["Applicability of Reverse charge"] = 'N'
    cdnr["blank1"] = None
    cdnr["blank2"] = None

    cdnr = cdnr[(cdnr['Type of supply'] != 'Export with payment') & (cdnr['Type of supply'] != 'Export without payment')]
    cdnr = cdnr[(cdnr['Status of recipient'] == 'Registered')]

    selected_columns_reordered7 = [7,37,10,11,9,8,36,2,21,38,15,16,17,18,19,20,5,6,1]  # Indices of the columns you want to select (0-based index)
    cdnr = cdnr.iloc[:, selected_columns_reordered7]

    cdnur = df4.copy()
    cdnur["Applicability of Reverse charge"] = 'N'
    cdnur["blank1"] = None
    cdnur["blank2"] = None

    cdnur = cdnur[
    (cdnur['Type of supply'] == 'Export with payment') |
    (cdnur['Type of supply'] == 'Export without payment') |
    (cdnur['Type of supply'] == 'Regular') |
    (cdnur['Type of supply'] == 'Regular B2B')
    ]
    
    cdnur = cdnur[(cdnur['Status of recipient'] == 'Unregistered')]

    cdnur['Type'] = cdnur['Type of supply'].apply(lambda x: 'Unreg/Regular' if 'Regular' in x else 
                                            'Unreg/Regular' if 'Regular B2B' in x else 
                                            'EWP/EWOP' if 'Export with payment' in x else 
                                            'EWP/EWOP' if 'Export without payment' in x else 
                                            None)
    
    cdnur['Custom'] = cdnur.apply(lambda row: row['Type'] if row['Type'] == 'EWP/EWOP' else
                                       row['Type'] if row['Invoice value (Rs.)'] >= 250000 else
                                       None, axis=1)
    
    cdnur.dropna(subset=['Custom'], inplace=True)

    selected_columns_reordered8 = [2,10,11,9,8,21,37,15,16,20,17,18,19,5,6,1]  # Indices of the columns you want to select (0-based index)
    cdnur = cdnur.iloc[:, selected_columns_reordered8]

    cdnur_b2cs = df4.copy()
    cdnur_b2cs["Applicability of Reverse charge"] = 'N'
    cdnur_b2cs["blank1"] = None
    cdnur_b2cs["blank2"] = None

    cdnur_b2cs = cdnur_b2cs[
    (cdnur_b2cs['Type of supply'] == 'Export with payment') |
    (cdnur_b2cs['Type of supply'] == 'Export without payment') |
    (cdnur_b2cs['Type of supply'] == 'Regular') |
    (cdnur_b2cs['Type of supply'] == 'Regular B2B')
    ]

    cdnur_b2cs = cdnur_b2cs[(cdnur_b2cs['Status of recipient'] == 'Unregistered')]

    cdnur_b2cs['Type'] = cdnur_b2cs['Type of supply'].apply(lambda x: 'Unreg/Regular' if 'Regular' in x else 
                                            'Unreg/Regular' if 'Regular B2B' in x else 
                                            'EWP/EWOP' if 'Export with payment' in x else 
                                            'EWP/EWOP' if 'Export without payment' in x else 
                                            None)
    
    cdnur_b2cs['Custom'] = cdnur_b2cs.apply(lambda row: row['Type'] if row['Type'] == 'EWP/EWOP' else
                                       row['Type'] if row['Invoice value (Rs.)'] <= 250000 else
                                       None, axis=1)

    cdnur_b2cs = cdnur_b2cs[cdnur_b2cs['Custom'] == 'Unreg/Regular']

    selected_columns_reordered9 = [2,10,11,9,8,21,37,15,16,20,17,18,19,5,6,1]  # Indices of the columns you want to select (0-based index)
    cdnur_b2cs = cdnur_b2cs.iloc[:, selected_columns_reordered9]

    cdnra = df5.copy()

    cdnra['Place of Supply'] = cdnra['Revised GSTIN of recipient'].str[:2]
    cdnra["Reverse Charge"] = None
    cdnra["Blank1"] = None
    cdnra["Applicable Rate"] = None

    cdnra = cdnra[(cdnra['Original status of recipient'] == 'Unregistered')]

    selected_columns_reordered10 = [13,48,11,12,15,16,14,46,47,2,26,49,20,21,25,22,23,24,1]  # Indices of the columns you want to select (0-based index)
    cdnra = cdnra.iloc[:, selected_columns_reordered10]

    cdnura = df5.copy()
    cdnura['Place of Supply'] = None
    cdnura["blank1"] = None

    cdnura = cdnura[
        ((cdnura['Original type of supply'] == 'Export with payment') |
        (cdnura['Original type of supply'] == 'Export without payment') |
        (cdnura['Original type of supply'] == 'Regular') |
        (cdnura['Original type of supply'] == 'Regular B2B'))
    ]
    cdnura = cdnura[(cdnura['Original status of recipient'] == 'Unregistered')]

    cdnura['Type'] = cdnura['Original type of supply'].apply(lambda x: 'Unreg/Regular' if x in ['Regular', 'Regular B2B'] else
                                                           'EWP/EWOP' if x in ['Export with payment', 'Export without payment'] else
                                                           None)
    
    cdnura['Custom'] = cdnura.apply(lambda row: row['Type'] if row['Type'] == 'EWP/EWOP' else
                                       row['Type'] if row['Invoice value (Rs.)'] >= 250000 else
                                       None, axis=1)

    cdnura.dropna(subset=['Custom'], inplace=True)

    selected_columns_reordered11 = [2,9,10,15,16,14,46,26,47,25,20,21,22,1]  # Indices of the columns you want to select (0-based index)
    cdnura = cdnura.iloc[:, selected_columns_reordered11]

    at = df6.copy()
    at["Applicable Rate"] = None

    selected_columns_reordered12 = [9,52,13,14,22,16,18,20,2,4,5]  # Indices of the columns you want to select (0-based index)
    at = at.iloc[:, selected_columns_reordered12]

    atadj = df6.copy()
    atadj["Applicable Rate"] = None

    selected_columns_reordered13 = [9,13,52,15,35,29,32,21,2,24,26,25]  # Indices of the columns you want to select (0-based index)
    atadj = atadj.iloc[:, selected_columns_reordered13]

    ata = df7.copy()
    ata["Applicable Rate"] = None

    selected_columns_reordered14 = [5,15,44,19,20,24,21,22,23,6,4]  # Indices of the columns you want to select (0-based index)
    ata = ata.iloc[:, selected_columns_reordered14]

    atadja = df7.copy()
    atadja["Applicable Rate"] = None

    selected_columns_reordered15 = [5,15,19,44,20,24,21,22,23,4]  # Indices of the columns you want to select (0-based index)
    atadja = atadja.iloc[:, selected_columns_reordered15]

    docs = pd.concat([docs_1, docs_2, docs_3], ignore_index=True)
    hsn = pd.concat([hsn_CDNR, hsn_OUTWARD], ignore_index=True)

    hsn.iloc[:, 1:] = hsn.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
    
    transposed_hsn = hsn.T
    transposed_hsn.columns = transposed_hsn.iloc[0]  # Promote the first row to headers
    transposed_hsn = transposed_hsn[1:]

    transposed_hsn['Total HSN'] = transposed_hsn['HSN'] + transposed_hsn['Credit note'] - transposed_hsn['Debit note']
    selected_columns_reordered16 = ['HSN', 'Credit note', 'Debit note', 'Total HSN']  # Indices of the columns you want to select (0-based index)
    transposed_hsn = transposed_hsn[selected_columns_reordered16]
    hsn = transposed_hsn.T
    hsn['Type'] = ['HSN', 'Credit note', 'Debit note', 'Total HSN']
    selected_columns_reordered17 = ['Type', 'Taxable_count', 'IGST_count', 'CGST_count', 'SGST_count', 'Total duty', 'Total value']  # Indices of the columns you want to select (0-based index)
    hsn = hsn[selected_columns_reordered17]
    hsn.reset_index(drop=True, inplace=True)

    source_outward_supply = df1.copy()
    selected_columns_reordered18 = [14,15,16,17,18,19,21,23,25,27,29]
    source_outward_supply = source_outward_supply.iloc[:, selected_columns_reordered18]

    # Group by multiple columns and calculate sum for each group
    grouped_df = source_outward_supply.groupby(['HSN', 'Unit Quantity Code', 'GST Rate (%)']).agg({
        source_outward_supply.columns[3]: 'sum',   # Quantity
        source_outward_supply.columns[5]: 'sum',   # Taxable Value (Rs.)
        source_outward_supply.columns[6]: 'sum',   # IGST #(lf)(Rs.)
        source_outward_supply.columns[7]: 'sum',   # CGST #(lf)(Rs.)
        source_outward_supply.columns[8]: 'sum',   # SGST/UTGST #(lf)(Rs.)
        source_outward_supply.columns[9]: 'sum'    # Cess Amount#(lf)(Rs.)
    }).reset_index()

    # Rename the columns to match the M code
    grouped_df.rename(columns={
        source_outward_supply.columns[3]: 'Sum of Quantity',
        source_outward_supply.columns[5]: 'Sum of Taxable Value (Rs.)',
        source_outward_supply.columns[6]: 'Sum of IGST (Rs.)',
        source_outward_supply.columns[7]: 'Sum of CGST (Rs.)',
        source_outward_supply.columns[8]: 'Sum of SGST/UGST (Rs.)',
        source_outward_supply.columns[9]: 'Sum of Cess Amount (Rs.)'
    }, inplace=True)

    # Update the original DataFrame with the grouped and aggregated data
    source_outward_supply = grouped_df.copy()

    source_debit_credit_notes = df4.copy()
    selected_columns_reordered19 = [9,12,13,14,15,16,17,18,19,20,21]
    source_debit_credit_notes = source_debit_credit_notes.iloc[:, selected_columns_reordered19]

    # Group by multiple columns and calculate sum for each group
    grouped_df2 = source_debit_credit_notes.groupby(['Note type','HSN ', 'UQC', 'Rate (%)']).agg({
        source_debit_credit_notes.columns[3]: 'sum',   # Quantity
        source_debit_credit_notes.columns[5]: 'sum',   # Taxable Value (Rs.)
        source_debit_credit_notes.columns[6]: 'sum',   # IGST #(lf)(Rs.)
        source_debit_credit_notes.columns[7]: 'sum',   # CGST #(lf)(Rs.)
        source_debit_credit_notes.columns[8]: 'sum',   # SGST/UTGST #(lf)(Rs.)
        source_debit_credit_notes.columns[9]: 'sum'    # Cess Amount#(lf)(Rs.)
    }).reset_index()

    # Rename the columns to match the M code
    grouped_df2.rename(columns={
        source_debit_credit_notes.columns[3]: 'Sum of Quantity',
        source_debit_credit_notes.columns[5]: 'Sum of Taxable Value (Rs.)',
        source_debit_credit_notes.columns[6]: 'Sum of IGST (Rs.)',
        source_debit_credit_notes.columns[7]: 'Sum of CGST (Rs.)',
        source_debit_credit_notes.columns[8]: 'Sum of SGST/UGST (Rs.)',
        source_debit_credit_notes.columns[9]: 'Sum of Cess Amount (Rs.)'
    }, inplace=True)

    # Update the original DataFrame with the grouped and aggregated data
    source_debit_credit_notes = grouped_df2.copy()



    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        source_outward_supply.to_excel(writer, sheet_name='Source_Outward_supply', index=False)
        source_debit_credit_notes.to_excel(writer, sheet_name='Source_Debit_Credit_Notes', index=False)
        hsn.to_excel(writer, sheet_name='HSN', index=False)
        b2b.to_excel(writer, sheet_name='b2b,sez,de', index=False)
        b2ba.to_excel(writer, sheet_name='b2ba', index=False)
        cdnur_b2cs.to_excel(writer, sheet_name='cdnur_b2cs', index=False)
        b2cs.to_excel(writer, sheet_name='b2cs', index=False)
        b2cl.to_excel(writer, sheet_name='b2cl', index=False)
        b2cla.to_excel(writer, sheet_name='b2cla', index=False)
        cdnura.to_excel(writer, sheet_name='cdnura', index=False)
        cdnr.to_excel(writer, sheet_name='cdnr', index=False)
        cdnra.to_excel(writer, sheet_name='cdnra', index=False)
        cdnur.to_excel(writer, sheet_name='cdnur', index=False)
        exemp.to_excel(writer, sheet_name='exemp', index=False)
        exp.to_excel(writer, sheet_name='exp', index=False)
        expa.to_excel(writer, sheet_name='expa', index=False)
        docs.to_excel(writer, sheet_name='docs', index=False)
        at.to_excel(writer, sheet_name='at', index=False)
        ata.to_excel(writer, sheet_name='ata', index=False)
        atadj.to_excel(writer, sheet_name='atadj', index=False)
        atadja.to_excel(writer, sheet_name='atadja', index=False)
        # hsn_OUTWARD.to_excel(writer, sheet_name='hsn_OUTWARD', index=False)
        # hsn_CDNR.to_excel(writer, sheet_name='hsn_CDNR', index=False)
        # docs_1.to_excel(writer, sheet_name='docs (1)', index=False)
        # docs_2.to_excel(writer, sheet_name='docs (2)', index=False)
        # docs_3.to_excel(writer, sheet_name='docs (3)', index=False)

    return file_path

def compare_excel_files(company_file_path, government_file_path):
    company = pd.ExcelFile(company_file_path)
    government = pd.ExcelFile(government_file_path)
    
    df1 = pd.read_excel(company, sheet_name="b2b,sez,de")
    df2 = pd.read_excel(company, sheet_name="cdnur")
    df3 = pd.read_excel(company, sheet_name="cdnr")
    df4 = pd.read_excel(company, sheet_name="exp")
    df5 = pd.read_excel(government, sheet_name="b2b, sez, de")
    df6 = pd.read_excel(government, sheet_name="cdnur")
    df7 = pd.read_excel(government, sheet_name="cdnr")
    df8 = pd.read_excel(government, sheet_name="exp")

    df5 = df5.iloc[2:]
    # Step 3: Set the 6th row as the header.
    header_row = df5.iloc[0]
    df5 = df5[1:]
    df5.columns = header_row
    b2b_match = pd.merge(
        df1,
        df5,
        how='outer',
        left_on=["Document number", df1.columns[0]],
        right_on=["Invoice number", "GSTIN/UIN of Recipient"],
        suffixes=('_b2b', '_b2b_govt')
    )

    selected_columns_reordered = [0,2,8,11,13,14,15,17,19,25,28,29,30,31]
    b2b_match = b2b_match.iloc[:, selected_columns_reordered]

    # Define a function to apply the conditions and generate remarks
    def generate_remark(row):
        if pd.isnull(row.iloc[0]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row.iloc[7]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[0] == row.iloc[7]:
            return "Matched"
        else:
            return "Unmatched / GSTIN"

    # Add a new column "Remark GSTIN of recipient" using the custom function
    b2b_match["Remark GSTIN of recipient"] = b2b_match.apply(generate_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark Taxable Value"
    def generate_taxable_value_remark(row):
        if pd.isnull(row["Taxable Value (Rs.)"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Taxable Value"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Taxable Value (Rs.)"] == row["Taxable Value"]:
            return "Matched"
        else:
            return "Unmatched / Taxable value"

    # Add a new column "Remark Taxable Value" using the custom function
    b2b_match["Remark Taxable Value"] = b2b_match.apply(generate_taxable_value_remark, axis=1)

    def generate_igst_remark(row):
        if pd.isnull(row.iloc[13]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Integrated Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[13] == row["Integrated Tax"]:
            return "Matched"
        else:
            return "Unmatched / IGST"

    # Add a new column "Remark IGST" using the custom function
    b2b_match["Remark IGST"] = b2b_match.apply(generate_igst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark CGST"
    def generate_cgst_remark(row):
        if pd.isnull(row.iloc[14]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Central Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[13] == row["Central Tax"]:
            return "Matched"
        else:
            return "Unmatched / CGST"

    # Add a new column "Remark CGST" using the custom function
    b2b_match["Remark CGST"] = b2b_match.apply(generate_cgst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark SGST"
    def generate_sgst_remark(row):
        if pd.isnull(row.iloc[15]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["State/UT Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[15] == row["State/UT Tax"]:
            return "Matched"
        else:
            return "Unmatched / SGST"

    # Add a new column "Remark SGST" using the custom function
    b2b_match["Remark SGST"] = b2b_match.apply(generate_sgst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark Type of Supply"
    def generate_type_of_supply_remark(row):
        if pd.isnull(row["Type of supply"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Invoice Type"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Type of supply"] == row["Invoice Type"]:
            return "Matched"
        else:
            return "Unmatched / Type of supply"

    # Add a new column "Remark Type of Supply" using the custom function
    b2b_match["Remark Type of Supply"] = b2b_match.apply(generate_type_of_supply_remark, axis=1)

    b2b_match["Taxable Value (Rs.)"] = pd.to_numeric(b2b_match["Taxable Value (Rs.)"], errors='coerce')
    b2b_match["Taxable Value"] = pd.to_numeric(b2b_match["Taxable Value"], errors='coerce')
    b2b_match.iloc[:, 13] = pd.to_numeric(b2b_match.iloc[:, 13], errors='coerce')
    b2b_match["Integrated Tax"] = pd.to_numeric(b2b_match["Integrated Tax"], errors='coerce')
    b2b_match.iloc[:, 14] = pd.to_numeric(b2b_match.iloc[:, 14], errors='coerce')
    b2b_match["Central Tax"] = pd.to_numeric(b2b_match["Central Tax"], errors='coerce')
    b2b_match.iloc[:, 15] = pd.to_numeric(b2b_match.iloc[:, 15], errors='coerce')
    b2b_match["State/UT Tax"] = pd.to_numeric(b2b_match["State/UT Tax"], errors='coerce')

    # Define a function to calculate the "Amount difference"
    def calculate_amount_difference(row):
        taxable_value_diff = row["Taxable Value (Rs.)"] - row["Taxable Value"]
        igst_diff = row.iloc[13] - row["Integrated Tax"]
        cgst_diff = row.iloc[14] - row["Central Tax"]
        sgst_diff = row.iloc[15] - row["State/UT Tax"]
        
        # Check for NaN values before rounding
        taxable_value_diff = 0 if pd.isna(taxable_value_diff) else taxable_value_diff
        igst_diff = 0 if pd.isna(igst_diff) else igst_diff
        cgst_diff = 0 if pd.isna(cgst_diff) else cgst_diff
        sgst_diff = 0 if pd.isna(sgst_diff) else sgst_diff
        
        # Round the sum of differences to the nearest integer
        amount_difference = round(taxable_value_diff + igst_diff + cgst_diff + sgst_diff)
        return amount_difference

    # Add a new column "Amount difference" using the custom function
    b2b_match["Amount difference"] = b2b_match.apply(calculate_amount_difference, axis=1)

    df6 = df6.iloc[2:]
    # Step 3: Set the 6th row as the header.
    header_row = df6.iloc[0]
    df6 = df6[1:]
    df6.columns = header_row

    cdnur_match = pd.merge(
        df2,
        df6,
        how='outer',
        left_on=["Type of supply", "Note number"],
        right_on=["UR Type", "Note Number"],
        suffixes=('_cdnur', '_cdnur_govt')
    )

    selected_columns_reordered2 = [0,1,3,8,10,11,12,16,17,19,24,25]
    cdnur_match = cdnur_match.iloc[:, selected_columns_reordered2]

    # Define a function to apply the conditions and generate remarks for "Remark Taxable Value"
    def generate_taxable_value_remark(row):
        if pd.isnull(row["Taxable Value (Rs.)"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Taxable Value"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Taxable Value (Rs.)"] == row["Taxable Value"]:
            return "Matched"
        else:
            return "Unmatched / Taxable value"

    # Add a new column "Remark Taxable Value" using the custom function
    cdnur_match["Remark Taxable Value"] = cdnur_match.apply(generate_taxable_value_remark, axis=1)

    def generate_igst_remark(row):
        if pd.isnull(row.iloc[4]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Integrated Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[4] == row["Integrated Tax"]:
            return "Matched"
        else:
            return "Unmatched / IGST"

    # Add a new column "Remark IGST" using the custom function
    cdnur_match["Remark IGST"] = cdnur_match.apply(generate_igst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark Type of Supply"
    def generate_type_of_supply_remark(row):
        if pd.isnull(row["Type of supply"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["UR Type"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Type of supply"] == row["UR Type"]:
            return "Matched"
        else:
            return "Unmatched / Type of supply"

    # Add a new column "Remark Type of Supply" using the custom function
    cdnur_match["Remark Type of Supply"] = cdnur_match.apply(generate_type_of_supply_remark, axis=1)

    cdnur_match["Taxable Value (Rs.)"] = pd.to_numeric(cdnur_match["Taxable Value (Rs.)"], errors='coerce')
    cdnur_match["Taxable Value"] = pd.to_numeric(cdnur_match["Taxable Value"], errors='coerce')
    cdnur_match.iloc[:, 4] = pd.to_numeric(cdnur_match.iloc[:, 4], errors='coerce')
    cdnur_match["Integrated Tax"] = pd.to_numeric(cdnur_match["Integrated Tax"], errors='coerce')
    cdnur_match.iloc[:, 5] = pd.to_numeric(cdnur_match.iloc[:, 5], errors='coerce')
    # cdnur_match["Central Tax"] = pd.to_numeric(cdnur_match["Central Tax"], errors='coerce')
    cdnur_match.iloc[:, 6] = pd.to_numeric(cdnur_match.iloc[:, 6], errors='coerce')
    # cdnur_match["State/UT Tax"] = pd.to_numeric(cdnur_match["State/UT Tax"], errors='coerce')

    # Define a function to calculate the "Amount difference"
    def calculate_amount_difference(row):
        taxable_value_diff = row["Taxable Value (Rs.)"] - row["Taxable Value"]
        igst_diff = row.iloc[4] - row["Integrated Tax"]
        cgst_diff = row.iloc[5]
        sgst_diff = row.iloc[6]
        
        # Check for NaN values before rounding
        taxable_value_diff = 0 if pd.isna(taxable_value_diff) else taxable_value_diff
        igst_diff = 0 if pd.isna(igst_diff) else igst_diff
        cgst_diff = 0 if pd.isna(cgst_diff) else cgst_diff
        sgst_diff = 0 if pd.isna(sgst_diff) else sgst_diff
        
        # Round the sum of differences to the nearest integer
        amount_difference = round(taxable_value_diff + igst_diff + cgst_diff + sgst_diff)
        return amount_difference

    # Add a new column "Amount difference" using the custom function
    cdnur_match["Amount difference"] = cdnur_match.apply(calculate_amount_difference, axis=1)

    df7 = df7.iloc[2:]
    # Step 3: Set the 6th row as the header.
    header_row = df7.iloc[0]
    df7 = df7[1:]
    df7.columns = header_row

    cdnr_match = pd.merge(
        df3,
        df7,
        how='outer',
        left_on=["GSTIN of recipient", "Note number"],
        right_on=["GSTIN/UIN of Recipient", "Note Number"],
        suffixes=('_cdnr', '_cdnr_govt')
    )

    selected_columns_reordered3 = [0,2,7,11,12,13,14,19,21,23,26,30,31,32,33]
    cdnr_match = cdnr_match.iloc[:, selected_columns_reordered3]

    def generate_taxable_value_remark(row):
        if pd.isnull(row["Taxable Value (Rs.)"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Taxable Value"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Taxable Value (Rs.)"] == row["Taxable Value"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark Taxable Value" using the custom function
    cdnr_match["Remark Taxable Value"] = cdnr_match.apply(generate_taxable_value_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark GSTIN of recipient"
    def generate_gstin_remark(row):
        if pd.isnull(row["GSTIN of recipient"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["GSTIN/UIN of Recipient"]):
            return "Unmatched / Its not present in Govt template"
        elif row["GSTIN of recipient"] == row["Govt cdnr.GSTIN/UIN of Recipient"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark GSTIN of recipient" using the custom function
    cdnr_match["Remark GSTIN of recipient"] = cdnr_match.apply(generate_gstin_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark IGST"
    def generate_igst_remark(row):
        if pd.isnull(row.iloc[4]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Integrated Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[4] == row["Integrated Tax"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark IGST" using the custom function
    cdnr_match["Remark IGST"] = cdnr_match.apply(generate_igst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark IGST"
    def generate_cgst_remark(row):
        if pd.isnull(row.iloc[5]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Central Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[5] == row["Central Tax"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark IGST" using the custom function
    cdnr_match["Remark CGST"] = cdnr_match.apply(generate_cgst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark IGST"
    def generate_sgst_remark(row):
        if pd.isnull(row.iloc[6]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["State/UT Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[6] == row["State/UT Tax"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark IGST" using the custom function
    cdnr_match["Remark SGST"] = cdnr_match.apply(generate_sgst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark Type of Supply"
    def generate_type_of_supply_remark(row):
        if pd.isnull(row["Type of supply"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Note Supply Type"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Type of supply"] == row["Note Supply Type"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark Type of Supply" using the custom function
    cdnr_match["Remark Type of Supply"] = cdnr_match.apply(generate_type_of_supply_remark, axis=1)

    cdnr_match["Taxable Value (Rs.)"] = pd.to_numeric(cdnr_match["Taxable Value (Rs.)"], errors='coerce')
    cdnr_match["Taxable Value"] = pd.to_numeric(cdnr_match["Taxable Value"], errors='coerce')
    cdnr_match.iloc[:, 4] = pd.to_numeric(cdnr_match.iloc[:, 4], errors='coerce')
    cdnr_match["Integrated Tax"] = pd.to_numeric(cdnr_match["Integrated Tax"], errors='coerce')
    cdnr_match.iloc[:, 5] = pd.to_numeric(cdnr_match.iloc[:, 5], errors='coerce')
    cdnr_match["Central Tax"] = pd.to_numeric(cdnr_match["Central Tax"], errors='coerce')
    cdnr_match.iloc[:, 6] = pd.to_numeric(cdnr_match.iloc[:, 6], errors='coerce')
    cdnr_match["State/UT Tax"] = pd.to_numeric(cdnr_match["State/UT Tax"], errors='coerce')

    # Define a function to calculate the "Amount difference"
    def calculate_amount_difference(row):
        taxable_value_diff = row["Taxable Value (Rs.)"] - row["Taxable Value"]
        igst_diff = row.iloc[4] - row["Integrated Tax"]
        cgst_diff = row.iloc[5] - row["Central Tax"]
        sgst_diff = row.iloc[6] - row["State/UT Tax"]
        
        # Check for NaN values before rounding
        taxable_value_diff = 0 if pd.isna(taxable_value_diff) else taxable_value_diff
        igst_diff = 0 if pd.isna(igst_diff) else igst_diff
        cgst_diff = 0 if pd.isna(cgst_diff) else cgst_diff
        sgst_diff = 0 if pd.isna(sgst_diff) else sgst_diff
        
        # Round the sum of differences to the nearest integer
        amount_difference = round(taxable_value_diff + igst_diff + cgst_diff + sgst_diff)
        return amount_difference

    # Add a new column "Amount difference" using the custom function
    cdnr_match["Amount difference"] = cdnr_match.apply(calculate_amount_difference, axis=1)

    df8 = df8.iloc[2:]
    # Step 3: Set the 6th row as the header.
    header_row = df8.iloc[0]
    df8 = df8[1:]
    df8.columns = header_row

    # Convert "Invoice Number" to int64 data type
    df8["Invoice Number"] = pd.to_numeric(df8["Invoice Number"], errors='coerce').astype(pd.Int64Dtype())

    # Perform the merge
    exp_match = pd.merge(
        df4,
        df8,
        how='outer',
        left_on=["Document number"],
        right_on=["Invoice Number"],
        suffixes=('_exp', '_exp_govt')
    )

    selected_columns_reordered4 = [0,1,8,10,11,12,14,15,22,23]
    exp_match = exp_match.iloc[:, selected_columns_reordered4]

    def generate_taxable_value_remark(row):
        if pd.isnull(row["Taxable Value (Rs.)"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Taxable Value"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Taxable Value (Rs.)"] == row["Taxable Value"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark Taxable Value" using the custom function
    exp_match["Remark Taxable Value"] = exp_match.apply(generate_taxable_value_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark IGST"
    def generate_igst_remark(row):
        if pd.isnull(row.iloc[4]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Integrated Tax"]):
            return "Unmatched / Its not present in Govt template"
        elif row.iloc[4] == row["Integrated Tax"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark IGST" using the custom function
    exp_match["Remark IGST"] = exp_match.apply(generate_igst_remark, axis=1)

    # Define a function to apply the conditions and generate remarks for "Remark Type of Supply"
    def generate_type_of_supply_remark(row):
        if pd.isnull(row["Type of supply"]):
            return "Unmatched / Its not present in Optitax's data"
        elif pd.isnull(row["Export Type"]):
            return "Unmatched / Its not present in Govt template"
        elif row["Type of supply"] == row["Export Type"]:
            return "Matched"
        else:
            return "Unmatched"

    # Add a new column "Remark Type of Supply" using the custom function
    exp_match["Remark Type of Supply"] = exp_match.apply(generate_type_of_supply_remark, axis=1)

    exp_match["Taxable Value (Rs.)"] = pd.to_numeric(exp_match["Taxable Value (Rs.)"], errors='coerce')
    exp_match["Taxable Value"] = pd.to_numeric(exp_match["Taxable Value"], errors='coerce')
    exp_match.iloc[:, 4] = pd.to_numeric(exp_match.iloc[:, 4], errors='coerce')
    exp_match["Integrated Tax"] = pd.to_numeric(exp_match["Integrated Tax"], errors='coerce')
    exp_match.iloc[:, 5] = pd.to_numeric(exp_match.iloc[:, 5], errors='coerce')
    # cdnur_match["Central Tax"] = pd.to_numeric(cdnur_match["Central Tax"], errors='coerce')
    exp_match.iloc[:, 6] = pd.to_numeric(exp_match.iloc[:, 6], errors='coerce')
    # cdnur_match["State/UT Tax"] = pd.to_numeric(cdnur_match["State/UT Tax"], errors='coerce')

    # Define a function to calculate the "Amount difference"
    def calculate_amount_difference(row):
        taxable_value_diff = row["Taxable Value (Rs.)"] - row["Taxable Value"]
        igst_diff = row.iloc[4] - row["Integrated Tax"]
        cgst_diff = row.iloc[5]
        sgst_diff = row.iloc[6]
        
        # Check for NaN values before rounding
        taxable_value_diff = 0 if pd.isna(taxable_value_diff) else taxable_value_diff
        igst_diff = 0 if pd.isna(igst_diff) else igst_diff
        cgst_diff = 0 if pd.isna(cgst_diff) else cgst_diff
        sgst_diff = 0 if pd.isna(sgst_diff) else sgst_diff
        
        # Round the sum of differences to the nearest integer
        amount_difference = round(taxable_value_diff + igst_diff + cgst_diff + sgst_diff)
        return amount_difference

    # Add a new column "Amount difference" using the custom function
    exp_match["Amount difference"] = exp_match.apply(calculate_amount_difference, axis=1)

    # Write the differences DataFrame to a new Excel file
    with pd.ExcelWriter(company_file_path, engine='openpyxl') as writer:
        b2b_match.to_excel(writer, sheet_name='b2b_match', index=False)
        cdnur_match.to_excel(writer, sheet_name='cdnur_match', index=False)
        cdnr_match.to_excel(writer, sheet_name='cdnr_match', index=False)
        exp_match.to_excel(writer, sheet_name='exp_match', index=False)


    return company_file_path

def summary_excel_files(s2_file_path, s3_file_path):
    s2 = pd.ExcelFile(s2_file_path)
    s3 = pd.ExcelFile(s3_file_path)
    
    df1 = pd.read_excel(s2, sheet_name="b2b,sez,de")
    df2 = pd.read_excel(s2, sheet_name="b2cl")
    df3 = pd.read_excel(s2, sheet_name="b2cs")
    df4 = pd.read_excel(s2, sheet_name="cdnr")
    df5 = pd.read_excel(s2, sheet_name="at")
    df6 = pd.read_excel(s2, sheet_name="atadj")
    df7 = pd.read_excel(s2, sheet_name="exp")
    df8 = pd.read_excel(s2, sheet_name="cdnur")
    df9 = pd.read_excel(s2, sheet_name="cdnur_b2cs")
    df10 = pd.read_excel(s2, sheet_name="exemp")
    df11 = pd.read_excel(s2, sheet_name="b2ba")
    df12 = pd.read_excel(s2, sheet_name="b2cla")
    df13 = pd.read_excel(s2, sheet_name="b2cla")
    df14 = pd.read_excel(s2, sheet_name="ata")
    df15 = pd.read_excel(s2, sheet_name="atadja")
    df16 = pd.read_excel(s2, sheet_name="expa")
    df17 = pd.read_excel(s2, sheet_name="cdnur")
    df18 = pd.read_excel(s2, sheet_name="docs")
    df19 = pd.read_excel(s2, sheet_name="HSN")

    df1 = df1.loc[(df1['Type of supply'] != "SEZ supplies with payment") & (df1['Type of supply'] != "SEZ supplies without payment")]


    df1['Applicable % of Tax Rate'] = df1['Applicable % of Tax Rate'].fillna('Null')

    # Group by 'Applicable % of Tax Rate' and sum the specified columns
    b2b = df1.groupby(['Applicable % of Tax Rate']).agg({
        df1.columns[11]: 'sum',
        df1.columns[13]: 'sum',
        df1.columns[14]: 'sum',
        df1.columns[15]: 'sum'
    }).reset_index()

    # Rename specific columns
    b2b = b2b.rename(columns={
        b2b.columns[1]: 'Taxable Value',
        b2b.columns[2]: 'Integrated Tax',
        b2b.columns[3]: 'Central Tax',
        b2b.columns[4]: 'State/UT Tax'
    })

    def determine_type(row):
        if row["Applicable % of Tax Rate"] == 'Null':
            return "B2B"
        else:
            return None

    # Add a new column 'Type' using the custom function
    b2b['Type'] = b2b.apply(determine_type, axis=1)

    b2b['Total duty'] = b2b['Integrated Tax'] + b2b['Central Tax'] + b2b['State/UT Tax']
    b2b['Total value'] = b2b['Taxable Value'] + b2b['Integrated Tax'] + b2b['Central Tax'] + b2b['State/UT Tax']
    
    def determine_particular(row):
        if "B2B" in row["Type"]:
            return "Taxable Outward Supply"
        else:
            return None

    # Add a new column 'PARTICULAR' using the custom function
    b2b['PARTICULAR'] = b2b.apply(determine_particular, axis=1)

    column_order_b2b = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    b2b = b2b[column_order_b2b]

    df2['Applicable % of Tax Rate'] = df2['Applicable % of Tax Rate'].fillna('Null')

    # Group by 'Applicable % of Tax Rate' and sum the specified columns
    b2cl = df2.groupby(['Applicable % of Tax Rate']).agg({
        df2.columns[6]: 'sum',
        df2.columns[9]: 'sum',
        df2.columns[10]: 'sum',
        df2.columns[11]: 'sum'
    }).reset_index()

    # Rename specific columns
    b2cl = b2cl.rename(columns={
        b2cl.columns[1]: 'Taxable Value',
        b2cl.columns[2]: 'Integrated Tax',
        b2cl.columns[3]: 'Central Tax',
        b2cl.columns[4]: 'State/UT Tax'
    })

    def determine_type(row):
        if row["Applicable % of Tax Rate"] == 'Null':
            return "B2CL"
        else:
            return None

    # Add a new column 'Type' using the custom function
    b2cl['Type'] = b2cl.apply(determine_type, axis=1)

    b2cl['Total duty'] = b2cl['Integrated Tax'] + b2cl['Central Tax'] + b2cl['State/UT Tax']
    b2cl['Total value'] = b2cl['Taxable Value'] + b2cl['Integrated Tax'] + b2cl['Central Tax'] + b2cl['State/UT Tax']
    
    def determine_particular(row):
        if "B2CL" in row["Type"]:
            return "Taxable Outward Supply"
        else:
            return None

    # Add a new column 'PARTICULAR' using the custom function
    b2cl['PARTICULAR'] = b2cl.apply(determine_particular, axis=1)

    column_order_b2cl = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    b2cl = b2cl[column_order_b2cl]

    df3['Applicable % of Tax Rate'] = df3['Applicable % of Tax Rate'].fillna('Null')

    # Group by 'Applicable % of Tax Rate' and sum the specified columns
    b2cs = df3.groupby(['Applicable % of Tax Rate']).agg({
        df3.columns[4]: 'sum',
        df3.columns[7]: 'sum',
        df3.columns[8]: 'sum',
        df3.columns[9]: 'sum'
    }).reset_index()

    # Rename specific columns
    b2cs = b2cs.rename(columns={
        b2cs.columns[1]: 'Taxable Value',
        b2cs.columns[2]: 'Integrated Tax',
        b2cs.columns[3]: 'Central Tax',
        b2cs.columns[4]: 'State/UT Tax'
    })

    def determine_type(row):
        if row["Applicable % of Tax Rate"] == 'Null':
            return "B2CS"
        else:
            return None

    # Add a new column 'Type' using the custom function
    b2cs['Type'] = b2cs.apply(determine_type, axis=1)

    b2cs['Total duty'] = b2cs['Integrated Tax'] + b2cs['Central Tax'] + b2cs['State/UT Tax']
    b2cs['Total value'] = b2cs['Taxable Value'] + b2cs['Integrated Tax'] + b2cs['Central Tax'] + b2cs['State/UT Tax']
    
    def determine_particular(row):
        if "B2CS" in row["Type"]:
            return "Taxable Outward Supply"
        else:
            return None

    # Add a new column 'PARTICULAR' using the custom function
    b2cs['PARTICULAR'] = b2cs.apply(determine_particular, axis=1)

    column_order_b2cs = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    b2cs = b2cs[column_order_b2cs]

    df4['Note type'] = df4['Note type'].fillna('Null')

    cdnr = df4.groupby(['Note type']).agg({
        df4.columns[11]: 'sum',
        df4.columns[12]: 'sum',
        df4.columns[13]: 'sum',
        df4.columns[14]: 'sum'
    }).reset_index()

    # Rename specific columns
    cdnr = cdnr.rename(columns={
        cdnr.columns[1]: 'Taxable Value',
        cdnr.columns[2]: 'Integrated Tax',
        cdnr.columns[3]: 'Central Tax',
        cdnr.columns[4]: 'State/UT Tax'
    })

    cdnr['Total duty'] = cdnr['Integrated Tax'] + cdnr['Central Tax'] + cdnr['State/UT Tax']
    cdnr['Total value'] = cdnr['Taxable Value'] + cdnr['Integrated Tax'] + cdnr['Central Tax'] + cdnr['State/UT Tax']
    
    def determine_type(row):
        if row["Note type"] == 'Credit note':
            return "Credit note"
        elif row["Note type"] == 'Debit note':
            return "Debit note 1"
        else:
            return None

    # Add a new column 'Type' using the custom function
    cdnr['Type'] = cdnr.apply(determine_type, axis=1)

    def determine_particular(row):
        if "Credit note" in row["Note type"]:
            return "Taxable Outward Supply-CDNR"
        elif "Debit note" in row["Note type"]:
            return "Taxable Outward Supply-CDNR"
        else:
            return "Null"

    # Add a new column 'PARTICULAR' using the custom function
    cdnr['PARTICULAR'] = cdnr.apply(determine_particular, axis=1)
    
    column_order_cdnr = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    cdnr = cdnr[column_order_cdnr]

    df5['Applicable Rate'] = df5['Applicable Rate'].fillna('Null')

    at = df5.groupby(['Applicable Rate']).agg({
        df5.columns[3]: 'sum',
        df5.columns[5]: 'sum',
        df5.columns[6]: 'sum',
        df5.columns[7]: 'sum'
    }).reset_index()

    # Rename specific columns
    at = at.rename(columns={
        at.columns[1]: 'Taxable Value',
        at.columns[2]: 'Integrated Tax',
        at.columns[3]: 'Central Tax',
        at.columns[4]: 'State/UT Tax'
    })

    def determine_type(row):
        if row["Applicable Rate"] == "Null":
            return "RECEIVED"
        else:
            return None

    # Add a new column 'Type' using the custom function
    at['Type'] = at.apply(determine_type, axis=1)

    # Define a function to determine the 'PARTICULAR' column
    def determine_particular(row):
        if "RECEIVED" in row["Type"]:
            return "Advance Received/ Adjusted"
        else:
            return None

    # Add a new column 'PARTICULAR' using the custom function
    at['PARTICULAR'] = at.apply(determine_particular, axis=1)

    at['Total duty'] = at['Integrated Tax'] + at['Central Tax'] + at['State/UT Tax']
    at['Total value'] = at['Taxable Value'] + at['Integrated Tax'] + at['Central Tax'] + at['State/UT Tax']

    column_order_at = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    at = at[column_order_at]

    df6['Applicable Rate'] = df6['Applicable Rate'].fillna('Null')
    df6['Type of document'] = df6['Type of document'].fillna('Null')

    # Group by 'Applicability rate' and 'Type of document' and sum the specified columns
    atadj = df6.groupby(['Applicable Rate', 'Type of document']).agg({
        df6.columns[3]: 'sum',
        df6.columns[5]: 'sum',
        df6.columns[6]: 'sum',
        df6.columns[7]: 'sum'
    }).reset_index()

    # Rename columns to match the M code
    atadj = atadj.rename(columns={
        atadj.columns[2]: 'Taxable Value',
        atadj.columns[3]: 'Integrated Tax',
        atadj.columns[4]: 'Central Tax',
        atadj.columns[5]: 'State/UT Tax'
    })

    atadj['Total duty'] = atadj['Integrated Tax'] + atadj['Central Tax'] + atadj['State/UT Tax']
    atadj['Total value'] = atadj['Taxable Value'] + atadj['Integrated Tax'] + atadj['Central Tax'] + atadj['State/UT Tax']

    conditions = [
        (atadj['Type of document'] == 'Refund voucher'),
        (atadj['Type of document'].isin(['Invoice', 'Invoice-cum-bill of supply', 'Bill of supply']))
    ]

    values = ['REFUND', 'ADJUSTED']

    # Create a new column 'Type' using numpy.select
    atadj['Type'] = np.select(conditions, values, default=None)

    atadj = atadj.groupby('Type').agg({
        'Taxable Value': 'sum',
        'Integrated Tax': 'sum',
        'Central Tax': 'sum',
        'State/UT Tax': 'sum',
        'Total duty': 'sum',
        'Total value': 'sum'
    }).reset_index()

    def determine_particular(row):
        if "ADJUSTED" in row["Type"]:
            return "Advance Received/ Adjusted"
        elif "REFUND" in row["Type"]:
            return "Advance Received/ Adjusted"
        else:
            return None

    # Add a new column 'PARTICULAR' using the custom function
    atadj['PARTICULAR'] = atadj.apply(determine_particular, axis=1)

    column_order_atadj = ['Type','Taxable Value','Integrated Tax','Central Tax', 'State/UT Tax', 'Total duty', 'Total value', 'PARTICULAR']

    # Reorder the columns in the DataFrame
    atadj = atadj[column_order_atadj]

    with pd.ExcelWriter(s2_file_path, engine='openpyxl') as writer:
        b2b.to_excel(writer, sheet_name='GSTR-1 summary', index=False)
        b2cl.to_excel(writer, sheet_name='b2cl', index=False)
        b2cs.to_excel(writer, sheet_name='b2cs', index=False)
        cdnr.to_excel(writer, sheet_name='cdnr', index=False)
        at.to_excel(writer, sheet_name='at', index=False)
        atadj.to_excel(writer, sheet_name='atadj', index=False)
        # b2cl.to_excel(writer, sheet_name='GSTR-1 E-Invoice summary', index=False)


    return s2_file_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    step = request.form['step']
    if file:
        original_filename, file_extension = os.path.splitext(file.filename)
        processed_filename = f"{original_filename}_{step}.xlsx"
        # Save the uploaded file to the 'uploads' folder
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)
        if step == 'step1':
            process_step1(filename)
        elif step == 'step2':
            process_step2(filename)
            
        # Send the processed file for download
        return send_file(filename, as_attachment=True, download_name=processed_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return 'File upload failed.'

@app.route('/upload1', methods=['POST'])
def upload1():
    company_file = request.files['companyFile']
    government_file = request.files['governmentFile']

    if company_file and government_file:
        original_filename, file_extension = os.path.splitext(company_file.filename)
        processed_filename = f"{original_filename}_step3.xlsx"
        # Save the uploaded files to the 'uploads' folder
        company_filename = secure_filename(company_file.filename)
        government_filename = secure_filename(government_file.filename)

        company_file_path = os.path.join(app.config['UPLOAD_FOLDER'], company_filename)
        government_file_path = os.path.join(app.config['UPLOAD_FOLDER'], government_filename)

        company_file.save(company_file_path)
        government_file.save(government_file_path)


        differences_file_path = compare_excel_files(company_file_path, government_file_path)

        # Send the differences file for download
        return send_file(differences_file_path, as_attachment=True, download_name=processed_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return 'File upload failed.'

@app.route('/upload2', methods=['POST'])
def upload2():
    s2_file = request.files['s2File']
    s3_file = request.files['s3File']

    if s2_file and s3_file:
        original_filename, file_extension = os.path.splitext(s2_file.filename)
        processed_filename = f"{original_filename}_Summary.xlsx"
        # Save the uploaded files to the 'uploads' folder
        s2_filename = secure_filename(s2_file.filename)
        s3_filename = secure_filename(s3_file.filename)

        s2_file_path = os.path.join(app.config['UPLOAD_FOLDER'], s2_filename)
        s3_file_path = os.path.join(app.config['UPLOAD_FOLDER'], s3_filename)

        s2_file.save(s2_file_path)
        s3_file.save(s3_file_path)


        differences_file_path = summary_excel_files(s2_file_path, s3_file_path)

        # Send the differences file for download
        return send_file(differences_file_path, as_attachment=True, download_name=processed_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return 'File upload failed.'

if __name__ == '__main__':
    app.run(host = '0.0.0.0', port = 3750)
