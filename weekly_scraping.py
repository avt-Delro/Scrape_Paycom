from playwright.sync_api import sync_playwright
import environment as env
from dotenv import load_dotenv
import win32com.client as win32
import os
from datetime import datetime
import calendar
import pandas as pd
from openpyxl import load_workbook




load_dotenv()
paycom_user = env.paycom_username
paycom_pass = env.paycom_password  
client_code = env.paycom_clientcode 

local_path = env.paycom_local
email = env.sendemail

datetoday = datetime.today()
first_day_of_month = datetoday.replace(day=1)




outlook = win32.Dispatch("Outlook.Application")
outlook_ap = outlook.GetNamespace("MAPI")
config = pd.read_csv('config/config.csv')

err_email = 'vjdelrosario@avatco.com'
err_cc = 'vjdelrosario@avatco.com'


def create_sheet(filepath, data_row, sheetname):
    if isinstance(data_row, list):
        df = pd.DataFrame(data_row)
    elif isinstance(data_row, dict):
        df = pd.DataFrame([data_row])
    
    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer,  sheet_name=sheetname, index=False)

def create_report_summary_we_month(file, sheet):
    df = pd.read_excel(file, sheet_name=sheet)
 
    cols = ["Scheduled Hours", "Actual Hours", "Variance"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce").fillna(0)
 
    df['Punch Date'] = pd.to_datetime(df['Punch Date'], errors='coerce')
 
    df_filtered = df[df['Punch Date']<= datetoday.strftime('%m/%d/%Y')]
 
    #Since Duplicate column headers, Pandas renamed the second column .1
    summary = (
        df_filtered.groupby("Employee", as_index=False)
        .agg({
            "Scheduled Hours": "sum",
            "Actual Hours": "sum",
            "Variance": "sum",
        })
    )
    create_sheet(file, summary.to_dict(orient="records"), 'Summary')
 
def create_report_we_month(file):
    df = pd.read_excel(file)
 
    cols = ["Scheduled Hours", "Actual Hours", "Variance"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce").fillna(0)
 
    df['Punch Date'] = pd.to_datetime(df['Punch Date'], errors='coerce')
 
    df_filtered = df[df['Punch Date']<= datetoday.strftime('%m/%d/%Y')]
 
 
    summary = (
        df_filtered.groupby("Employee", as_index=False)
        .agg({
            "Scheduled Hours": "sum",
            "Actual Hours": "sum",
            "Variance": "sum",
        })
    )
    create_sheet(file, summary.to_dict(orient="records"), "Summary")
 
 
def send_email_we_month (toemail,ccemail, filepath):
    outlook = win32.Dispatch("Outlook.Application")
    outlook_ap = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(0)
 
    df = pd.read_excel(filepath, sheet_name="Summary")
 
    df_without_scheduled_hours = df.loc[df['Scheduled Hours'] == 0, ['Employee', 'Scheduled Hours', 'Actual Hours']]
    df_negative_variance = df.loc[df['Variance'] < 0, ['Employee', 'Variance']]
    df_positive_variance = df.loc[df['Variance'] > 0, ['Employee', 'Variance']]
 
    html_without = df_without_scheduled_hours.to_html(index=False)
    html_without = html_without.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
    html_negative = df_negative_variance.to_html(index=False)
    html_negative = html_negative.replace("<thead", "<thead style='background-color:#1CFF77; color:white;'")
    html_positive = df_positive_variance.to_html(index=False)
    html_positive = html_positive.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
    html_df = df.to_html(index=False)
    html_df = html_df.replace("<thead", "<thead style='background-color:#1CFF77; color:white;'")
   
    mail.To = toemail
    mail.CC = ccemail
 
    mail.Attachments.Add(filepath)
 
    mail.Subject = f'OT Report, from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}'
    mail.HTMLBody = f"""
        <html>
        <head>
        <style>
        table {{
            border-collapse: collapse;
            width: 75%;
        }}
        th, td {{
            padding: 8px;
            text-align: left;
        }}
        </style>
        </head>
        <body>
            <p>Good day, Here are the <b>Summary of work hours</b> from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}</p>
 
            <p>Attached is the report file: {os.path.basename(filepath)}</p>
 
            <h1>Employees with Positive Variance</h1>
            {html_positive}
            <br>
 
            <h1>Employees with Negative Variance</h1>
            {html_negative}
            <br>
 
            <h2>Summary of the Report:</h2>
            {html_df}
            <br>
 
            <h1>Employees without Scheduled Hours</h1>
            {html_without}
            <br>
 
            <p>Thank you,<br>
            Automated Reporting System</p>
        </body>
        </html>
        """
   
    mail.Send()
    print('Email Sent')
 
def create_clp_summary_we_month(file):
    df = pd.read_excel(file)
    df.rename(columns = {'InPunchTime':'Date', 'EarnHours': 'Meal Penalty Hours'}, inplace = True)
    
 
    #Since Duplicate column headers, Pandas renamed the second column .1
    summary = (
        df.groupby(['EECode', 'Lastname', 'Firstname'], as_index=False)['Meal Penalty Hours'].sum(min_count = 1)
    )
    new_df_clp = df.loc[df['Meal Penalty Hours'] > 0, ['Date',  'EECode', 'Lastname', 'Firstname', 'Meal Penalty Hours',"EarnCode", "HomeDepartment", "HomeAllocation", "Pay Class","Home Job Desc","Badge","Employee Approved","Supervisor Approved"]]

    create_sheet(file, new_df_clp.to_dict(orient="records"), 'Sheet1')
    create_sheet(file, summary.to_dict(orient="records"), 'Summary')
 
def send_email_missing_we_month(toemail,ccemail,filepath):
    outlook = win32.Dispatch("Outlook.Application")
   
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet_missing')
   
    mail = outlook.CreateItem(0)
 
    mail.To = toemail
    mail.CC = ccemail
 
    mail.Attachments.Add(filepath)
 
    df = pd.read_excel(filepath)
 
    df.sort_values(by='Date', ascending=True, inplace=True)
    df_missing_summary = df[['EE Code', 'Date' ,'Last Name', 'First Name', 'In Punch Time', 'Out Punch Time']]
    df_missing_summary_to_html = df_missing_summary.to_html(index=False)
 
    mail.Subject = f'Missing Punches Report from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}'
    mail.HTMLBody = f"""
        <html>
        <head>
        <style>
        table {{
            border-collapse: collapse;
            width: 75%;
        }}
        th, td {{
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #FF2B3F;
        }}
        </style>
        </head>
        <body>
            <p>Good day, Here are the summary of <b>Missing Punches</b> from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}</p>
            <p>Attached is from the report file: {os.path.basename(filepath)}</p>
 
            <h2>Summary of the Report:</h2>
            {df_missing_summary_to_html}
            <br>
            <p>Thank you,<br>
            Automated Reporting System</p>
        </body>
        </html>
        """
   
    mail.Send()
    print('Email Sent')


def send_email_clp_we_month(to_email, cc_email, filepath):
    outlook = win32.Dispatch("Outlook.Application")
    
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet_clp')

    mail = outlook.CreateItem(0)
    df = pd.read_excel(filepath, sheet_name='Summary')

    

    df.sort_values(by = 'Meal Penalty Hours', ascending=False, inplace= True)
    df_clp = df.loc[df['Meal Penalty Hours'] > 0, ['EECode', 'Lastname', 'Firstname', 'Meal Penalty Hours']]
    if len(df_clp)>0:
        df_clp_summary_to_html = df_clp.to_html(index=False)
        df_clp_summary_to_html = df_clp_summary_to_html.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
    else:
        df_clp_summary_to_html = "<p><b style = text-transform:uppercase;>No meal penalties for this period.</b></p>"

    mail.Attachments.Add(filepath)

    mail.To = to_email
    mail.CC = cc_email
    mail.Subject = f'CA Meal Penalty Report from {first_day_of_month.strftime("%m/%d/%Y")} -  {datetoday.strftime("%m/%d/%Y")}'
    mail.HTMLBody = f"""
        <html>
        <head>
        <style>
        table {{
            border-collapse: collapse;
            width: 75%;
        }}
        th, td {{
            padding: 8px;
            text-align: left;
        }}
        </style>
        </head>
        <body>
            <p>GOOD DAY, HERE IS THE SUMMARY OF <b>CA MEAL PENALTY</b> FROM {first_day_of_month.strftime("%m/%d/%Y")} -  {datetoday.strftime("%m/%d/%Y")}</p>
            <p>Attached is from the report file: {os.path.basename(filepath)}</p>

            <h2>Summary of the Report:</h2>
            {df_clp_summary_to_html}
            <br>
            <p>Thank you,<br>
            Automated Reporting System</p>
        </body>
        </html>
        """
    
    mail.Send()
    print('Email Sent')
       
