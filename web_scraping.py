
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



def paycom_scraping(weblink, username, password, client_code, int_choice = 1):
    try:
        with sync_playwright() as p:
            date = datetoday.strftime("%m/%d/%Y")
            
            
            context = p.chromium.launch_persistent_context(
                    user_data_dir="edge_automation_profile",
                    channel="msedge",
                    headless=False
                )


            page = context.new_page()
            
            page.goto(weblink, wait_until='load') 
            page.locator("#clientcode").fill(client_code)
            page.locator("#username").fill(username)
            page.locator("#password").fill(password)
            page.get_by_role("button", name="Log In").click()
            page.wait_for_load_state("networkidle")
            page.goto("https://www.paycomonline.net/v4/cl/rpt-center.php",wait_until='networkidle')
            page.get_by_role("tab", name="Push Reporting™").click()
            page.get_by_role("tab", name="Saved Reports").click()
            if int_choice == 1:
                page.get_by_role("row", name="Favorite Actual v Scheduled w").locator("input[type=\"button\"]").click()
            elif int_choice == 2:
                page.get_by_role("row", name="Favorite Missing Punches w").locator("input[type=\"button\"]").click()
            elif int_choice == 3:
                page.get_by_role("row", name="Favorite CLP w Groups Time").locator("input[type=\"button\"]").click()
            page.get_by_role("button", name="Download").wait_for(timeout= 600000)
            
            
            with page.expect_download() as download_info:
                page.get_by_role("button", name="Download").click()
            download = download_info.value


            file_folder = os.path.join(local_path, 'HR Files')
            os.makedirs(file_folder, exist_ok=True)

            if int_choice == 1:
                file_name =  f'OT_report_{datetoday.strftime("%Y%m%d")}.xlsx'
            elif int_choice == 2:
                file_name =  f'MissingPunches_report_{datetoday.strftime("%Y%m%d")}.xlsx'
            elif int_choice == 3:
                file_name =  f'CLP_report_{datetoday.strftime("%Y%m%d")}.xlsx'

            download.save_as(os.path.join(file_folder,file_name))

            return f"{str(file_folder)}\{file_name}"

    except Exception as e:
        print(e)

def create_sheet(filepath, data_row, sheetname):
    if isinstance(data_row, list):
        df = pd.DataFrame(data_row)
    elif isinstance(data_row, dict):
        df = pd.DataFrame([data_row])
    
    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer,  sheet_name=sheetname, index=False)


def create_report_summary(file, sheet):
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

def create_report(file):
    df = pd.read_excel(file)

    cols = ["Scheduled Hours", "Actual Hours", "Variance"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    df['Punch Date'] = pd.to_datetime(df['Punch Date'], errors='coerce')

    df_filtered = df[df['Punch Date']<= datetoday.strftime('%m/%d/%Y')]
    sheet_folder = os.path.join(local_path, 'sheet')

    #Since Duplicate column headers, Pandas renamed the second column .1

    for group, df_group in df_filtered.groupby("Schedule Group.1"):
        exc_filepath = f'{group}.xlsx'

        df_group.to_excel(os.path.join(sheet_folder, exc_filepath))
        print(f'Created file for {group}')
    
    for file in os.listdir(sheet_folder):
        create_report_summary(os.path.join(sheet_folder, file), 'Sheet1')


def send_email (filepath):
    outlook = win32.Dispatch("Outlook.Application")
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet')

    for file in os.listdir(sheet_folder):
        mail = outlook.CreateItem(0)
        df = pd.read_excel(os.path.join(sheet_folder, file), sheet_name='Summary')

        sheet_number = str(os.path.basename(file)).replace('.xlsx', '')

        try:
            subject_header = config.loc[config['Group_Code'] == int(sheet_number), 'NAME_sched'].values[0]
            emailto = config.loc[config['Group_Code']== int(sheet_number), 'SEND_to'].values[0]
            ccto = config.loc[config['Group_Code']== int(sheet_number), 'CC_S'].values[0]
        except Exception as e:
            subject_header = 'Schedule Group Not in Config'
            emailto = err_email
            ccto = err_cc
            print(f'Exception: {e}')

        df_without_scheduled_hours = df.loc[df['Scheduled Hours'] == 0, ['Employee', 'Scheduled Hours', 'Actual Hours']]
        df_negative_variance = df.loc[df['Variance'] < 0, ['Employee', 'Variance']]
        df_positive_variance = df.loc[df['Variance'] > 0, ['Employee', 'Variance']]


        if len(df_without_scheduled_hours)>0:
            df_without_scheduled_hours.sort_values(by ='Actual Hours', ascending=False, inplace= True)
            html_without = df_without_scheduled_hours.to_html(index=False)
            html_without = html_without.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
        else:
            html_without = "<p><b style = text-transform:uppercase;>No employees without scheduled hours for this period.</b></p>"

        #///
        if len(df_negative_variance)>0:
            df_negative_variance.sort_values(by ='Variance', ascending=True, inplace= True)
            html_negative = df_negative_variance.to_html(index=False)
            #Change thead color to green for negative
            html_negative = html_negative.replace("<thead", "<thead style='background-color:#1CFF77; color:white;'")
        else:
            html_negative = "<p><b style = text-transform:uppercase;>No employees with negative variance for this period.</b></p>"

        #///
        if len(df_positive_variance)>0:
            df_positive_variance.sort_values(by ='Variance', ascending=False, inplace= True)
            html_positive = df_positive_variance.to_html(index=False)
            #Change thead color to red for positive
            html_positive = html_positive.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
        else:
            html_positive = "<p><b style = text-transform:uppercase;>No employees with positive variance for this period.</b></p>"

        #///
        df.sort_values(by = 'Variance', ascending=True, inplace= True)
        html_df = df.to_html(index=False)
        html_df = html_df.replace("<thead", "<thead style='background-color:#1CFF77; color:white;'")

        mail.Attachments.Add(os.path.join(sheet_folder, file))

        mail.To = emailto
        mail.CC = ccto
        mail.Subject = f'OT Report, from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")} {subject_header}'
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
                <p>GOOD DAY, HERE IS THE <b>SUMMARY OF WORK HOURS</b> FROM {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}</p>
                <p>Attached is from the report file: {os.path.basename(filepath)}</p>

                <h1>Employees with Positive Variance</h1>
                {html_positive}
                <br>

                <h1>Employees with Negative Variance</h1>
                {html_negative}
                <br>

                <h1>Summary of the Report:</h1>
                <p>Employees with <b>Positive variance:</b> {len(df_positive_variance)}</p>
                <p>Employees with <b>Negative variance:</b> {len(df_negative_variance)}</p>
                <p>Employees without <b>Scheduled Hours:</b> {len(df_without_scheduled_hours)}</p>
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
        
    os.remove(filepath)

def create_missing_report(file):
    df = pd.read_excel(file)

    sheet_folder = os.path.join(local_path, 'sheet_missing')
    os.makedirs(sheet_folder, exist_ok=True)

    for group, df_group in df.groupby('Schedule Group.1'):
        exc_filepath = f'{group}.xlsx'
        df_group.to_excel(os.path.join(sheet_folder, exc_filepath))
        print(f'Created file for {group}')

def create_clp_summary(file):
    df = pd.read_excel(file)

    #Since Duplicate column headers, Pandas renamed the second column .1
    summary = (
        df.groupby(['EECode', 'Lastname', 'Firstname'], as_index=False)['EarnHours'].sum(min_count = 1)
    )
    summary.rename(columns = {'EarnHours': 'Meal Penalty Hours'}, inplace = True)
    create_sheet(file, summary.to_dict(orient="records"), 'Summary')

def create_clp_report(file):
    df = pd.read_excel(file)

    sheet_folder = os.path.join(local_path, 'sheet_clp')
    os.makedirs(sheet_folder, exist_ok=True)

    for group, df_group in df.groupby('Schedule Group.1'):
        exc_filepath = f'{group}.xlsx'
        df_group.to_excel(os.path.join(sheet_folder, exc_filepath))
        print(f'Created file for {group}')
    
    for file in os.listdir(sheet_folder):
        create_clp_summary(os.path.join(sheet_folder, file))

def send_email_clp(filepath):
    outlook = win32.Dispatch("Outlook.Application")
    
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet_clp')

    for file in os.listdir(sheet_folder):
        mail = outlook.CreateItem(0)
        df = pd.read_excel(os.path.join(sheet_folder, file), sheet_name='Summary')

        sheet_number = str(os.path.basename(file)).replace('.xlsx', '')

        try:
            subject_header = config.loc[config['Group_Code'] == int(sheet_number), 'NAME_sched'].values[0]
            #Commented because of testing
            emailto = config.loc[config['Group_Code']== int(sheet_number), 'SEND_to'].values[0]
            ccto = config.loc[config['Group_Code']== int(sheet_number), 'CC_S'].values[0]
        except Exception as e:
            subject_header = 'Schedule Group Not in Config'
            emailto = err_email
            ccto = err_cc
            print(f'Exception: {e}')

        df.sort_values(by = 'Meal Penalty Hours', ascending=False, inplace= True)
        df_clp = df.loc[df['Meal Penalty Hours'] > 0, ['EECode', 'Lastname', 'Firstname', 'Meal Penalty Hours']]
        if len(df_clp)>0:
            df_clp_summary_to_html = df_clp.to_html(index=False)
            df_clp_summary_to_html = df_clp_summary_to_html.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")
        else:
            df_clp_summary_to_html = "<p><b style = text-transform:uppercase;>No meal penalties for this period.</b></p>"

        mail.Attachments.Add(os.path.join(sheet_folder, file))

        mail.To = emailto
        mail.CC = ccto
        mail.Subject = f'CA Meal Penalty Report from {first_day_of_month.strftime("%m/%d/%Y")} -  {datetoday.strftime("%m/%d/%Y")}, {subject_header}'
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
        
    os.remove(filepath)

def send_email_missing(filepath):
    outlook = win32.Dispatch("Outlook.Application")
    
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet_missing')

    for file in os.listdir(sheet_folder):
        mail = outlook.CreateItem(0)
        df = pd.read_excel(os.path.join(sheet_folder, file))

        sheet_number = str(os.path.basename(file)).replace('.xlsx', '')

        try:
            subject_header = config.loc[config['Group_Code'] == int(sheet_number), 'NAME_sched'].values[0]
            emailto = config.loc[config['Group_Code']== int(sheet_number), 'SEND_to'].values[0]
            ccto = config.loc[config['Group_Code']== int(sheet_number), 'CC_S'].values[0]
            # emailto = 'vjdelrosario@avatco.com'
            # ccto = 'vjdelrosario@avatco.com;TTPhan@avatco.com'
        except Exception as e:
            subject_header = 'Schedule Group Not in Config'
            emailto = err_email
            ccto = err_cc
            print(f'Exception: {e}')
        

        #Sort df date by ascending order
        df.sort_values(by= 'Date', ascending = True, inplace = True)
        df_missing_summary = df[['EE Code', 'Date' ,'Last Name', 'First Name', 'In Punch Time', 'Out Punch Time']]
        df_missing_summary_to_html = df_missing_summary.to_html(index=False)
        df_missing_summary_to_html = df_missing_summary_to_html.replace("<thead", "<thead style='background-color:#FF1A1A; color:white;'")


        mail.Attachments.Add(os.path.join(sheet_folder, file))

        mail.To = emailto
        mail.CC = ccto
        mail.Subject = f'Missing Punches Report from {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")} {subject_header}'
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
                <p>GOOD DAY, HERE IS THE SUMMARY OF <b>MISSING PUNCHES</b> FROM {first_day_of_month.strftime("%m/%d/%Y")} - {datetoday.strftime("%m/%d/%Y")}</p>
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
        
    os.remove(filepath)


paycom_filepath = paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 1)
create_report(paycom_filepath)
send_email(paycom_filepath)
missingpunches_filepath = paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 2)
create_missing_report(missingpunches_filepath)
send_email_missing(missingpunches_filepath)
clp_filepath = paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 3)
create_clp_report(clp_filepath)
send_email_clp(clp_filepath)





