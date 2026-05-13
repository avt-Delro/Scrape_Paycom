
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



outlook = win32.Dispatch("Outlook.Application")
outlook_ap = outlook.GetNamespace("MAPI")
config = pd.read_csv('config/config.csv')



def paycom_scraping(weblink, username, password, client_code):
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
            page.get_by_role("row", name="Favorite Actual v Scheduled w").locator("input[type=\"button\"]").click()
            page.wait_for_load_state("load") 
            
            with page.expect_download() as download_info:
                page.get_by_role("button", name="Download").click()
            download = download_info.value


            file_folder = os.path.join(local_path, 'HR Files')
            os.makedirs(file_folder, exist_ok=True)

            file_name =  f'OT_report_{datetoday.strftime("%Y%m%d")}.xlsx'

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


def send_email (email, filepath):
    outlook = win32.Dispatch("Outlook.Application")
    outlook_ap = outlook.GetNamespace("MAPI")
    sheet_folder = os.path.join(local_path, 'sheet')

    for file in os.listdir(sheet_folder):
        mail = outlook.CreateItem(0)
        df = pd.read_excel(os.path.join(sheet_folder, file), sheet_name='Summary')

        sheet_number = str(os.path.basename(file)).replace('.xlsx', '')

        subject_header = config.loc[config['Group_Code'] == int(sheet_number), 'NAME_sched'].values[0]
        emailto = config.loc[config['Group_Code']== int(sheet_number), 'SEND_to'].values[0]

        # Commented because this is for the summary 
        df_without_scheduled_hours = df.loc[df['Scheduled Hours'] == 0, ['Employee', 'Scheduled Hours', 'Actual Hours']]
        df_negative_variance = df.loc[df['Variance'] < 0, ['Employee', 'Variance']]
        df_positive_variance = df.loc[df['Variance'] > 0, ['Employee', 'Variance']]

        html_without = df_without_scheduled_hours.to_html(index=False)
        html_negative = df_negative_variance.to_html(index=False)
        html_positive = df_positive_variance.to_html(index=False)
        html_df = df.to_html(index=False)

        mail.Attachments.Add(os.path.join(sheet_folder, file))

        mail.To = emailto
        mail.Subject = f'OT Report: {datetoday.strftime("%m/%d/%Y")} {subject_header}'
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
                <p>Good day, Here are the summary of work hours as of: {datetoday.strftime("%m/%d/%Y")}</p>
                <p>Attached is the report file: {os.path.basename(filepath)}</p>

                <h2>Summary of the Report:</h2>
                {html_df}
                <br>

                <h1>Employees with Positive Variance</h1>
                {html_positive}
                <br>

                <h1>Employees with Negative Variance</h1>
                {html_negative}
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



paycom_filepath = paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code)
create_report(paycom_filepath)
send_email(email, paycom_filepath)




