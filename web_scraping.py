
from playwright.sync_api import sync_playwright
import environment as env
from dotenv import load_dotenv
import win32com.client as win32
import os
from datetime import datetime
import calendar


load_dotenv()
paycom_user = env.paycom_username
paycom_pass = env.paycom_password  
client_code = env.paycom_clientcode 

local_path = env.paycom_local
email = env.sendemail

datetoday = datetime.today()


outlook = win32.Dispatch("Outlook.Application")
outlook_ap = outlook.GetNamespace("MAPI")

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
            page.get_by_role("row", name=f"{date} Actual v Scheduled").get_by_label("More options dropdown").click()
            page.get_by_role("link", name="View Files").click()
            with page.expect_download() as download_info:
                page.get_by_role("link", name="Download").click()
                download = download_info.value

            file_folder = os.path.join(local_path, 'HR Files')
            os.makedirs(file_folder, exist_ok=True)

            download.save_as(os.path.join(file_folder, download.suggested_filename))

            return f"{str(file_folder)}\{download.suggested_filename}"

    except Exception as e:
        print(e)

def send_email (email, filepath):
    outlook = win32.Dispatch("Outlook.Application")
    outlook_ap = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(0)

    mail.Attachments.Add(filepath)

    mail.To = email
    mail.Subject = f'See attached file, for Paycom Data for the date: {datetoday.strftime("%m/%d/%Y")}'
    mail.Send()
    print('Email Sent')



paycom_filepath = paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code)
send_email(email, paycom_filepath)