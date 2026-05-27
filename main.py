import web_scraping
import environment as env
from dotenv import load_dotenv
from datetime import datetime
import calendar
import pandas as pd
import weekly_scraping
load_dotenv()
paycom_user = env.paycom_username
paycom_pass = env.paycom_password  
client_code = env.paycom_clientcode 

local_path = env.paycom_local
email = env.sendemail

datetoday = datetime.today()
day_today = datetime.now()


first_day_of_month = datetoday.replace(day=1)
last_day = calendar.monthrange(datetoday.year, datetoday.month)[1]


config = pd.read_csv('config/config.csv')

err_email = 'vjdelrosario@avatco.com'
err_cc = 'vjdelrosario@avatco.com'



def main():
    paycom_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 1)
    web_scraping.create_report(paycom_filepath)
    web_scraping.send_email(paycom_filepath)
    missingpunches_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 2)
    web_scraping.create_missing_report(missingpunches_filepath)
    web_scraping.send_email_missing(missingpunches_filepath)
    clp_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 3)
    web_scraping.create_clp_report(clp_filepath)
    web_scraping.send_email_clp(clp_filepath)
    
    if day_today.weekday() == 0:  # Check if it's Monday
        weekly_scraping.create_report_we_month(paycom_filepath)
        weekly_scraping.create_report_summary_we_month(paycom_filepath)
        weekly_scraping.create_report_clp_we_month(clp_filepath)
        weekly_scraping.send_email_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com',paycom_filepath)
        weekly_scraping.send_email_missing_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com', missingpunches_filepath)
        weekly_scraping.send_email_clp_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com', clp_filepath)

    if datetoday.day == last_day:  # Check if it's the last day of the month
        weekly_scraping.send_email_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com',paycom_filepath)
        weekly_scraping.send_email_missing_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com', missingpunches_filepath)
        weekly_scraping.send_email_clp_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com', clp_filepath)
    



if __name__ == "__main__":
    main()
    
    
    




