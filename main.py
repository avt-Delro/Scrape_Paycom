import traceback

import web_scraping
import environment as env
from dotenv import load_dotenv
from datetime import datetime
import calendar
import pandas as pd
import weekly_scraping
from logging_conf import setup_logger, send_email_err_report

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

logger, log_file = setup_logger('paycom_scraping')


def main():
    try:
        paycom_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 1)
        web_scraping.create_report(paycom_filepath)
        web_scraping.send_email(paycom_filepath)
        logger.info("Daily OT report created and email sent.")
        missingpunches_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 2)
        web_scraping.create_missing_report(missingpunches_filepath)
        web_scraping.send_email_missing(missingpunches_filepath)
        logger.info("Daily Missing Punches report created and email sent.")
        clp_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 3)
        web_scraping.create_clp_report(clp_filepath)
        web_scraping.send_email_clp(clp_filepath)
        logger.info("Daily CLP report created and email sent.")
        
        if day_today.weekday() == 0:  # Check if it's Monday
            weekly_scraping.create_report_summary_we_month(paycom_filepath, None)
            weekly_scraping.create_clp_summary_we_month(clp_filepath, None)
            weekly_scraping.send_email_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com',paycom_filepath)
            weekly_scraping.send_email_missing_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', missingpunches_filepath)
            weekly_scraping.send_email_clp_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', clp_filepath)

        if datetoday.day == last_day:  # Check if it's the last day of the month
            weekly_scraping.send_email_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com',paycom_filepath)
            weekly_scraping.send_email_missing_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com', missingpunches_filepath)
            weekly_scraping.send_email_clp_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com', clp_filepath)
    except Exception as e:
        print(f"⚠️ Error processing message: {e}")
        error_type = type(e).__name__
        error_message = str(e)
        stack_trace = traceback.format_exc()
        logger.error(f"Error Type: {error_type}\nError Message: {error_message}\nStack Trace: {stack_trace}")
        send_email_err_report(err_email, error_message, error_type, stack_trace)


if __name__ == "__main__":
    main()
    # paycom_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 1)
    # missingpunches_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 2)
    # clp_filepath = web_scraping.paycom_scraping('https://www.paycomonline.net/v4/cl/cl-login.php', paycom_user, paycom_pass, client_code, 3)
    # weekly_scraping.create_report_summary_we_month(paycom_filepath, None)
    # weekly_scraping.create_clp_summary_we_month(clp_filepath, None)
    # # # weekly_scraping.send_email_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com',paycom_filepath)
    # # # weekly_scraping.send_email_missing_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', missingpunches_filepath)
    # # # weekly_scraping.send_email_clp_we_month('raquel@avatco.com;mason@avatco.com;Frank.Shi@avatco.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', clp_filepath)
    # weekly_scraping.create_report_summary_we_month(paycom_filepath, {'Schedule Group.1':[8606,8607,8608,8609,8610,8619,8620]})
    # # weekly_scraping.create_clp_summary_we_month(clp_filepath, {'Schedule Group.1':[8606,8607,8608,8609,8610,8619,8620]})
    # weekly_scraping.send_email_we_month('ever@americantirestores.com;leon@tireoutletus.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', paycom_filepath, 'Retail Stores')
    # # weekly_scraping.send_email_we_month('vjdelrosario@avatco.com', 'vjdelrosario@avatco.com', paycom_filepath, 'Retail Stores')
    # weekly_scraping.send_email_missing_we_month('ever@americantirestores.com;leon@tireoutletus.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', missingpunches_filepath, {'Schedule Group.1':[8606,8607,8608,8609,8610,8619,8620]}, 'Retail Stores')
    # weekly_scraping.send_email_clp_we_month('ever@americantirestores.com;leon@tireoutletus.com', 'vjdelrosario@avatco.com;TTPhan@avatco.com', clp_filepath, 'Retail Stores') 
    

    
    
    




