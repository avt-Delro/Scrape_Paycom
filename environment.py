import os
from dotenv import load_dotenv


load_dotenv()

paycom_username = os.getenv('web_username')
paycom_password = os.getenv('web_password')
paycom_clientcode = os.getenv('client_code')
paycom_local = os.getenv('localroot')

sendemail = os.getenv('send_em')