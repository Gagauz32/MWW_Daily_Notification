import pyodbc
import pandas
import smtplib
import os
from datetime import datetime, timedelta
from pretty_html_table import build_table
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_email(text):
    pwd = os.environ['EMAIL_PWD']
    my_email = os.environ['EMAIL_SENDER']
    receivers = os.environ['EMAIL_RECEIVERS']

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "MWW Notification:"
    msg["From"] = my_email
    msg["To"] = receivers
    msg_start = """
    <html>
        <body>
             <h2>FL programate în urmatoarele 60 zile:</h2><br/>
    """
    msg_body = text
    msg_end = """
        </body>
    </html>
    """
    email_text = msg_start + msg_body + msg_end
    html_message = MIMEText(email_text, "html")
    msg.attach(html_message)
    with smtplib.SMTP(host="smtp.office365.com", port=587) as smtp_connection:
        smtp_connection.starttls()
        smtp_connection.login(user=my_email, password=pwd)
        smtp_connection.sendmail(my_email, receivers, msg.as_string())


def get_db_data():
    db_ip = os.environ["DB_IP"]
    db_name = os.environ["DB_NAME"]
    db_uid = os.environ["DB_UID"]
    db_pwd = os.environ["DB_PWD"]
    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={db_ip};DATABASE={db_name};UID={db_uid};PWD={db_pwd};')
    sql = ('''SELECT TOP (1000) [DatProCld]
      ,[CodOrdClb]
      ,[CodObjeto]
      ,[DscOrdClb]
      ,[CodFunCnr] FROM [MWW_Imcomvil].[dbo].[Calibra] ORDER BY [DatProCld]''')
    raw_data = pandas.read_sql(sql, conn)
    conn.close()
    return raw_data


today = datetime.today()
next_week = []
num_of_days = 0

for i in range(60):
    next_day = today + timedelta(days=num_of_days)
    next_week.append(next_day.strftime("%Y-%m-%d"))
    num_of_days += 1

dataframe = get_db_data()
new_dataframe = pandas.DataFrame()
new_list = []
row_number = 0

for row in dataframe.DatProCld:
    new_row = str(row).split(sep=" ")[0]
    if new_row in next_week:
        new_list.append(dataframe.loc[row_number].tolist())
    row_number += 1

for item in new_list:
    item[0] = item[0].strftime("%d.%m.%Y - %H:%M")

new_df = pandas.DataFrame(new_list, columns=["Data și ora", "Nr.FL", "Echipament", "Descrierea FL", "Persoana Asignată"])
html_table = build_table(new_df, 'green_dark', font_family='Nunito', font_size='14px', width='200px')
send_email(html_table)
