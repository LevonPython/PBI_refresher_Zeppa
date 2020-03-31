import pypyodbc
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sql_server_pbixrefresher as pref
import pyautogui
import subprocess


def send_mail(new_records):
    fromaddr = "levon.python@gmail.com"
    toaddr = "levon.python@gmail.com"
    # instance of MIMEMultipart
    msg = MIMEMultipart()
    # storing the senders email address
    msg['From'] = fromaddr
    # storing the receivers email address
    msg['To'] = toaddr
    # storing the subject
    msg['Subject'] = f"SQL Server zeppa records {new_records}"
    # string to store the body of the mail
    body = "New records have been updated"
    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))
    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)
    # start TLS for security
    s.starttls()
    # Authentication
    s.login(fromaddr, "lqglrrtepepxamvz")
    # for specific password
    # https://myaccount.google.com/security?rapt=AEjHL4PlAh6rv38PgIokqo4MNOKXlPRfjcyZlgvwF4by81rOQ8XUhwWJnn12AzOa5
    # VS-vGqITVfOyHZQzutJNc-grjyjZtI4sQ
    print("login success")
    # Converts the Multipart msg into a string
    text = msg.as_string()
    # sending the mail
    s.sendmail(fromaddr, toaddr, text)
    # terminating the session
    s.quit()
    print("sent . . .")


def update_zeppa_apps(rows_count=361):
    """
    This script automatically executes the stored procedure inside the asReal Database,
    which updates table zeppa_table for Power BI report
    :return:
    """
    myuser = "acramon"
    gaghtnagir = "2{sD6@.[SsF0[#jq"
    sql = "EXEC dbo.ZeppaApps @productId = 1"
    sql2 = "SELECT COUNT(acraId) FROM zeppa_table"
    while True:
        db_connection = pypyodbc.connect(
            driver='{SQL Server}', server='192.168.0.54', database='asReal', UID=myuser, PWD=gaghtnagir)
        cursor = db_connection.cursor()
        print(f"Rows_count start-point is: {rows_count}")
        try:
            cursor.execute(sql)
            rows_new = list(cursor.execute(sql2))[0][0]
            if rows_count == rows_new:
                cursor.close()
                db_connection.close()
                print("connection closed, new apps NOT found, sleep ...")
                time.sleep(300)
                print("trying to find new apps ...")
            else:
                print("new apps found")
                print(f"Here are already {rows_new}")
                db_connection.commit()
                cursor.close()
                db_connection.close()
                print("connection closed with NEW apps, sleep ...")
                send_mail(rows_new)
                # pbi block
                path_pbi_file = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
                subprocess.Popen(path_pbi_file, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE,
                                 universal_newlines=True, shell=False)
                time.sleep(18)
                pref.PbiRefresher().open_file()
                # pref.open_file()

                time.sleep(30)
                update_zeppa_apps(rows_new)
        except Exception as e:
            print("Oops!")
            print(e)
            if e.args[0] in ("01000", "08S01", "08001"):
                print("*" * 100)
                time.sleep(5)
                try:
                    print("trying to connect again ...")
                    db_connection = pypyodbc.connect(
                        driver='{SQL Server}', server='192.168.0.54', database='asReal', UID=myuser, PWD=gaghtnagir)
                    cursor = db_connection.cursor()
                finally:
                    try:
                        print("plan b started")
                        time.sleep(10)
                        update_zeppa_apps()
                        break
                    except Exception as e2:
                        print("fail occurred, trying again ...")
                        print(e2.args)
                        time.sleep(1200)
                        db_connection = pypyodbc.connect(
                            driver='{SQL Server}', server='192.168.0.54', database='asReal', UID=myuser,
                            PWD=gaghtnagir)
                        cursor = db_connection.cursor()
                        print("plan c started")
                        time.sleep(10)
                        update_zeppa_apps()


if __name__ == "__main__":
    update_zeppa_apps()
