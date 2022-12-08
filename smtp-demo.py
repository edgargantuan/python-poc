import cx_Oracle
import zipfile
import csv
import smtplib
from os.path import basename
from os import remove
from email.mime.base import MIMEBase
from email import message
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email import encoders

"""
Enable port forwarding (putty)

110.173.155.69:22
L21521 127.0.0.1:1521
"""

def generate_data():

    with cx_Oracle.connect(user="etan", password="tsitcot9p", dsn="localhost:21521/XE") as conn:
        with conn.cursor() as cur:
            cur.execute(
            """
            SELECT FIN_YEAR, 
                REPORT_ID, 
                COUNCIL_ID, 
                COUNCIL, 
                NUMBEROFSTAFF, 
                DATRACKER, 
                PUBLIC_DATRACKER, 
                APPLICATION_ID, 
                APPLICATIONTYPE, 
                COUNCILREFERENCENUMBER, 
                CONSTRUCTIONVALUEESTIMATE, 
                VALUE_CODE, 
                LODGEMENTDATE, 
                CATEGORYOFDEVELOPMENT, 
                CONCURRENCEREQUIRED, 
                DESIGNATEDDEVELOPMENT, 
                INTEGRATEDDEVELOPMENT, 
                STOPTHECLOCKDAYS, 
                REFERRALDAYS, 
                YES_STC, 
                YES_REFERRAL, 
                YES_BOTH, 
                DETERMINED, 
                DETERMINED_YEAR, 
                DETERMINED_MONTH, 
                DETERMINATIONLEVEL, 
                DETERMINATIONTYPE, 
                IN_OUT, 
                SEPP_COUNCIL_CONTROL, 
                CDC_VALID_DATE, 
                GROSS_DAYS, 
                NET_DAYS, 
                NEW_DWELLINGS, 
                EXISTING_DWELLINGS, 
                DEMOLISHED_DWELLING, 
                DONATION_BY_DEVELOPER, 
                DONATION_BY_SUB, 
                ARH_CATEGORYOFDEVELOPMENT, 
                NUMBER_OF_DWELLING, 
                SITE_COMPATIBILITY_CERT
            FROM allan_dba.qa_allfinaldata
            """
            )

            with open(r"D:\LDPM\auto\output\ldpm_allfin.csv", 'w', newline='', encoding='UTF8') as f:
                tof = csv.writer(f)
                hdr = ("FIN_YEAR", "REPORT_ID", "COUNCIL_ID", "COUNCIL", "NUMBEROFSTAFF", "DATRACKER", "PUBLIC_DATRACKER", "APPLICATION_ID", "APPLICATIONTYPE", "COUNCILREFERENCENUMBER", "CONSTRUCTIONVALUEESTIMATE", "VALUE_CODE", "LODGEMENTDATE", "CATEGORYOFDEVELOPMENT", "CONCURRENCEREQUIRED", "DESIGNATEDDEVELOPMENT", "INTEGRATEDDEVELOPMENT", "STOPTHECLOCKDAYS", "REFERRALDAYS", "YES_STC", "YES_REFERRAL", "YES_BOTH", "DETERMINED", "DETERMINED_YEAR", "DETERMINED_MONTH", "DETERMINATIONLEVEL", "DETERMINATIONTYPE", "IN_OUT", "SEPP_COUNCIL_CONTROL", "CDC_VALID_DATE", "GROSS_DAYS", "NET_DAYS", "NEW_DWELLINGS", "EXISTING_DWELLINGS", "DEMOLISHED_DWELLING", "DONATION_BY_DEVELOPER", "DONATION_BY_SUB", "ARH_CATEGORYOFDEVELOPMENT", "NUMBER_OF_DWELLING", "SITE_COMPATIBILITY_CERT")
                tof.writerow(hdr)
                for rec in cur:
                    tof.writerow(rec)
            
            remove(r"D:\LDPM\auto\output\ldpm_allfin.zip")
            with zipfile.ZipFile(r"output\ldpm_allfin.zip", "a", zipfile.ZIP_DEFLATED) as zip:
                zip.write(r"D:\LDPM\auto\output\ldpm_allfin.csv", "ldpm_allfin.csv")


def send_data():

    from_addr = "ePlanning@eplanning.planning.nsw.gov.au"
    to_addr = "edgar.tan@planning.nsw.gov.au"
    subject = 'LDPM daily allfin extract'
    content = "Please find attached allfin extract.\n\nThis email is auto-generated."

    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr 
    msg['Subject'] = subject

    body = MIMEText(content, 'plain')
    msg.attach(body)

    filename = r"output\ldpm_allfin.zip"
    with open(filename, 'rb') as f:
        attachment = MIMEBase('application', 'zip')
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment', filename=basename(filename))

    msg.attach(attachment)

    server = smtplib.SMTP('email-smtp.us-east-1.amazonaws.com', 587)
    server.ehlo()
    server.starttls()
    server.login('A******Q', 'A*****x')
    server.send_message(msg, from_addr=from_addr, to_addrs=[to_addr])


if __name__ == "__main__":
    generate_data()
    send_data()

