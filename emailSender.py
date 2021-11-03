import win32com.client
import os
import glob
import dateutil.parser
from datetime import datetime, timedelta
import openpyxl
import colorama
from colorama import Fore, Back, Style
from emailTeamParser import EXCEL_NAME

# Mail inputs
MAIL_SUBJECT = "Candidature bien reçue!"

def mailHTMLBodyBuilder(firstName, lastName):
    customBody = f"""<h3>Bonjour {firstName} {lastName},</h3>

Merci pour votre candidature à la <b>Accuracy Business Cup</b>.<br>
Elle a bien été reçue et nous reviendrons vers vous dans les plus brefs délais"""
    return customBody

colorama.init(autoreset=True)

def emailSender():
    outlook = win32com.client.Dispatch('outlook.application')

    # Get the candidates sheet
    path = os.path.join(os.getcwd(), EXCEL_NAME)
    if not os.path.exists(path):
        print(Fore.RED + "\nCandidates excel doesn't exist")
        return
    candidates_wb = openpyxl.load_workbook(path)
    sheet = candidates_wb.active

    listOfMails = []
    for row in sheet.rows:
        if row[5].value == 0:
            mail=outlook.CreateItem(0)
            mail.To = row[3].value
            mail.Subject = MAIL_SUBJECT
            mail.HTMLBody = mailHTMLBodyBuilder(row[1].value, row[2].value)
            listOfMails.append(mail)

            print(Fore.GREEN + f"Adding {row[3].value} to the list")
        elif row[5].value == 1:
            print(Fore.YELLOW + f"Already sent email to {row[3].value}")

    print("Mails will be sent to the following targets:")
    for i, mail in enumerate(listOfMails):
        print(i + 1, mail.To)
    _input = input("Send to these targets? ( y / n )")
    if _input == "y":
        for i, mail in enumerate(listOfMails):
            target = mail.To
            mail.Send()
            for row in sheet.rows:
                if row[3].value == target and row[5].value == 0:
                    row[5].value = 1
                    print(Fore.GREEN + f"Sent mail to {row[3].value}")
        print(Fore.GREEN + f"DONE")
    else :
        print(Fore.YELLOW + f"Aborted the send")
    candidates_wb.save(EXCEL_NAME)

if __name__ == "__main__":
    emailSender()