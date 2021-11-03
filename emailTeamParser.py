import win32com.client
import os
import glob
import dateutil.parser
from datetime import datetime, timedelta
import openpyxl
import colorama
from colorama import Fore, Back, Style

colorama.init(autoreset=True)

SAVE_DIR = os.path.join(os.getcwd(), "teams")
EXCEL_NAME = "candidates.xlsx"
STANDARD_EMAIL_SENDER_ADDRESS = "julien.foucault@accuracy.com"

def EmailGetterSaver():
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    # Connecting to the right inbox
    inbox = mapi.Folders("Accuracy Business Cup").Folders("Inbox").Folders("ABC 2021").Folders("03. Candidatures")

    # Building the message list of all messages sent by the applying server
    standardMessages = []
    for message in inbox.Items:
        if message.Class == 43:
            if message.SenderEmailType == 'EX':
                if STANDARD_EMAIL_SENDER_ADDRESS == message.Sender.GetExchangeUser().PrimarySmtpAddress:
                    standardMessages.append(message)
            else:
                if STANDARD_EMAIL_SENDER_ADDRESS == message.SenderEmailAddress:
                    standardMessages.append(message)

    # Save attachments and create team folders
    ## Create the teams directory if not already there
    attachmentsDir = SAVE_DIR
    if not os.path.isdir(attachmentsDir):
        os.makedirs(attachmentsDir)
        print(f"Created folder {attachmentsDir}")

    ## Iterate over the messages
    messages_list = []
    try:
        for message in list(standardMessages):
            # Get the latest team folder name
            listOfTeams = glob.glob(os.path.join(attachmentsDir, "*"))
            try:    
                latestTeam = os.path.basename(max(listOfTeams, key=os.path.getmtime))
            except Exception as e:
                latestTeam = "team0"

            try:
                s = message.subject
                teamDir = os.path.join(attachmentsDir, s + "-team" + str(int(latestTeam[-1:]) + 1))

                # Create the team folder if it doesn't exist and the captain hasn't created a team before, and save all attachments         
                if not s in map(lambda team : os.path.basename(team)[:os.path.basename(team).find("-")], listOfTeams):
                    # Build the dictionnary with this message's info to be sent
                    message_dict = dict.fromkeys(["receivedTime", "team", "body"])
                    for key in message_dict.keys():
                        if key == "receivedTime":
                            message_dict[key] = str(getattr(message, key.capitalize())).rstrip("+00:00").strip()
                        elif key == "team":
                            message_dict[key] = str(int(latestTeam[-1:]) + 1)
                        else:
                            message_dict[key] = getattr(message, key.capitalize())
                    messages_list.append(message_dict)

                    if not os.path.isdir(teamDir):
                        os.makedirs(teamDir)
                        print(Fore.GREEN + f"Created folder {teamDir}")

                    for attachment in message.Attachments:
                        if attachment.Filename[:5] != "image":
                            attachment.SaveASFile(os.path.join(teamDir, attachment.FileName))
                            print(f"Attachment {attachment.FileName} from captain {s} saved")
                else:
                    print(Fore.YELLOW + f"Captain {s} has already created a team")

            except Exception as e:
                print(Fore.RED + "Error when saving the attachment:" + str(e))
    except Exception as e:
            print(Fore.RED + "Error when processing emails messages:" + str(e))

    return messages_list

def messagesListParser(messages_list):
    # Check if candidates list excel exists, creating it if not
    path = os.path.join(os.getcwd(), EXCEL_NAME)
    if not os.path.exists(path):
        print(Fore.GREEN + "\nCandidates excel doesn't exist, creating it")
        workbook = openpyxl.Workbook()
        headers = ("Team", "First Name", "Last Name", "email", "School", "Email Sent?")
        sheet = workbook.active
        sheet.append(headers)
        workbook.save(filename=EXCEL_NAME)
    candidates_wb = openpyxl.load_workbook(path)
    sheet = candidates_wb.active

    # Parse each message for team info
    for i, message in enumerate(messages_list):    
        team = message["team"]
        #print(team)
        candidateList = message["body"].split("CANDIDATE")[1:]
        for candidate in candidateList:
            firstName = candidate.split("First name : ")[1][:candidate.split("First name : ")[1].find("\r")]
            if " Last" in firstName:    
                firstName = firstName[:firstName.find(" Last")]
            #print(firstName, " ", len(firstName))
        
            lastName = candidate.split("Last name : ")[1][:candidate.split("Last name : ")[1].find("\r")]
            if " Email" in lastName:    
                lastName = lastName[:lastName.find(" Email")]
            #print(lastName, " ", len(lastName))

            email = candidate.split("Email address : ")[1][:candidate.split("Email address : ")[1].find("\r")]
            if " <mailto" in email:    
                email = email[:email.find(" <mailto")]
            #print(email, " ", len(email))

            school = candidate.split("University : ")[1][:candidate.split("University : ")[1].find("\r")]
            if "<http://www.accuracy.com>" in school:    
                school = school[:school.find("<http://www.accuracy.com>")]
            if " SECOND" in school:
                school = school[:school.find(" SECOND")]
            if " THIRD" in school:
                school = school[:school.find(" THIRD")]
            #print(school, " ", len(school))

            # Append data to excel
            candidateData = (team, firstName, lastName, email, school, 0)
            sheet.append(candidateData)
    candidates_wb.save(EXCEL_NAME)
    return None

if __name__ == "__main__":
    messages_list = EmailGetterSaver()
    messagesListParser(messages_list)