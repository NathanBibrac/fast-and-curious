import matplotlib.pyplot as plt
import win32com.client
import os
from datetime import datetime
import time

class timestamp():
    def __init__(self):
        self.now = datetime.now()
        self.ts_filename = self.now.strftime("%Y%m%d_%H%M%S")
        self.ts_display = self.now.strftime("%Y-%m-%d %H:%M:%S")
    
    def update(self):
        self.now = datetime.now()
        self.ts_filename = self.now.strftime("%Y%m%d_%H%M%S")
        self.ts_display = self.now.strftime("%Y-%m-%d %H:%M:%S")

    def ts_filename(self):
        self.update()
        return self.ts_filename
    
    def ts_display(self):
        self.update()
        return self.ts_display
    

class stats_email_per_month():
    def __init__(self,year):
        self.year = year
        self.stats = {}

    def plot_graph(self):
        ts = timestamp()
        plt.figure(figsize=(20, 10),label = f"Number of emails received per month in {self.year}")
        plt.bar(self.stats.keys(), self.stats.values())
        plt.title = f"Number of emails received per month in {self.year}"
        plt.ylabel = "Number of emails"
        plt.xlabel = f"Month ({self.year})"
        plt.show()
        ts_str = ts.ts_filename
        plt.savefig(f'D:\\Lab\\data\\images\\graphs\\stats_email_by_tear_month_{str(self.year)}-{ts_str}.png')
        

    def print_upper_stats(self,K):
        # Affichez les statistiques
        for sender, count in self.stats.items():
            if count > K:
                print(f"{sender}: {count}")



def send_mail_attachment():

    ts = timestamp()
    print(f'{ts.ts_display()}: Launching send_mail_attachment()')

    ol = win32com.client.Dispatch('Outlook.Application')
    # size of the new email
    olmailitem = 0x0

    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.To = 'n.bibrac.pro@gmail.com'
    newmail.CC = 'nathan.bibrac@amundi.com'

    newmail.Body= 'Hi, here\'s the sheatcode Alto PDF.'

    # attach_img = 'D:\\Lab\\data\\images\\test\\PP_FB.jpg'
    # attach_pdf = r'D:\Lab\Python\documentation\Finxter_OpenAI_Glossary.pdf'
    attach_pysheat_dir = r'D:\Lab\Python\documentation\finxter'
    for file in os.listdir(attach_pysheat_dir):
        if file.endswith(".pdf"):
            attach_pysheat = os.path.join(attach_pysheat_dir, file)
            newmail.Attachments.Add(attach_pysheat)

    newmail.Display()
    newmail.Send()

    print(f'{ts.ts_display()}: Email sent')


def read_mail(attachment = True):
    ts = timestamp()
    print(f'{ts.ts_display()}: Launching read_mail()')

    #  retrieve last email received in the inbox
    ol = win32com.client.Dispatch('Outlook.Application')
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    messages = inbox.Items
    message = messages.GetLast()
    body_content = message.body

    if attachment:
        attachments = message.Attachments
        i = 0
        for attachment in attachments:
            attachment.SaveAsFile(r'D:\Lab\data\pdf\test\outlook_mail' + f' {i} ' + '.pdf')
            i = i+1
    
    print(body_content)
    print(f'{ts.ts_display}: Email read')


# get the object and the sender of the 20 last emails received in a dict
def get_last_emails(k=20):

    ts = timestamp()
    print(f'{ts.ts_display}: Launching get_last_emails()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", True)  # Trie les emails par date de réception, True signifie en ordre décroissant

    last_emails = []
    for i in range(k):
        if i < len(emails):
            email = emails[i]
            sender = email.SenderName
            subject = email.Subject
            last_emails.append({"Sender": sender, "Subject": subject})
    
    print(f'{ts.ts_display}: Emails retrieved')
    return last_emails

def move_emails_to_folder():

    ts = timestamp()
    print(f'{ts.ts_display}: Launching move_emails_to_folder()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox
    linkedin_folder_name = "Linkedin"  # Le nom du dossier de destination
    uber_eats_folder_name = "Uber Eats"
    lydia_folder_name = "Lydia"
    
    # Créez ou obtenez le dossier "Linkedin" s'il n'existe pas déjà
    try:
        linkedin_folder = inbox.Folders.Item(linkedin_folder_name)
    except:
        linkedin_folder = inbox.Folders.Add(linkedin_folder_name)
    
    try:
        Uber_Eats_folder = inbox.Folders.Item(uber_eats_folder_name)
    except:
        Uber_Eats_folder = inbox.Folders.Add(uber_eats_folder_name)

    try:
        lydia_folder = inbox.Folders.Item(lydia_folder_name)
    except:
        lydia_folder = inbox.Folders.Add(lydia_folder_name)

    # Récupérez les e-mails de la boîte de réception
    emails = inbox.Items
    
    # Parcourez tous les e-mails
    for email in emails:
        try:
            if email.SenderName == "LinkedIn" :
                # Déplacez l'e-mail vers le dossier "Linkedin"
                email.Move(linkedin_folder)
            elif email.SenderName == "Uber Eats" :
                # Déplacez l'e-mail vers le dossier "Uber Eats"
                email.Move(Uber_Eats_folder)
            elif email.SenderName == "Lydia App" :
                # Déplacez l'e-mail vers le dossier "Uber Eats"
                email.Move(lydia_folder)
        except:
            pass

    print(f'{ts.ts_display}: Emails moved')
    

# Déplacer tous les mails de 2020 dans un dossier appelé 2020
def move_emails_to_folder_archive(year):

    ts = timestamp()
    print(f'{ts.ts_display}: Launching move_emails_to_folder_archive()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox
    folder_name = str(year)  # Le nom du dossier de destination
    
    print(folder_name)
    # Créez ou obtenez le dossier "2020" s'il n'existe pas déjà
    try:
        folder = inbox.Folders.Item(folder_name)
    except:
        folder = inbox.Folders.Add(folder_name)
        print(f"{ts.ts_display}: Folder {folder_name} created")
    
    # Récupérez les e-mails de la boîte de réception
    emails = inbox.Items
    i=0
    # Parcourez tous les e-mails
    for email in emails:
        i = i+1
        print(f"{ts.ts_display}:Progression : {i} / {emails.Count}")
        try:
            # Vérifiez si l'e-mail a été reçu en 2020
            if email.ReceivedTime.year == year:
                ReceivedTime = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                print(f"{ts.ts_display}: Email received on the {ReceivedTime} == {year}")
                # Déplacez l'e-mail vers le dossier "2020"
                email.Move(folder)

        except:
            pass

def archivage_past_years():
    ts = timestamp()
    print(f'{ts.ts_display}: Launching archivage_past_years()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", False)  # Trie les emails par date de réception, False signifie en ordre décroissant
    
    
    i=1
    for email in emails:
        print(f"{ts.ts_display}: {i}/{emails.Count} Done")
        if i < emails.Count:
                try:
                    year = email.ReceivedTime.year
                    IsNotRead = email.UnRead
                    Sender = email.SenderName
                    ReceivedTime = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    print(f"{ts.ts_display}: Reception Date : {ReceivedTime}, from {Sender} | {i/emails.Count*100}% Status : {IsNotRead}")
                    if year < 2023 and IsNotRead:
                        try:
                            email.MarkAsRead()
                            print( f"{ts.ts_display}: Reception Date : {ReceivedTime}, from {Sender} | {i/emails.Count*100}% Done ")
                            break
                        except:
                            pass
                except:
                    pass
        else:
            break
        i = i+1
        



def print_last_email_info(k=20):
    last_mail_info = get_last_emails(k)
    for i, email in enumerate(last_mail_info, start=1):
        print(f"Email {i}:")
        print(f"Sender: {email['Sender']}")
        print(f"Subject: {email['Subject']}")
        print()


def stats_email_by_sender():
    ts = timestamp()
    print(f'{ts.ts_display}: Launching stats_email_by_sender()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", False)  # Trie les emails par date de réception, False signifie en ordre décroissant

    # Créez un dictionnaire vide pour stocker les statistiques
    stats = {}

    # Parcourez tous les e-mails
    for email in emails:
        try:
            # Vérifiez si l'expéditeur est déjà dans le dictionnaire
            if email.SenderName in stats:
                # Si oui, augmentez le nombre de mails de 1
                print(f"{email.SenderName}  : + 1")
                stats[email.SenderName] += 1
            else:
                # Sinon, créez une nouvelle entrée dans le dictionnaire
                print(f"{email.SenderName}  : Enters the list")
                stats[email.SenderName] = 1
        except:
            pass

    # Affichez les statistiques
    for sender, count in stats.items():
        print(f"{sender}: {count}")


def stats_email_by_year_month(year = 2023):
    ts = timestamp()
    print(f'{ts.ts_display}: Launching stats_email_by_sender()')

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", False)  # Trie les emails par date de réception, False signifie en ordre décroissant

    # Créez un dictionnaire vide pour stocker les statistiques
    graph = stats_email_per_month(year)
    stats = {}
    i = 0
    # Parcourez tous les e-mails
    for email in emails:
        i = i+1
        print(f"{ts.ts_display}: Progression : {i} / {emails.Count}")
        try:
            #Vérifier si le mois / année est déjà dans le dictionnaire
            if email.ReceivedTime.strftime("%Y-%m") in stats: 
                # Si oui, augmentez le nombre de mails de 1
                stats[email.ReceivedTime.strftime("%Y-%m")] += 1
            elif email.ReceivedTime.strftime("%Y-%m") not in stats and email.ReceivedTime.year == year:
                # Sinon, créez une nouvelle entrée dans le dictionnaire
                stats[email.ReceivedTime.strftime("%Y-%m")] = 1
            else:
                pass
        except:
            pass

    graph.stats = stats

    return(graph)


#

def main():

    # send_mail_attachment()
    # read_mail(False)
    # print_last_email_info(20)
    # move_emails_to_folder()
    # print_last_email_info(20)
    archivage_past_years()
    move_emails_to_folder_archive(2021)
    # stats_email_by_sender()
    
    # stats_email_by_year_month()

if __name__ == "__main__":
    main()