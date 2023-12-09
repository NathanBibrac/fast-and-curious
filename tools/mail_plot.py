import matplotlib.pyplot as plt
import win32com.client
from datetime import datetime




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

    def plot_graph(self,Label="Email analysis"):
        ts = timestamp()
        plt.figure(figsize=(20, 10),label = f"Number of emails received per month in {self.year}")
        plt.bar(self.stats.keys(), self.stats.values())
        plt.show()
        ts_str = ts.ts_filename
        plt.savefig(f'D:\\Lab\\data\\images\\graphs\\stats_email_by_tear_month_{str(self.year)}-{ts_str}.png')

class stats_email_per_sender():
    def __init__(self,threshold):
        self.threshold = threshold
        self.stats = {}
    
    def plot_graph(self,Label="Email analysis"):
        ts = timestamp()
        plt.figure(figsize=(20, 10),label = f"Number of emails received per sender")
        plt.bar(self.stats.keys(), self.stats.values())
        plt.show()
        ts_str = ts.ts_filename
        plt.savefig(f'D:\\Lab\\data\\images\\graphs\\stats_email_by_sender_{ts_str}.png')

    def print_upper_stats(self,K):
        # Affichez les statistiques
        for sender, count in self.stats.items():
            if count > K:
                print(f"{sender}: {count}")


def stats_email_by_year_month(emails,year = 2023):
    ts = timestamp()
    print(f'{ts.ts_display}: Launching stats_email_by_sender()')

    # Créez un dictionnaire vide pour stocker les statistiques
    graph = stats_email_per_month(year)
    stats = {}
    i = 0
    # Parcourez tous les e-mails
    for email in emails:
        i = i+1
        pct = round(i/emails.Count*100,2)
        print(f"{ts.ts_display}: Progression : { pct} %")
        try:
            #Vérifier si le mois / année est déjà dans le dictionnaire
            if email.ReceivedTime.strftime("%Y-%m") in stats: 
                # Si oui, augmentez le nombre de mails de 1
                print(f"{ts.ts_display}: {email.ReceivedTime.strftime('%Y-%m')} already in stats")
                stats[email.ReceivedTime.strftime("%Y-%m")] += 1
            elif email.ReceivedTime.strftime("%Y-%m") not in stats and email.ReceivedTime.year == year:
                # Sinon, créez une nouvelle entrée dans le dictionnaire
                print(f"{ts.ts_display}: {email.ReceivedTime.strftime('%Y-%m')} not in stats")
                stats[email.ReceivedTime.strftime("%Y-%m")] = 1
            else:
                pass
        except:
            pass
    graph.stats = stats
    return graph

def graph_email_by_year_month(graph):

    graph.plot_graph()


def stats_email_by_sender(emails,min):

    ts = timestamp()
    print(f'{ts.ts_display}: Launching stats_email_by_sender()')    

    # Créez un dictionnaire vide pour stocker les statistiques
    graph = stats_email_per_sender(min)

    
    stats = {}



def main():

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", False)  # Trie les emails par date de réception, True signifie en ordre croissant
    stats = stats_email_by_year_month(emails,2023)
    graph_email_by_year_month(stats)

if __name__ == "__main__":
    main()

