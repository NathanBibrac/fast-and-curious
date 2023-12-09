import matplotlib.pyplot as plt
import win32com.client
import pandas as pd
import numpy as np
from datetime import datetime
import os

class timestamp():
    def __init__(self):
        self.now = datetime.now()
        self.ts_filename = self.now.strftime("%Y%m%d_%H%M%S")
        self.ts_display = self.now.strftime("%Y-%m-%d %H:%M:%S")
    
    def update(self):
        self.now = datetime.now()
        self.ts_filename = self.now.strftime("%Y%m%d_%H%M%S")
        self.ts_display = self.now.strftime("%Y-%m-%d %H:%M:%S")




def inbox_to_df():

    ts = timestamp()

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 correspond au dossier Inbox

    emails = inbox.Items
    emails.Sort("[ReceivedTime]", True) 

    email_list = []
    i = 0
    count = 0
    pct = 0

    for email in emails:
        ts.update()
        pct = round(i/emails.Count*100,2)
        print(f"{ts.ts_display}: Progression : {pct} %. {count} emails collected.")
        i = i+1
        try:
            if email.ReceivedTime.year == 2023:
                email_list.append([i,email.Subject, email.SenderName, email.ReceivedTime.strftime("%m-%y"),email.ReceivedTime.strftime("%Y%m%d_%H%M%S"),email.UnRead])
                count = count + 1
        except:
            pass

    df = pd.DataFrame(email_list, columns=['id','subject','sender','YearMonth','reiceived_time','IsRead'])
    
    '''
    try:
        df.to_excel(file_name)
    except:
        print("Excel export failed.")

    '''

    return df

def df_to_excel(df, file_name):
    ts = timestamp()
    print(f"{ts.ts_display}: Exporting to Excel")    
    try:
        df.to_excel(pd.ExcelWriter(os.path.join(r"D:\Lab\data\excel\outlook",file_name)))
    except:
        print(f"{ts.ts_display}: Excel export failed.")

def analyse():
   
    ts = timestamp()
    print(f"{ts.ts_display}: Analysis Started")

    # file_name = f"email_list_{ts.ts_filename}.xlsx"    

    df = inbox_to_df()

    df = df.sort_values(by=['reiceived_time'], ascending=True)

    print(df.head(10))

    df_count_by_sender = df[["id","sender"]].groupby(['sender']).count()
    df_count_by_sender = df_count_by_sender.sort_values(by=['id'], ascending=False)
    df_count_by_sender = df_count_by_sender.rename(columns={"id": "count"})
    df_count_by_sender = df_count_by_sender.iloc[:5]

    print(df_count_by_sender.head(10))

    df_count_by_read = df[["id","IsRead"]].groupby(['IsRead']).count()
    df_count_by_read = df_count_by_read.rename(columns={"id": "count"})
    print(df_count_by_read.head())

    df_count_by_month = df[["id","YearMonth"]].groupby(['YearMonth']).count()
    df_count_by_month = df_count_by_month.rename(columns={"id": "count"})
    
    print(df_count_by_month.head())

    

    fig, axes = plt.subplots(2, 2,figsize=(18, 10))
    # axes = axes.flatten() 
    
    # plt.figure(figsize=(20, 10),label = "Number of emails received per month in 2023")

    p = axes[0,0].bar(df_count_by_month.index.to_list(), df_count_by_month["count"].to_list())
    axes[0,0].set_title = "Number of emails received per month in 2023"
    axes[0,0].set_ylim(0,800)
    axes[0,0].set_ylabel = "Number of emails"
    axes[0,0].set_xlabel = "Month (2023)"
    axes[0,0].bar_label(p, padding=3)

    q = axes[0,1].bar(df_count_by_sender.index.to_list(), df_count_by_sender["count"].to_list())
    axes[0,1].set_title = "Number of emails received per sender"
    axes[0,1].set_ylim(0,300)
    axes[0,1].set_ylabel = "Number of emails"
    axes[0,1].set_xlabel = "Sender"
    axes[0,1].bar_label(q, padding=3)

    r = axes[1,0].bar(df_count_by_read.index.to_list(), df_count_by_read["count"].to_list())
    axes[1,0].set_title = "Number of emails received per read status"
    axes[1,0].set_ylim(0,5000)
    axes[1,0].set_ylabel = "Number of emails"
    axes[1,0].set_xlabel = "Read status"
    axes[1,0].bar_label(r, padding=3)
    
    fig.tight_layout()
    fig.show()



    fig.tight_layout()
    fig.show()
    plt.pause(60)
    fig.savefig(f'D:\\Lab\\data\\images\\graphs\\stats_email_by_sender_{ts.ts_filename}.png')
 

def analyse_2():
   
    ts = timestamp()
    print(f"{ts.ts_display}: Stats email per Sender")

    file_name = f"email_list_{ts.ts_filename}.xls"    

    df = inbox_to_df()
    df_to_excel(df[["id","sender","IsRead","YearMonth","reiceived_time"]], file_name)

    df = df.sort_values(by=['reiceived_time'], ascending=True)

    print(df.head(10))

    df_count_by_sender = df[["id","sender","IsRead"]].groupby(['sender','IsRead']).count()
    df_count_by_sender = df_count_by_sender.sort_values(by=['id'], ascending=False)
    df_count_by_sender = df_count_by_sender.rename(columns={"id": "count"})
    df_count_by_sender = df_count_by_sender.reset_index()
    df_count_by_sender = df_count_by_sender.iloc[:5]

    # import pdb; pdb.set_trace()

    df_read_sender = df_count_by_sender.loc[df_count_by_sender['IsRead']==True]
    df_unread_sender = df_count_by_sender.loc[df_count_by_sender['IsRead']==False]

    print(df_count_by_sender.head(10))

    df_count_by_month = df[["id","YearMonth","IsRead"]].groupby(['YearMonth','IsRead']).count()
    df_count_by_month = df_count_by_month.rename(columns={"id": "count"})
    df_count_by_month = df_count_by_month.reset_index()

    df_read_month = df_count_by_month.loc[df_count_by_month['IsRead']==True]
    df_unread_month = df_count_by_month.loc[df_count_by_month['IsRead']==False]
    
    print(df_count_by_month.head())
    print(df_read_month.head())
    print(df_unread_month.head())

    fig, axes = plt.subplots(1, 2,figsize=(18, 10))
    # axes = axes.flatten() 

    width = 0.6
    bot = np.zeros(len(df_read_month["YearMonth"].to_list()))

    # import pdb; pdb.set_trace()
    try:
        p = axes[0].bar(df_read_month["YearMonth"].to_list(), df_read_month["count"].to_list(),width,label = "Read",bottom= bot)
        bot = bot + df_read_month["count"].to_list()
        p = axes[0].bar(df_unread_month["YearMonth"].to_list(), df_unread_month["count"].to_list(),width,label = "UnRead",bottom= bot)
    except:
        p = axes[0].bar(df_count_by_month["YearMonth"].to_list(), df_count_by_month["count"].to_list())

    axes[0].set_title = "Number of emails received per month in 2023"
    axes[0].set_ylim(0,800)
    axes[0].set_ylabel = "Number of emails"
    axes[0].set_xlabel = "Month (2023)"
    axes[0].bar_label(p, padding=3)

    # import pdb; pdb.set_trace()

    bot_2 = np.zeros(len(df_read_sender["sender"].to_list()))
    try :
        q = axes[1].bar(df_read_sender["sender"].to_list(), df_read_sender["count"].to_list(),width,label = "Read",bottom=bot_2)
        bot_2 = bot_2 + df_read_sender["count"].to_list()
        # import pdb; pdb.set_trace()
        q = axes[1].bar(df_unread_sender["sender"].to_list(), df_unread_sender["count"].to_list(),width,label = "UnRead",bottom=bot_2)
    except:
        q = axes[1].bar(df_count_by_sender["sender"].to_list(), df_count_by_sender["count"].to_list())

    axes[1].set_title = "Number of emails received per sender"
    axes[1].set_ylim(0,300)
    axes[1].set_ylabel = "Number of emails"
    axes[1].set_xlabel = "Sender"
    axes[1].bar_label(q, padding=3)

    fig.set_label = "Email Analysis"
    fig.legend()
    fig.tight_layout()
    fig.show()
    plt.pause(60)
    fig.savefig(f'D:\\Lab\\data\\images\\graphs\\stats_email_{ts.ts_filename}.png')

def main():
    analyse_2()

if __name__ == "__main__":
    main()
