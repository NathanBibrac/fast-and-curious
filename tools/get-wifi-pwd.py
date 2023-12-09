import subprocess
import re
import pandas as pd
from datetime import datetime
import os
import argparse

parser = argparse.ArgumentParser(description='Wifi password recovery.')
parser.add_argument('-e', '--export', help='Export dataset to excel', action='store_true')
parser.add_argument('-p', '--path', help='Chose an excel path for  export.', action='store_true')
# parser.add_argument('-i', '--info', help='Reminder of the tools fonctionning.', action='store_true')
args = parser.parse_args()

def get_wifi_pwd_list():

    command_output = subprocess.run(["netsh","wlan","show","profiles"], capture_output = True).stdout.decode('iso-8859-1')

    profile_names = (re.findall(("Profil Tous les utilisateurs    ÿ: (.*)\r"),command_output))
    wifi_list = list()

    if len(profile_names) != 0:
        
        for name in profile_names:

            wifi_profil = dict()
            command_output_temp = subprocess.run(["netsh","wlan","show","profiles",name,"key=clear"], capture_output = True).stdout.decode('iso-8859-1')
            profile_key = (re.findall(("Contenu de la cl\x82            : (.*)\r"),command_output_temp))

            if re.search("Cl\x82 de s\x82curit\x82ÿÿÿÿÿÿÿÿ: Absente",command_output_temp) or len(profile_key) == 0:

                continue

            else:

                wifi_profil["1-ssid"]=name
                wifi_profil["2-password"]=profile_key[0]

            wifi_list.append(wifi_profil)

    df_wifi_list = pd.DataFrame(wifi_list)
    

    return(df_wifi_list)

def export_to_excel(df,path=None):

    ts = datetime.now()
    ts_str = ts.strftime("%Y%m%d%H%M%S")
    file_name = "wifi_pwd_list"+ts_str+".xlsx"
    if path==None:
        root_dir = os.getcwd()
        root_directory = os.path.join(root_dir,"output_files")
        file_name_dir = os.path.join(root_directory,file_name)
    else:
        file_name_dir = os.path.join(path,file_name)

    with pd.ExcelWriter(file_name_dir) as writer:

        df.to_excel(writer,sheet_name="Results")

    # df.to_excel(file_name_dir,engine='xlsxwriter')

def main():
    
    df_wifi_list = get_wifi_pwd_list()
    print(df_wifi_list[5:])
    if args.export and args.path:
        export_to_excel(df_wifi_list,args.path)
    elif args.export:
        export_to_excel(df_wifi_list)  
    print("Done")

if __name__ == '__main__' :

    main()
    
