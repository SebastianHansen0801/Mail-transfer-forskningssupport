from win32com.client import Dispatch
import os
import re
from  timeit import default_timer as timer

# directory to work in
PARENT_DIR = os.path.join('//byg-cserver1', 'Grupper', 'Adm_support', 'Forskningssupport', '0_outlook_arkiv_overført_primo_2022')
SECTION = '1100'

# opening application
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# opening inbox
inbox = outlook.Folders['BYG-Forskningssupport'].Folders['Indbakke'].Folders[7]

# create path to section folder
path = os.path.join(PARENT_DIR, SECTION)

# function to save messages
def save_messages(messages, s):
    count = 1
    for message in messages:
        global mail_count
        mail_count += 1
        name = str(message.Subject)
        # removing all special charaters and spaces from the name
        name = re.sub('[^A-Za-z0-9æøå]', '_', name)
        name = re.sub('____', '_', name)
        name = re.sub('___', '_', name)
        name = re.sub('__', '_', name)
        # adding file extension and mail number
        name = str(count) + '_' + name
        name.encode("utf8")
        full_path = os.path.join(path, s, name)
        full_path = full_path[:238]
        if full_path[-1] == '_':
            full_path = full_path[:-1]
        full_path += '.msg'
        print(full_path)
        
        # try to save message, if error, add the message to error log
        try:
            message.SaveAs(full_path)
        
        except Exception as e:
            log.write((full_path + '\n').encode("utf8"))
            log.write((str(e) + '\n\n').encode("utf8"))
            global error_count
            error_count += 1
        
        count += 1

if __name__ == "__main__":
    start = timer()
    # open the error log
    log = open('log.txt', 'wb')
    error_count = 0
    mail_count = 0
    
    #get the messages
    messages = inbox.Items
    
    # function call to save messages
    save_messages(messages, '')
    
    # looping through folders in the inbox
    for folder in inbox.Folders:
        #create a folder on the drive (byg-cserver1) arccordig to folders from mailbox
        folder_name = str(folder.Name)
        folder_name = re.sub('[^A-Za-z0-9æøå]', '_', folder_name)
        folder_name = re.sub('_____|____|___|__', '_', folder_name)
        os.makedirs(os.path.join(path, folder_name))

        #get the messages
        messages = folder.Items
        # function call to save messages
        save_messages(messages, folder_name)
        
        # looping through subfolders of folders in inbox
        for subfolder in folder.Folders:
            #create a folder on the drive (byg-cserver1) arccordig to folders from mailbox
            subfolder_name = str(subfolder.Name)
            subfolder_name = re.sub('[^A-Za-z0-9æøå]', '_', subfolder_name)
            subfolder_name = re.sub('_____|____|___|__', '_', subfolder_name)
            os.mkdir(os.path.join(path, folder_name, subfolder_name))
            
            folder_path = os.path.join(folder_name, subfolder_name)

            #get the messages
            messages_sub = subfolder.Items
            # function call to save messages
            save_messages(messages_sub, folder_path)
            
            for subsubfolder in subfolder.Folders:
                #create a folder on the drive (byg-cserver1) arccordig to folders from mailbox
                subsubfolder_name = str(subsubfolder.Name)
                subsubfolder_name = re.sub('[^A-Za-z0-9æøå]', '_', subsubfolder_name)
                subsubfolder_name = re.sub('_____|____|___|__', '_', subsubfolder_name)
                os.mkdir(os.path.join(path, folder_name, subfolder_name, subsubfolder_name))
                
                folder_path = os.path.join(folder_name, subfolder_name, subsubfolder_name)

                #get the messages
                messages_subsub = subsubfolder.Items
                # function call to save messages
                save_messages(messages_subsub, folder_path)
    
    log.close()

    elapsed_time = timer() - start
    print("The program took ", elapsed_time/60, 'min')
    print("Mails: ", mail_count)
    print("Errors: ", error_count)