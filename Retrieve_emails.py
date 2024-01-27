# -*- coding: utf-8 -*-

import os
import datetime as dt
import csv
import win32com.client as win32 
import io
import datetime

domainmapping={'polimi.it':'Politecnico di Milano'}


#If outlook cannot be dispatched it means the script is not running in a local folder
try:
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
except:   
    os.chdir(r'C:') #it's correct, DO NOT add back/forward slash
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')



mailfolders = []

#here, to add something to check if the folder exists
def scan_folders(folders, excluding_folders):
    ''' Function that loops through all folders, subfolders, subfolders of subfolders etc.
        
    Args: 
        folders (MAPI object): Outlook folder 
        
    Returns:
        A list with objects (folders and all the subfolders)
    '''

    for folder in folders.Folders:
        if not folder.Name in excluding_folders:
            print('Folder %s was added in the parsing list' %folder.Name)
            print(mailfolders.__len__())
#            count=count+1
#            if count<400:
            mailfolders.append(folder)
#            else:
#                continue
            
            scan_folders(folders=folders.Folders[folder.Name], excluding_folders=excluding_folders)
    if not mailfolders: #In case the selected folder has no subfolders, return the folder so, at least, the loop can run in that one.
        mailfolders.append(folders)    
    #mailfolders = [mf for mf in mailfolders_r if mf.Name not in excluding_folders]
    return None


def initiate_csv(filename):

    if not os.path.isfile(filename):
        with io.open(filename, 'w', newline='', encoding='utf-8') as csvfile: #Always use with open in order to ensure that the files is closed after it is appended
            writer = csv.writer(csvfile, delimiter= ',')
            writer.writerow(['Received Time', 'ID','Subject', 'Sender', 'Recipients',
                             'Body', 'CC', 'Full path', 'conversationID', 'is_final_reply', 
                             'is_direct_reply', 'is_media', 'is_crm', 'months']) 
            print('Columns were added')
    else:
        pass
    
    return None
        

def parse_emails(filename, parent_folder):
    #if not os.path.isfile(filename):
        
    with io.open(filename, 'a', newline='', encoding='utf-8') as csvfile: #Always use with open in order to ensure that the files is closed after it is appended
        writer = csv.writer(csvfile, delimiter= ',')     
        for folder in mailfolders:
            print('Parsing folder %s' %folder.Name)
            for mail in folder.Items:
              try: 
                if mail.Class == 43:
                    try:
                        mail_address=mail.Sender.GetExchangeUser().PrimarySmtpAddress
                        if mail_address == '':
                            try:
                                mail_address = mail.Sender.Address
                                mail_address = mail_address.split('/')[-1]
                                mail_address = mail_address.split('=')[-1]
                                if mail_address[:4].upper() == 'ESCM':
                                    mail_address[5:7]
                                else:
                                    mail_address='none@ECB.INT'
                            except AttributeError:
                                mail_address = 'tipota re'
                    except AttributeError:
                        try:
                            mail_address = mail.Sender.GetExchangeDistributionList().PrimarySmtpAddress
                        except AttributeError:
                            try:
                                mail_address = mail.Sender.Address
                            except AttributeError:
                                try:
                                    mail_address = mail.Sender.Address
                                    mail_address = mail_address.split('/')[-1]
                                    mail_address = mail_address.split('=')[-1]
                                    if mail_address[:4].upper() == 'ESCM':
                                        mail_address[5:7]
                                    else:
                                        mail_address='none@ECB.INT'
                                except AttributeError:
                                    mail_address = 'tipota re' 
                    
                    mail_recipients= mail.Recipients
                    cc_contacts = mail.cc

                    recipients_list = list()
                    for r in mail_recipients:
                        recipients_list.append(r)
                    receivers = list()
                    for recipient in recipients_list:
                        try:
                            receivers.append(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
                        except: receivers.append(mail.to)
                    mail_recipients = '' + '; '.join(receivers)
                    mail_recipients = mail_recipients.replace(mail_address, '').replace(" ;", '')
                        
                    mail_Body=mail.Body.replace('"', '')

                    subject=mail.Subject.replace(",","").replace("'","")\
                                .replace('"','').replace('[EXT] ', '')\
                                .replace('RE:', '').replace('FW:', '')\
                                .replace('Input needed: ', '')\
                                .replace('Re: ', '')\
                                .replace('R:', '')\
                                .replace('[EXT]: ', '')
                                
                    ##### please cheange the following when new CRM goes live
                    if '(#' in subject:
                        is_crm = True
                    else: is_crm = False

                    if "(#" in subject:
                        ID_query = subject.split("(#")[1].split("- ")[1].split(")")[0]
                        subject = subject.split('(#')[0]
                    elif "REQ-" in subject:
                        ID_query = "REQ-" + subject.split("REQ-")[1][:8]
                        subject = subject.split('REQ-')[0]
                    else: ID_query = "NA"

                    if 'statistics@ecb' in mail_address and "@" in mail.To:
                        reply_user = True
                    else: reply_user = False

                    if 'Direct Replies' in folder.FolderPath:
                        direct_reply = True
                    else: direct_reply = False

                    if 'Media' in folder.FolderPath:
                        is_media = True
                    else: is_media = False

                    months = str(mail.SentOn.strftime('%Y/%m/%d %H:%M:%S')[5:7])

                    path_folder = folder.FolderPath.replace("%2F",'|')
                    
                    if "Undeliverable:" not in mail.Subject and "Message Recall" not in mail.Subject : #undeliverable messages 
                        # General documentation https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem
                        writer.writerow([mail.SentOn.strftime('%Y/%m/%d %H:%M:%S'), 
                                        ID_query,
                                        subject,
                                        mail_address,              # https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.sender,
                                        mail_recipients,            # https://docs.microsoft.com/en-us/office/vba/api/outlook.addressentry
                                        mail_Body,                 # https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.body
                                        cc_contacts,                   # https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.cc
                                        path_folder,         # https://docs.microsoft.com/en-us/office/vba/api/outlook.folder.folderpath
                                        mail.ConversationID,
                                        reply_user,
                                        direct_reply,
                                        is_media,
                                        is_crm,
                                        months
                                        ]) 
              except: pass
        
    return None


def main(folders, filename, excluding_folders):
    parent_folder = folders.FolderPath.split('\\')[2] #This takes the parent folder. In our case, the name of the parent folder is the 3d item of the returned list
    scan_folders(folders=folders, 
                 excluding_folders=excluding_folders)
    parse_emails(filename=filename, parent_folder=parent_folder)
    
    return None

def retrieve_emails(path, year):   
    start_time = dt.datetime.now()
    initiate_csv(filename = path + "\\data\\outlook_emails_" + year + ".csv")
    mailfolders=[]
    main(folders = outlook.Folders['statistics@ecb.europa.eu'].Folders['Inbox'].Folders['* Archive ' + year], filename = path + "\\data\\outlook_emails_" + year + ".csv",
         excluding_folders=[])

    end_time = dt.datetime.now()
    print('DONE. It took %s mins to run' %((end_time-start_time)/60))

# if __name__ == '__main__':   
#     start_time = dt.datetime.now()
#     initiate_csv(filename = 'All emails.csv')
#     mailfolders=[]
#     main(folders = outlook.Folders['statistics@ecb.europa.eu'].Folders['Inbox'].Folders['* Archive 2022'], filename = 'All emails.csv',
#          excluding_folders=[])

#     end_time = dt.datetime.now()
#     print('DONE. It took %s mins to run' %((end_time-start_time)/60))

