from turtle import fd
import win32com.client
import os
import PyPDF2
import PdfReader


from TextAnalyse import TextAbrufAnalyse

path = 'C:\\Users\\S-Ste\\Desktop'
pdf_Cache_Path = 'C:\\Users\\S-Ste\\Documents\\Dokumente_Sascha\\Bildung\\Programmiern\\Python\\Automate Outlook\\ChacheForDocument'

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace('MAPI')
'''
for account in mapi.Accounts: 
    print(account.DeliveryStore.DisplayName)

for idx, folder in enumerate(mapi.Folders):
    #index starts from 1
    print(idx+1, folder)

# or using index to access the folder
for idx, folder in enumerate(mapi.Folders(1).Folders): 
    print(idx+1, folder)
'''

# Get Message from Folder
messages = mapi.Folders("s-steger@live.de").Folders("Einsortieren").Items
folder_Archiv = mapi.Folders("s-steger@live.de").Folders("Archiv")

for message in messages:
        if message.Class == 43:
            if message.SenderEmailAddress == 's-steger@live.de':
                for attached in message.Attachments:
                    
                    attached.SaveASFile(os.path.join(pdf_Cache_Path, attached.Filename))
        
                    pdfFileObj = open(os.path.join(pdf_Cache_Path, attached.FileName), 'rb')
                    

                    
                    # creating a pdf reader object 
                    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 

                    pdf_Text ="" 
                    pages = pdfReader.getNumPages()
                    for page_number in range(pages):   # use xrange in Py2
                        page = pdfReader.getPage(page_number)
                        page_content = page.extractText()
                        pdf_Text = pdf_Text + page_content
                    
                    pdf_Text.splitlines()
                    print(pdf_Text)
                    
                    pdfFileObj.close() # Beende Zugriff auf Dokument 
                    os.remove(os.path.join(pdf_Cache_Path, attached.FileName))