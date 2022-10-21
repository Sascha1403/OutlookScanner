import win32com.client
from zipfile import ZipFile

from ExcelSheet import ExcelSheetAbruf
import PDFFiles
import MessageTyp


test_Abruf = False
test_Wareneingang = True 
test_MeldungInst = False 

email_Acc = 's-steger@live.de'
folder_Skript = 'Python Skript'
path_Test_Sheet = 'C:\\Users\\S-Ste\\Documents\\Dokumente_Sascha\\Bildung\\Programmiern\\Python\\Automate Outlook\\test.xlsx'


def main():
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')
    messages = mapi.Folders(email_Acc).Folders(folder_Skript).Items
    
    folder_Archiv = mapi.Folders(email_Acc).Folders("Archiv")
    folder_Abnahme_IO = mapi.Folders(email_Acc).Folders(folder_Skript).Folders("Abnahme Vorgeprüft")
    excel_Sheet_Abruf = ExcelSheetAbruf(path_Test_Sheet)

    for message in messages:

        # get Message Type 
        message_Typ = MessageTyp.get_Message_Typ(message)

        # Based on the Typ of the Document execute a routine

        if message_Typ == 'Abruf HIL' and test_Abruf == True:
            for attached in message.Attachments:
                pdf_Abruf = PDFFiles.PdfAbruf(attached)
                pdf_Abruf.saveToFileserver() 
                excel_Sheet_Abruf.add_New_Abruf(pdf_Abruf)
                
        elif message_Typ == 'Wareneingang' and test_Wareneingang == True:
            for attached in message.Attachments:
                pdf_Wareneingang = PDFFiles.PdfWareneingang(message, attached) 
                pdf_Wareneingang.saveToFileserver()
                excel_Sheet_Abruf.add_New_Wareneingang()

        elif message_Typ == 'Meldung Instandsetzung' and test_MeldungInst == True:
            for attached in message.Attachments:
                doc_Meld_Inst = ZIP.MeldungInst(message, attached) # To Do
                doc_Meld_Inst.save_To_Fileserver() # To Do
                
                ergebnis_Prüfung = DocPrüfer.Abruf(doc_Meld_Inst.abruf_Nr, doc_Meld_Inst.kontrakt_Nr) # To Do 
                if  ergebnis_Prüfung == True:
                    excel_Sheet_Abruf.add_Geprüfter_Abruf(doc_Meld_Inst.abruf) # To Do
                    message.move()
                
                else:
                    SendMail.Zurückweisung(ergebnis_Prüfung) # To Do

        else:
            print(f'{message.Subject} hat einen Unbekannt Typ')

    excel_Sheet_Abruf.save_Excelsheet()

                



                

    
    excel_Sheet_Abruf.save_Excelsheet()

if __name__ == "__main__":
    main()