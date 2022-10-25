from abc import ABC, abstractmethod
import win32com.client
import os
import shutil

from PDF.Textanalyse import PdfReader
from Excel import DictonaryContract
from Message import SubjectAnalyse


class PdfFile(ABC):
    def __init__(self, attachment: win32com.client.CDispatch) -> None:
        super().__init__()
        self.attachment = attachment
        self.text = self.get_Attachment_text()

 

    def saveToFileserver(self):
        """ Diese Methode ist der grundsätzliche Ablauf zur Bearbeitung der Mail in Abhängigkeit vom Typ (also der Subklasse)"""
        if os.path.exists(self.path):
            self.attachment.SaveASFile(os.path.join(self.path, self.attachment.FileName))

        else: 
            ## Wenn der Ordner nicht existiert wird ein neuer Muster Abruf Ordner erstellt 
            src = os.path.dirname(self.path) + "\\00_Musterabruf"
            shutil.copytree(src, self.path) 
            self.attachment.SaveAsFile(os.path.join(self.path, self.attachment.FileName))

    def get_Attachment_text(self)-> str:
        """ Diese Methode return den Text vom Attachment"""
        
        #Document temporare speicher
        pdf_Chache_Path = os.path.join("/TestDocuments", self.attachment.FileName)
        self.attachment.SaveAsFile(pdf_Chache_Path)
        
        # Text extrahieren 
        pdf_Text = PdfReader.readPDF(pdf_Chache_Path)
        
        # Dokument wieder löschen
        os.remove(pdf_Chache_Path)
        
        return pdf_Text
    
    @abstractmethod 
    def get_PDF_Path(self):
        """ Return the Attachment Path where it will be saved on the Fileserver

        Args:
        -----

        Return:
            Path of the Attachment where it will be saved
        """
        pass



class Abruf(PdfFile):
    def __init__(self, attachment: win32com.client.CDispatch) -> None:
        super().__init__(attachment)
        Text_Analyse = TextAbrufAnalyse(self.text)
        self.abruf_Nr = Text_Analyse.get_Abruf_Nr()
        self.contract_Nr = Text_Analyse.get_Contract_Nr()
        self.versorgungs_Nr =Text_Analyse.get_Vers_Nr()
        self.menge = Text_Analyse.get_Menge()
        self.zufuehr_Nr = Text_Analyse.get_ZuführNr()
        self.lieferdatum = Text_Analyse.get_Lieferdatum()
        self.path = self.get_PDF_Path()

    def __str__(self):
        str = (f"Abruf: {self.abruf_Nr}\nKontraktnummer: {self.contract_Nr}\nVersorgungsnummer: {self.versorgungs_Nr}\nmenge: {self.menge}\nZuführnummer: {self.zufuehr_Nr}\nLieferdatum: {self.lieferdatum}\nAbruf Path:{self.path}\n")
        return str

        

    def get_PDF_Path(self): ## TO DO ##
        """ Return the Attachment Path where it will be saved on the Fileserver

        Args:
        -----

        Return:
            Path of the Attachment where it will be saved
        """
        contract_path = DictonaryContract.get_Contract_Path(self.contract_Nr)
        path_abruf = f'C:\\Users\\S-Ste\\Desktop\\GLS\\{contract_path}\\Abrufe\\{self.abruf_Nr}'

        return path_abruf

class Wareneingang(PdfFile):
    def __init__(self, message: win32com.client.CDispatch, attached:win32com.client.CDispatch) -> None:
        self.path = SubjectAnalyse.get_Path(attached.Filename)

        



if __name__ == "__main__":
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')
    messages = mapi.Folders("s-steger@live.de").Folders("Python Skript").Items
    folder_Archiv = mapi.Folders("s-steger@live.de").Folders("Archiv")
    for message in messages:
        if message.SenderEmailAddress == 's-steger@live.de':
            for attached in message.Attachments:
                Pdf_Abruf = Abruf(attachment=attached)
                print(Pdf_Abruf)
                






