import win32com.client
from PDF.Files.PDF import Abruf



class ProcedureAbrufMail():
    def __init__(self, message: win32com.client.CDispatch) -> None:
        self.mail = message
        self.attachments = message.Attachments

    def procedure(self):
        list_PDF_Abrufe = []
        for abruf in self.attachments:
            pdf_Abruf = Abruf(abruf)

            



if __name__ == "__main__":
    pass

    



         
