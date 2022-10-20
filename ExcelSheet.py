import PDFFiles
import pandas as pd
import win32com.client
import PDFFiles

class ExcelSheetAbruf():
    def __init__(self, path_ExcelSheet):
        self.path_ExcelSheet = path_ExcelSheet
    
        self.df_Abruf = pd.read_excel(path_ExcelSheet, sheet_name=0)
        self.df_Wareneingaenge = pd.read_excel(path_ExcelSheet, sheet_name=1)
        self.df_AbrufErledigt = pd.read_excel(path_ExcelSheet, sheet_name=2)
    
    


    def save_Excelsheet(self):
        with pd.ExcelWriter(self.path_ExcelSheet) as writer:
            self.df_Abruf.to_excel(writer, sheet_name='Abrufe', index=False)
            self.df_Wareneingaenge.to_excel(writer, sheet_name='Wareneingaenge')
            self.df_AbrufErledigt.to_excel(writer, sheet_name='Abrufe abschließen')
            

    
    def access_df_column(self, df, column_name):
        df_column = df[column_name]
        return df_column


    def check_If_Value_In_Df_Column(self, df, column_name, value):
        return (int(value) in list(df[column_name]))


    def get_Column_Index(self, df, Abrufnummer):
        return df.index[df[df.columns[0]] == Abrufnummer].tolist()[0]


    def change_value_df(self, df, column_name: str, row_index: int, value):
        df.at[row_index, column_name] = value


    def add_New_Abruf(self, pdf_Abruf: PDFFiles.PdfAbruf):

        # Spaltennamen von Tabelle für neuen Dataframe ermitteln, damit mit pd.concat die zwei Dataframes zusammen geführt werden können
        column_names = list(self.df_Abruf.columns.values)
        index = column_names.index('Lieferdatum')
        column_names_Abruf =column_names[:index+1]

        # Wenn Abruf schon in Tabelle mache nicht  
        if (pdf_Abruf.abruf_Nr in list(self.df_Abruf['Abrufe'])):
            return None

        # Erstellen neuen Dataframe mit Abruf und mit Methode pd.concat füge die Dataframes zusammen 
        df2 = pd.DataFrame([[pdf_Abruf.abruf_Nr, pdf_Abruf.contract_Nr, pdf_Abruf.versorgungs_Nr, pdf_Abruf.menge, pdf_Abruf.zufuehr_Nr, pdf_Abruf.lieferdatum]], columns=column_names_Abruf)
        self.df_Abruf = pd.concat([self.df_Abruf, df2])
        print(self.df_Abruf)


if __name__ == '__main__':
    
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')
    messages = mapi.Folders("s-steger@live.de").Folders("Python Skript").Items
    folder_Archiv = mapi.Folders("s-steger@live.de").Folders("Archiv")
    for message in messages:
        # Test ab hier
        path_Test_Sheet = 'C:\\Users\\S-Ste\\Documents\\Dokumente_Sascha\\Bildung\\Programmiern\\Python\\Automate Outlook\\test.xlsx'
        excel_Sheet_Abruf = ExcelSheetAbruf(path_Test_Sheet)
        if message.SenderEmailAddress == 's-steger@live.de':
            for attached in message.Attachments:
                pdf_Abruf = PDFFiles.PdfAbruf(attachment=attached)

                
                excel_Sheet_Abruf.add_New_Abruf(pdf_Abruf)
                excel_Sheet_Abruf.save_Excelsheet()
    
   
