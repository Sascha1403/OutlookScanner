import dateutil.parser as dparser
import re 
import PdfReader


class TextAbrufAnalyse():
    def __init__(self, text_Abruf) -> None:
        self.text_Abruf = text_Abruf
    
    def get_Abruf_Nr(self)->int:
        regexp = re.compile('[4][6][0][1][0-9][0-9][0-9][0-9][0-9][0-9]')
        for line in self.text_Abruf:
            if regexp.search(line):
                vers_Nr = re.findall('[4][6][0][1][0-9][0-9][0-9][0-9][0-9][0-9]',line)
                return int(vers_Nr[0])
        raise Exception('Versorgungsnummer in PDF nicht gefunden')

        
        
    def get_ZuführNr(self)->str:
        length_zufuehr_Number = 14
        for line in self.text_Abruf:
            if line.find('ZuführNr.') != -1 and line.find('Z2975'):
                index = line.index('Z2975')
                zufuehrNr = line[index:index+length_zufuehr_Number]
                return zufuehrNr
        raise Exception('Zuführnummer in PDF nicht gefunden')
        

    def get_Lieferdatum(self)->str:
        for line in self.text_Abruf:
            if line.find('Liefertermin') != -1:
                test = dparser.parse(line, fuzzy=True)
                date = f"{test.day}.{test.month}.{test.year}"
                return date
        raise Exception('Lieferdatum in PDF nicht gefunden')

    def get_Vers_Nr(self)->int:
        regexp = re.compile('[0-9][0-9][0-9][0-9][-][0-9][0-9][-][0-9][0-9][0-9][-][0-9][0-9][0-9][0-9]')
        for line in self.text_Abruf:
            if regexp.search(line):
                vers_Nr = re.findall('[0-9][0-9][0-9][0-9][-][0-9][0-9][-][0-9][0-9][0-9][-][0-9][0-9][0-9][0-9]',line)
                vers_Nr = vers_Nr[0].replace('-','')
                return int(vers_Nr)
        raise Exception('Versorgungsnummer in PDF nicht gefunden')


    def get_Menge(self)->int:
        regexp = re.compile('[0-9][0-9][0-9[,][0-9][0-9]')
        for i, line in enumerate(self.text_Abruf):
            if line.find('Instandsetzung') != -1 and line.find('Baugruppe') != -1:
                nextline = self.text_Abruf[i+1]
                if regexp.search(nextline):
                    string = re.findall('[0-9]+\s+[L][E]',nextline) #\s igonres whitespaces     
                    menge = re.findall('[0-9]+', string[0])
                    return int(menge[0])
        raise Exception('Menge in PDF nicht gefunden')

    def get_Contract_Nr(self)->int:
        for line in self.text_Abruf:
            if line.find('Kontrakt') != -1 and line.find('Position') != -1:
                contract_Nr = re.findall('[3][0][0][0-9][0-9][0-9][0-9]',line)
                return int(contract_Nr[0])
        raise Exception('Contract Nr in PDF nicht gefunden')





if __name__ == "__main__":
    pdf_Text_Abruf = PdfReader.readPDF('TestDocuments\\4601476457_891.pdf')

    TextAbrufAnalyse = TextAbrufAnalyse(pdf_Text_Abruf)
    abruf_Nr = (TextAbrufAnalyse.get_Abruf_Nr())
    Kontrakt_Nr = (TextAbrufAnalyse.get_Contract_Nr())
    vers_Nr = (TextAbrufAnalyse.get_Vers_Nr())
    menge = (TextAbrufAnalyse.get_Menge())
    zufuehr_Nr = TextAbrufAnalyse.get_ZuführNr()
    lieferdatum = TextAbrufAnalyse.get_Lieferdatum()
    print()
