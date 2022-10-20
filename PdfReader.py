# importing required modules 
import PyPDF2 
import os   


def readPDF(pdf_Path):
    # creating a pdf file object 
    pdfFileObj = open(pdf_Path, 'rb')

    # creating a pdf reader object 
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
 
    pdf_Text ="" 
    pages = pdfReader.getNumPages()
    for page_number in range(pages):   # use xrange in Py2
        page = pdfReader.getPage(page_number)
        page_content = page.extractText()
        pdf_Text = pdf_Text + page_content
    
    pdf_Text = pdf_Text.splitlines()
    pdfFileObj.close()
    

    return pdf_Text



if __name__ == "__main__":
    pdf_Text = readPDF('4601476457_891.pdf')






