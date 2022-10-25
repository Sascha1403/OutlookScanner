import win32com.client

def get_Message_Typ(message: win32com.client.CDispatch):
    # Prüfe Sender der Mail
    # Wenn Mail von GLS kommt oder von s-steger@live.de weiter Verarbeitung

    # !!! Prüfen @gls.de !!!
    if '@gls.de' in message.SenderEmailAddress or 's-steger@live.de' == message.SenderEmailAddress:
        
        # Überprüfe Titel der Mail 
        
        if 'Wareneingang' in message.Subject:
            return 'Wareneingang'
        
        elif 'Meldung Instandsetzung' in message.Subject:
            return 'Meldung Instandsetzung'
    

    # !!! Prüfen @hil.de !!!
    if '@hil' in message.SenderEmailAddress or 's-steger@live.de' == message.SenderEmailAddress:
        substring_list = ['Abruf', 'Beauftragung']
        if any(substring in message.Subject for substring in substring_list):
            return 'Abruf HIL'

    else:
        return 'Unknown'

        
