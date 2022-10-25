def get_Contract_Number(filename):
    length_contract_Number = 7
    index = filename.index('300')
    contract_Number = filename[index:index+length_contract_Number]
    return contract_Number

def get_Abruf_Number(filename):
    length_Abruf_Number = 11
    index = filename.index('4601')
    abruf_Number = filename[index:index+length_Abruf_Number]
    return abruf_Number

def get_Contract_Path(contract_Number):
    my_dict={'3003008' : '3003008_BG1286_Periskop_UAN Tescon_ZtQ4.1' ,
             'Ava': '002' , 'Joe': '003'}
    return my_dict.get(contract_Number)

def get_Path(filename):
    abruf_Number = get_Abruf_Number(filename)
    contract_Number = get_Contract_Number(filename)
    contract_Path= get_Contract_Path(contract_Number)
    path = f"C:\\Users\S-Ste\Desktop\\GLS\{contract_Path}\{abruf_Number}"
    return path

print(get_Path('Wareneingang 3003008 4601444555'))

