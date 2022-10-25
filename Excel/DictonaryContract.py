my_dict={3003008 : '3003008_BG1286_Periskop_UAN Tescon_ZtQ4.1', 3002929: '3002929_BG1286_Steuergerät_UAN Firecraft_ZtQ 2.4'}

def get_Contract_Path(contract_Number):
    
    if my_dict.get(contract_Number) != None:
        return my_dict.get(contract_Number) 
    
    else:
        return str(contract_Number)


if __name__ == "__main__":
    print(get_Contract_Path(contract_Number=3003008))
    print(get_Contract_Path(contract_Number=3003435))
    

