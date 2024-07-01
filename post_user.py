from libraries import *
from get_data import *



def post_user(city_text, workplace_text, fname_text, lname_text, position_text, department_text):
    num_of_row = len(data) +2
    fname = fname_text
    lname = lname_text
    position = position_text
    department = department_text
    fname_clear = ""
    lname_clear = ""
        
    for letter_in_fname in fname :
        if letter_in_fname == "Ł":
            fname_clear = fname_clear + "L"
        elif letter_in_fname == "ł":
            fname_clear = fname_clear + "l"
        else:
            fname_clear = fname_clear + letter_in_fname
    fname_clear = unicodedata.normalize('NFKD', fname_clear).encode('ascii', 'ignore')
    fname_clear = fname_clear.decode('UTF-8')
    
    for letter_in_lname in lname :
        if letter_in_lname == "Ł":
            lname_clear = lname_clear + "L"
        elif letter_in_lname == "ł":
            lname_clear = lname_clear + "l"
        else:
            lname_clear = lname_clear + letter_in_lname
    lname_clear = unicodedata.normalize('NFKD', lname_clear).encode('ascii', 'ignore')
    lname_clear = lname_clear.decode('UTF-8')
    mail = fname_clear[0].lower() + "." + lname_clear.lower() + "@" + DOMAIN
    mail_microsoft = fname_clear[0].lower() + "." + lname_clear.lower() + f"@{COMPANY_SHORT.lower()}.onmicrosoft.com"
    small = 'abcdefghjkmnopqrstuvwxyz'
    big = 'ABCDEFGHJKLMNOPQRSTUVWXYZ'
    password = big[random.randint(0, 24)] + big[random.randint(0, 24)] + small[random.randint(0, 23)] + small[random.randint(0, 23)] + str(random.randint(1000, 9999))

    insert_row = [city_text, workplace_text, fname, lname, position, department, mail, password, mail_microsoft]
    data_sheet_name_main.insert_row(insert_row,num_of_row)
    print("User succesful created ;) ")



