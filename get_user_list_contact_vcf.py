from libraries import *
from get_data import *

def get_user_list_contact_vcf():
    with open("contact/contact.vcf", 'w', encoding='utf-8') as kontakt:
        size = len(data)
        for k in range(0,size): #to num of row
                row = data[k]
                sim_num = row["SIM"]
                position = row["Stanowisko"]
                #if position == "Brygadzista":
                if sim_num != "" and sim_num != "-" and position != "Doradca ds. OZE":
                    fname = row["Imie"]
                    lname = row["Nazwisko"]
                    mail = row["Email"]
                    kontakt.writelines("BEGIN:VCARD\n")
                    kontakt.write("VERSION:2.1\n")
                    kontakt.write("N:%s;%s;;%s |;\n" %(fname,lname,COMPANY_SHORT))
                    kontakt.write("FN:%s | %s %s\n" %(COMPANY_SHORT, lname,fname))
                    kontakt.write("EMAIL:%s\n" %(mail))
                    kontakt.write("TEL;CELL:%s\n" %(sim_num))
                    kontakt.write("ORG:%s;%s\n" %(COMPANY_SHORT,position))
                    kontakt.write("END:VCARD\n")
    kontakt.close()
    print("Contact.vcf file has been created.")