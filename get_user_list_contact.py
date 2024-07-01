from libraries import *
from get_data import *

def get_user_list_contact():
    size = len(data) + 1
    for k in range(0,size):
            row = data[k]
            sim_num = row["SIM"]
            if sim_num != "" and sim_num != "-" :
                imie = row["Imie"]
                nazwisko = row["Nazwisko"]
                mail = row["Email_Arago"]
                print(nazwisko + "," + imie + "," + str(sim_num) + "," + mail) # label of telephone number's and e-mail's