from libraries import *
from get_data import *

def get_user_list_data(self,search_box): #wyszukiwanie ludzi po kolumnach
    for row in data:
        
        if search_box in row["Email"]:
            print(row["Email"]  + " " + search_box)
            fname = row["Imie"]
            lname = row["Nazwisko"]
            position = row["Stanowisko"]
            department = row["Dzial"]
            mail = row["Email"]
            password = row["Passwd"]
            sim_num = row["SIM"]
            pc_name = row["PC_name"]

            self.label_user_info.destroy()
            self.label_user_info = customtkinter.CTkLabel(
            master=self.frame_info,  
            text="Imie\t\t: %s \n\nNazwisko\t: %s \n\nStanowisko\t: %s \n\nDzial\t\t: %s \n\nE-mail\t\t: %s \n\nHasło\t\t: %s \n\nTel\t\t: %s \n\nNr_PC\t\t: %s\n\n"%(fname,lname,position,department,mail,password,sim_num,pc_name),
            height=100,
            corner_radius=15,
            fg_color=("white", "#333333"),
            justify=tkinter.LEFT)
            self.label_user_info.grid(column=0, row=0, sticky="nwe", padx=5, pady=40)
            
        
            
