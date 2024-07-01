from libraries import *
from get_data import *
from get_last_protocol import *
from get_first_protocol import *
from get_footer import *
from print_data import *
from get_user_list_data import *
from get_user_list_contact import *
from get_user_list_contact_vcf import *
from send_mail import *
from post_user import *
from qr_gen import *

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):

    WIDTH = 900
    HEIGHT = 700

    def __init__(self):
        super().__init__()

        self.title("TAU Data Base Management")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed

        # ============ create two frames ============

        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self, width=180, corner_radius=15, border_width=1,border_color="#00B2B5")
        self.frame_left.grid(row=0, column=0, sticky="nswe")

       
        # ============ frame_left ============

        # configure grid layout (1x11)
        self.frame_left.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(14, weight=1)  # empty row as spacing
        self.frame_left.grid_rowconfigure(14, minsize=20)    # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(16, minsize=10)  # empty row with minsize as spacing
        self.label_1 = customtkinter.CTkLabel(master=self.frame_left,text="TAU DBM")  # font name and size in px
        self.label_1.grid(row=2, column=0, pady=10, padx=10)
        #SEARCH ON TOP WIDTH SCROLL LIST TO CHOSE USER AFTER SEARCH UNDER SWITCH TO SELECT OPTION ACCAUNT (ENABLE,DISABLE)
        
        #CREATE NEW USER
        self.button_4 = customtkinter.CTkButton(master=self.frame_left,text="New User",fg_color=("#00B2B5", "#2E2E2E"),corner_radius=15,border_width=1,border_color="#00B2B5",command=self.new_user_window)
        self.button_4.grid(row=4, column=0, pady=10, padx=20)
        #GENERATE QR CODE
        self.button_3 = customtkinter.CTkButton(master=self.frame_left,text="QR Code",fg_color=("#00B2B5", "#2E2E2E"),corner_radius=15,border_width=1,border_color="#00B2B5",command=self.qr_window)
        self.button_3.grid(row=5, column=0, pady=10, padx=20)
        #GENERATE VCF FILE WITH CONTACT FOOTER , PRINT(LOGIN,PASS)
        self.button_4 = customtkinter.CTkButton(master=self.frame_left,text="Addons",fg_color=("#00B2B5", "#2E2E2E"),corner_radius=15,border_width=1,border_color="#00B2B5",command=self.generator_window)
        self.button_4.grid(row=6, column=0, pady=10, padx=20)
        #USER INFO
        self.button_2 = customtkinter.CTkButton(master=self.frame_left,text="User Info",fg_color=("#00B2B5", "#2E2E2E"),corner_radius=15,border_width=1,border_color="#00B2B5",command=self.user_list)
        self.button_2.grid(row=7, column=0, pady=10, padx=20)
        #SELECT PROTOCOL TYPE (FIRST,CHANGE,LAST)
        self.button_1 = customtkinter.CTkButton(master=self.frame_left , text="Protocol" ,fg_color=("#00B2B5", "#2E2E2E"), corner_radius=15,border_width=1,border_color="#00B2B5",command=self.protocol_window)
        self.button_1.grid(row=8, column=0, pady=10, padx=20)


        #GET USER MAIL LIST FILTERS OPTIONS(GENDER,DEPEND,POSITION)
        #self.button_4 = customtkinter.CTkButton(master=self.frame_left,text="User Mail",command=self.button_event)
        #self.button_4.grid(row=5, column=0, pady=10, padx=20)
        #SELECT SCROLL LIST DEVAICES(PHONE,SIM_CARD,PC,DISPLAY)
        #self.button_5 = customtkinter.CTkButton(master=self.frame_left,text="Devices",command=self.button_event)
        #self.button_5.grid(row=6, column=0, pady=10, padx=20)
        #GENERATE CONTACT LIST EXCEL FILE
        #self.button_6 = customtkinter.CTkButton(master=self.frame_left,text="Contact Excel",command=self.button_event)
        #self.button_6.grid(row=7, column=0, pady=10, padx=20)
        #DOCX TO PDF
        #self.button_8 = customtkinter.CTkButton(master=self.frame_left,text="Docx to Pdf",command=self.button_event)
        #self.button_8.grid(row=9, column=0, pady=10, padx=20)

        #self.label_mode = customtkinter.CTkLabel(master=self.frame_left, text="Appearance Mode:")
        #self.label_mode.grid(row=15, column=0, pady=0, padx=20, sticky="w")
        #self.optionmenu_1 = customtkinter.CTkOptionMenu(master=self.frame_left,values=["Light", "Dark", "System"],corner_radius=15,border_width=1,border_color="#00B2B5",command=self.change_appearance_mode)
        #self.optionmenu_1.grid(row=15, column=0, pady=10, padx=20, sticky="w")
        #self.optionmenu_1.set("Dark")



        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=5, pady=0)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** Welcome ***",height=50,fg_color=("white", "#00B2B5"),corner_radius=15,justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(column=0, row=0, columnspan = 2 ,sticky="nwe", padx=15, pady=15)

    def new_user_window(self):   

        kill_window = self.frame_right
        kill_window = kill_window.destroy()

        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe",padx=5, pady=0)
        self.frame_right.rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.columnconfigure(1, weight=0)
        self.frame_right.columnconfigure(3, weight=1)

        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** Create new user ***",height=50,fg_color=("white", "#00B2B5"),corner_radius=20,justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(row=0, column=1, columnspan=2, sticky="we", pady=30)

        #self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="First name  :",height=30,justify=tkinter.LEFT)  # <- custom tuple-color  
        #self.label_user_info.grid(column=1, row=1, columnspan = 1 ,sticky="we")
        self.fname_text = customtkinter.CTkEntry(master=self.frame_right,width=80,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="First name")
        self.fname_text.grid(row=1, column=1, columnspan=2, sticky="we")

        #self.lname_label = customtkinter.CTkLabel(master=self.frame_right, text="Last name  :",height=30,justify=tkinter.LEFT)  # <- custom tuple-color  
        #self.lname_label.grid(column=1, row=2, columnspan = 1 ,sticky="we")
        self.lname_text = customtkinter.CTkEntry(master=self.frame_right,width=80,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Last name")
        self.lname_text.grid(row=2, column=1, columnspan=2, sticky="we")

        #self.position_label = customtkinter.CTkLabel(master=self.frame_right, text=" Position  :",height=30,justify=tkinter.LEFT)  # <- custom tuple-color  
        #self.position_label.grid(column=1, row=3, columnspan = 1 ,sticky="we")
        self.position_text = customtkinter.CTkEntry(master=self.frame_right,width=80,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Position")
        self.position_text.grid(row=3, column=1, columnspan=2, sticky="we")

        #self.department_label = customtkinter.CTkLabel(master=self.frame_right, text="Department  :",height=30,justify=tkinter.LEFT)  # <- custom tuple-color  
        #self.department_label.grid(row=4, column=1, columnspan = 1 ,sticky="we")
        self.department_text = customtkinter.CTkEntry(master=self.frame_right,width=80,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Department")
        self.department_text.grid(row=4, column=1, columnspan=2, sticky="we")

        #self.workplace_label = customtkinter.CTkLabel(master=self.frame_right, text="Work Place  :",height=30,justify=tkinter.LEFT)
        #self.workplace_label.grid(row=5, column=1, columnspan = 1 ,sticky="we")
        self.workplace = customtkinter.CTkComboBox(master=self.frame_right,border_width=1,corner_radius=15,border_color="#00B2B5",values=["Biuro", "Magazyn", "Teren", "Handlowiec"])
        self.workplace.grid(row=5, column=1, columnspan=2, sticky="we")

        #self.city_label = customtkinter.CTkLabel(master=self.frame_right, text="City  :",height=30,justify=tkinter.LEFT)
        #self.city_label.grid(row=6, column=1, columnspan = 1 ,sticky="we")
        self.city = customtkinter.CTkComboBox(master=self.frame_right,border_width=1,corner_radius=15,border_color="#00B2B5",values=["Brodnica", "Warszawa", "Katowice", "Kalisz", "Rzeszów", "Legnica"])
        self.city.grid(row=6, column=1, columnspan=2, sticky="we")

        self.button_search = customtkinter.CTkButton(master=self.frame_right,text="Add user to database",border_width=1,fg_color=None,corner_radius=15,border_color="#00B2B5",command = self.create_user)
        self.button_search.grid(row=7, column=1, columnspan=2, pady=30, sticky="we")

    def protocol_window(self):   

        kill_window = self.frame_right
        kill_window = kill_window.destroy()

        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=5, pady=0)

        self.frame_right.rowconfigure((0, 1, 2, 3), weight=0)
        self.frame_right.rowconfigure(7, weight=10)

        self.frame_right.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
        self.frame_right.columnconfigure(6, weight=0)

        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** Protocol ***",height=50,fg_color=("white", "#00B2B5"),corner_radius=20,justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(column=1, row=0, columnspan = 4 ,sticky="nwe", padx=15, pady=15)

        self.id_protocol = customtkinter.CTkEntry(master=self.frame_right,width=120,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Username")
        self.id_protocol.grid(row=1, column=1, columnspan=4, pady=15, padx=20, sticky="we")

        self.radio_var = tkinter.IntVar(value=0)
        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.frame_right,text="Protokół Wydający",variable=self.radio_var,value=0)
        self.radio_button_1.grid(row=3, column=1, pady=15, padx=20, sticky="n")
        
        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.frame_right,text="Protokół Zdający",variable=self.radio_var,value=1)
        self.radio_button_2.grid(row=3, column=4, pady=15, padx=20, sticky="n")
        
        self.button_search = customtkinter.CTkButton(master=self.frame_right,text="Generate Protocol",border_width=1,corner_radius=15,border_color="#00B2B5",hover_color= "#808080",fg_color=None,command = self.protocol_gen)
        self.button_search.grid(row=4, column=1, columnspan=4, pady=15, padx=20, sticky="we")

    def qr_window(self):
        
        kill_window = self.frame_right
        kill_window = kill_window.destroy()

           
        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=5, pady=0)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)

        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** New QR Code ***",height=50,fg_color=("white", "#00B2B5"),corner_radius=20,justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(column=0, row=0, columnspan = 2 ,sticky="nwe", padx=15, pady=15)

        self.link_qr = customtkinter.CTkEntry(master=self.frame_right,width=120,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="QR code link")
        self.link_qr.grid(row=1, column=0, columnspan=2, pady=5, padx=20, sticky="we")

        self.qr_name = customtkinter.CTkEntry(master=self.frame_right,width=120,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Name qr code")
        self.qr_name.grid(row=2, column=0, columnspan=2, pady=5, padx=20, sticky="we")

        self.radio_var_color = tkinter.IntVar(value=0)
        self.radio_button_1_color = customtkinter.CTkRadioButton(master=self.frame_right,text="Kolory Firmowe",variable=self.radio_var_color,value=0)
        self.radio_button_1_color.grid(row=3, column=0, pady=10, padx=20, sticky="n")
        self.radio_button_2_color = customtkinter.CTkRadioButton(master=self.frame_right,text="Czarno-Biały",variable=self.radio_var_color,value=1)
        self.radio_button_2_color.grid(row=3, column=1, pady=10, padx=20, sticky="n")

        self.button_qr_gen = customtkinter.CTkButton(master=self.frame_right,text="Generate QR code",border_width=1,corner_radius=15,border_color="#00B2B5",fg_color=None,command = self.qr_gen)
        self.button_qr_gen.grid(row=4, column=0, columnspan=2, pady=20, padx=20, sticky="we")

    def generator_window(self):   
        
        kill_window = self.frame_right
        kill_window = kill_window.destroy()

        
        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe",padx=5, pady=0)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** Mega generator ***",height=70,fg_color=("white", "#00B2B5"),justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(column=0, row=0, columnspan = 2 ,sticky="nwe", padx=15, pady=15)
        self.id = customtkinter.CTkEntry(master=self.frame_right,width=120,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Username")
        self.id.grid(row=1, column=0, columnspan=2, pady=5, padx=20, sticky="we")
       
        self.box_gen_footer = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Generuj stopke")
        self.box_gen_footer.grid(row=2, column=0, pady=10, padx=20, sticky="w")

        self.box_gen_vcf = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Generuj .vcf")
        self.box_gen_vcf.grid(row=2, column=1, pady=10, padx=20, sticky="w")
        
        self.box_send_mail = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Wyślij mail")
        self.box_send_mail.grid(row=3, column=0, columnspan=2, pady=10, padx=20, sticky="w")

        self.box_print_data = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Drukuj dane logowania")
        self.box_print_data.grid(row=3, column=1, columnspan=2, pady=10, padx=20, sticky="w")

        self.box_send_bookmarks = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Zakładki")
        self.box_send_bookmarks.grid(row=4, column=0, columnspan=2, pady=10, padx=20, sticky="w")

        self.box_user_data_txt = customtkinter.CTkCheckBox(master=self.frame_right,corner_radius=15,fg_color=("white", "#00B2B5"),border_width=1,border_color="#00FF00",text="Dane uzytkownika")
        self.box_user_data_txt.grid(row=4, column=1, columnspan=2, pady=10, padx=20, sticky="w")

        self.button_search = customtkinter.CTkButton(master=self.frame_right,text="Compleat",border_width=1,corner_radius=15,border_color="#00B2B5",fg_color=None,command = self.addon_gen)
        self.button_search.grid(row=5, column=0, columnspan=2, pady=20, padx=20, sticky="we")
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right ,border_width=1,border_color="#00FF00")
        self.frame_info.grid(row=6, column=0, columnspan=2, rowspan=4, pady=20, padx=20, sticky="nsew")
        self.frame_info.rowconfigure(0, weight=1)
        self.frame_info.columnconfigure(0, weight=1)

    def user_list(self):

        kill_window = self.frame_right
        kill_window = kill_window.destroy()

        self.frame_right = customtkinter.CTkFrame(master=self, corner_radius=20, border_width=1, border_color="#00B2B5")
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=5, pady=0)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.label_user_info = customtkinter.CTkLabel(master=self.frame_right, text="*** User Info ***",height=70,fg_color=("white", "#00B2B5"),justify=tkinter.LEFT)  # <- custom tuple-color  
        self.label_user_info.grid(column=0, row=0, columnspan = 2 ,sticky="nwe", padx=15, pady=15)
        self.search_box = customtkinter.CTkEntry(master=self.frame_right,width=120,height=30,corner_radius=15,border_width=1,border_color="#00B2B5",placeholder_text="Entry user name or data")
        self.search_box.grid(row=1, column=0, columnspan=2, pady=5, padx=20, sticky="we")
        self.button_search = customtkinter.CTkButton(master=self.frame_right,text="Search user",border_width=2,corner_radius=15,border_color="#00B2B5",fg_color=None,command = self.search_user_data)
        self.button_search.grid(row=2, column=0, columnspan=2, pady=20, padx=20, sticky="we")

        self.frame_info = customtkinter.CTkFrame(master=self.frame_right,border_width=1,border_color="#00FF00",)
        self.frame_info.grid(row=3, column=0, columnspan=2, rowspan=5, pady=20, padx=20, sticky="nsew")
        self.frame_info.rowconfigure(0, weight=1)
        self.frame_info.columnconfigure(0, weight=1)
        self.label_user_info = customtkinter.CTkLabel(master=self.frame_info,  
            text="Imie\t\t:\n\nNazwisko\t:\n\nStanowisko\t:\n\nDzial\t\t:\n\nE-mail\t\t:\n\nHasło\t\t:\n\nTel\t\t:\n\nNr_PC\t\t:\n\n",
            height=100,
            corner_radius=15,
            fg_color=("white", "#333333"),

            justify=tkinter.LEFT) 
        self.label_user_info.grid(column=0, row=0, sticky="nwe", padx=5, pady=40)

##########################################################################################################################################

    def create_user(self): 
        fname_text = self.fname_text.get()
        lname_text = self.lname_text.get()
        position_text = self.position_text.get()
        department_text = self.department_text.get()
        workplace_text = self.workplace.get()
        city_text = self.city.get()
        post_user(city_text, workplace_text, fname_text, lname_text, position_text, department_text)

    def search_user_data(self):
        search_box = self.search_box.get()
        get_user_list_data(self,search_box)
        print("searching")

    def protocol_gen(self):
        id_protocol = self.id_protocol.get()
        protocol_select = self.radio_var.get()
        protocol_select = int(protocol_select)
        print(protocol_select)
        if protocol_select == 0 :
            get_first_protocol(self,id_protocol)
        else:
            get_last_protocol(self,id_protocol)
   
    def addon_gen(self):
        id = self.id.get()
        box_gen_footer = self.box_gen_footer.get()
        box_gen_vcf = self.box_gen_vcf.get()
        box_send_mail = self.box_send_mail.get()
        box_print_data = self.box_print_data.get()
        box_send_bookmarks = self.box_send_bookmarks.get()
        box_user_data_txt = self.box_user_data_txt.get()
        print(box_send_bookmarks)
        if box_gen_footer == True:
            get_footer(id)
        if box_gen_vcf == True:
            get_user_list_contact_vcf()
        if box_send_mail == True:
            send_mail_attachment(id,box_gen_vcf,box_gen_footer,box_send_bookmarks,box_user_data_txt)
        if box_print_data == True:
            print_data(id)
            print("Wydrukowano dane logowania")

    def button_event(self): # wybranie opcji protokołu
        id = self.id.get()
        num_protocol = self.num_protocol.get()
        protocol_select = self.radio_var.get()
        protocol_select = str(protocol_select)
        print("Button pressed" + id + " " + num_protocol + " " + protocol_select)
    
    def qr_gen(self):
        link_qr = self.link_qr.get()
        qr_name = self.qr_name.get()
        color_select = self.radio_var_color.get()
        color_select = int(color_select)
        get_qr(link_qr,qr_name,color_select)

    def change_appearance_mode(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def on_closing(self, event=0):
        self.destroy()
    
if __name__ == "__main__":
    app = App()
    app.mainloop()