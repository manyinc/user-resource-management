from libraries import *
from get_data import *


def get_first_protocol(self,id):

    self.label_info_1 = customtkinter.CTkLabel(master=self.frame_right,  text="***   Generating ...   ***",  height=80,fg_color=("white", ORANGE),corner_radius=15,justify=tkinter.CENTER)  # <- custom tuple-color  
    self.label_info_1.grid(column=1 , row=5 , columnspan =4 , rowspan = 3 , sticky="nwe" , padx=15 , pady=30)
    self.update()
    
    data_row = []
    for row in data:
        if id in row["Email"]:
            data_row = row
            break

    try:
        fname = data_row["Imie"]
        lname = data_row["Nazwisko"]
        pc_brand = data_row["PC_vendor"]
        pc_model = data_row["PC_model"]
        pc_sn = data_row["PC_SN"]
        laptop_bag = data_row["Torba_laptop"]
        mouse = data_row["Myszka"]
        keyboard = data_row["Klawiatura"]
        headphones = data_row["Sluchawki"]
        dock_station = data_row["Stacja_dokujaca"]
        speaker = data_row["Glosniki"]
        monitor_1_brand = data_row["Monitor_1_vendor"]
        monitor_1_model = data_row["Monitor_1_model"]
        monitor_1_sn = data_row["Monitor_1_SN"]
        monitor_2_brand = data_row["Monitor_2_vendor"]
        monitor_2_model = data_row["Monitor_2_model"]
        monitor_2_sn = data_row["Monitor_2_SN"]
        tel_brand = data_row["Mobile_vendor"]
        tel_model = data_row["Mobile_model"]
        tel_sn = data_row["Mobile_SN"]
        tel_imei = data_row["Mobile_IMEI"]
        sim_num = data_row["SIM"]
        tab_brand = data_row["Tablet_vendor"]
        tab_model = data_row["Tablet_model"]
        tab_sn = data_row["Tablet_SN"]
        tab_imei = data_row["Tablet_IMEI"]
        additional_devaice = data_row["Additional_device"]
        dev_type = data_row["Device_type"]
        dev_vendor = data_row["Device_vendor"]
        dev_model = data_row["Device_model"]
        dev_sn = data_row["Device_SN"]
        dev_imei = data_row["Device_IMEI"]
        additional_devaice_2 = data_row["Additional_device_2"]
        dev_type_2 = data_row["Device_type_2"]
        dev_vendor_2 = data_row["Device_vendor_2"]
        dev_model_2 = data_row["Device_model_2"]
        dev_sn_2 = data_row["Device_SN_2"]
        dev_imei_2 = data_row["Device_IMEI_2"]

        lp = 1

        #nazwa protokołu
        YEAR = 2024
        mypatch_raw = f"protocol" 
        files = next(walk(mypatch_raw), (None,None ,[]))[2] #2 files[0] | 1 folder
        length = len(files) - 1
        lastest_file = files[length]
        lastest_month = lastest_file[4:6]
        
        lastest_num = lastest_file[6:9]
        lastest_num = int(lastest_num)


        #pobranie aktualnej daty
        now_date = datetime.datetime.now()
        d_day =str(now_date.day)
        d_month = str(now_date.month)
        d_year =str(now_date.year)

        if int(d_day) < 10:
            d_day = str(0) + str(now_date.day) 

        if int(d_month) < 10:
            d_month = str(0) + str(now_date.month)
        
        #nazwa protokołu
        if d_month == lastest_month:
            num_protocol = lastest_num + 1
        else:
            num_protocol = 1
        num_protocol = str(num_protocol)

        #numer protokolu
        if int(num_protocol) > 9:
            name_protocol = d_year + "/" + d_month + "/" + "0" + num_protocol
            num_protocol = d_year + d_month + "0" + num_protocol
        else:
            name_protocol = d_year + "/" + d_month + "/" + "00" + num_protocol
            num_protocol = d_year + d_month + "00" + num_protocol



        #otwieranie dokumentu
        doc_new = Document()
        doc_tmp = Document("template/protokol_temp_wydający.docx")
        length_s = len(doc_tmp.paragraphs)

        #formatowanie arkusza
        styles = doc_new.styles
        p = styles.add_style("Paragraph",WD_STYLE_TYPE.PARAGRAPH)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.font.name = 'Calibri'
        p.font.size = Pt(10)

        h1 = styles.add_style("H1",WD_STYLE_TYPE.PARAGRAPH)
        h1.base_style = styles["Heading 1"]
        h1.font.name = 'Calibri'
        h1.font.size = Pt(16)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(0,0,0)

        h2 = styles.add_style("H2",WD_STYLE_TYPE.PARAGRAPH)
        h2.base_style = styles["Heading 2"]
        h2.font.name = 'Calibri'
        h2.font.size = Pt(10)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(0,0,0) 

        #nagłówek
        section = doc_new.sections[0]
        header = section.header
        p = header.add_paragraph()
        p.alignment = 2
        r = p.add_run()
        pic = r.add_picture("img/Header.png")

        #stopka
        footer = section.footer
        p = footer.add_paragraph()
        p.alignment = 1
        r = p.add_run()
        r.add_picture("img/Footer.png",width = Pt(400))

        #dodaj wiersz
        def add_table_row(c1,c2,c3):
            cells = table.add_row()
            cells = table.rows[lp].cells
            cells[0].text = c1
            cells[1].text = c2
            cells[2].text = c3
            
        #szerokość kolumn
        def set_col_widths(table):
            widths = (Pt(20), Pt(230), Pt(190))
            for row in table.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width

        row_doc = doc_tmp.paragraphs[0].text
        center_row_header = doc_new.add_paragraph(row_doc , style="H1")
        center_row_header.alignment = 1

        for i in range(1,length_s):
            
            if doc_tmp.paragraphs[i].text == "Paste_nr":
                center_row = doc_new.add_paragraph(name_protocol,style="Paragraph")
                center_row.alignment = 1
            
            elif doc_tmp.paragraphs[i].text == "Paste_date":
                center_row = doc_new.add_paragraph("Sporządzony w dniu %s.%s.%sr w Gdańsku pomiędzy"%(d_day,d_month,d_year),style = "Paragraph")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "Paste_imie_nazwisko":
                center_row = doc_new.add_paragraph("pracownikiem firmy - „pobierający” – %s %s"%(fname,lname),style = "Paragraph")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "Tabela":

                table = doc_new.add_table(1,3)
                table.style = 'Table Grid'
                table.style.font.name = 'Calibri'
                table.style.font.size = Pt(10)
                table.alignment = 1

                heading_cells = table.rows[0].cells
                heading_cells[0].text = 'Lp'
                heading_cells[1].text = 'Nazwa przedmiotu'
                heading_cells[2].text = 'Numer identyfikacyjny przedmiotu'
                
                    
                if pc_sn != '' and pc_sn != '-':
                    add_table_row(str(lp),"Laptop %s %s"%(pc_brand,pc_model),"S/N %s"%(pc_sn)) 
                    lp += 1

                if laptop_bag == 'tak':
                    add_table_row(str(lp),"Torba do laptopa","-")
                    lp += 1

                if mouse == 'tak':
                    add_table_row(str(lp),"Mysz bezprzewodowa","-")
                    lp += 1

                if keyboard == 'tak':
                    add_table_row(str(lp),"Klawiatura bezprzewodowa","-")
                    lp += 1

                if headphones == 'tak':
                    add_table_row(str(lp),"Słuchawki przewodowe Sennheiser PC 7 USB","-")
                    lp += 1

                if speaker == 'tak':
                    add_table_row(str(lp),"Głośnik","-")
                    lp += 1

                if dock_station == 'tak':
                    add_table_row(str(lp),"Stacja dokująca","-")
                    lp += 1

                if monitor_1_sn != '' and monitor_1_sn != '-':
                    add_table_row(str(lp),"Monitor %s %s"%(monitor_1_brand,monitor_1_model),"S/N %s"%(monitor_1_sn))
                    lp += 1
                
                if monitor_2_sn != '' and monitor_2_sn != '-':
                    add_table_row(str(lp),"Monitor %s %s"%(monitor_2_brand,monitor_2_model),"S/N %s"%(monitor_2_sn))
                    lp += 1

                if tel_sn != '' and tel_sn != '-':
                    add_table_row(str(lp),"Smartfon %s %s"%(tel_brand,tel_model),"S/N %s\nIMEI %s"%(tel_sn,tel_imei))
                    lp += 1

                if sim_num != '' and sim_num != '-':
                    add_table_row(str(lp),"Karta SIM",str(sim_num))
                    lp += 1

                if tab_sn != '' and tab_sn != '-':
                    add_table_row(str(lp),"Tablet %s %s"%(tab_brand,tab_model ),"S/N %s\nIMEI %s"%(tab_sn,tab_imei))
                    lp += 1

                status_additional_1 = "" #"\n(przedmiot wydany bieżącym protokołem)"
                if additional_devaice != 'nie' and additional_devaice != '':
                    if dev_imei != '' and dev_imei != '-' and dev_sn != '' and dev_sn != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type,dev_vendor,dev_model,status_additional_1 ),"S/N %s\nIMEI %s"%(dev_sn,dev_imei))
                        lp += 1
                    elif dev_sn != '' and dev_sn != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type,dev_vendor,dev_model,status_additional_1 ),"S/N %s"%(dev_sn))
                        lp += 1
                    elif dev_imei == '' or dev_imei == '-' and dev_sn == '' or dev_sn == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type,dev_vendor,dev_model,status_additional_1 ),"-")
                        lp += 1

                status_additional_2 = ""
                if additional_devaice_2 != 'nie' and additional_devaice_2 != '':
                    if dev_imei_2 != '' and dev_imei_2 != '-' and dev_sn_2 != '' and dev_sn_2 != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type_2,dev_vendor_2,dev_model_2,status_additional_2 ),"S/N %s\nIMEI %s"%(dev_sn_2,dev_imei_2))
                        lp += 1
                    elif dev_sn_2 != '' and dev_sn_2 != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type_2,dev_vendor_2,dev_model_2,status_additional_2 ),"S/N %s"%(dev_sn_2))
                        lp += 1
                    elif dev_imei_2 == '' or dev_imei_2 == '-' and dev_sn_2 == '' or dev_sn_2 == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(dev_type_2,dev_vendor_2,dev_model_2,status_additional_2 ),"-")
                        lp += 1

            else:
                row_doc = doc_tmp.paragraphs[i].text
                center_row = doc_new.add_paragraph(row_doc,style = "Paragraph")
                center_row.alignment = 1

        set_col_widths(table)
        
        doc_new.save("protocol\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(num_protocol,fname,lname))
        convert("protocol\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(num_protocol,fname,lname))
        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_right,  text="Protocol compleat\n %s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(num_protocol,fname,lname),  height=80,fg_color=("white", GREEN),corner_radius=15,justify=tkinter.CENTER)  # <- custom tuple-color  
        self.label_info_1.grid(column=1 , row=5 , columnspan =4 , rowspan = 3 , sticky="nwe" , padx=15 , pady=30)
        self.update()
    except:
        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_right,  text="Error: No data found",  height=80,fg_color=("white", RED),corner_radius=15,justify=tkinter.CENTER)
        self.label_info_1.grid(column=1 , row=5 , columnspan =4 , rowspan = 3 , sticky="nwe" , padx=15 , pady=30)
        self.update()
    