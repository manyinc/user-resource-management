from libraries import *
from get_data import *
import csv

def find_user(login):
    global row
    for row in data:
        if login in row["Email_Arago"]:
            return row
            
class Person:
    def __init__(self, row):
        self.city = row['Miasto']
        self.work_place = row['Miejsce_pracy']
        self.fname = row['Imie']
        self.lname = row['Nazwisko']
        self.position = row['Stanowisko']
        self.department = row['Dzial']
        self.mail = row['Email_Arago']
        self.password = row['PSW_Home']
        self.ms_mail = row['Email_Microsoft']

        #self.equipment = row['Has_equipment']

        self.pc = row['PC']
        self.system = row['System_Operacyjny']
        self.win_key = row['Windows_key']
        self.pc_type = row['PC_type']
        self.pc_name= row['PC_name']
        self.pc_vendor = row['PC_vendor']
        self.pc_model = row['PC_model']
        self.pc_sn = row['PC_SN']

        self.bag = row['Torba_laptop']
        self.mouse = row['Myszka']
        self.keyboard = row['Klawiatura']
        self.headphones = row['Sluchawki']
        self.headphones_model = row['Sluchawki_model']
        self.dock = row['Stacja_dokujaca']
        self.speaker = row['Glosniki']

        self.mon_1 = row['Monitor_1']
        self.mon_1_vendor = row['Monitor_1_vendor']
        self.mon_1_model = row['Monitor_1_model']
        self.mon_1_sn = row['Monitor_1_SN']

        self.mon_2 = row['Monitor_2']
        self.mon_2_vendor = row['Monitor_2_vendor']
        self.mon_2_model = row['Monitor_2_model']
        self.mon_2_sn = row['Monitor_2_SN']

        self.phone = row['Mobile']
        self.phone_type = row['Mobile_type']
        self.phone_vendor = row['Mobile_vendor']
        self.phone_model = row['Mobile_model']
        self.phone_sn = row['Mobile_SN']
        self.phone_imei = row['Mobile_IMEI']
        self.phone_sim = row['SIM']
        self.phone_pin = row['SIM_PIN']
        self.phone_puk = row['SIM_PUK']

        self.tab = row['Tablet']
        self.tab_name = row['Tablet_name']
        self.tab_vendor = row['Tablet_vendor']
        self.tab_model = row['Tablet_model']
        self.tab_sn = row['Tablet_SN']
        self.tab_imei = row['Tablet_IMEI']

        self.add_dev = row['Additional_device']
        self.dev_type = row['Device_type']
        self.dev_vendor = row['Device_vendor']
        self.dev_model = row['Device_model']
        self.dev_sn = row['Device_SN']
        self.dev_imei = row['Device_IMEI']
        self.dev_desc = row['Description']

        self.add_dev_2 = row['Additional_device_2']
        self.dev_type_2 = row['Device_type_2']
        self.dev_vendor_2 = row['Device_vendor_2']
        self.dev_model_2 = row['Device_model_2']
        self.dev_sn_2 = row['Device_SN_2']
        self.dev_imei_2 = row['Device_IMEI_2']
        self.dev_desc_2 = row['Description_2']

        self.fname_nu = ""
        self.lname_nu = ""
        self.secondary_table = False
        self.first_width_update = False
        self.second_width_update = False

    def protocol_name(self):
        #nazwa protokołu

        self.YEAR = 2023
        MYPATH = f"C:\\Users\m.zarzycki\\OneDrive - ARAGO Sp. z o.o\\Dysk\\DIT - Zespół Wsparcia i Utrzymania IT\\04_Ewidencja\\Spis komputerow i telefonow\\protokoły zdawczo-odbiorcze\\{self.YEAR}"

        files = next(walk(MYPATH), (None,None ,[]))[2] #2 files[0] | 1 folder
        length = len(files) - 1

        lastest_file = files[length]

        lastest_month = lastest_file[4:6]
        
        lastest_num = lastest_file[6:9]
        lastest_num = int(lastest_num)


        #pobranie aktualnej daty
        now_date = datetime.datetime.now()
        self.d_day = str(now_date.day)
        self.d_month = str(now_date.month)
        self.d_year =str(now_date.year)

        if int(self.d_day) < 10:
            self.d_day = str(0) + str(now_date.day) 

        if int(self.d_month) < 10:
            self.d_month = str(0) + str(now_date.month)
        
        #nazwa protokołu
        if self.d_month == lastest_month:
            num_protocol = lastest_num + 1
        else:
            num_protocol = 1
        num_protocol = str(num_protocol)

        #numer protokolu
        if int(num_protocol) > 9:
            self.name_protocol = self.d_year + "/" + self.d_month + "/" + "0" + num_protocol
            self.num_protocol = self.d_year + self.d_month + "0" + num_protocol
        else:
            self.name_protocol = self.d_year + "/" + self.d_month + "/" + "00" + num_protocol
            self.num_protocol = self.d_year + self.d_month + "00" + num_protocol

    def protocol(self,protocol_type):

        self.protocol_name()
        self.protocol_type = protocol_type

        lp = 1

        #otwieranie dokumentu
        doc_new = Document()
        if self.protocol_type == 1:
            doc_tmp = Document("template/protokol_temp_zdawczy.docx")
        else:
            doc_tmp = Document("template/protokol_temp_wydający.docx")
        length_s = len(doc_tmp.paragraphs)

        #formatowanie arkusza
        styles = doc_new.styles
        p = styles.add_style("Paragraph",WD_STYLE_TYPE.PARAGRAPH)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.font.name = 'Graphik Regular'
        p.font.size = Pt(10)

        h1 = styles.add_style("H1",WD_STYLE_TYPE.PARAGRAPH)
        h1.base_style = styles["Heading 1"]
        h1.font.name = 'Graphik Regular'
        h1.font.size = Pt(16)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(0,0,0)

        h2 = styles.add_style("H2",WD_STYLE_TYPE.PARAGRAPH)
        h2.base_style = styles["Heading 2"]
        h2.font.name = 'Graphik Regular'
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

        if self.protocol_type == 1:
            footer = section.footer
            p = footer.add_paragraph()
            p.alignment = 1
            r = p.add_run()
            r.add_picture("img/footer_last.png",width = Pt(400))
        else:
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
            if self.protocol_type == 1:
                cells[3].text = "nieuszkodzony"
        
        def set_col_widths_last(table):
            widths = (Pt(20), Pt(230), Pt(140),Pt(100))
            for row in table.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width
    
        def set_col_widths_first(table):
            widths = (Pt(20), Pt(230), Pt(190))
            for row in table.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width

        row_doc = doc_tmp.paragraphs[0].text
        center_row_header = doc_new.add_paragraph(row_doc , style="H1")
        center_row_header.alignment = 1

        for i in range(1,length_s):
            
            if doc_tmp.paragraphs[i].text == "Paste_nr":
                center_row = doc_new.add_paragraph(self.name_protocol,style="Paragraph")
                center_row.alignment = 1
            
            elif doc_tmp.paragraphs[i].text == "Paste_date":
                center_row = doc_new.add_paragraph("Sporządzony w dniu %s.%s.%sr w Brodnicy pomiędzy"%(self.d_day,self.d_month,self.d_year),style = "Paragraph")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "Paste_imie_nazwisko":
                if self.protocol_type == 1:  
                    center_row = doc_new.add_paragraph("pracownikiem firmy - „zdający” – %s %s"%(self.fname,self.lname),style = "Paragraph")
                    center_row.alignment = 1
                else:
                    center_row = doc_new.add_paragraph("pracownikiem firmy - „pobierający” – %s %s"%(self.fname,self.lname),style = "Paragraph")
                    center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "Tabela":

                if self.protocol_type == 1:  
                    table = doc_new.add_table(1,4)
                else:
                    table = doc_new.add_table(1,3)
                table.style = 'Table Grid'
                table.style.font.name = 'Graphik Regular'
                table.style.font.size = Pt(10)
                table.alignment = 1

                heading_cells = table.rows[0].cells
                heading_cells[0].text = 'Lp'
                heading_cells[1].text = 'Nazwa przedmiotu'
                heading_cells[2].text = 'Numer identyfikacyjny przedmiotu'
                if self.protocol_type == 1: 
                    heading_cells[3].text = 'Stan przedmiotu'  
                
                    
                if self.pc_sn != '' and self.pc_sn != '-':
                    add_table_row(str(lp),"Laptop %s %s"%(self.pc_vendor,self.pc_model),"S/N %s"%(self.pc_sn)) 
                    lp += 1

                if self.bag == 'tak':
                    add_table_row(str(lp),"Torba do laptopa","-")
                    lp += 1

                if self.mouse == 'tak':
                    add_table_row(str(lp),"Mysz bezprzewodowa","-")
                    lp += 1

                if self.keyboard == 'tak':
                    add_table_row(str(lp),"Klawiatura bezprzewodowa","-")
                    lp += 1

                if self.headphones == 'tak':
                    add_table_row(str(lp),"Słuchawki przewodowe Sennheiser PC 7 USB","-")
                    lp += 1

                if self.speaker == 'tak':
                    add_table_row(str(lp),"Głośnik","-")
                    lp += 1

                if self.dock == 'tak':
                    add_table_row(str(lp),"Stacja dokująca","-")
                    lp += 1

                if self.mon_1_sn != '' and self.mon_1_sn != '-':
                    add_table_row(str(lp),"Monitor %s %s"%(self.mon_1_vendor,self.mon_1_model),"S/N %s"%(self.mon_1_sn))
                    lp += 1
                
                if self.mon_2_sn != '' and self.mon_2_sn != '-':
                    add_table_row(str(lp),"Monitor %s %s"%(self.mon_2_vendor,self.mon_2_model),"S/N %s"%(self.mon_2_sn))
                    lp += 1

                if self.phone_sn != '' and self.phone_sn != '-':
                    add_table_row(str(lp),"Smartfon %s %s"%(self.phone_vendor,self.phone_model),"S/N %s\nIMEI %s"%(self.phone_sn,self.phone_imei))
                    lp += 1

                if self.phone_sim != '' and self.phone_sim != '-':
                    add_table_row(str(lp),"Karta SIM",str(self.phone_sim))
                    lp += 1

                if self.tab_sn != '' and self.tab_sn != '-':
                    add_table_row(str(lp),"Tablet %s %s"%(self.tab_vendor,self.tab_model ),"S/N %s\nIMEI %s"%(self.tab_sn,self.tab_imei))
                    lp += 1

                
                if self.add_dev != 'nie' and self.add_dev != '':
                    if self.dev_imei != '' and self.dev_imei != '-' and self.dev_sn != '' and self.dev_sn != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"S/N %s\nIMEI %s"%(self.dev_sn,self.dev_imei))
                        lp += 1
                    elif self.dev_sn != '' and self.dev_sn != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"S/N %s"%(self.dev_sn))
                        lp += 1
                    elif self.dev_imei == '' or self.dev_imei == '-' and self.dev_sn == '' or self.dev_sn == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"-")
                        lp += 1

                status_additional_2 = ""
                if self.add_dev_2 != 'nie' and self.add_dev_2 != '':
                    if self.dev_imei_2 != '' and self.dev_imei_2 != '-' and self.dev_sn_2 != '' and self.dev_sn_2 != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"S/N %s\nIMEI %s"%(self.dev_sn_2,self.dev_imei_2))
                        lp += 1
                    elif self.dev_sn_2 != '' and self.dev_sn_2 != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"S/N %s"%(self.dev_sn_2))
                        lp += 1
                    elif self.dev_imei_2 == '' or self.dev_imei_2 == '-' and self.dev_sn_2 == '' or self.dev_sn_2 == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"-")
                        lp += 1

            else:
                row_doc = doc_tmp.paragraphs[i].text
                center_row = doc_new.add_paragraph(row_doc,style = "Paragraph")
                center_row.alignment = 1
        
        if self.protocol_type == 1:  
            set_col_widths_last(table)
        else:
            set_col_widths_first(table)
        
        doc_new.save("C:\\Users\m.zarzycki\\OneDrive - ARAGO Sp. z o.o\\Dysk\\DIT - Zespół Wsparcia i Utrzymania IT\\04_Ewidencja\\Spis komputerow i telefonow\\protokoły zdawczo-odbiorcze\\%s\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(self.YEAR,self.num_protocol,self.fname,self.lname))
        convert("C:\\Users\m.zarzycki\\OneDrive - ARAGO Sp. z o.o\\Dysk\\DIT - Zespół Wsparcia i Utrzymania IT\\04_Ewidencja\\Spis komputerow i telefonow\\protokoły zdawczo-odbiorcze\\%s\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(self.YEAR,self.num_protocol,self.fname,self.lname))

    def transform_non_unicode(self):
        for letter_in_fname in self.fname :
            if letter_in_fname == "Ł":
                self.fname_nu = self.fname_nu + "L"
            elif letter_in_fname == "ł":
                self.fname_nu = self.fname_nu + "l"
            else:
                self.fname_nu = self.fname_nu + letter_in_fname
        self.fname_nu = unicodedata.normalize('NFKD', self.fname_nu).encode('ascii', 'ignore')
        self.fname_nu = self.fname_nu.decode('UTF-8')
        

        for letter_in_lname in self.lname :
            if letter_in_lname == "Ł":
                self.lname_nu = self.lname_nu + "L"
            elif letter_in_lname == "ł":
                self.lname_nu = self.lname_nu + "l"
            else:
                self.lname_nu = self.lname_nu + letter_in_lname
        self.lname_nu = unicodedata.normalize('NFKD', self.lname_nu).encode('ascii', 'ignore')
        self.lname_nu = self.lname_nu.decode('UTF-8')
        return self.fname_nu,self.lname_nu

    def footer_generate(self):
        self.transform_non_unicode()
        with open("template/Stopka.html", 'r',encoding='utf-8') as stopka:
            with open("footer\%s%s_arago_green (%s.%s@arago.green).htm"%(self.fname_nu[0].lower(),self.lname_nu.lower(),self.fname_nu[0].lower(),self.lname_nu.lower()), 'w',encoding='utf-8') as new_footer:
                for line in stopka:

                    if line.strip() == 'User_Data':
                        if self.phone_sim == '' or self.phone_sim == '-':
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>%s %s</b><br><span style='font-size: 11'>%s</span><br></p>"%(self.fname,self.lname,self.position))
                        else:
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>%s %s</b><br><span style='font-size: 11'>%s</span><br><span style='font-size: 12'>Tel.: +48 %s</span></p>"%(self.fname,self.lname,self.position,self.phone_sim))
                    elif line.strip() == 'Location_data':
                        if self.city == 'Warszawa':
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Biuro w Warszawie</b><br><span style='font-size: 12'>ul. Franciszka Klimczaka 1 <br>02-797 Warszawa</span></p>")
                        elif self.city == 'Rzeszów':
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Magazyn w Kolbuszowej</b><br><span style='font-size: 12'>ul. Sokołowska 28 G<br>36-100 Kolbuszowa</span></p>")
                        elif self.city == 'Kalisz':
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Magazyn w Kaliszu</b><br><span style='font-size: 12'>ul. Złota 44<br>62-800 Kalisz</span></p>")
                        elif self.city == 'Katowice':
                            new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Biuro w Katowicach</b><br><span style='font-size: 12'>Al. Roździeńskiego 188C<br>40-203 Katowice</span></p>")
                    else:
                        new_footer.write(line)


            new_footer.close()
        stopka.close()

    def send_email(self,box_gen_vcf,box_gen_footer,box_send_bookmarks,box_user_data_txt):
            toaddr = self.mail
            user = 'charon@arago.green'
            passw = 'XCm2VORTEX'
            port = 465
            subject = "Stopka mailowa + instrukcja zmiany stopki"

            attachment_contact_name = "contact\contact.vcf"
            attachment_footer_name =  "footer\%s%s_arago_green (%s.%s@arago.green).htm"%(self.fname_nu[0].lower(),self.lname_nu.lower(),self.fname_nu[0].lower(),self.lname_nu.lower())
            attachment_bookmarks =  "template\Bookmarks.html"
            attachment_footer_man =  "template\IT-009 Instrukcja zmiany stopki mailowej.pdf"
            #attachment_user_data =  "user\%s.%s"%(self.fname_nu[0].lower(),self.lname_nu.lower())

            msg = EmailMessage()
            msg['From'] = user
            msg['To'] = toaddr
            #msg['Bcc']="t.czepek@arago.green"
            msg['Subject'] = subject

            msg.set_type('text/html')
            msg.set_content(" This is the Data Message that we want to send")
            html_msg = """
            <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
            <html xmlns="http://www.w3.org/1999/xhtml">
            <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
            </head>
            <body bgcolor="#FFFFFF">
            <style>
            /* unvisited link */
            a:link {
            color: black;
            text-decoration: none;
            }
            /* visited link */
            a:visited {
            color: black;
            text-decoration: none;
            }
            /* mouse over link */
            a:hover {
            color: black;
            text-decoration: none;
            }
            /* selected link */
            a:active {
            color: black;
            text-decoration: none;
            }
            </style>
            <strong><h3>Zapisz i zamień stopkę w folderze :</h3></strong>
            <strong><h3>%appdata%\Microsoft\Signatures</h3></strong>
            <br>
            <br>
            <table width="435" border="0" cellpadding="0" cellspacing="0">
            <tr valign="top">
            <td width="208" height="156">
                <p style="font-family:arial;color:black;text-align:left;font-size:14px;">Z wyrazami szacunku<br><br></p>
            <p style="font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;"><b>Charon AI</b><br><span style="font-size: 11">Artificial Intelligence</span><br><span style="font-size: 12"></span></p>
            </td>
            <td width="177" height="156">
            <img src="http://arago.green/stopka/2023_znaczek.jpg" width="177" height="156" alt="" />
            </td>
            </tr>
            </table>
            <table width="475" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td width="270" height="88">
            <a href="http://www.arago.green">
            <img src="http://arago.green/stopka/2023_logo.jpg" width="270" height="88" alt="" /></a>
            </td>
            <td width="205" height="88" align="left">
            <p style='font-family:arial;color:black;text-align:left;font-size:13px;'><b>ARAGO Sp. z o.o.</b><br>
                    <span style="font-size: 12">ul. Podgórna 82A<br>87-300 Brodnica<br>
                    <a href="tel:600991359">Tel. +48 600 991 359</a><br>
                    <a href="mailto:kontakt@arago.green">e-mail: kontakt@arago.green</a></span></p>
            </td>
            </tr>
            </table>
            <table width="444" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td width="270" height="42">
            <a href="http://www.arago.green">
            <img src="http://arago.green/stopka/2023_www.jpg" width="270" height="42" alt="" /></a>
            </td>
            <td width="24" height="42">
            <a href="https://www.linkedin.com/company/aragogreen/">
            <img src="http://arago.green/stopka/2023_linkedin.png" width="42" height="42" alt="" /></a>
            </td>
            <td width="50" height="42">
            <img src="http://arago.green/stopka/2023_break.jpg" width="24" height="42" alt="" />
            </td>
            <td width="24" height="42">
            <a href="https://www.facebook.com/arago.green">
            <img src="http://arago.green/stopka/2023_facebook.png" width="42" height="42" alt="" /></a>
            </td>
            <td width="50" height="42">
            <img src="http://arago.green/stopka/2023_break.jpg" width="24" height="42" alt="" />
            </td>
            <td width="23" height="42">
            <a href="https://www.instagram.com/arago.green/">
            <img src="http://arago.green/stopka/2023_instagram.png" width="42" height="42" alt="" /></a>
            </td>
            </tr>
            </table>
            <br>
            <table width="450" border="0" cellpadding="0" cellspacing="0">
                <tr valign="middle">
                    <td width="100%" valign="middle">
                        <p style='font-family:arial;color:#9d9d9c;text-align:left;font-size:16px;line-height:21px;'>Nowe znaczenie komfortu dzięki inteligentnej technologii.</p>
                    </td>
                </tr>
                <tr valign="middle">
                    <td width="100%" valign="middle">
                        <p style='font-family:arial;color:#9d9d9c;text-align:left;font-size:16px;line-height:21px;'>Bądź <b><font color="#00686D">com</font><font color="#009AA1">tech</font></b>.</p>
                    </td>
                </tr>
            </table>
            <br>
            <table width="450" border="0" cellpadding="0" cellspacing="0">
                <tr valign="middle">
                    <td width="450" height="3">
                        <img src="http://arago.green/stopka/2023_linia.jpg" width="450" height="3" alt="" />
                    </td>
                </tr>
            </table>
            <br>
            <table width="auto" border="0" cellpadding="0" cellspacing="0">
                <tr valign="middle">
                    <td>
                        <img src="http://arago.green/stopka/2023_nagrody.jpg" width="450" height="auto" alt="" style="margin-right: 30px; margin-bottom: 4px;"/>
                    </td>
            </table>
            <br>
            <table width="450" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td width="450" height="3">
            <img src="http://arago.green/stopka/2023_linia.jpg" width="450" height="3" alt="" />
            </td>
            </tr>
            </table>
            <br>
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td width="100%" valign="middle">
                <p style="font-family:arial;color:#9d9d9c;text-align:left;font-size:8px;padding-right:22px;line-height:12px;">
                    UWAGA: Informacja zawarta w niniejszej wiadomości lub którymkolwiek z jej załączników podlega ochronie i jest objęta zakazem ujawniania. Jeśli czytelnik niniejszej wiadomości nie jest jej zamierzonym adresatem lub pośrednikiem upoważnionym do jej przekazania adresatowi, niniejszym informujemy, że wszelkie ujawnianie, w tym przekazanie osobom trzecim, rozprowadzanie, dystrybucja, powielanie niniejszej wiadomości lub jej załączników, bądź inne działanie o podobnym charakterze jest zabronione. Jeżeli otrzymałeś tę wiadomość omyłkowo, prosimy niezwłocznie zawiadomić nadawcę wysyłając odpowiedź na niniejszą wiadomość, a następnie usunąć ją z komputera bez otwierania załączników.
                    <br>Dziękujemy, ARAGO Sp. z o.o.
                    <br><br>ARAGO Sp. z o.o. z siedzibą w Brodnicy, ul. Podgórna 82A, 87-300 Brodnica wpisana do rejestru przedsiębiorców prowadzonego przez Sąd Rejonowy w Toruniu, VII Wydział Gospodarczy pod numerem <u>KRS: 0000686904, NIP: 9562326208, Regon: 367787398</u>.
                    <br>Wartość kapitału zakładowego: 2 100 000 PLN. BDO 000507551. FGAZ-P/03/0439/21.
                    <br><br>Administratorem danych osobowych jest Arago sp. z o.o. z siedzibą w Brodnicy (87- 300), ul. Podgórna 82A. Dane osobowe są przetwarzane w związku z realizacją prawnie uzasadnionego interesu Administratora polegającego na analizie przesłanego przez Ciebie zapytania lub zgłoszonej sprawy, w celu udzielenia odpowiedzi, a także w związku z koniecznością zapewnienia płynności komunikacji w ramach prowadzonej korespondencji (podstawa z art. 6 ust. 1 lit. f) RODO). Twoje dane osobowe przetwarzamy przez okres ważności naszego prawnie uzasadnionego interesu albo do czasu, aż zgłosisz swój sprzeciw. Więcej informacji dotyczących przetwarzania danych, w tym o przysługujących prawach znajduje się w <u><a href="https://arago.green/polityka-prywatnosci/" target="_blank">Polityce prywatności</a></u>.
            </p>
            </td>
            </tr>
            </table>
            <br>
            <table width="450" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td width="450" height="3">
            <img src="http://arago.green/stopka/2023_linia.jpg" width="450" height="3" alt="" />
            </td>
            </tr>
            </table>
            <br>
            <table width="530" border="0" cellpadding="0" cellspacing="0">
            <tr valign="middle">
            <td align="left">
            <img src="http://arago.green/stopka/2023_partnerzy.jpg" alt="" />
            </td>
            </tr>
            </table>
            </body>
            </html>
            """
            msg.add_alternative(html_msg, subtype="html")

            if box_gen_vcf == True :
                with open(attachment_contact_name, 'rb') as content_file:
                    content = content_file.read()
                    msg.add_attachment(content, maintype='application', subtype='pdf', filename=attachment_contact_name[8:len(attachment_contact_name)])

            if box_gen_footer == True :
                with open(attachment_footer_name, 'rb') as content_file:
                    content = content_file.read()
                    msg.add_attachment(content, maintype='application', subtype='pdf', filename=attachment_footer_name[7:len(attachment_footer_name)])
        
            if box_send_bookmarks == True :
                with open(attachment_bookmarks, 'rb') as content_file:
                    content = content_file.read()
                    msg.add_attachment(content, maintype='application', subtype='pdf', filename=attachment_bookmarks[9:len(attachment_bookmarks)])
                    
            if box_user_data_txt == True :
                with open(attachment_footer_man, 'rb') as content_file:
                    content = content_file.read()
                    msg.add_attachment(content, maintype='application', subtype='pdf', filename=attachment_footer_man[9:len(attachment_footer_man)])
                    #print(attachment_footer_man)
                    #print(attachment_footer_name)

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("serwer2077752.home.pl", port, context=context) as server:
                server.login(user, passw)
                server.send_message(msg)
            

            print("Soft >>> Mail sent to : " + toaddr)
            
    def protocol_new(self):

        self.protocol_name()
        lp = 1

        #otwieranie dokumentu
        doc_new = docx.Document()
        doc_tmp = docx.Document("template/new_protocol_temp.docx")
        
        sections = doc_new.sections
        for section in sections:
            section.top_margin = Cm(0)
            section.bottom_margin = Cm(0)
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)
            section.header_distance = Cm(0)
            section.footer_distance = Cm(0)

        length_s = len(doc_tmp.paragraphs)
        #formatowanie arkusza
        styles = doc_new.styles
        p = styles.add_style("Paragraph",WD_STYLE_TYPE.PARAGRAPH)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.font.name = 'Graphik Regular'
        p.font.size = Pt(9)

        h1 = styles.add_style("H1",WD_STYLE_TYPE.PARAGRAPH)
        h1.base_style = styles["Heading 1"]
        h1.font.name = 'Graphik Regular'
        h1.font.size = Pt(13)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(0,0,0)

        h2 = styles.add_style("H2",WD_STYLE_TYPE.PARAGRAPH)
        h2.base_style = styles["Heading 2"]
        h2.font.name = 'Graphik Regular'
        h2.font.size = Pt(9)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(0,0,0) 

        #nagłówek
        section = doc_new.sections[0]
        header = section.header
        p = header.add_paragraph()
        p.alignment = 2
        r = p.add_run()
        pic = r.add_picture("img/Header.png")

       
        footer = section.footer
        p = footer.add_paragraph()
        p.alignment = 1
        r = p.add_run()
        r.add_picture("img/Footer.png",width = Pt(400))
        

        #dodaj wiersz
        def add_table_row(c1,c2,c3):
            cells = table.add_row()
            cells = table.rows[lp].cells
            
            if self.secondary_table == True:
                cells[0].text = c1
                cells[1].text = f'{c2}\n{c3}'
                cells[2].text = ' '
                cells[3].text = ' '
            else:
                cells[0].text = c1
                cells[1].text = c2
                cells[2].text = c3
        
        def set_col_widths_second(table):
            widths = (Pt(20), Pt(220), Pt(200),Pt(100))
            for row in table.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width
                    row.cells[idx].height = 70
                    
    
        def set_col_widths_first(table):
            widths = (Pt(20), Pt(230), Pt(190))
            for row in table.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width

        for i in range(0,length_s):
            
            if "[PROTOCOL]" in doc_tmp.paragraphs[i].text:
                center_row = doc_new.add_paragraph('Protokół zdawczo-odbiorczy',style="H1")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "[PROTOCOL_NUMBER]":
                center_row = doc_new.add_paragraph(self.name_protocol,style="Paragraph")
                center_row.alignment = 1
            
            elif doc_tmp.paragraphs[i].text == "[DATE]":
                center_row = doc_new.add_paragraph(f'Sporządzony w dniu {self.d_day}.{self.d_month}.{self.d_year}r w Brodnicy pomiędzy',style = "Paragraph")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "[NAME]":
                center_row = doc_new.add_paragraph(f'współpracownikiem {self.fname} {self.lname} -  dalej jako „Pobierający”',style = "Paragraph")
                center_row.alignment = 1
                
            elif doc_tmp.paragraphs[i].text == "[FIRST_TABLE]":
                self.first_width_update = True
                table = doc_new.add_table(1,3)
                table.style = 'Table Grid'
                table.style.font.name = 'Graphik Regular'
                table.style.font.size = Pt(10)
                table.alignment = 1

                heading_cells = table.rows[0].cells
                heading_cells[0].text = 'Lp'
                heading_cells[1].text = 'Nazwa przedmiotu'
                heading_cells[2].text = 'Numer identyfikacyjny przedmiotu'

                if self.pc_sn != '' and self.pc_sn != '-':
                    add_table_row(str(lp),f'Laptop {self.pc_vendor} {self.pc_model}',f'S/N {self.pc_sn}') 
                    lp += 1
                    add_table_row(str(lp),'Ładowarka do laptopa','-')
                    lp += 1

                if self.bag == 'tak':
                    add_table_row(str(lp),'Torba do laptopa','-')
                    lp += 1

                if self.mouse == 'tak':
                    add_table_row(str(lp),'Mysz bezprzewodowa','-')
                    lp += 1

                if self.keyboard == 'tak':
                    add_table_row(str(lp),'Klawiatura bezprzewodowa','-')
                    lp += 1

                if self.headphones == 'tak':
                    add_table_row(str(lp),f'Słuchawki przewodowe {self.headphones_model}','-')
                    lp += 1

                if self.speaker == 'tak':
                    add_table_row(str(lp),'Głośnik','-')
                    lp += 1

                if self.dock == 'tak':
                    add_table_row(str(lp),'Stacja dokująca','-')
                    lp += 1

                if self.mon_1_sn != '' and self.mon_1_sn != '-':
                    add_table_row(str(lp),f'Monitor {self.mon_1_vendor} {self.mon_1_model}',f'S/N {self.mon_1_sn}')
                    lp += 1
                
                if self.mon_2_sn != '' and self.mon_2_sn != '-':
                    add_table_row(str(lp),f'Monitor {self.mon_2_vendor} {self.mon_2_model}',f'S/N {self.mon_2_sn}')
                    lp += 1

                if self.phone_sn != '' and self.phone_sn != '-':
                    add_table_row(str(lp),f'Smartfon {self.phone_vendor} {self.phone_model}',f'S/N {self.phone_sn}\nIMEI {self.phone_imei}')
                    lp += 1
                    add_table_row(str(lp),'Ładowarka do smartfona','-')
                    lp += 1

                if self.phone_sim != '' and self.phone_sim != '-':
                    add_table_row(str(lp),'Karta SIM',str(self.phone_sim))
                    lp += 1

                if self.tab_sn != '' and self.tab_sn != '-':
                    add_table_row(str(lp),f'Tablet {self.tab_vendor} {self.tab_model}',f'S/N {self.tab_sn}\nIMEI {self.tab_imei}')
                    lp += 1
                    add_table_row(str(lp),'Ładowarka do tableta','-')
                    lp += 1

                if self.add_dev != 'nie' and self.add_dev != '':
                    if self.dev_imei != '' and self.dev_imei != '-' and self.dev_sn != '' and self.dev_sn != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"S/N %s\nIMEI %s"%(self.dev_sn,self.dev_imei))
                        lp += 1
                    elif self.dev_sn != '' and self.dev_sn != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"S/N %s"%(self.dev_sn))
                        lp += 1
                    elif self.dev_imei == '' or self.dev_imei == '-' and self.dev_sn == '' or self.dev_sn == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type,self.dev_vendor,self.dev_model),"-")
                        lp += 1

                if self.add_dev_2 != 'nie' and self.add_dev_2 != '': 
                    if self.dev_imei_2 != '' and self.dev_imei_2 != '-' and self.dev_sn_2 != '' and self.dev_sn_2 != '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"S/N %s\nIMEI %s"%(self.dev_sn_2,self.dev_imei_2))
                        lp += 1
                    elif self.dev_sn_2 != '' and self.dev_sn_2 != '-' :
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"S/N %s"%(self.dev_sn_2))
                        lp += 1
                    elif self.dev_imei_2 == '' or self.dev_imei_2 == '-' and self.dev_sn_2 == '' or self.dev_sn_2 == '-':
                        add_table_row(str(lp),"%s %s %s %s"%(self.dev_type_2,self.dev_vendor_2,self.dev_model_2),"-")
                        lp += 1
            
            elif "[PROTOCOL_DESTORY]" in doc_tmp.paragraphs[i].text:
                center_row = doc_new.add_paragraph('Protokół zniszczenia mienia',style="H1")
                center_row.alignment = 1

            elif doc_tmp.paragraphs[i].text == "[SECOND_TABLE]":
                self.second_width_update = True
                self.secondary_table = True
                lp = 1
                table = doc_new.add_table(1,4)
                table.style = 'Table Grid'
                table.style.font.name = 'Graphik Regular'
                table.style.font.size = Pt(10)
                table.alignment = 1

                heading_cells = table.rows[0].cells
                heading_cells[0].text = 'Lp'
                heading_cells[1].text = 'Nazwa przedmiotu\n(Dane identyfikacyjne)'
                heading_cells[2].text = 'Opis usterki'
                heading_cells[3].text = 'Wartość szkody'
                lp = 0
                for i in range(1,11):
                    lp = lp + 1
                    add_table_row(str(lp),"","")

            else:
                row_doc = doc_tmp.paragraphs[i].text
                center_row = doc_new.add_paragraph(row_doc,style = "Paragraph")
                center_row.alignment = 1

            if self.first_width_update == True: 
                self.first_width_update = False 
                set_col_widths_first(table)

            if self.second_width_update == True: 
                self.second_width_update = False
                set_col_widths_second(table)

        doc_new.save("C:\\Users\m.zarzycki\\OneDrive - ARAGO Sp. z o.o\\Dysk\\DIT - Zespół Wsparcia i Utrzymania IT\\04_Ewidencja\\Spis komputerow i telefonow\\protokoły zdawczo-odbiorcze\\%s\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(self.YEAR,self.num_protocol,self.fname,self.lname))
        convert("C:\\Users\m.zarzycki\\OneDrive - ARAGO Sp. z o.o\\Dysk\\DIT - Zespół Wsparcia i Utrzymania IT\\04_Ewidencja\\Spis komputerow i telefonow\\protokoły zdawczo-odbiorcze\\%s\\%s_Protokol_zdawczo-odbiorczy_sprzetu_IT - %s %s.docx"%(self.YEAR,self.num_protocol,self.fname,self.lname))       
         

while True:
    login = input("User >>> Login : ")

    find_user(login)
    user = Person(row)

    user.protocol_new()
    #user.protocol(1)
    #user.protocol(0)

    #user.footer_generate()
    #user.send_email(False,True,False,True)


#generowanie plikow excel z działem w nazwie
"""
size = len(data)
print(size)
for i in range(0,size):
    row = data[i]
    user = Person(row)
    ud = user.department

    if "PV - B" not in ud and "PC - B" not in ud and ud != "-" and "S - 1" not in ud and "S - 2" not in ud and "Pracownicy Budowy" not in ud:
        user_dep_finall = ""
        for udl in ud:
            if udl == " ":
                user_dep_finall = user_dep_finall + "_"
            else:
                user_dep_finall = user_dep_finall + udl

        user_dep_finall = "Ewidencja_czasu_pracy_" + user_dep_finall + "_2023.xlsx"
        os.system(f'copy new_ewd\\template.xlsx new_ewd\\{user_dep_finall}')
"""

#wysyłka stopki do konkretnej osoby
"""
while True:
    login = input("User >>> Login : ")

    find_user(login)
    user = Person(row)

    #user.protocol_new()
    #user.protocol(1)
    #user.protocol(0)

    user.footer_generate()
    user.send_email(False,True,False,True)
"""

#generowanie danych do drukarki dla wszystkich z firmy (biuro , magazyn)
"""
with open("address_scan.txt", 'w',encoding='utf-8') as new_scan_list:
    new_scan_list.writelines("Abbreviated name\tDestination type\tSearch key\tFax: Fax number\tSIP Fax: Fax number\tIP Address Fax: IP address\tIP Address Fax: Port Number\tInternet Fax: Internet Fax address: E-Mail Address\tSMB: Host Address\tSMB: File Path\tSMB: User ID\tSMB: Password\tFTP: Host Address\tFTP: File Path\tFTP: Port Number\tFTP: User ID\tFTP: Password\tWebDAV: Host Address\tWebDAV: File Path\tWebDAV: Port Number\tWebDAV: User ID\tWebDAV: Password\tUser box:User box Name\n")
    new_scan_list.writelines("LICZNIKI\tJkl\t\t\t\t\t\tliczniki@copiersservice.pl\n")
    new_scan_list.writelines("Adrian Wiśniewski\tABC\ta.wisniewski@arago.greenn\n")
    size = len(data)
    for i in range(0,size):
        row = data[i]
        user = Person(row)
        user.gen_printer_address_book()
new_scan_list.close()
"""

#generowanie listy kontaktów
"""
with open("contact_list.txt", 'w',encoding='utf-8') as new_contact_list:
    size = len(data)
    next_num = 1
    for i in range(0,size):
        row = data[i]
        user = Person(row)
        if user.phone_sim != "-" and user.phone_sim != " ":
            new_contact_list.writelines(f'{next_num};{user.fname} {user.lname};{user.phone_sim[0:3]}{user.phone_sim[3:6]}{user.phone_sim[6:11]};{user.mail}\n')
            next_num = next_num + 1
new_contact_list.close()

"""

#masowe wysłanie maili do wszystkich z arkusza google
"""
size = len(data)
for i in range(0,size):
    row = data[i]
    user = Person(row)

    if user.lname != "Ślosarek" and user.lname != "Czapiewski" and user.lname != "Pustelnik" and user.position != "Doradca ds. OZE" and user.position != "Monter Instalacji Fotowoltaicznych" and user.position != "Monter Pomp Ciepła":
        user.footer_generate()
        if user.mail != "" and user.mail != "-":
            print(user.mail)
            user.send_email(False,True,False,True)
            sleep(random.randint(30,60))

"""

#kobieta | mezczyzna
"""
size = len(data)
for i in range(0,size):
    row = data[i]
    user = Person(row)
    name_len = len(user.fname) - 1
    if user.fname[name_len] == "a" and user.position != "Doradca ds. OZE" :
        print(user.mail)
"""

#tworzenie plików bez unicode
"""
excel_files = next(walk('new_ewd\\'), (None,None ,[]))[2]
for ex_file in excel_files:
    ex_done = ex_file[22:-10]
    ex_done_nu = ""
    for lt_in_exd in ex_done :
            if lt_in_exd == "Ł":
                ex_done_nu = ex_done_nu + "L"
            elif lt_in_exd == "ł":
                ex_done_nu = ex_done_nu + "l"
            else:
                ex_done_nu = ex_done_nu + lt_in_exd
    ex_done_nu = unicodedata.normalize('NFKD', ex_done_nu).encode('ascii', 'ignore')
    ex_done_nu = ex_done_nu.decode('UTF-8')
    ex_done_nu = ex_done_nu.lower()
    ex_done_nu = "GZ_" + ex_done_nu
    exf = ""
    for exl in ex_done:
        if exl == "_":
            exf = exf + " "
        else:
            exf = exf + exl

    print(f'{ex_done_nu}\t{exf}')

    #os.system(f'copy new_ewd\\template.xlsx new_ewd\\ok\\Ewidencja_czasu_pracy_{ex_done_nu}_2023.xlsx')
"""

#tworzenie nazw grup zabezpieczeń z plików
"""
excel_files = next(walk('new_ewd\\'), (None,None ,[]))[2]
os.system('cd new_ewd\\ok')
enum = 1
for ex_file in excel_files:
    ex_done = ex_file[22:-10]
    exf = ""
    
    for exl in ex_done:
        if exl == "_":
            exf = exf + " "
        else:
            exf = exf + exl
    
    if enum < 10 :
        exf = "0" + str(enum) + "_" + exf
    else:
        exf = str(enum) + "_" + exf

    enum = enum + 1
    
    os.system(f'mkdir \"{exf}\"')
"""