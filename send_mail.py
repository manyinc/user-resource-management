from libraries import *
from get_data import *


def send_mail_attachment(id,box_gen_vcf,box_gen_footer,box_send_bookmarks,box_user_data_txt):
    size = len(data)
    for i in range(0,size):
        row = data[i]
        fname = row["Imie"]
        lname = row["Nazwisko"]
        mail = row["Email"]
        combo_name = fname.lower() + " " + lname.lower()    
        combo_name_revers = lname.lower() + " " +  fname.lower()
        id_str = id.lower()
        
        if id_str in combo_name or id_str in combo_name_revers : #and id_str != ""
            fname = ""
            lname = ""
            num_row = int(i)
            row = data[num_row]
            imie = row["Imie"]
            for letter_in_imie in imie :
                if letter_in_imie == "Ł":
                    fname = fname + "L"
                elif letter_in_imie == "ł":
                    fname = fname + "l"
                else:
                    fname = fname + letter_in_imie
            fname = unicodedata.normalize('NFKD', fname).encode('ascii', 'ignore')
            fname = fname.decode('UTF-8')

            nazwisko = row["Nazwisko"]
            for letter_in_nazwisko in nazwisko :
                if letter_in_nazwisko == "Ł":
                    lname = lname + "L"
                elif letter_in_nazwisko == "ł":
                    lname = lname + "l"
                else:
                    lname = lname + letter_in_nazwisko
            lname = unicodedata.normalize('NFKD', lname).encode('ascii', 'ignore')
            lname = lname.decode('UTF-8')

            toaddr = mail
            user = MAIL_TO_SEND_MSG
            passw = PASS_TO_SEND_MSG
            smtp_server = SMTP_SERVER
            port = 465
            subject = "Stopka mailowa / dane konfiguracyjne"
            attachment_contact_name = "contact\contact.vcf"
            attachment_footer_name =  f'footer\{fname[0].lower()}{lname.lower()}_{COMPANY_SPACE} ({fname[0].lower()}.{lname.lower()}@{DOMAIN}).htm'
            attachment_bookmarks =  "template\Bookmarks.html"
            attachment_user_data =  "user\%s.%s"%(fname[0].lower(),lname.lower())

            msg = EmailMessage()
            msg['From'] = user
            msg['To'] = toaddr
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
            <strong><h3>Zapisz stopkę w folderze :</h3></strong>
            <strong><h3>%appdata%\Microsoft\Signatures</h3></strong>
            <br>
            <br>
            <table width="435" border="0" cellpadding="0" cellspacing="0">
                <tr valign="top">
                    <td width="208" height="156">
                        <h1>HTML signature</h1>
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
                with open(attachment_user_data, 'rb') as content_file:
                    content = content_file.read()
                    msg.add_attachment(content, maintype='application', subtype='pdf', filename=attachment_user_data[5:len(attachment_user_data)])

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
                server.login(user, passw)
                server.send_message(msg)
            

            print("Mail sent to : " + toaddr)
            

