from libraries import *
from get_data import *

def get_footer(id):

    for row in data:
        if id in row["Email"]:

            city = row["Miasto"]
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

            position = row["Stanowisko"]
            print(f'Wygenerowano stopkę dla : {fname} {lname}')
            numer = row["SIM"]
            with open("template/Stopka.html", 'r',encoding='utf-8') as stopka:
                with open(f'footer\{fname[0].lower()}{lname.lower()}_{COMPANY_SPACE} ({fname[0].lower()}.{lname.lower()}@{DOMAIN}).htm', 'w',encoding='utf-8') as new_footer:
                    for line in stopka:

                        if line.strip() == 'User_Data':
                            if numer == '' or numer == '-':
                                new_footer.writelines(f'<p style="font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;"><b>{imie} {nazwisko}</b><br><span style="font-size: 11">{position}</span><br></p>')
                            else:
                                new_footer.writelines(f'<p style="font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;"><b>{imie} {nazwisko}</b><br><span style="font-size: 11">{position}</span><br><span style="font-size: 12">Tel.: +48 {numer}</span></p>')
                        elif line.strip() == 'Location_data':
                            if city == 'Warszawa':
                                new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Biuro w Warszawie</b><br><span style='font-size: 12'>ul. Mazowiecka 123 <br>02-797 Warszawa</span></p>")
                            elif city == 'Rzeszów':
                                new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Magazyn w Kolbuszowej</b><br><span style='font-size: 12'>ul. Marcepanowa 321<br>36-100 Kolbuszowa</span></p>")
                            elif city == 'Kalisz':
                                new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Magazyn w Kaliszu</b><br><span style='font-size: 12'>ul. Złota 123<br>62-800 Kalisz</span></p>")
                            elif city == 'Katowice':
                                new_footer.writelines("<p style='font-family:arial;color:black;text-align:left;font-size:12px;line-height:21px;'><b>Biuro w Katowicach</b><br><span style='font-size: 12'>ul. Matejki 12<br>40-203 Katowice</span></p>")
                        else:
                            new_footer.write(line)
                new_footer.close()
            stopka.close()