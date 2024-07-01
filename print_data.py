from libraries import *
from get_data import *

def print_data(id):
    for row in data:
        if id in row["Email"]:

            mail = row["Email_Arago"]
            password = row["PSW_Home"]
            err = 0
            login =""
            while mail[err] != "@":
                login = login + mail[err]
                err = err + 1
            anytext = login + "\n" + password
            filename = tempfile.mktemp(".txt")
            open(filename, "w").write(anytext)
            os.startfile(filename, "print")
            
            