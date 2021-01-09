import pyodbc,os,sys,pyautogui
from datetime import datetime
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter.ttk import *

def main():
    #Prompting sign-in from user input
    server = pyautogui.prompt("What server are you using?") #server you intend to use
    db = pyautogui.prompt("What database are you using?") #database you intend to use
    table = pyautogui.prompt("What table are you loading?") #table you intend to load data into
    splitting = pyautogui.prompt("How is the file separated(, tab |)?") #what separates the data
    

    userid = 'username'
    pwd = 'password'

    cn = pyodbc.connect('DRIVER={SQL Server};\
                        SERVER='+ server +'; \
                        DATABASE='+ db +';\
                        UID='+ userid + ';\
                        PWD='+ pwd +';',
                        autocommit=True  )


    csr = cn.cursor()



    #Asking user to delete data in SQL table
    while True:
        delete = pyautogui.prompt("Delete data in table?(Y/N)")

        if delete == 'y' or delete == 'Y':
            sql = f"""DELETE FROM {table};"""
            csr.execute(sql)
            pyautogui.confirm(f"{table} data has been deleted")
            break


        elif delete == 'n' or delete == 'N':
            break

        else:
            pyautogui.alert("Your input isn't Y or N. Enter Y or N")


    #Pulling file
    root = Tk()
    root.withdraw()
    f1 = askopenfilename(initialdir="Desktop",filetypes=[("Text file", "*.txt"),("CSV file", "*.csv")]) #add any file type into the file type list
    num_lines = sum(1 for line in open(os.path.join(sys.path[0],f1)))-1
    num = 0

    #Parsing data
    with open(os.path.join(sys.path[0],f1),"r") as f:
        for j,i in enumerate(f):

            if j == 0:
                #print(i)
                x = i.strip('\n')
                x = x.replace('\t','|')
                x = x.strip("'")
                #x = x.replace("'","''")
                if splitting == 'tab' or splitting == 'Tab' or splitting == 'TAB':
                    x = x.split('|')
                else:
                    x = x.split(splitting)
                len_of_header = len(x)
                header = ["["+d+"]" for d in x]
                header = str(header)[1:-1]
                header = header.replace("'","")
                #print(header)
                


            else:

                x = i.strip('\n')
                x = x.replace('\t','|')
                x = x.strip("'")
                #x = x.replace("'","''")
                x = x.replace("'","")
                if splitting == 'tab'or splitting == 'Tab' or splitting == 'TAB':
                    x = x.split('|')
                else:
                    x = x.split(splitting)
                len_of_column = len(x)
                column = [d[:8000] for d in x]
                column = str(column)[1:-1]



                insert_sql = f"INSERT INTO {db}.dbo.{table}({header}) values({column})"
                #print(insert_sql)
                try:
                    csr.execute(insert_sql)
                except pyodbc.Error as e:
                    pyautogui.alert("Length of column doesn't fit into specified storage of SQL column")
                    print(e)
                    print(f"Num of Headers: {len_of_header};Num of Columns: {len_of_column}")
                    print("error on line: "+str(num + 2))
                    #print(insert_sql)
                    #print(f"Header Name:{header}; Length:{len(i) for i in column}")
                    break

                num +=1
                print(f"{num}/{num_lines}")



if __name__ == "__main__":
    main()
