# Importing the libraries
import win32com.client as win32
import pandas as pd

# Connecting with Outlook
outlook = win32.Dispatch('outlook.application')

# reading the file using the pandas library
read = pd.read_excel("The path of your file here.xlsx")

# Setting the e-mail informations
for data, row in read.iterrows():

    email = outlook.CreateItem(0)
    email.to = (row["EMAIL"])
    email.Subject = "Olá, " + (row["NAME"])

    email.HTMLBody = """
    <p>
    Good morning, """ + (row["NOME"]) + """! How are you?<br>
    <br>
    I am looking foward to see you! <br>
    <br>
    If you have any issues, email admin@test.com<br>
    <br>
        </p>

    """

    email.Send()

# Checking if everything worked
print("Email Sent")
