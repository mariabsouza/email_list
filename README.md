# Sending e-mails with python from an e-mail list on Excel :computer:

### What this will do?

This script in python will get contact info from a spreadsheet on Excel using ```pandas``` and it will send e-mails to all the people on the spreadsheet in a an automatic way. 
It´s also possible to make personalized e-mail with different names, last name and any other data that you put in the spreasheet.

### How to use it

* In line 9, type the path of your file with a .xlsx at the end
```read = pd.read_excel("The path of your file here.xlsx")```

* In your excel file, it´s important to name the columns so you can call them and personalize your e-mail

![alt text](https://www.imagemhost.com.br/images/2021/07/22/Captura-de-tela-2021-07-22-145710.png)

* If you want to use a data from the spreadsheet to personalize each e-mail, you should use ```(linha["THE NAME OF YOUR COLUMN"])```
