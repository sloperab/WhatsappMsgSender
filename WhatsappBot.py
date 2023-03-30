import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import date
from time import sleep
import tkinter as tk
from webdriver_manager.chrome import ChromeDriverManager


#filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

wrkbk = openpyxl.load_workbook("WhatsappBotExcel.xlsx",data_only=True)
sh = wrkbk.active


def abrir_pagina():
    global driver
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get('https://web.whatsapp.com/')

def send_messages():
    today = date.today()
    d1 = today.strftime("%d/%m/%Y")
    dia = d1[:2]
    # iterate through excel and display data
    for row in sh.iter_rows(min_row=2, min_col=1):
        lista = []
        for cell in row:
            lista.append(cell.value)
        lista = [str(i) for i in lista]

        if len(lista[4]) == 1:
            lista[4] = '0' + lista[4]
        elif len(lista[4]) == 0:
            lista[4] = ''
        if len(lista[3]) == 1:
            lista[3] = '0' + lista[3]
        if lista[4] == 'None':
            lista[4] = ''
        if lista[5] == 'None':
            lista[5] = ''
        if lista[0] != 'None':

            if int(dia) == (int(lista[1]) - 1):
                user = driver.find_element_by_xpath('//*[@id="side"]/div[1]/div/label/div/div[2]')
                user.send_keys('57' + lista[2])
                user.send_keys(Keys.RETURN)
                chatbox = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
                chatbox.send_keys(
                    "Recuerda cita de fisioterapia mañana " + (lista[3] + ':' + lista[4]) + ' ' + lista[5]  + " . Confirmar asistencia. Gracias.")
                chatbox.send_keys(Keys.RETURN)
                sleep(3)
                print("Successfully Sent!")
            else:
                pass

parent = tk.Tk()
frame = tk.Frame(parent)
frame.pack()

text_disp= tk.Button(frame,
                   text="Abrir WhatsApp",
                   command=abrir_pagina
                   )

text_disp.pack(side=tk.LEFT)

exit_button = tk.Button(frame,
                   text="QR scaneado",
                   fg="green",
                   command=send_messages)
exit_button.pack(side=tk.RIGHT)
exit_button2 = tk.Button(frame,
                   text="Exit",
                   fg="green",
                   command = quit)
exit_button2.pack(side=tk.BOTTOM)

parent.mainloop()
