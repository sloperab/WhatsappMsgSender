import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import date
from time import sleep
import tkinter as tk
from webdriver_manager.chrome import ChromeDriverManager

# Load the Excel workbook
workbook = openpyxl.load_workbook("WhatsappBotExcel.xlsx", data_only=True)
sheet = workbook.active

def open_webpage():
    global driver
    # Initialize the Chrome WebDriver
    driver = webdriver.Firefox()
    driver.get('https://web.whatsapp.com/')

def send_messages():
    today = date.today()
    formatted_date = today.strftime("%d/%m/%Y")
    day = formatted_date[:2]

    # Iterate through Excel data and send messages
    for row in sheet.iter_rows(min_row=2, min_col=1):
        data_list = [str(cell.value) for cell in row]

        if len(data_list[4]) == 1:
            data_list[4] = '0' + data_list[4]
        elif len(data_list[4]) == 0:
            data_list[4] = ''

        if len(data_list[3]) == 1:
            data_list[3] = '0' + data_list[3]

        if data_list[4] == 'None':
            data_list[4] = ''

        if data_list[5] == 'None':
            data_list[5] = ''

        if data_list[0] != 'None':
            if int(day) == (int(data_list[1]) - 1):
                user = driver.find_element_by_xpath('//*[@id="side"]/div[1]/div/label/div/div[2]')
                user.send_keys('57' + data_list[2])
                user.send_keys(Keys.RETURN)
                chatbox = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
                chatbox.send_keys(
                    "Reminder: Physical therapy appointment tomorrow at " + (data_list[3] + ':' + data_list[4]) + ' ' + data_list[5]  + ". Please confirm your attendance. Thank you.")
                chatbox.send_keys(Keys.RETURN)
                sleep(3)
                print("Successfully Sent!")
            else:
                pass

# Create a Tkinter window
parent = tk.Tk()
frame = tk.Frame(parent)
frame.pack()

# Create buttons for opening WhatsApp and sending messages
open_button = tk.Button(frame,
                   text="Open WhatsApp",
                   command=open_webpage
                   )
open_button.pack(side=tk.LEFT)

send_button = tk.Button(frame,
                   text="Send Messages",
                   fg="green",
                   command=send_messages)
send_button.pack(side=tk.RIGHT)

# Create an exit button to close the application
exit_button = tk.Button(frame,
                   text="Exit",
                   fg="green",
                   command=parent.quit)
exit_button.pack(side=tk.BOTTOM)

# Start the Tkinter main loop
parent.mainloop()
