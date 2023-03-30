Whatsapp Appointment Reminder
=============================

This is a Python script that reads an Excel file with input and sends appointment reminders to contacts that meet a condition, using Whatsapp Web.

Dependencies
------------

The script uses the following Python libraries:

*   openpyxl
*   selenium
*   chromedriver-autoinstaller
*   datetime
*   tkinter

To install the required libraries, run `pip install -r requirements.txt` in your terminal or command prompt.

Usage
-----

1.  Clone or download the repository.
    
2.  Open the Excel file `WhatsappBotExcel.xlsx` and fill in the details for each appointment you want to send a reminder for. The file should have the following columns:
    
    *   `Name`: The name of the patient.
    *   `Date`: The date of the appointment, in the format dd/mm/yyyy.
    *   `Cellphone`: The phone number of the patient, including the country code.
    *   `Hour`: The hour of the appointment, in 12-hour format.
    *   `Min`: The minute of the appointment.
    *   `AM/PM`: The time of day of the appointment, either "AM" or "PM".
3.  Run the script `WhatsappBot.py` in your Python environment.
    
4.  The current condition requires the appointment being the day after the script is run.
    
5.  The script will then launch the Chrome browser and open Whatsapp Web. Use the GUI that appears to confirm that Whatsapp Web has logged in.
    
6.  Once logged in, the script will start sending reminders to the selected contacts. The reminder message will include the name of the patient, the date and time of the appointment, and a pre-written message.
    
    

Note
----

*   The script will send reminders to contacts only if the appointment is scheduled for the current day and time, and the condition specified by the user is met.
*   The script will not send reminders to contacts outside of the appointment schedule.
*   Tested only for personal use with Colombian phonenumbers.
