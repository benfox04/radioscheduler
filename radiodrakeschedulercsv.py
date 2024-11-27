import PySimpleGUI as sg
import openpyxl 

def open_window():
    layout = [
              [sg.Text("Hosts: ", size=(7,1)), sg.Input(key='input_hosts'), sg.Push()],
              [sg.Text("Year: ", size=(7,1)), sg.Input(key='input_year')],
              [sg.Text("Class: ", size=(7,1)), sg.Input(key='input_classname')],
              [sg.Text("Concept: ", size=(7,1)), sg.Input(key='input_concept')],
              [sg.Text("Features: ", size=(7,1)), sg.Input(key='input_features')],
              [sg.Text("Time: ", size=(7,1)), sg.Input(key='input_time')],
              [sg.Text("Date: ", size=(7,1)), sg.Input(key='input_date')],
              [sg.Text("Producer: ", size=(7,1)), sg.Input(key='input_producer')],
              [sg.Button("Submit", key='Submit')]]

    window = sg.Window("Booking Form", layout, modal=True, finalize=True)
    # window.maximize() 

    choice = None
    while True:
        event, values = window.read()
        if event == 'Submit': 
            count = int(sheet['A50'])
            hostName = str(values['input_hosts'])
            year = int(values['input_year'])
            classname = str(values['input_classname'])
            concept = str(values['input_concept'])
            form = str(values['input_features'])
            time = str(values['input_time'])
            date = str(values['input_date'])
            producer = str(values['input_producer'])
            if len(date) > 10 or len(date) < 9:
                sg.popup("Invalid year format. Please use the following formatting: 01/01/2024")
            # if time != "morning" or time != "Morning" or time != "afternoon" or time != "Afternoon":
                # sg.popup("Invalid time format. Please type 'morning' or 'afternoon'.")
            else: 
                sheet.updateRow(count, [hostName, year, classname, concept, form, time, date, producer])
                count = count + 1
                sheet['A50'] = count  
                sg.popup("That's all booked in! Remember to arrive in the " + time + " on " + date + "!")      
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
       
    window.close()

def open_window_two():
    layout = [
              [sg.Text("Enter the date you want to check: ", size=(20,1))], 
               [sg.Input(key='checking_date')],
              [sg.Button("Submit", key = "Submit")]]
    
    window = sg.Window("Checking...", layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == 'Submit':
            searchedDate = str(values['checking_date'])
            dateColumn = sheet.getColumn(7)
            for x in dateColumn: 
                if x == searchedDate: 
                    print("Found it!")
                    foundRow = sheet.getRow(dateColumn.index(x) + 1)
                    sg.popup(
                        "Hosts: " + foundRow[0] + "\n"
                        "Year Group: " + foundRow[1] + "\n"
                        "Class: " + foundRow[2] + "\n"
                        "Concept: " + foundRow[3] + "\n"
                        "Features: " + foundRow[4] + "\n"
                        "Time: " + foundRow[5] + "\n"
                        "Date: " + foundRow[6] + "\n"
                        "Producer: " + foundRow[7] + "\n")
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
       
    window.close()

def main():
    layout = [
        [sg.Text("Welcome to Radio Drake's Scheduler!")],
        [sg.Button("Book", key="book"), sg.Button("Check", key = "check"), sg.Button("Exit", key = "Exit")]]
    window = sg.Window("Main Window", layout)
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        if event == "book":
            open_window()
        if event == "check":
            open_window_two()
        
    window.close()

main()


