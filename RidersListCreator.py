import csv
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

from xlwt import Workbook

root = Tk()
root.title("BEM Riders List Converter")

fileName = ""
fileSave = ""
numberOfRiders = 0


def openFile():
    global fileName
    fileName = filedialog.askopenfilename(initialdir='/', title="Select a csv file",
                                          filetypes=(("CSV Files", ".csv"), ("All files", "*.*")))
    buttonSave.config(state=NORMAL)
    buttonOpen.config(state=DISABLED)


def saveFile():
    global fileSave
    fileSave = filedialog.asksaveasfilename(initialfile="RidersList", defaultextension=".xls",
                                            filetypes=[('Excel file', '.xls'), ('all files', '.*')],
                                            title="Choose location")
    buttonConvert.config(state=NORMAL)
    buttonSave.config(state=DISABLED)


def convert():
    global fileName
    global fileSave
    global numberOfRiders
    print(fileName)
    print(fileSave)

    try:
        with open(fileName) as f:
            readCSV = csv.reader(f, delimiter=',')
            data = list(readCSV)

            print(len(data))

            # sheet name

            sheet1 = wb.add_sheet('EM5_EXT')

            # sheet label on first line

            sheet1.write(0, 0, 'Licence_num')
            sheet1.write(0, 1, 'UCI_ID')
            sheet1.write(0, 2, 'UCIcode')
            sheet1.write(0, 3, 'FederationID')
            sheet1.write(0, 4, 'International Licence Code')
            sheet1.write(0, 5, 'Expiry_date')
            sheet1.write(0, 6, 'Licence_type')
            sheet1.write(0, 7, 'Dob')
            sheet1.write(0, 8, 'First_name')
            sheet1.write(0, 9, 'Surname')
            sheet1.write(0, 10, 'Sex')
            sheet1.write(0, 11, 'Emergency Contact Person')
            sheet1.write(0, 12, 'Emergency Contact Number')
            sheet1.write(0, 13, 'CLUB')
            sheet1.write(0, 14, 'State')
            sheet1.write(0, 15, 'UCI_Country')
            sheet1.write(0, 16, 'Class')
            sheet1.write(0, 17, 'Class2')
            sheet1.write(0, 18, 'Class3')
            sheet1.write(0, 19, 'Class4')
            sheet1.write(0, 20, 'Plate')
            sheet1.write(0, 21, 'Plate2')
            sheet1.write(0, 22, 'Plate3')
            sheet1.write(0, 23, 'Plate4')
            sheet1.write(0, 24, 'Ranking')
            sheet1.write(0, 25, 'Ranking2')
            sheet1.write(0, 26, 'Ranking3')
            sheet1.write(0, 27, 'Ranking4')
            sheet1.write(0, 28, 'Transponder')
            sheet1.write(0, 29, 'Transponder2')
            sheet1.write(0, 30, 'Transponder3')
            sheet1.write(0, 31, 'Transponder4')
            sheet1.write(0, 32, 'Tlabel')
            sheet1.write(0, 33, 'Tlabel2')
            sheet1.write(0, 34, 'Tlabel3')
            sheet1.write(0, 35, 'Tlabel4')
            sheet1.write(0, 36, 'Reference')
            sheet1.write(0, 37, 'Team_No')
            sheet1.write(0, 38, 'Team_No2')
            sheet1.write(0, 39, 'Team_No3')
            sheet1.write(0, 40, 'Team_No4')
            sheet1.write(0, 41, 'Sponsor')

            i = 1
            numberOfRiders = (len(data))
            progbar.config(mode="determinate", maximum=(len(data)), value=1)

            # fill sheet

            for row in data[1:len(data)]:
                sheet1.write(i, 0, row[1])  # Licence_num
                sheet1.write(i, 1, row[1])  # UCI_ID
                sheet1.write(i, 2, row[1])  # UCIcode
                sheet1.write(i, 3, row[1])  # FederationID
                sheet1.write(i, 4, row[1])  # International Licence Code
                sheet1.write(i, 5, row[5])  # Expiry date
                sheet1.write(i, 6, row[6])  # Licence type
                sheet1.write(i, 7, row[7])  # Date of Birth
                sheet1.write(i, 8, row[8])  # First Name
                row[9] = row[9].upper()
                sheet1.write(i, 9, row[9])  # Surname
                sheet1.write(i, 10, row[10])  # Sex
                sheet1.write(i, 11, row[0])  # Emergency Contact Person
                sheet1.write(i, 12, row[0])  # Emergency Contact Number
                if row[11] == "BMX &amp; 4X Team BRNO":
                    row[11] = "BMX 4X Team BRNO"
                if row[11] == "BMX &amp; 4X TEAM OLYMPUS":
                    row[11] = "BMX 4X TEAM OLYMPUS"
                sheet1.write(i, 13, row[11])  # Club
                sheet1.write(i, 14, row[12])  # State
                sheet1.write(i, 15, row[13])  # UCI_Country
                sheet1.write(i, 16, row[16])  # Class
                sheet1.write(i, 17, row[17])  # Class2
                sheet1.write(i, 18, row[18])  # Class3
                sheet1.write(i, 19, row[19])  # Class4
                sheet1.write(i, 20, row[20])  # Plate
                sheet1.write(i, 21, row[20])  # Plate2
                sheet1.write(i, 22, row[20])  # Plate3
                sheet1.write(i, 23, row[20])  # Plate4
                sheet1.write(i, 28, row[29])  # Transponder
                sheet1.write(i, 29, row[30])  # Transponder2
                sheet1.write(i, 32, row[33])  # Tlabel
                sheet1.write(i, 33, row[34])  # Tlabel2
                sheet1.write(i, 41, row[11])  # Sponsor - the same as club

                i = i + 1
                progbar.config(value=i)
                labelRiders.config(text="Converted was {} riders.".format(numberOfRiders))
            wb.save(fileSave)
            wb.close()



    except:
        print("Chyba v názvu souboru")


wb = Workbook()

labelTitle = ttk.Label(text="BEM Rides List Converter", font=(('Arial'), 22))
labelTitle.grid(row=0, column=1)

buttonOpen = ttk.Button(root, text="Open csv riders list", command=openFile)
buttonOpen.grid(row=1, column=0)

buttonSave = ttk.Button(root, text="Save Riders list as ...", command=saveFile)
buttonSave.grid(row=1, column=1)
buttonSave.config(state=DISABLED)

buttonConvert = ttk.Button(root, text="Convert", command=convert)
buttonConvert.grid(row=1, column=2)
buttonConvert.config(state=DISABLED)

progbar = ttk.Progressbar(root, orient=HORIZONTAL, length=400)
progbar.grid(row=2, column=0, columnspan=3)

labelRiders = ttk.Label(font=(('Arial'), 16), text="Converted was {} riders.".format(numberOfRiders))
labelRiders.grid(row=3, column=1)

# status = Label(root, text="(c)2019 David Průša, Asociace klubů BMX.", anchor=E)
# status.pack(side=BOTTOM, fill=X)

root.geometry("550x150+350+250")
root.mainloop()
