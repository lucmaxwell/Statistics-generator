from tkinter import *
from tkinter import messagebox
import pandas as pd
import datetime
from datetime import timedelta

root = Tk()
root.iconbitmap(r'Icon.ico')
root.geometry("1340x430")
root.resizable(False, False)
root.title("Statistics Tracker for Bilingual Service Centers")
root.configure(bg = "#cdcdcd")

date_start = datetime.datetime.now() - timedelta(days = datetime.datetime.now().weekday())
date_finish = datetime.datetime.now() - timedelta(days = datetime.datetime.now().weekday() - 4)
date = date_start.strftime("%b") + " " + date_start.strftime("%d") + " – " + date_finish.strftime("%b") + " " + date_finish.strftime("%d")

data = pd.read_excel("Statistics Summary.xlsx")


def save_excel(): 
    button_info = []
    button_info.append(int(l0_entry.get()))
    button_info.append(int(l1_entry.get()))
    for i in range(22): 
        button_info.append(int(entry_list[i].get()))
    data[date] = button_info

    messagebox.showinfo(title = "Saved successfully", message = "Data saved successfully")

def click(entry_slot, flag):
    try:
        int(entry_slot.get())
        if flag:
            test_value = 0
            addorsub = 1

        else:
            test_value = 1
            addorsub = -1
                
        if int(entry_slot.get()) >= test_value:
            entry = int(entry_slot.get()) + addorsub
            entry_slot.delete(0, END)
            entry_slot.insert(0, entry)
        else:
            messagebox.showwarning(title = "Invalid entry", message = "Negative numbers are not a valid entry. ")
    except ValueError:
        messagebox.showwarning(title = "Invalid entry", message = "Please enter a positive integer. ")
        pass

spacer1 = Label(root, text = "     ", bg = "#cdcdcd")
spacer1.grid(row = 0, column = 4)
spacer2 = Label(root, text = " ", bg = "#cdcdcd")
spacer2.grid(row = 0, column = 9)


line = Label(root, text = "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", bg = "#cdcdcd")
line.grid(row = 1, column = 0, columnspan = 14)

l0_entry = Entry(root, width = 8, bd = 2, justify = "center", font = "TkDefault 10")
l0_entry.grid(row = 0, column = 2, ipady = 2, pady = 4, padx = 2)
l0_label = Label(root, text = "Services français: ", justify = "left", font = "TkDefault 11", bg = "#cdcdcd", anchor = "w", width = 20)
l0_label.grid(row = 0, column = 0, pady = 2)
l0_sub = Button(root, text = "–", command = lambda: click(l0_entry, False), justify = "center", height = 1, font = "TkDefault 10", width = 2, bg = "#d7d7d7")
l0_sub.grid(row = 0, column = 1, pady = 2)
l0_add = Button(root, text = "+", command = lambda: click(l0_entry, True), justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7")
l0_add.grid(row = 0, column = 3, pady = 2)

l1_entry = Entry(root, width = 8, bd = 2, justify = "center", font = "TkDefault 10")
l1_entry.grid(row = 0, column = 7, ipady = 2, pady = 4, padx = 2)
l1_label = Label(root, text = "          Services anglais: ", font = "TkDefault 11", bg = "#cdcdcd", anchor = "w", width = 20)
l1_label.grid(row = 0, column = 5, pady = 2)
l1_sub = Button(root, text = "–", command = lambda: click(l1_entry, False), justify = "center", height = 1, font = "TkDefault 10", width = 2, bg = "#d7d7d7")
l1_sub.grid(row = 0, column = 6, pady = 2)
l1_add = Button(root, text = "+", command = lambda: click(l1_entry, True), justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7")
l1_add.grid(row = 0, column = 8, pady = 2)

write_excel_button = Button(root, text = "Save to excel", command = save_excel, justify = "center", height = 1, font = "TkDefault 11", width = 72, bg = "#d7d7d7")
write_excel_button.place(x = 670, y = 6)

Services = (
"Réception/aiguillage - autre (qui ne figure pas dans les catégories ci-dessous)", 
"Agent(e) d'information", 
"Recherche d’emploi (CV, Lettre de Motivation, etc.)", 
"Tâches administratives pour les partenaires (réservation de la salle de conf., etc. )", 
"Photocopies/Télécopies/Scanneur", 
"Ordinateurs d'accès public",
"Appui technique au client à l’ordinateur", 
"Téléphone public", 
"Salle de conférence(nombre de participants)", 
"Commissaire d'assermentation", 
"Notaire public", 
"Service Canada", 
"Gouvernement du Canada(CRA, Immigration, etc.)", 
"Aide à l'emploi et au revenu (EIA)",
"Ville de Winnipeg", 
"ORS (Office régional de santé)", 
"Carte Assurance santé", 
"Logement (Rent Assist, Manitoba Housing)", 
"Services à la famille", 
"Items promos et dépliants des CSB (clé USB, écouteurs etc.)", 
"Bureau de l'état civil/Vital Statistics MB (certificat de naissance, etc.)]",
"Autre", 
)

service_num = 0
row_num = 2
label_list = []
for i in range(round((len(Services))/2)):
   label_list.append(Label(root, text = Services[i], justify = "left", font = "TkDefault 11", bg = "#cdcdcd", anchor = "w", width = 60))
   label_list[i].grid(row = row_num, column = 0, pady = 2, columnspan = 6)
   row_num += 1
   service_num += 1
   

row_num = 2
for i in range((service_num), len(Services)): 
    label_list.append(Label(root, text = Services[i], justify = "left", font = "TkDefault 11", bg = "#cdcdcd", anchor = "w", width = 60))
    label_list[i].grid(row = row_num, column = 10, pady = 2)
    row_num += 1


row_num = 2
inc_button = []
for i in range(round((len(Services))/2)): 
	inc_button.append(Button(root, text = "+", command = lambda i=i: click(entry_list[i], True),  justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7"))
	inc_button[i].grid(row = row_num, column = 8, pady = 2)
	row_num += 1

row_num = 2
for i in range((service_num),len(Services)): 
    inc_button.append(Button(root, text = "+", command = lambda i=i: click(entry_list[i], True),  justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7"))
    inc_button[i].grid(row = row_num, column = 13, pady = 2)
    row_num += 1

row_num = 2
entry_list = []
for i in range(round((len(Services))/2)): 
	entry_list.append(Entry(root, width = 8, bd = 2, justify = "center", font = "TkDefault 10"))
	entry_list[i].grid(row = row_num, column = 7, ipady = 2, pady = 4, padx = 2)
	row_num += 1

row_num = 2
for i in range((service_num), len(Services)): 
    entry_list.append(Entry(root, width = 8, bd = 2, justify = "center", font = "TkDefault 10"))
    entry_list[i].grid(row = row_num, column = 12, ipady = 2, pady = 4, padx = 2)
    row_num += 1

row_num = 2
dec_button = []
for i in range(round((len(Services))/2)): 
	dec_button.append(Button(root, text = "-", command = lambda i=i: click(entry_list[i], False), justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7"))
	dec_button[i].grid(row = row_num, column = 6, pady = 2)
	row_num += 1

row_num = 2
for i in range((service_num),len(Services)): 
    dec_button.append(Button(root, text = "-", command = lambda i=i: click(entry_list[i], False), justify = "center", font = "TkDefault 10", width = 2, bg = "#d7d7d7"))
    dec_button[i].grid(row = row_num, column = 11, pady = 2)
    row_num += 1


column_titles = list(data.columns)
if date in column_titles:  
    load = data[date].to_list()
    for i in range(0, 22):
        entry_list[i].delete(0, END)
        entry_list[i].insert(0, load[i+2])
    l0_entry.delete(0, END)
    l0_entry.insert(0, load[0])
    l1_entry.delete(0, END)
    l1_entry.insert(0, load[1])
   
else: 
    default = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    data[date] = default
    load = default
    for i in range(0, 22):
        entry_list[i].delete(0, END)
        entry_list[i].insert(0, load[i+2])
    l0_entry.delete(0, END)
    l0_entry.insert(0, load[0])
    l1_entry.delete(0, END)
    l1_entry.insert(0, load[1])
        


root.mainloop()

data.to_excel("Statistics summary.xlsx", index = False)