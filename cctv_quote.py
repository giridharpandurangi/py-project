import tkinter
from tkinter import ttk
import os
import openpyxl



def enter_data():
    #try:
       # camera_2mp = int(camera_number_2mp_entry.get())
    #except:
        #camera_2mp = 0
    camera_4mp = int(camera_number_4mp_entry.get())
    camera_ptz = int(camera_number_PTZ_entry.get())
    dvr_channel = dvr_channel_combobox.get()
    powersupply_channel = power_supply_combobox.get()
    hdd = hdd_label_entry.get()
    cable= cable_label_entry.get()
    rack=int(rack_label_entry.get())
    filepath = "E:\Onedrive\OneDrive - Pandurangi\Documents\py-project\data.xlsx"
    if not os.path.exists(filepath):
        workbook=openpyxl.Workbook()
        sheet = workbook.active
        #heading = ["2mp camera","4mp camera","PTZ camera","dvr","power supply","HDD","Cable","rack"]
        heading = ["ITEMS","QUANTITY","UNIT"]
        sheet.append(heading)
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    #sheet.append(["2MP camera",camera_2mp,"NO"])
    if not camera_2mp == 0 and camera_4mp == 0 and camera_ptz == 0:
        sheet.append(["2MP camera",camera_2mp,"NO"])
    elif not camera_4mp == 0 and camera_2mp == 0 and camera_ptz == 0:
        sheet.append(["4MP camera",camera_4mp,"NO"])
    else:
        sheet.append(["2MP camera",camera_2mp,"NO"])
        sheet.append(["4MP camera",camera_4mp,"NO"])
        sheet.append(["PTZ camera",camera_ptz,"NO"])
    sheet.append(["DVR channel",dvr_channel,"CHANNEL"])
    sheet.append(["power supply",powersupply_channel,"CHANNEL"])
    sheet.append(["HDD",hdd,"NO"])
    sheet.append(["Cable",cable,"METER"])
    if not rack == 0:
        sheet.append(["Rack",rack,"U"])
    elif rack == 0:
        workbook.save(filepath)
        
    workbook.save(filepath)

window = tkinter.Tk()
window.title("CCTV QUOTATION")

frame = tkinter.Frame(window)
frame.pack()

user_info_frame = tkinter.LabelFrame(frame, text="CCTV details")
user_info_frame.grid(row= 0, column=0,padx=20,pady=20)

camera_number_2mp_label  = tkinter.Label(user_info_frame, text="Number of 2MP cameras")
camera_number_2mp_label.grid(row=0, column=0)
camera_number_4mp_label  = tkinter.Label(user_info_frame, text="Number of 4MP cameras")
camera_number_4mp_label.grid(row=0, column=1)
camera_number_PTZ_label  = tkinter.Label(user_info_frame, text="Number of PTZ cameras")
camera_number_PTZ_label.grid(row=0, column=2)


camera_number_2mp_entry = tkinter.Entry(user_info_frame)
camera_number_2mp_entry.insert(0,"0")
camera_number_2mp_entry.grid(row=1,column=0)
camera_number_4mp_entry = tkinter.Entry(user_info_frame)
camera_number_4mp_entry.insert(0,"0")
camera_number_4mp_entry.grid(row=1,column=1)
camera_number_PTZ_entry = tkinter.Entry(user_info_frame)
camera_number_PTZ_entry.insert(0,"0")
camera_number_PTZ_entry.grid(row=1,column=2)





dvr_channel_label = tkinter.Label(user_info_frame, text="Dvr Channel")
dvr_channel_label.grid(row=2,column=0)
power_supply_label=tkinter.Label(user_info_frame, text="Power Supply Channel")
power_supply_label.grid(row=2,column=2)


dvr_channel_combobox= ttk.Combobox(user_info_frame, values=[4,8,16,32,64])
dvr_channel_combobox.grid(row=3,column=0)
power_supply_combobox= ttk.Combobox(user_info_frame,values=[4,8,16,32,64])
power_supply_combobox.grid(row=3,column=2)



cable_label=tkinter.Label(user_info_frame, text="cable in Meter")
rack_label=tkinter.Label(user_info_frame, text="Rack in U")
cable_label.grid(row=4,column=1)
rack_label.grid(row=4,column=2)

hdd_label=tkinter.Label(user_info_frame, text="HDD in TB")
hdd_label.grid(row=4,column=0)
hdd_label_entry=tkinter.Spinbox(user_info_frame,from_=1,to=12)
hdd_label_entry.grid(row=5,column=0)
cable_label_entry=tkinter.Entry(user_info_frame)
cable_label_entry.grid(row=5,column=1)
rack_label_entry=tkinter.Entry(user_info_frame)
rack_label_entry.grid(row=5,column=2)


button=tkinter.Button(user_info_frame, text="Create Excel File", command= enter_data)
button.grid(row=6,column=1, sticky="news", padx=50,pady=20)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



window.mainloop()