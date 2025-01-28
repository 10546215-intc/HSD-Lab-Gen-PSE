import tkinter as tk
from tkinter import ttk
from tkinter import *
import tkinter.messagebox
import tkinter.font as tkFont
from ttkwidgets import tooltips
import datetime
#from datetime import datetime, timedelta
from datetime import timedelta
import time
import csv
import json
import requests
from requests_kerberos import HTTPKerberosAuth
import urllib3
import os
 
# this is to ignore the ssl insecure warning as we are passing in 'verify=false'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
os.environ["no_proxy"] = "127.0.0.1,localhost,intel.com"
#headers = { 'Content-type': 'application/json' }
 
vernum = "0.977" #application version number
configver = "" #config version number
dynamicver = "" #dynamic version number
staticver = "" #static version number
 
menu_color = "#e4e5ee"
option_color = "#D4D6E4"
 
 
class App():
    def __init__(self, root):
        root.title("HSD-Lab Gen - Version " + str(vernum))
        root.eval("tk::PlaceWindow . center")
 
        #setting window size
        width=492
        height=680
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
 
        #Load Icon
        def UpdateIcon():
            try:
                p1 = PhotoImage(file=icon)
                root.iconphoto(False, p1)
            except:
                print('Could not find icon.')
        #icon = 'dependencies/B.png' #Production Mode
        #UpdateIcon()
 
        #Set Size for GUI
        cv = Canvas(root,width=892,height=680)
        cv.pack()
        cv.create_rectangle(2,2,485,90,fill=menu_color)
        cv.create_line(165,2,165,90)
        cv.create_line(326,2,326,90)
 
        #ProgressSuccess Label
        Label_ProgressSuccess=tk.Label(root)
        ft = tkFont.Font(family='Times',size=14)
        Label_ProgressSuccess["font"] = ft
        Label_ProgressSuccess["fg"] = "#333333"
        Label_ProgressSuccess["justify"] = "left"
        Label_ProgressSuccess["text"] = ""
        Label_ProgressSuccess.place(x=0,y=340,width=492,height=20)
 
        #Validate field aginist HSD-ES DB
        hsdcheck = True
        def validateField(field, value):
            isvalid = False
            headers = { 'Content-type': 'application/json' }
            Label_ProgressSuccess["text"]= "Validating Menu Options with HSD-ES DB"
            root.after(1,Label_ProgressSuccess.update())
            url_validate = f'https://hsdes-api.intel.com/rest/query/autocomplete/support/services_sys_val/{field}'
            try:
                response = requests.get(url_validate, verify=False, auth=HTTPKerberosAuth(), headers = headers)
            except:
                hsdcheck=False
    
            for i in response.json()["data"]:
                if value.lower() in i[field]:
                    isvalid = True
                #print(value, ':', isvalid)Validate 
            return isvalid
 
        # Create variables for the project and site option selected, etc.
        project_option_selected = tk.StringVar(root)
        site_option_selected = tk.StringVar(root)
        customer_option_selected = tk.StringVar(root)
        notify_option_selected = tk.StringVar(root)
 
        project_options=['']
        sites_options=['']
        customer_options=['']
        notify_options=['']
        lab_dict={}
        notify_dict={}
        lead_dict={}
        MS_options=['']
        l_MS_options=['']
 
        #Close GUI and Python
        def close():
            exec
            root.destroy()
 
        #Load config CSV file
        try:
            with open(r"dependencies/config.csv", encoding="utf8") as f:
                csv_config = csv.DictReader(f)
                print('\nloading config and verifing options')
                for line in csv_config:
                    if(line['Version']):
                        configver = line['Version']
    ##******## This seems to look at config file for Dynamic value, is hardcoded to "New"
    ##******## If has value then checks HSD to see if field values exist if passes the writes to ticket variables
                    if(line['Dynamic']):
                        dynamic_CSV_file = line['Dynamic']
 
                    if(line['Program']):
 
                        if validateField('services_sys_val.support.program',line['Program']):
                            project_options.append(line['Program'])
 
                    if(line['Site']):
                       if validateField('support.site',line['Site']):
                           sites_options.append(line['Site'])
                           lab_dict.update({line['Site']: line['Lab']})
                           notify_dict.update({line['Site']: line['Notify']})
                           lead_dict.update({line['Site']: line['Customer']})
 
                    if(line['Customer']):
                       customer_options.append(line['Customer'])
 
                    if(line['Notify']):
                       notify_options.append(line['Notify'])
 
            config_open=True
        except:
            config_open=False
 
        print('config_open= '+ str(config_open))
        #dynamic_CSV_file = "VV" #Options: 'Original' 'PO' 'VV' 'NEW'
        dynamic_CSV_file = 'NEW'
 
        Original_dynamic_vals = {}
        PO_dynamic_vals = {}
        PO_dict_exist = True
        VV_dynamic_vals = {}
        VV_dict_exist = True
        dynamic_vals = {}
 
        #Load dynamic CSV file
        print('Load NEW Track dynamic CSV File')
        ### NEW Track CSV Code for file dynamic_vals.csv ###
        try:
            with open("dependencies/dynamic_vals_new.csv", encoding="utf8") as data_file:
                print('\nloading NEW track dynamic_vals')
                data = csv.reader(data_file)
                dynamic_headers = next(data)[1:]            
                for row in data:
                    temp_dict = {}
                    name = row[0]
                    values = []

                    for x in row[1:]:
                        values.append(x)

                    for i in range(len(values)):
                        if values[i]:
                            temp_dict[dynamic_headers[i]] = values[i]

                    dynamic_vals[name] = temp_dict
                    
            with open("dependencies/dynamic_vals_new.csv", encoding="utf8") as f:
                csv_dynamic_= csv.DictReader(f)
                for line in csv_dynamic_:
                    if(line['Version']):                        
                        dynamicver = line['Version']
                        print('dynamicver: ', dynamicver)

                for key in dynamic_vals.keys():
                        l_MS_options.append(key[0] + '.' + key[1])

                MS_options=sorted([*set(l_MS_options)])

            dynamic_vals_open=True
        except:
            print('Did not load dynamic_vals_new')
            dynamic_vals_open=False
 
 
        #Load static CSV file
        static_vals = {}
        try:
            with open("dependencies/static_vals.csv", encoding="utf8") as f:
                print('\nloading static_vals')
                csv_static_vals = csv.DictReader(f)
 
                for line in csv_static_vals:
                    if(line['Version']):
                        staticver = line['Version']
 
                    static_vals[line['Item']] = line['Value']
            static_vals_open=True
        except:
            static_vals_open=False
        print('static_vals_open= ' + str(static_vals_open))
 
        #CSV file load error handeling
        if hsdcheck:
            if not config_open or not dynamic_vals_open or not static_vals_open:
                FileOpenMessage="Can not find file(s):\n"
 
                if not config_open:
                    FileOpenMessage = FileOpenMessage + '\ndependencies/config.csv'
 
                if not dynamic_vals_open:
                    FileOpenMessage = FileOpenMessage + '\ndependencies/dynamic_vals.csv'
 
                if not static_vals_open:
                    FileOpenMessage = FileOpenMessage + '\ndependencies/static_vals.csv'
 
                FileOpenMessage = FileOpenMessage + '\n\nWill now exit ' + str(root.title())
                print(FileOpenMessage)
                tkinter.messagebox.showwarning('File Not Found',FileOpenMessage)
                close()
        else:
            tkinter.messagebox.showwarning('Connection Error','Not connected to HSDE-ES DB.')
            close()
 
        clipboard = static_vals['createClipboard']
 
        pre_production = static_vals['pre_production']
        print('pre_production =', pre_production)
 
        if pre_production == "TRUE":
            url = 'https://hsdes-api-pre.intel.com/rest/article'
            linkUrl = 'https://hsdes-pre.intel.com/appstore/article/#/'
            icon = 'dependencies/Y.png'
            print('pre_production mode')
        else:
            url = 'https://hsdes-api.intel.com/rest/article'
            linkUrl = 'https://hsdes.intel.com/appstore/article/#/'
            icon = 'dependencies/B.png'
            print('production mode')
        UpdateIcon()
 
        #Create List for WorkWeek OptionMenu
        i = 1
        WorkWeekList = []
        for i in range(53):
            WorkWeekList.insert(i,i)
            i += 1
 
        #Get the Date
 ##       from datetime import date
        today = datetime.date.today()
        iso_calendar = today.isocalendar()
        WorkWeek = iso_calendar[1]
 #       WorkWeek = today.isocalendar().week
        year = iso_calendar[0]
#        year = today.year
        
 
        WorkWeekValue_Inside = tk.StringVar(root)
        YearValue_Inside = tk.StringVar(root)
 
        #Create List for Year OptionMenu
        i = 1
        YearList = []
        for i in range(2021,2051):
            YearList.insert(i,i)
            i += 1
 
        project_options_267=ttk.OptionMenu(root,project_option_selected,*project_options)
        project_options_267['tooltip'] = 'Choose program / project.'
        project_options_267.place(x=10,y=33,width=150,height=25)
 
        site_options267=ttk.OptionMenu(root,site_option_selected,*sites_options)
        site_options267['tooltip'] = 'Choose Site.'
        site_options267.place(x=168,y=33,width=150,height=25)
        
        Option_101=ttk.OptionMenu(root,WorkWeekValue_Inside,*WorkWeekList)
        Option_101['tooltip'] = 'Choose todays known work week for power on.'
        Option_101.place(x=400,y=33,width=80,height=25)
        WorkWeekValue_Inside.set(WorkWeek)
 
        Option_403=ttk.OptionMenu(root,YearValue_Inside,*YearList)
        Option_403['tooltip'] = 'Choose todays known year for power on.'
        Option_403.place(x=400,y=63,width=80,height=25)
        YearValue_Inside.set(year)
 
        def AdvancedCheck():
                if var101.get() == 1:
                    print('Awake')
                    MS_option1.configure(state="enabled")
                    MS_option2.configure(state="enabled")
                    MS_option3.configure(state="enabled")
                    MS_option4.configure(state="enabled")
                    MS_option5.configure(state="enabled")
                    MS_option6.configure(state="enabled")
                    MS_option7.configure(state="enabled")
                    MS_option8.configure(state="enabled")
                    MS_option9.configure(state="enabled")
 
                    entry_1.configure(state= "normal")
                    entry_2.configure(state= "normal")
                    entry_3.configure(state= "normal")
                    entry_4.configure(state= "normal")
                    entry_5.configure(state= "normal")
                    entry_6.configure(state= "normal")
                    entry_7.configure(state= "normal")
                    entry_8.configure(state= "normal")
                    entry_9.configure(state= "normal")
 
                    button_clear1["state"] = "normal"
                    button_clear2["state"] = "normal"
                    button_clear3["state"] = "normal"
                    button_clear4["state"] = "normal"
                    button_clear5["state"] = "normal"
                    button_clear6["state"] = "normal"
                    button_clear7["state"] = "normal"
                    button_clear8["state"] = "normal"
                    button_clear9["state"] = "normal"
 
                elif var101.get() == 0:
                    print('Sleeping')
                    MS_option1.configure(state="disabled")
                    MS_option2.configure(state="disabled")
                    MS_option3.configure(state="disabled")
                    MS_option4.configure(state="disabled")
                    MS_option5.configure(state="disabled")
                    MS_option6.configure(state="disabled")
                    MS_option7.configure(state="disabled")
                    MS_option8.configure(state="disabled")
                    MS_option9.configure(state="disabled")
 
                    entry_1.configure(state= "disabled")
                    entry_2.configure(state= "disabled")
                    entry_3.configure(state= "disabled")
                    entry_4.configure(state= "disabled")
                    entry_5.configure(state= "disabled")
                    entry_6.configure(state= "disabled")
                    entry_7.configure(state= "disabled")
                    entry_8.configure(state= "disabled")
                    entry_9.configure(state= "disabled")
                    button_clear1["state"] = "disabled"
 
                    button_clear1["state"] = "disabled"
                    button_clear2["state"] = "disabled"
                    button_clear3["state"] = "disabled"
                    button_clear4["state"] = "disabled"
                    button_clear5["state"] = "disabled"
                    button_clear6["state"] = "disabled"
                    button_clear7["state"] = "disabled"
                    button_clear8["state"] = "disabled"
                    button_clear9["state"] = "disabled"
                else:
                    messagebox.showerror('OOPS', 'Something went wrong!')
 
 
 
        Label_optional_MS=tk.Label(root)
        Label_optional_MS["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=8, weight='bold')
        Label_optional_MS["font"] = ft
        Label_optional_MS["fg"] = "#333333"
        Label_optional_MS["justify"] = "left"
        Label_optional_MS["text"] = "Milestone"
        Label_optional_MS["relief"] = "flat"
        Label_optional_MS.place(x=497,y=118,width=55,height=15)
 
        Label_optional_text=tk.Label(root)
        Label_optional_text["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=8, weight='bold')
        Label_optional_text["font"] = ft
        Label_optional_text["fg"] = "#333333"
        Label_optional_text["justify"] = "left"
        Label_optional_text["text"] = "Text to append to ticket title."
        Label_optional_text["relief"] = "flat"
        Label_optional_text.place(x=557,y=118,width=250,height=15)
 
        self.varSelectAll=tk.IntVar(root)
        self.CheckBox_SelectAll=tk.Checkbutton(root)
        self.CheckBox_SelectAll["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=10)
        self.CheckBox_SelectAll["font"] = ft
        self.CheckBox_SelectAll["fg"] = "#333333"
        self.CheckBox_SelectAll["justify"] = "left"
        self.CheckBox_SelectAll["text"] = "Select All"
        self.CheckBox_SelectAll.place(x=10,y=120,width=80,height=25)
        self.CheckBox_SelectAll["offvalue"] = "0"
        self.CheckBox_SelectAll["onvalue"] = "1"
        self.CheckBox_SelectAll["command"] = self.CheckBox_SelectAll_command
        self.CheckBox_SelectAll["variable"] = self.varSelectAll
        self.CheckBox_SelectAll.select()
 
 
        #print(dynamic_vals)
        print('')
 
        try:
            var11=tk.IntVar(root)
            self.CheckBox_11=tk.Checkbutton(root)
            self.CheckBox_11["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_11["font"] = ft
            self.CheckBox_11["fg"] = "#333333"
            self.CheckBox_11["justify"] = "left"
            self.CheckBox_11["text"] = dynamic_vals["112"]["cb_title"]
            self.CheckBox_11.place(x=10,y=160,width=490,height=25)
            self.CheckBox_11["offvalue"] = "0"
            self.CheckBox_11["onvalue"] = "1"
            self.CheckBox_11["variable"] = var11
            self.CheckBox_11.select()
        except:
            print("No 1.1 Milestone")

        try:
            var21=IntVar()
            self.CheckBox_21=tk.Checkbutton(root)
            self.CheckBox_21["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_21["font"] = ft
            self.CheckBox_21["fg"] = "#333333"
            self.CheckBox_21["justify"] = "left"
            self.CheckBox_21["text"] = dynamic_vals["212"]["cb_title"]
            self.CheckBox_21.place(x=10,y=190,width=490,height=25)
            self.CheckBox_21["offvalue"] = "0"
            self.CheckBox_21["onvalue"] = "1"
            self.CheckBox_21["variable"] = var21
            self.CheckBox_21.select()
        except:
            print("N0 2.1 Milestone")
    
        try:
            var22=IntVar()
            self.CheckBox_22=tk.Checkbutton(root)
            self.CheckBox_22["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_22["font"] = ft
            self.CheckBox_22["fg"] = "#333333"
            self.CheckBox_22["justify"] = "left"
            self.CheckBox_22["text"] = dynamic_vals["224"]["cb_title"]
            self.CheckBox_22.place(x=10,y=220,width=490,height=25)
            self.CheckBox_22["offvalue"] = "0"
            self.CheckBox_22["onvalue"] = "1"
            self.CheckBox_22["variable"] = var22
            self.CheckBox_22.select()
        except:
            print("N0 2.2 Milestone")
    
        try:
            var31=IntVar()
            self.CheckBox_31=tk.Checkbutton(root)
            self.CheckBox_31["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_31["font"] = ft
            self.CheckBox_31["fg"] = "#333333"
            self.CheckBox_31["justify"] = "left"
            self.CheckBox_31["text"] = dynamic_vals["311"]["cb_title"]
            self.CheckBox_31.place(x=10,y=250,width=490,height=25)
            self.CheckBox_31["offvalue"] = "0"
            self.CheckBox_31["onvalue"] = "1"
            self.CheckBox_31["variable"] = var31
            self.CheckBox_31.select()
        except:
            print("N0 3.1 Milestone")
    
        try:
            var32=IntVar()
            self.CheckBox_32=tk.Checkbutton(root)
            self.CheckBox_32["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_32["font"] = ft
            self.CheckBox_32["fg"] = "#333333"
            self.CheckBox_32["justify"] = "left"
            self.CheckBox_32["text"] = dynamic_vals["321"]["cb_title"]
            self.CheckBox_32.place(x=10,y=280,width=490,height=25)
            self.CheckBox_32["offvalue"] = "0"
            self.CheckBox_32["onvalue"] = "1"
            self.CheckBox_32["variable"] = var32
            self.CheckBox_32.select()
        except:
            print("N0 3.2 Milestone")

        try:
            var41=IntVar()
            self.CheckBox_41=tk.Checkbutton(root)
            self.CheckBox_41["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_41["font"] = ft
            self.CheckBox_41["fg"] = "#333333"
            self.CheckBox_41["justify"] = "left"
            self.CheckBox_41["text"] = dynamic_vals["411"]["cb_title"]
            self.CheckBox_41.place(x=10,y=310,width=490,height=25)
            self.CheckBox_41["offvalue"] = "0"
            self.CheckBox_41["onvalue"] = "1"
            self.CheckBox_41["variable"] = var41
            self.CheckBox_41.select()
        except:
            print("N0 4.1 Milestone")
        try:
            var42=IntVar()
            self.CheckBox_42=tk.Checkbutton(root)
            self.CheckBox_42["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_42["font"] = ft
            self.CheckBox_42["fg"] = "#333333"
            self.CheckBox_42["justify"] = "left"
            self.CheckBox_42["text"] = dynamic_vals["421"]["cb_title"]
            self.CheckBox_42.place(x=10,y=340,width=490,height=25)
            self.CheckBox_42["offvalue"] = "0"
            self.CheckBox_42["onvalue"] = "1"
            self.CheckBox_42["variable"] = var42
            self.CheckBox_42.select()
        except:
            print("N0 4.2 Milestone")
    
        try:
            var43=IntVar()
            self.CheckBox_43=tk.Checkbutton(root)
            self.CheckBox_43["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_43["font"] = ft
            self.CheckBox_43["fg"] = "#333333"
            self.CheckBox_43["justify"] = "left"
            self.CheckBox_43["text"] = dynamic_vals["431"]["cb_title"]
            self.CheckBox_43.place(x=10,y=370,width=490,height=25)
            self.CheckBox_43["offvalue"] = "0"
            self.CheckBox_43["onvalue"] = "1"
            self.CheckBox_43["variable"] = var43
            self.CheckBox_43.select()
        except:
            print("N0 4.3 Milestone")
    
        try:
            var51=IntVar()
            self.CheckBox_51=tk.Checkbutton(root)
            self.CheckBox_51["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_51["font"] = ft
            self.CheckBox_51["fg"] = "#333333"
            self.CheckBox_51["justify"] = "left"
            self.CheckBox_51["text"] = dynamic_vals["511"]["cb_title"]
            self.CheckBox_51.place(x=10,y=400,width=490,height=25)
            self.CheckBox_51["offvalue"] = "0"
            self.CheckBox_51["onvalue"] = "1"
            self.CheckBox_51["variable"] = var51
            self.CheckBox_51.select()
        except:
            print("N0 5.1 Milestone")
    
        try:
            var61=IntVar()
            self.CheckBox_61=tk.Checkbutton(root)
            self.CheckBox_61["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_61["font"] = ft
            self.CheckBox_61["fg"] = "#333333"
            self.CheckBox_61["justify"] = "left"
            self.CheckBox_61["text"] = dynamic_vals["611"]["cb_title"]
            self.CheckBox_61.place(x=10,y=430,width=490,height=25)
            self.CheckBox_61["offvalue"] = "0"
            self.CheckBox_61["onvalue"] = "1"
            self.CheckBox_61["variable"] = var61
            self.CheckBox_61.select()
        except:
            print("N0 6.1 Milestone")
    
        try:
            var62=IntVar()
            self.CheckBox_62=tk.Checkbutton(root)
            self.CheckBox_62["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_62["font"] = ft
            self.CheckBox_62["fg"] = "#333333"
            self.CheckBox_62["justify"] = "left"
            self.CheckBox_62["text"] = dynamic_vals["621"]["cb_title"]
            self.CheckBox_62.place(x=10,y=460,width=490,height=25)
            self.CheckBox_62["offvalue"] = "0"
            self.CheckBox_62["onvalue"] = "1"
            self.CheckBox_62["variable"] = var62
            self.CheckBox_62.select()
        except:
            print("N0 6.2 Milestone")
    
        try:
            var71=IntVar()
            self.CheckBox_71=tk.Checkbutton(root)
            self.CheckBox_71["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_71["font"] = ft
            self.CheckBox_71["fg"] = "#333333"
            self.CheckBox_71["justify"] = "left"
            self.CheckBox_71["text"] = dynamic_vals["711"]["cb_title"]
            self.CheckBox_71.place(x=10,y=490,width=490,height=25)
            self.CheckBox_71["offvalue"] = "0"
            self.CheckBox_71["onvalue"] = "1"
            self.CheckBox_71["variable"] = var71
            self.CheckBox_71.select()
        except:
            print("N0 7.1 Milestone")
    
    
        try:
            var81=IntVar()
            self.CheckBox_81=tk.Checkbutton(root)
            self.CheckBox_81["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_81["font"] = ft
            self.CheckBox_81["fg"] = "#333333"
            self.CheckBox_81["justify"] = "left"
            self.CheckBox_81["text"] = dynamic_vals["811"]["cb_title"]
            self.CheckBox_81.place(x=10,y=520,width=490,height=25)
            self.CheckBox_81["offvalue"] = "0"
            self.CheckBox_81["onvalue"] = "1"
            self.CheckBox_81["variable"] = var81
            self.CheckBox_81.select()
        except:
            print("N0 8.1 Milestone")

        try:
            var91=IntVar()
            self.CheckBox_91=tk.Checkbutton(root)
            self.CheckBox_91["anchor"] = "w"
            ft = tkFont.Font(family='Times',size=10)
            self.CheckBox_91["font"] = ft
            self.CheckBox_91["fg"] = "#333333"
            self.CheckBox_91["justify"] = "left"
            self.CheckBox_91["text"] = dynamic_vals["911"]["cb_title"]
            self.CheckBox_91.place(x=10,y=550,width=490,height=25)
            self.CheckBox_91["offvalue"] = "0"
            self.CheckBox_91["onvalue"] = "1"
            self.CheckBox_91["variable"] = var91
            self.CheckBox_91.select()
        except:
            print("No 9.1 Milestone")
        ###/NEW
 
        var101=IntVar()
        self.CheckBox_101=tk.Checkbutton(root)
        self.CheckBox_101["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=10)
        self.CheckBox_101["font"] = ft
        self.CheckBox_101["fg"] = "#333333"
        self.CheckBox_101["justify"] = "left"
        self.CheckBox_101["text"] = "Optional: Add additional tickets per milestone."
        self.CheckBox_101.place(x=495,y=98,width=490,height=20)
        self.CheckBox_101["variable"] = var101
        self.CheckBox_101["command"] = AdvancedCheck
        self.CheckBox_101.setvar = False
        
        #Program Label
        Label_Program=tk.Label(root)
        Label_Program["anchor"] = "c"
        ft = tkFont.Font(family='Times',size=14)
        Label_Program["font"] = ft
        Label_Program["fg"] = "#333333"
        Label_Program["justify"] = "left"
        Label_Program["text"] = "Program"
        Label_Program["background"] = menu_color
        Label_Program.place(x=10,y=5,width=150,height=25)
 
        #Site Label
        Label_Site=tk.Label(root)
        Label_Site["anchor"] = "c"
        ft = tkFont.Font(family='Times',size=14)
        Label_Site["font"] = ft
        Label_Site["fg"] = "#333333"
        Label_Site["justify"] = "center"
        Label_Site["text"] = "Site"
        Label_Site["background"] = menu_color
        Label_Site.place(x=168,y=5,width=150,height=25)
 
        #Power On Date Label
        Label_PODate=tk.Label(root)
        Label_PODate["anchor"] = "c"
        ft = tkFont.Font(family='Times',size=14)
        Label_PODate["font"] = ft
        Label_PODate["fg"] = "#333333"
        Label_PODate["justify"] = "left"
        Label_PODate["text"] = "Power On Date"
        Label_PODate["background"] = menu_color
        Label_PODate.place(x=330,y=5,width=150,height=25)
 
        #WorkWeek Label
        Label_WorkWeek=tk.Label(root)
        Label_WorkWeek["anchor"] = "e"
        ft = tkFont.Font(family='Times',size=10)
        Label_WorkWeek["font"] = ft
        Label_WorkWeek["fg"] = "#333333"
        Label_WorkWeek["justify"] = "right"
        Label_WorkWeek["text"] = "Work Week"
        Label_WorkWeek["background"] = menu_color
        Label_WorkWeek.place(x=330,y=33,width=70,height=25)
 
        #Year Label
        Label_Year=tk.Label(root)
        Label_Year["anchor"] = "c"
        ft = tkFont.Font(family='Times',size=10)
        Label_Year["font"] = ft
        Label_Year["fg"] = "#333333"
        Label_Year["justify"] = "right"
        Label_Year["text"] = "Year"
        Label_Year["background"] = menu_color
        Label_Year.place(x=370,y=63,width=30,height=25)
        
        #Milestones Label
        Label_MS=tk.Label(root)
        Label_MS["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=14)
        Label_MS["font"] = ft
        Label_MS["fg"] = "#333333"
        Label_MS["justify"] = "left"
        Label_MS["text"] = dynamic_CSV_file+" Milestones"
        Label_MS["relief"] = "flat"
        Label_MS.place(x=10,y=95,width=186,height=30)
 
        #Progress Label
        Label_ProgressSuccess.place(x=30,y=635,width=246,height=20)
        Label_ProgressSuccess["text"] = "Ready"
        root.after(1,Label_ProgressSuccess.update())
 
 
        def saveClipboard(my_string):
            if clipboard == 'TRUE':
                text_file = open(r'dependencies/Clipboard.txt', 'w')
                text_file.write(my_string)
                text_file.close()
                command = 'clip < dependencies/Clipboard.txt'
                os.system(command)
 
        total_tickets = 0
        ticket_number = 0
 
 
        def postnewHSD(fields):
            headers = { 'Content-type': 'application/json' }
            subject = "support"
            tenant = "services_sys_val"
 
            title = fields["title"]
            description = fields["description"]
            service_type = fields["service_type"]
            service_sub_type = fields["service_sub_type"]
            lab_org = fields["lab_org"]
            category = fields["category"]
            component = fields["component"]
            priority = fields["priority"]
            site = fields["site"]
            notify = fields["notify"]
            org_unit = fields["org_unit"]
            customer_contact = fields["customer_contact"]
            program = fields["program"]
            required_by_milestone = fields["required_by_milestone"]
            survey_comment = fields["survey_comment"]
            send_mail = "false"
 
            if pre_production == "TRUE":
                lab = "JF4-4108 (Debug Lab)"
                print('Pre-Perduction Mode')
                payload = {
                    "subject": subject,
                    "tenant": tenant,
                    "fieldValues": [
                        {
                        "title": title
                        },
                        {
                        "description": description
                        },
                        {
                        "services_sys_val.support.service_type": service_type
                        },
                        {
                        "services_sys_val.support.service_sub_type": service_sub_type
                        },
                        {
                        "services_sys_val.support.lab_org": lab_org
                        },
                        {
                        "services_sys_val.support.category": category
                        },
                        {
                        "component": component
                        },
                        {
                        "priority": priority
                        },
                        {
                        "support.customer_contact": customer_contact
                        },
                        {
                        "support.site": site
                        },
                        {
                        "services_sys_val.support.lab": lab
                        },
                        {
                        "notify": notify
                        },
                        {
                        "services_sys_val.support.org_unit": org_unit
                        },
                        {
                        "services_sys_val.support.program": program
                        },
                        {
                        "send_mail": send_mail
                        },
                        {
                        "services_sys_val.support.required_by_milestone": required_by_milestone
                        },
                        {
                        "services_sys_val.support.survey_comment": survey_comment
                        }
                    ]
                    }
            else:
                milestone_eta = fields["milestone_eta"]
                print('Production Mode')
                lab = fields['lab']
                payload = {
                    "subject": subject,
                    "tenant": tenant,
                    "fieldValues": [
                        {
                        "title": title
                        },
                        {
                        "description": description
                        },
                        {
                        "services_sys_val.support.service_type": service_type
                        },
                        {
                        "services_sys_val.support.service_sub_type": service_sub_type
                        },
                        {
                        "services_sys_val.support.lab_org": lab_org
                        },
                        {
                        "services_sys_val.support.category": category
                        },
                        {
                        "component": component
                        },
                        {
                        "priority": priority
                        },
                        {
                        "support.customer_contact": customer_contact
                        },
                        {
                        "support.site": site
                        },
                        {
                        "services_sys_val.support.lab": lab
                        },
                        {
                        "notify": notify
                        },
                        {
                        "services_sys_val.support.org_unit": org_unit
                        },
                        {
                        "services_sys_val.support.program": program
                        },
                        {
                        "services_sys_val.support.milestone_eta": milestone_eta
                        },
                        {
                        "send_mail": send_mail
                        },
                        {
                        "services_sys_val.support.required_by_milestone": required_by_milestone
                        },
                        {
                        "services_sys_val.support.survey_comment": survey_comment
                        }
                    ]
                    }
            
            readyMessage=''
            total_tickets=0
            exitFunction=False
            if customer_contact == '':
                print('no customer')
                readyMessage = "No Customer chosen.\n"
                exitFunction=TRUE
 
            if notify == '':
                print('no notify')
                readyMessage = readyMessage + "No Notify chosen.\n"
                exitFunction=TRUE
 
            if lab == '':
                print('no Lab')
                readyMessage = readyMessage + "No Lab chosen.\n"
                exitFunction=TRUE
 
            readyMessage = readyMessage + "\nCheck config.csv file!"
            if exitFunction==TRUE:
                ticketInterval = self.ticketInterval
                total_tickets = app.total_tickets
                
                if ticketInterval == total_tickets:
                    print(readyMessage)
                    tkinter.messagebox.showinfo('Config Error',readyMessage)
 
                return None
            else:
                data = json.dumps(payload)
    
                response = requests.post(url, verify=False,auth=HTTPKerberosAuth(), headers = headers, data = data)
                return response.json()
                return {'new_id':2}

        #--------------------------------------------------------------------------------------------------------------------
        #Start Here------------------------------------------------------
        def build_ticket_detials():
            not_ready=''
            not_ready1=''
            not_ready2=''
 
            if (len(project_option_selected.get()) == 0):
                not_ready1 = 'Select Program\n'
 
            if (len(site_option_selected.get()) == 0):
                not_ready2 = 'Select Site\n'
                
            not_ready = (not_ready1 + not_ready2)
            print(not_ready)
            
            if (len(not_ready) == 0):
                result=tkinter.messagebox.askquestion('Create Tickets?','Are you sure you want to create tickets?')
 
                if result=='yes':
                    Label_ProgressSuccess["text"]= "Connecting to HSD-ES DB"
                    root.after(1,Label_ProgressSuccess.update())
                    
                    selected_site=site_option_selected.get()
                    selected_lab=lab_dict[selected_site]
                    selected_notify=notify_dict[selected_site]
                    selected_lead=lead_dict[selected_site]
 
                    print("Selected Program: {}".format(project_option_selected.get()))
                    print("Selected Site: " + selected_site)
                    print("Selected WW: {}".format(WorkWeekValue_Inside.get()))
                    print("Selected Year: {}".format(YearValue_Inside.get()))
                    print("Selected Customer: " + selected_lead)
                    print("Selected Notify PDL: " + selected_notify)
                    print("Selected Lab: " + selected_lab)
                    print("11 CheckBox: {}".format(var11.get()))
                    print("21 CheckBox: {}".format(var21.get()))
                    print("22 CheckBox: {}".format(var22.get()))
                    print("31 CheckBox: {}".format(var31.get()))
                    print("32 CheckBox: {}".format(var32.get()))
                    print("41 CheckBox: {}".format(var41.get()))
                    print("42 CheckBox: {}".format(var42.get()))
                    print("43 CheckBox: {}".format(var43.get()))
                    print("51 CheckBox: {}".format(var51.get()))
                    print("61 CheckBox: {}".format(var61.get()))
                    print("62 CheckBox: {}".format(var62.get()))
                    #print("63 CheckBox: {}".format(var63.get()))
                    print("71 CheckBox: {}".format(var71.get()))
                    print("81 CheckBox: {}".format(var81.get()))
                    print("91 CheckBox: {}".format(var91.get()))
 
                    #Create a Dict to store the Checkbox values for each milestone
                    checkbox_dict={}
                    if dynamic_CSV_file == "Original":
                        checkbox_dict.update({"11":int(var11.get())})
                        checkbox_dict.update({"21":int(var21.get())})
                        checkbox_dict.update({"22":int(var22.get())})
                        checkbox_dict.update({"31":int(var31.get())})
                        checkbox_dict.update({"32":int(var32.get())})
                        checkbox_dict.update({"41":int(var41.get())})
                        checkbox_dict.update({"42":int(var42.get())})
                        checkbox_dict.update({"43":int(var43.get())})
                        checkbox_dict.update({"51":int(var51.get())})
                        checkbox_dict.update({"61":int(var61.get())})
                        checkbox_dict.update({"62":int(var62.get())})
                        #checkbox_dict.update({"63":int(var63.get())})
                        checkbox_dict.update({"71":int(var71.get())})
                        checkbox_dict.update({"81":int(var81.get())})
                        checkbox_dict.update({"91":int(var91.get())})
                    else: # PO and VV Tickets Also default for "New"
                        checkbox_dict.update({"111":int(var11.get())})
                        checkbox_dict.update({"211":int(var21.get())})
                        checkbox_dict.update({"221":int(var22.get())})
                        checkbox_dict.update({"311":int(var31.get())})
                        checkbox_dict.update({"321":int(var32.get())})
                        checkbox_dict.update({"411":int(var41.get())})
                        checkbox_dict.update({"421":int(var42.get())})
                        checkbox_dict.update({"431":int(var43.get())})
                        checkbox_dict.update({"511":int(var51.get())})
                        checkbox_dict.update({"611":int(var61.get())})
                        checkbox_dict.update({"621":int(var62.get())})
                        #checkbox_dict.update({"631":int(var63.get())})
                        checkbox_dict.update({"711":int(var71.get())})
                        checkbox_dict.update({"811":int(var81.get())})
                        checkbox_dict.update({"911":int(var91.get())})
                    #add Advance milestones
                    checkbox_dict.update({"101":int(1)})
          
                    #Create a List of Dicts to store the complete tictet information. One ticket per row
                    fieldlist=[]
 
                    dictionaryloop=1
                    if dictionaryloop==1:
 
                        for milestone in dynamic_vals.keys(): #Original Loop
 
                            #Check if Milestones Checkbox is selected
                            if len(milestone) > 2:
                                milestone_check = milestone[:2]+"1"
                            else:
                                milestone_check = milestone
 
                            if checkbox_dict.get(milestone_check) == 1:
                                _title={}
                                _title["title"]=dynamic_vals[milestone]["title"]
                                        
                                _description={}
                                _description["description"]=dynamic_vals[milestone]["description"]
 
                                _required_by_milestone={}
                                _required_by_milestone["required_by_milestone"]=dynamic_vals[milestone]["required_by_milestone"]
 
                                _lab_org={}
                                _lab_org["lab_org"]=static_vals.get("lab_org")
 
                                _org_unit={}
                                _org_unit["org_unit"]=static_vals.get("org_unit")
 
                                _category={}
                                _category["category"]=static_vals.get("category")
 
                                _component={}
                                _component["component"]=static_vals.get("component")
 
                                _priority={}
                                _priority["priority"]=static_vals.get("priority")
 
                                _status={}
                                _status["status"]=static_vals.get("status")
 
                                _reason={}
                                _reason["reason"]=static_vals.get("reason")
 
                                _customer_contact={}
                                _customer_contact["customer_contact"]=selected_lead
 
                                _site={}
                                _site["site"]=selected_site
 
                                _program={}
                                _program["program"]=project_option_selected.get()
 
                                _milestone_eta={}
                                d = YearValue_Inside.get()+WorkWeekValue_Inside.get()
                                r = datetime.strptime(d + '-1', "%Y%W-%w")
                                x = r - timedelta(weeks = int(dynamic_vals[milestone]["ETA_WW"]))
                                year = str(x.isocalendar().year)
                                week = str(x.isocalendar().week).zfill(2)
                                milestoneww = (year + week)
                                _milestone_eta["milestone_eta"]=milestoneww
                                print("milestoneww = "+milestoneww + "milestone = "+milestone)
 
                                _service_type={}
                                _service_type["service_type"]=static_vals.get("service_type")
 
                                _service_sub_type={}
                                _service_sub_type["service_sub_type"]=static_vals.get("service_sub_type")
 
                                _survey_comment={}

                                print(dynamic_vals[milestone]["ticket"])
                                _survey_comment["survey_comment"]=static_vals.get("survey_comment")
 
                                _lab={'lab': selected_lab}
 
                                _notify={'notify': selected_notify}
                            
                                line_dict={}
                                line_dict.update(_title)
                                line_dict.update(_description)
                                line_dict.update(_lab_org)
                                line_dict.update(_org_unit)
                                line_dict.update(_category)
                                line_dict.update(_component)
                                line_dict.update(_priority)
                                line_dict.update(_status)
                                line_dict.update(_reason)
                                line_dict.update(_customer_contact)
                                line_dict.update(_notify)
                                line_dict.update(_site)
                                line_dict.update(_program)
                                line_dict.update(_milestone_eta)
                                line_dict.update(_service_type)
                                line_dict.update(_service_sub_type)
                                line_dict.update(_required_by_milestone)
                                line_dict.update(_survey_comment)
                                line_dict.update(_lab)
                                fieldlist.append(line_dict)
 
                    if var101.get() == 1: #create advance tickets
                        #********************
                        print('awake')
                        
                        option_advance_dict={}
                        entry_advance_dict={}
                        # Where Milestone option number and additional title both have data add values to specific DB
                        if not entry_1.get() == '':
                            if not MS_1_selected.get() == '':
                                option_advance_dict.update({"MS1":(MS_1_selected.get())})
                                entry_advance_dict.update({"MS1":(entry_1.get())})
 
                        if not entry_2.get() == '':
                            if not MS_2_selected.get() == '':
                                option_advance_dict.update({"MS2":(MS_2_selected.get())})
                                entry_advance_dict.update({"MS2":(entry_2.get())})
 
                        if not entry_3.get() == '':
                            if not MS_3_selected.get() == '':
                                option_advance_dict.update({"MS3":(MS_3_selected.get())})
                                entry_advance_dict.update({"MS3":(entry_3.get())})
 
                        if not entry_4.get() == '':
                            if not MS_4_selected.get() == '':
                                option_advance_dict.update({"MS4":(MS_4_selected.get())})
                                entry_advance_dict.update({"MS4":(entry_4.get())})
 
                        if not entry_5.get() == '':
                            if not MS_5_selected.get() == '':
                                option_advance_dict.update({"MS5":(MS_5_selected.get())})
                                entry_advance_dict.update({"MS5":(entry_5.get())})
 
                        if not entry_6.get() == '':
                            if not MS_6_selected.get() == '':
                                option_advance_dict.update({"MS6":(MS_6_selected.get())})
                                entry_advance_dict.update({"MS6":(entry_6.get())})
 
                        if not entry_7.get() == '':
                            if not MS_7_selected.get() == '':
                                option_advance_dict.update({"MS7":(MS_7_selected.get())})
                                entry_advance_dict.update({"MS7":(entry_7.get())})
 
                        if not entry_8.get() == '':
                            if not MS_8_selected.get() == '':
                                option_advance_dict.update({"MS8":(MS_8_selected.get())})
                                entry_advance_dict.update({"MS8":(entry_8.get())})
 
                        if not entry_9.get() == '':
                            if not MS_9_selected.get() == '':
                                option_advance_dict.update({"MS9":(MS_9_selected.get())})
                                entry_advance_dict.update({"MS9":(entry_9.get())})
 
                        for a in option_advance_dict:                            
                            milestone = option_advance_dict[a]
                            milestone =  milestone.replace(".","")
                            append_title = entry_advance_dict[a]
 
                            if dynamic_CSV_file == 'Original':
                                print(append_title)
                            if dynamic_CSV_file == 'PO':
                                milestone_check = milestone + "1"
                            if dynamic_CSV_file == 'VV':
                                milestone_check = milestone + "1"
 
                            _title={}
                            _title["title"]=dynamic_vals[milestone_check]["cb_title"]+" - "+append_title
                                        
                            _description={}
                            _description["description"]=dynamic_vals[milestone_check]["description"]
 
                            _required_by_milestone={}
                            _required_by_milestone["required_by_milestone"]=dynamic_vals[milestone_check]["required_by_milestone"]
 
                            _lab_org={}
                            _lab_org["lab_org"]=static_vals.get("lab_org")
 
                            _org_unit={}
                            _org_unit["org_unit"]=static_vals.get("org_unit")
 
                            _category={}
                            _category["category"]=static_vals.get("category")
 
                            _component={}
                            _component["component"]=static_vals.get("component")
 
                            _priority={}
                            _priority["priority"]=static_vals.get("priority")
 
                            _status={}
                            _status["status"]=static_vals.get("status")
 
                            _reason={}
                            _reason["reason"]=static_vals.get("reason")
 
                            _customer_contact={}
                            _customer_contact["customer_contact"]=selected_lead
 
                            _site={}
                            _site["site"]=selected_site
 
                            _program={}
                            _program["program"]=project_option_selected.get()
 
                            _milestone_eta={}
                            d = YearValue_Inside.get()+WorkWeekValue_Inside.get()
                            r = datetime.strptime(d + '-1', "%Y%W-%w")
                            x = r - timedelta(weeks = int(dynamic_vals[milestone_check]["ETA_WW"]))
                            year = str(x.isocalendar().year)
                            week = str(x.isocalendar().week).zfill(2)
                            milestone = (year + week)
                            _milestone_eta["milestone_eta"]=milestone
 
                            _service_type={}
                            _service_type["service_type"]=static_vals.get("service_type")
 
                            _service_sub_type={}
                            _service_sub_type["service_sub_type"]=static_vals.get("service_sub_type")
 
                            _survey_comment={}

                        _survey_comment["survey_comment"]=static_vals.get("survey_comment")
                        _lab={'lab': selected_lab}
                        _notify={'notify': selected_notify}
                        
                        line_dict={}
                        line_dict.update(_title)
                        line_dict.update(_description)
                        line_dict.update(_lab_org)
                        line_dict.update(_org_unit)
                        line_dict.update(_category)
                        line_dict.update(_component)
                        line_dict.update(_priority)
                        line_dict.update(_status)
                        line_dict.update(_reason)
                        line_dict.update(_customer_contact)
                        line_dict.update(_notify)
                        line_dict.update(_site)
                        line_dict.update(_program)
                        line_dict.update(_milestone_eta)
                        line_dict.update(_service_type)
                        line_dict.update(_service_sub_type)
                        line_dict.update(_required_by_milestone)
                        line_dict.update(_survey_comment)
                        line_dict.update(_lab)

                        fieldlist.append(line_dict)
 
                    #********************
                    print("")
 
                    #Count List lines
                    total_tickets = 0
                    ticketInterval = 0
 
                    for i in fieldlist:
                        total_tickets=total_tickets+1
                        App.total_tickets = total_tickets
 
                    App.ticketInterval = ticketInterval
 
                    if total_tickets==0:
                        print('No Tickets')
                        tkinter.messagebox.showinfo('No Milestones Selected', 'Select at least one Milestone.')
                    else:
                        #Create Ticket(s)
                        ProgressBar['value']=0
                        ticket_number=0
                        error_number=0
                        tickets = []
                        errors = []
                        HSD_ID = []
 
                        for fields in fieldlist:
                            Button_Create["state"] = "disabled"
                            app.ticketInterval = app.ticketInterval +1
                            #----------------------------------------------------------
                            #response = postnewHSD(fields) #Uncomment to create tickets
                            #----------------------------------------------------------
                            print('ticket_number',ticket_number)
 
                            try:                         
                                tickets.append(response['new_id'])
                                print("Created ticket:",response['new_id'])
                                HSD_ID.append(response['new_id'])
                                ticket_number = ticket_number + 1
                            except:
                                try:
                                    errors.append(response['message'])
                                    print('error:',response['message'])
                                    HSD_ID.append(response['message'])
                                    error_number = error_number + 1
                                except:
                                    print('error')
                                                        
                            ProgressBar['value']+=100/total_tickets
                            Label_ProgressSuccess["text"]= "Created: " + str(ticket_number) + " / " + str(total_tickets) + " tickets."
                            root.after(1,ProgressBar.update(),Label_ProgressSuccess.update())
 
                        display_tickets='\n\n'
 
                        for line in HSD_ID:
                            display_tickets=display_tickets + str(line) + '\n'
 
                        final_message=""
                        final_message="Successfully created " + str(ticket_number) + " of " + str(total_tickets) + " tickets. " + str(display_tickets)
 
                        if error_number>1:
                            final_message=final_message + str(error_number) + " tickets had errors" 
 
                        tkinter.messagebox.showinfo('Tickets Created?',final_message)
                        
                        linkCollection = ""
 
                        if error_number == 0:
 
                            for line in HSD_ID:
                                link = ""      
                                link = linkUrl + str(line)
                                linkCollection=linkCollection + link + " \n"
                            linkCollection = linkCollection[0:(len(linkCollection)-2)]
                            saveClipboard(linkCollection)
 
                    Label_ProgressSuccess["text"]= "Ready"
                    root.after(1,Label_ProgressSuccess.update())
                    Button_Create["state"] = "enabled"
 
                    return
                else:
                    return None
 
            tkinter.messagebox.showerror('Not Ready!',not_ready)
 
        def display_info():
            versiona_message = ('HSD-Lab GEN: ' + vernum + '\n\nconfig.csv: ' + configver + '\n\nstatic_vals.csv: ' + staticver + '\n\ndynamic_vals.csv: ' + dynamicver)
            tkinter.messagebox.showinfo('File Versions',versiona_message)  
 
        def advance_window(tog=[0]):
            tog[0] = not tog[0]
            if tog[0]:
                print("Advance Window Open")
                root.geometry("892x680")
                Button_advance["text"] = "<-"
            else:
                print("Advance Window Close")
                root.geometry("492x680")
                Button_advance["text"] = "->"
 
        MS_1_selected = tk.StringVar(root)
        MS_2_selected = tk.StringVar(root)
        MS_3_selected = tk.StringVar(root)
 
        entry_1 = tk.Entry(root)
        entry_2 = tk.Entry(root)
        entry_3 = tk.Entry(root)
 
        def clear_advance_row1():
            print ('Clear Row 1')
            MS_1_selected.set("")
            entry_1.delete(0, tk.END)
 
        def clear_advance_row2():
            print ('Clear Row 2')
            MS_2_selected.set("")
            entry_2.delete(0, tk.END)
 
        def clear_advance_row3():
            print ('Clear Row 3')
            MS_3_selected.set("")
            entry_3.delete(0, tk.END)
 
        def clear_advance_row4():
            print ('Clear Row 4')
            MS_4_selected.set("")
            entry_4.delete(0, tk.END)
 
        def clear_advance_row5():
            print ('Clear Row 5')
            MS_5_selected.set("")
            entry_5.delete(0, tk.END)
 
        def clear_advance_row6():
            print ('Clear Row 6')
            MS_6_selected.set("")
            entry_6.delete(0, tk.END)
 
        def clear_advance_row7():
            print ('Clear Row 7')
            MS_7_selected.set("")
            entry_7.delete(0, tk.END)
 
        def clear_advance_row8():
            print ('Clear Row 8')
            MS_8_selected.set("")
            entry_8.delete(0, tk.END)
 
        def clear_advance_row9():
            print ('Clear Row 9')
            MS_9_selected.set("")
            entry_9.delete(0, tk.END)
 
        ProgressBar=ttk.Progressbar(root)
        ProgressBar.place(x=6,y=665,width=480,height=10)
 
        Button_Create=ttk.Button(root)
        Button_Create["text"] = "Create"
        Button_Create["tooltip"] = "Create selected milestone tickets and copies links to them in the clipboard cache."
        Button_Create.place(x=320,y=630,width=158,height=30)
        Button_Create["command"] = build_ticket_detials
 
        Button_info=ttk.Button(root)
        Button_info["text"] = "info"
        Button_info["tooltip"] = "File version information."
        Button_info.place(x=417,y=96,width=65,height=30)
        Button_info["command"] = display_info
 
        Button_advance=ttk.Button(root)
        Button_advance["text"] = "->"
        Button_advance["tooltip"] = "Advance options."
        Button_advance.place(x=417,y=126,width=65,height=30)
        Button_advance["command"] = advance_window
 
        #-----------------------------------------------------------------------------------------------------------------------------
        #Advance Options
        MS_1_selected = tk.StringVar(root)
        MS_option1=ttk.OptionMenu(root,MS_1_selected,*MS_options)
        MS_option1['tooltip'] = 'Choose Milestone.'
        MS_option1.place(x=500,y=135,width=54,height=25)
        MS_option1.configure(state="disabled")
 
        
        MS_option2=ttk.OptionMenu(root,MS_2_selected,*MS_options)
        MS_option2['tooltip'] = 'Choose Milestone.'
        MS_option2.place(x=500,y=165,width=54,height=25)
        MS_option2.configure(state="disabled")
 
        
        MS_option3=ttk.OptionMenu(root,MS_3_selected,*MS_options)
        MS_option3['tooltip'] = 'Choose Milestone.'
        MS_option3.place(x=500,y=195,width=54,height=25)
        MS_option3.configure(state="disabled")
 
        MS_4_selected = tk.StringVar(root)
        MS_option4=ttk.OptionMenu(root,MS_4_selected,*MS_options)
        MS_option4['tooltip'] = 'Choose Milestone.'
        MS_option4.place(x=500,y=225,width=54,height=25)
        MS_option4.configure(state="disabled")
 
        MS_5_selected = tk.StringVar(root)
        MS_option5=ttk.OptionMenu(root,MS_5_selected,*MS_options)
        MS_option5['tooltip'] = 'Choose Milestone.'
        MS_option5.place(x=500,y=255,width=54,height=25)
        MS_option5.configure(state="disabled")
 
        MS_6_selected = tk.StringVar(root)
        MS_option6=ttk.OptionMenu(root,MS_6_selected,*MS_options)
        MS_option6['tooltip'] = 'Choose Milestone.'
        MS_option6.place(x=500,y=285,width=54,height=25)
        MS_option6.configure(state="disabled")
 
        MS_7_selected = tk.StringVar(root)
        MS_option7=ttk.OptionMenu(root,MS_7_selected,*MS_options)
        MS_option7['tooltip'] = 'Choose Milestone.'
        MS_option7.place(x=500,y=315,width=54,height=25)
        MS_option7.configure(state="disabled")
 
        MS_8_selected = tk.StringVar(root)
        MS_option8=ttk.OptionMenu(root,MS_8_selected,*MS_options)
        MS_option8['tooltip'] = 'Choose Milestone.'
        MS_option8.place(x=500,y=345,width=54,height=25)
        MS_option8.configure(state="disabled")
 
        MS_9_selected = tk.StringVar(root)
        MS_option9=ttk.OptionMenu(root,MS_9_selected,*MS_options)
        MS_option9['tooltip'] = 'Choose Milestone.'
        MS_option9.place(x=500,y=375,width=54,height=25)
        MS_option9.configure(state="disabled")
 
 
        #Advanced Entry
        entry_1.place(x=556,y=135,width=282,height=25)
        entry_1.configure(state= "disabled")
        
        entry_2.place(x=556,y=165,width=282,height=25)
        entry_2.configure(state= "disabled")
        
        entry_3.place(x=556,y=195,width=282,height=25)
        entry_3.configure(state= "disabled")
 
        entry_4 = tk.Entry(root)
        entry_4.place(x=556,y=225,width=282,height=25)
        entry_4.configure(state= "disabled")
 
        entry_5 = tk.Entry(root)
        entry_5.place(x=556,y=255,width=282,height=25)
        entry_5.configure(state= "disabled")
 
        entry_6 = tk.Entry(root)
        entry_6.place(x=556,y=285,width=282,height=25)
        entry_6.configure(state= "disabled")
 
        entry_7 = tk.Entry(root)
        entry_7.place(x=556,y=315,width=282,height=25)
        entry_7.configure(state= "disabled")
 
        entry_8 = tk.Entry(root)
        entry_8.place(x=556,y=345,width=282,height=25)
        entry_8.configure(state= "disabled")
 
        entry_9 = tk.Entry(root)
        entry_9.place(x=556,y=375,width=282,height=25)
        entry_9.configure(state= "disabled")
 
        #Advance Buttons
        Row = 1
        button_clear1 = tk.Button(root)
        button_clear1["text"] = "Clear"
        #button_clear1['tooltip'] = "Clear Entry."
        button_clear1.place(x=840,y=135,width=45,height=25)
        button_clear1["command"] = clear_advance_row1
        button_clear1["state"] = "disabled"
 
        button_clear2 = tk.Button(root)
        button_clear2["text"] = "Clear"
        #button_clear2["tooltip"] = "Clear Entry."
        button_clear2.place(x=840,y=165,width=45,height=25)
        button_clear2["command"] = clear_advance_row2
        button_clear2["state"] = "disabled"
 
        button_clear3 = tk.Button(root)
        button_clear3["text"] = "Clear"
        #button_clear3["tooltip"] = "Clear Entry."
        button_clear3.place(x=840,y=195,width=45,height=25)
        button_clear3["command"] = clear_advance_row3
        button_clear3["state"] = "disabled"
 
        button_clear4 = tk.Button(root)
        button_clear4["text"] = "Clear"
        #button_clear4["tooltip"] = "Clear Entry."
        button_clear4.place(x=840,y=225,width=45,height=25)
        button_clear4["command"] = clear_advance_row4
        button_clear4["state"] = "disabled"
 
        button_clear5 = tk.Button(root)
        button_clear5["text"] = "Clear"
        #button_clear5["tooltip"] = "Clear Entry."
        button_clear5.place(x=840,y=255,width=45,height=25)
        button_clear5["command"] = clear_advance_row5
        button_clear5["state"] = "disabled"
 
        button_clear6 = tk.Button(root)
        button_clear6["text"] = "Clear"
        #button_clear6["tooltip"] = "Clear Entry."
        button_clear6.place(x=840,y=285,width=45,height=25)
        button_clear6["command"] = clear_advance_row6
        button_clear6["state"] = "disabled"
 
        button_clear7 = tk.Button(root)
        button_clear7["text"] = "Clear"
        #button_clear7["tooltip"] = "Clear Entry."
        button_clear7.place(x=840,y=315,width=45,height=25)
        button_clear7["command"] = clear_advance_row7
        button_clear7["state"] = "disabled"
 
        button_clear8 = tk.Button(root)
        button_clear8["text"] = "Clear"
        #button_clear8["tooltip"] = "Clear Entry."
        button_clear8.place(x=840,y=345,width=45,height=25)
        button_clear8["command"] = clear_advance_row8
        button_clear8["state"] = "disabled"
 
        button_clear9 = tk.Button(root)
        button_clear9["text"] = "Clear"
        #button_clear9["tooltip"] = "Clear Entry."
        button_clear9.place(x=840,y=375,width=45,height=25)
        button_clear9["command"] = clear_advance_row9
        button_clear9["state"] = "disabled"
 
 
 
 
    def CheckBox_SelectAll_command(self):
        if (self.varSelectAll.get()==1):
            self.CheckBox_11.select()
            self.CheckBox_21.select()
            self.CheckBox_22.select()
            self.CheckBox_31.select()
            self.CheckBox_32.select()
            self.CheckBox_41.select()
            self.CheckBox_42.select()
            self.CheckBox_43.select()            
            self.CheckBox_51.select()
            self.CheckBox_61.select()
            self.CheckBox_62.select()
            #self.CheckBox_63.select()
            self.CheckBox_71.select()
            self.CheckBox_81.select()
            self.CheckBox_91.select()
        else:
            self.CheckBox_11.deselect()
            self.CheckBox_21.deselect()
            self.CheckBox_22.deselect()
            self.CheckBox_31.deselect()
            self.CheckBox_32.deselect()
            self.CheckBox_41.deselect()
            self.CheckBox_42.deselect()
            self.CheckBox_43.deselect()
            self.CheckBox_51.deselect()
            self.CheckBox_61.deselect()
            self.CheckBox_62.deselect()
            #self.CheckBox_63.deselect()
            self.CheckBox_71.deselect()
            self.CheckBox_81.deselect()
            self.CheckBox_91.deselect()
 
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    canvas = Canvas()
    root.mainloop()