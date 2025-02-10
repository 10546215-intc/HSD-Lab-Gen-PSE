import tkinter as tk
from tkinter import ttk
from tkinter import *
import tkinter.messagebox
import tkinter.font as tkFont
from ttkwidgets import tooltips
import datetime
from datetime import timedelta
import time
import csv
import json
import requests
from requests_kerberos import HTTPKerberosAuth
import urllib3
import os
import customtkinter
import logging


 
# this is to ignore the ssl insecure warning as we are passing in 'verify=false'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
os.environ["no_proxy"] = "127.0.0.1,localhost,intel.com"
#headers = { 'Content-type': 'application/json' }

# Declare global variables
url = ""
linkUrl = ""
auto_c = ""
icon = ""
dynamic_vals_open = ""
pull_down_mode = "dynamic" # "dynamic" for HSD based "static" for Config file
milestones = {}
checkboxes = {}
variables = {}
mode_var = {}
src_var = {}
#dynamic_vals = {}
project_options = ()
site_options = ()
dynamic_CSV_file = 'NEW'
###site_configs could be list
sites_options=['']
notify_dict={}
customer_options={}
backup_options={}
notify_options={}
lab_dict={}
lead_dict={}
static_vals = {}
config_open = ""
checkbox_dict={}
staticver = ""
static_vals_open = "" #static version number
linkCollection = ""
hsd_source = ""

#url_Prod = "https://hsdes-api.intel.com/"
#url_Pre = "https://hsdes-api-pre.intel.com/"

 
vernum = "0.98" #application version number
configver = "" #config version number
dynamicver = "" #dynamic version number
 
menu_color = "#e4e5ee"
option_color = "#D4D6E4"
 
root = customtkinter.CTk()
 
def get_milestones():
            ### NEW Track CSV Code for file dynamic_vals.csv ###
        global milestones
        global dynamic_vals_open
        try:
            with open("dependencies/dynamic_vals.csv", encoding="utf8") as data_file:
                print('\nloading NEW track dynamic_vals')
                data = csv.reader(data_file)
                dynamic_headers = next(data)[0:]                            
                for row in data:
                    temp_dict = {}
                    keystone = row[0]
                    name = row[2]
                    values = []
                    for x in row[0:]:    
                        values.append(x)

                    for i in range(len(values)):
                        if values[i]:
                            temp_dict[dynamic_headers[i]] = values[i]
                    milestones[name] = temp_dict
            dynamic_vals_open=True
        except:
            print('Did not load dynamic_vals_new')
            dynamic_vals_open=False

def get_checkboxes():
    #########  Create milestone check boxes
    global checkboxes
    global variables
    print('')

    cb_num=1
    start_y = 196
    keystones = 0

    # Loop through each dictionary in the milestone collection
    for key, item in milestones.items():
        # Check if the value of the "keystone" key is 1
        if item.get("keystone") == "1":
            # Increment the count
            keystones += 1
#    if keystones > 8:
#        tkinter.messagebox.showwarning('Keystone Error','Only 14 Milestones can be displayed in tool. Click "OK" to continue. All Milestones will be created just not listed')
#        print("Only 14 Milestones can be displayed in tool")

    try:

        # Loop to create checkboxes
        for key, item in milestones.items():
            # Check if the value of the "keystone" key is 1
            if item.get("keystone") == "1":
                var = tk.IntVar(root)
                checkbox = tk.Checkbutton(root)
                checkbox["anchor"] = "w"
                checkbox["font"] = tkFont.Font(family='Times', size=10)
                checkbox["fg"] = "#333333"
                checkbox["justify"] = "left"
                checkbox["text"] = milestones[key]["cb_title"]
                checkbox.place(x=10, y=196 + cb_num * 30, width=490, height=25)  # Adjust y position for each checkbox
                checkbox["offvalue"] = "0"
                checkbox["onvalue"] = "1"
                checkbox["variable"] = var
                
                # Select the Checkbutton by default
                checkbox.select()
                #Orig# variables[key] = var
                #orig# variables[key] = {'widget': checkbox, 'variable': var, 'text': checkbox["text"]}
                variables[key] = {'widget': checkbox, 'variable': var, 'text': checkbox["text"], 'mile': milestones[key]["mile"] }
                # Store the checkbox in the dictionary
                checkboxes[cb_num] = checkbox
                cb_num += 1
                
    except:
        print("fail")

def get_opt_menu_list(field, name, data_src):
    global project_options
    global site_options
    global static_vals
    tmp_dict=[]
    data_lower=[]
    global customer_options
    global backup_options
    global notify_dict
    global lab_dict
    global lead_dict
    
    src_file = f"dependencies/{name}.csv"   
    try:
        headers = {'Content-type': 'application/json'}
#        Label_ProgressSuccess["text"] = "Validating Menu Options with HSD-ES DB"
#        root.after(1, Label_ProgressSuccess.update())
        url_validate = f'{auto_c}/{field}'
        field_type = field
        response = requests.get(url_validate, verify=False, auth=HTTPKerberosAuth(), headers=headers)
        
        # Check if the response is successful
        if response.status_code != 200:
#           logger.error(f"Failed to fetch data: {response.status_code} - {response.text}")
            return []
        
        data = response.json().get("data", [])
        data_lower = [(item.get('' + field_type + '', None)) for item in data if item.get('' + field_type + '', None)]

        # Validate the structure of the data
        if not isinstance(data, list):
#           logger.error("Invalid data format received")
            return []

    except requests.RequestException as e:
#       logger.error(f"Request failed: {e}")
        tkinter.messagebox.showwarning('Connection Error','Not connected to HSDE-ES DB.')
        return []
    except Exception as e:
#       logger.error(f"An unexpected error occurred: {e}")
        tkinter.messagebox.showwarning('Connection Error','Not connected to HSDE-ES DB.')
        return []
### Load from CSV file    
    if data_src == 'CSV':     
        try:
        #with open(r"dependencies/config.csv", encoding="utf8") as f:
            with open(src_file, encoding="utf8") as f:    
                csv_config = csv.DictReader(f)
                print('\nloading config and verifing options')
                for line in csv_config:
                    #Checks if there is a value in the first line and if it exist in HSD
                    #print(f" Checking for {line[name]} ")
                    if(line[name]) and line.get(name).lower() in data_lower:
                        tmp_dict.append(line[name])
                        if name == 'Site':                     
                            sites_options.append(line['Site'])
                            customer_options.update({line['Site']: line['Customer']})
                            backup_options.update({line['Site']: line['Backup']})
                            notify_dict.update({line['Site']: line['Notify']})
                            lab_dict.update({line['Site']: line['Lab']})
                            lead_dict.update({line['Site']: line['Customer']})
                    else:
                        print(f" {line[name]} - Not Found In HSD, Check Name and update program.csv" )
                        FileOpenMessage="Can not find file(s):\n"
                        FileOpenMessage = FileOpenMessage + f'\ndependencies/{name}.csv'
            v_list = tmp_dict
            
            if name == 'program':
                project_options = v_list
                project_options.sort()            
                project_options.insert(0,"")
            elif name == 'Site':
                site_options = v_list
                site_options.sort()
                site_options.insert(0,"")
            else:
                print("Options list unassigned") 
            
            config_open=True
        except:
            config_open=False

### Load from HSD query and CSV File        
    if data_src == 'HSD':
        try:
            v_list = [(item.get('' + field_type + '', None)) for item in data if item.get('' + field_type + '', None)]   
            #with open(r"dependencies/config.csv", encoding="utf8") as f:
            with open(src_file, encoding="utf8") as f:    
                csv_config = csv.DictReader(f)
                print('\nloading config and verifing options')
                for line in csv_config:                
                    sites_options.append(line['Site'])
                    customer_options.update({line['Site']: line['Customer']})
                    backup_options.update({line['Site']: line['Backup']})
                    notify_dict.update({line['Site']: line['Notify']})
                    lab_dict.update({line['Site']: line['Lab']})
                    lead_dict.update({line['Site']: line['Customer']})
                    
                if name == 'program':
                    project_options = v_list
                    project_options.sort()            
                    project_options.insert(0,"")
                elif name == 'site':
                    site_options = v_list
                    site_options.sort()
                    site_options.insert(0,"")
                else:
                    print("Options list unassigned")          
            config_open=True
        except:
            config_open=False
            FileOpenMessage="Can not find file(s):\n"
            FileOpenMessage = FileOpenMessage + f'\ndependencies/{name}.csv'
    print('config_open= '+ str(config_open))
### Load Static_vals happens each time menu list are updated
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
        FileOpenMessage="Can not find file(s):\n"
        FileOpenMessage = FileOpenMessage + '\ndependencies/static_vals.csv'
    print('static_vals_open = ' + str(static_vals_open))
        
def get_hsd_url(mode):
    global url
    global linkUrl
    global icon
    global milestones
    global auto_c
    
    if mode == "Pre-Prod":
#        mode_label.configure(text="Pre-Production (HSD)")
        url = 'https://hsdes-api-pre.intel.com/rest/article'
        linkUrl = 'https://hsdes-pre.intel.com/appstore/article/#/'
        auto_c = 'https://hsdes-api-pre.intel.com/rest/query/autocomplete/support/services_sys_val/'
        icon = PhotoImage(file='dependencies/Y.png')
    else:
#        mode_label.configure(text="Production (HSD)")
        url = 'https://hsdes-api.intel.com/rest/article'
        linkUrl = 'https://hsdes.intel.com/appstore/article/#/'
        auto_c = 'https://hsdes-api.intel.com/rest/query/autocomplete/support/services_sys_val/'
        icon = PhotoImage(file='dependencies/B.png')


class App():
#    global dynamic_vals
    global dynamic_CSV_file
    global dynamic_vals_open
    global static_vals_open
    global pull_down_mode
    global project_options
    global site_options
    global lab_dict
    global sites_options
    global notify_dict
    global lead_dict
    global customer_options
    global notify_options
    global config_open
    global icon
    global hsd_source

    def __init__(self, root):
        self.root = root
        root.title("HSD-Lab Gen - Version " + str(vernum))
        root.eval("tk::PlaceWindow . center")
        #setting window size
        #width=492
        width=492
        height=740
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
 
        # #Load Icon
    # def UpdateIcon():
    #     global icon
    #     try:
    #         root.iconphoto(False, icon)
    #     except:
    #         print('Could not find icon.')
 
    #     UpdateIcon()
    
        #Set Size for GUI
        

    
        cv = Canvas(root,width=892,height=740)
        cv.pack()
        cv.create_rectangle(2,38,485,126,fill=menu_color)
        cv.create_line(165,38,165,126)
        cv.create_line(326,38,326,126)



        def update_hsd_url(self):
            source = self.hsd_source.get()
            global url, linkUrl, icon, auto_c
            global auto_c
            
            if source == "Pre-Prod":
        #        mode_label.configure(text="Pre-Production (HSD)")
                url = 'https://hsdes-api-pre.intel.com/rest/article'
                linkUrl = 'https://hsdes-pre.intel.com/appstore/article/#/'
                auto_c = 'https://hsdes-api-pre.intel.com/rest/query/autocomplete/support/services_sys_val/'
                icon = PhotoImage(file='dependencies/Y.png')
            if source == "Prod":
        #        mode_label.configure(text="Production (HSD)")
                url = 'https://hsdes-api.intel.com/rest/article'
                linkUrl = 'https://hsdes.intel.com/appstore/article/#/'
                auto_c = 'https://hsdes-api.intel.com/rest/query/autocomplete/support/services_sys_val/'
                icon = PhotoImage(file='dependencies/B.png')


        #Add Menu for Options
        menu = tk.Menu(root)

        hsd_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='HSD Sources', menu=hsd_menu)
        # Use a single StringVar to hold the value of the selected radio button
        self.hsd_source = tk.StringVar(value='Pre-Prod')
        # Add radio buttons to the menu
        hsd_menu.add_radiobutton(label='HSD-Prod', value='Prod', variable=self.hsd_source,command=lambda: update_hsd_url(self))
        hsd_menu.add_radiobutton(label='HSD-Pre', value='Pre-Prod', variable=self.hsd_source,command =lambda: update_hsd_url(self))
        
        prog_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='Program Sources', menu=prog_menu)
        # Use a single StringVar to hold the value of the selected radio button
        prog_source = tk.StringVar(value = 'CSV')
        # Add radio buttons to the menu
        prog_menu.add_radiobutton(label='HSD', value='HSD', variable=prog_source)
        prog_menu.add_radiobutton(label='CSV', value='CSV', variable=prog_source)

        site_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='Site Sources', menu=site_menu)
        # Use a single StringVar to hold the value of the selected radio button
        site_source = tk.StringVar(value = 'CSV')
        # Add radio buttons to the menu
        site_menu.add_radiobutton(label='HSD', value='HSD', variable=site_source)
        site_menu.add_radiobutton(label='CSV', value='CSV', variable=site_source)
        site_menu = tk.Menu(menu, tearoff=False)

        root.configure(menu = menu)
        











        
        # Create Toggle function
        def clicker():
            mode_switch.toggle()
            if switch_var.get() == "Pre-Prod":
                mode_label.configure(text="Pre-Production (HSD)")
            else:
                mode_label.configure(text="Production (HSD)")

        # Create a StringVar
        switch_var = customtkinter.StringVar(value="Pre-Prod")

        # Create Switch
        mode_switch = customtkinter.CTkSwitch(root, text="", command=get_hsd_url,
                                            variable=switch_var, onvalue="Pre-Prod", offvalue="Prod",
                                            switch_width=20,
                                            switch_height=10,
                                            font=("Times", 24),
                                            text_color="blue",
                                            state="normal",
                                            )
        mode_var = {'widget': mode_switch, 'variable': switch_var}

        # Create Toggle function
        def clicker_menu_src():
            src_switch.toggle()

        # Option Menu Source Switch
        switch_src = customtkinter.StringVar(value="CSV")

        # Create Switch
        src_switch = customtkinter.CTkSwitch(root, text="",
                                            variable=switch_src, onvalue="HSD", offvalue="CSV",
                                            switch_width=20,
                                            switch_height=10,
                                            font=("Times", 24),
                                            text_color="blue",
                                            state="normal",
                                            )
        src_var = {'widget': src_switch, 'variable': switch_src}
        # Create Label
        mode_label = customtkinter.CTkLabel(root, text="Pre-Production", font=("Times", 24), text_color="blue")
    #    src_label =
        # Place the label and switch
        mode_label.place(relx=0.5, rely=0.02, anchor=CENTER)  # Adjust x and y as needed
        mode_switch.place(relx=1, rely=0.02, anchor=CENTER)  # Adjust x and y as needed
        
    #    src_label.place(relx=0.5, rely=0.02, anchor=CENTER)  # Adjust x and y as needed
        src_switch.place(relx=1, rely=0.29, anchor=CENTER)  # Adjust x and y as needed       

        #ProgressSuccess Label
        Label_ProgressSuccess=tk.Label(root)
        ft = tkFont.Font(family='Times',size=18)
        Label_ProgressSuccess["font"] = ft
        Label_ProgressSuccess["fg"] = "#333333"
        Label_ProgressSuccess["justify"] = "left"
        Label_ProgressSuccess["text"] = ""
        Label_ProgressSuccess.place(x=0,y=340,width=492,height=20)



        get_hsd_url(switch_var.get())
        get_opt_menu_list('services_sys_val.support.program','program',switch_src.get()) 
        get_opt_menu_list('support.site','Site', switch_src.get()) 
        get_milestones()
        #UpdateIcon()
        
#### Update to populate pulldowns from query, to be removed after 
        #Validate field aginist HSD-ES DB
        # Create variables for the project and site option selected, etc.
        project_option_selected = tk.StringVar(root)
        site_option_selected = tk.StringVar(root)
        customer_option_selected = tk.StringVar(root)
        notify_option_selected = tk.StringVar(root)
 

        MS_options=['']
#        l_MS_options=['']
        
        #Close GUI and Python
        def close():
            exec
            root.destroy()
 
        clipboard = static_vals['createClipboard']

### Sets production http can be removed.

        #Create List for WorkWeek OptionMenu
        i = 1
        WorkWeekList = []
        for i in range(53):
            WorkWeekList.insert(i,i)
            i += 1
 
        #Get the Date
        today = datetime.date.today()
        iso_calendar = today.isocalendar()
        WorkWeek = iso_calendar[1]
        year = iso_calendar[0]
        WorkWeekValue_Inside = tk.StringVar(root)
        YearValue_Inside = tk.StringVar(root)
 
        #Create List for Year OptionMenu
        i = 1
        YearList = []
        for i in range(2021,2051):
            YearList.insert(i,i)
            i += 1

        # Create Pulldown OptionMenu boxes with tooltips
        # load full project list
        
        
        project_menu=ttk.OptionMenu(root,project_option_selected,*project_options)
        project_menu['tooltip'] = 'Choose program / project.'
        project_menu.place(x=10,y=69,width=150,height=25)
 
        site_menu=ttk.OptionMenu(root,site_option_selected,*site_options)
        site_menu['tooltip'] = 'Choose Site.'
        site_menu.place(x=168,y=69,width=150,height=25)
        
        Option_101=ttk.OptionMenu(root,WorkWeekValue_Inside,*WorkWeekList)
        Option_101['tooltip'] = 'Choose todays known work week for power on.'
        Option_101.place(x=400,y=69,width=80,height=25)
        WorkWeekValue_Inside.set(WorkWeek)
 
        Option_403=ttk.OptionMenu(root,YearValue_Inside,*YearList)
        Option_403['tooltip'] = 'Choose todays known year for power on.'
        Option_403.place(x=400,y=99,width=80,height=25)
        YearValue_Inside.set(year)
 
        # Create labels for pulldown Boxes

        # Program Label
        Label_Program = tk.Label(root, text="Program", font=("Times", 14), fg="#333333", bg=menu_color, anchor="c", justify="left")
        Label_Program.place(x=10, y=41, width=150, height=25)

        # Site Label
        Label_Site = tk.Label(root, text="Site", font=("Times", 14), fg="#333333", bg=menu_color, anchor="c", justify="center")
        Label_Site.place(x=168, y=41, width=150, height=25)

        # Power On Date Label
        Label_PODate = tk.Label(root, text="Power On Date", font=("Times", 14), fg="#333333", bg=menu_color, anchor="c", justify="left")
        Label_PODate.place(x=330, y=41, width=150, height=25)

        # WorkWeek Label
        Label_WorkWeek = tk.Label(root, text="WW#", font=("Times", 14), fg="#333333", bg=menu_color, anchor="e", justify="right")
        Label_WorkWeek.place(x=330, y=69, width=70, height=25)

        #Year Label
        Label_Year = tk.Label(root, text="Year", font=("Times", 14), fg="#333333", bg=menu_color, anchor="e", justify="right")
        Label_Year.place(x=330, y=99, width=70, height=25)

        self.varSelectAll=tk.IntVar(root)
        self.CheckBox_SelectAll=tk.Checkbutton(root)
        self.CheckBox_SelectAll["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=10)
        self.CheckBox_SelectAll["font"] = ft
        self.CheckBox_SelectAll["fg"] = "#333333"
        self.CheckBox_SelectAll["justify"] = "left"
        self.CheckBox_SelectAll["text"] = "Select All"
        self.CheckBox_SelectAll.place(x=10,y=156,width=110,height=25)
        self.CheckBox_SelectAll["offvalue"] = "0"
        self.CheckBox_SelectAll["onvalue"] = "1"
        self.CheckBox_SelectAll["command"] = self.CheckBox_SelectAll_command
        self.CheckBox_SelectAll["variable"] = self.varSelectAll
        self.CheckBox_SelectAll.select()
 
 ########################################################################################
 #########  Create milestone check boxes
        get_checkboxes()

        
        #Milestones Label
        Label_MS=tk.Label(root)
        Label_MS["anchor"] = "w"
        ft = tkFont.Font(family='Times',size=14)
        Label_MS["font"] = ft
        Label_MS["fg"] = "#333333"
        Label_MS["justify"] = "left"
        Label_MS["text"] = dynamic_CSV_file+" Milestones"
        Label_MS["relief"] = "flat"
        Label_MS.place(x=10,y=131,width=186,height=30)
 
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
 
 
        ### Required because Pre and Prod HSD dont have the same fields. Need to submit Ticket to 
        ### Have them synced then wont need to pass different sets of variables.
            if switch_var.get() == 'Pre-Prod':
                lab = fields['lab']
                print('Pre-Production Mode')
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
            global url
            global linkUrl
            global icon
            global checkboxes
            global variables
            global checkbox_dict
            global linkCollection
            
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
                    
                    selected_site = site_option_selected.get()
                    selected_lab = lab_dict[selected_site]
                    selected_notify = notify_dict[selected_site]
                    selected_lead = lead_dict[selected_site]

                    print("Selected Program: {}".format(project_option_selected.get()))
                    print("Selected Site: " + selected_site)
                    print("Selected WW: {}".format(WorkWeekValue_Inside.get()))
                    print("Selected Year: {}".format(YearValue_Inside.get()))
                    print("Selected Customer: " + selected_lead)
                    print("Selected Notify PDL: " + selected_notify)
                    print("Selected Lab: " + selected_lab)


                    ### Checkbox check for keystone check boxes, Comment out for production                    
                    # for key, value in variables.items():
                    #     status = value['variable'].get()
                    #     text = value['text']
                    #     print("Status of checkbox ({} - {}): {}".format(key, text, status))
                    ### End Checkbox check
                    
                    for key, value in variables.items():
                        checkbox_dict.update({value['mile'].split('.')[0]:int(value['variable'].get())})

                    #Create a List of Dicts to store the complete tictet information. One ticket per row
                    fieldlist=[]

                    ### Test check for tickets that will be created based on check boxes  Comment out for deployment
                    print("")
                    print("##########################")
                    print("## HSD Milestone Status ##")
                    print("##########################")
                    for name, milestone in milestones.items():
                        # Retrieve the value associated with the key 'mile'
                        mile_value = milestone.get('mile').split('.')[0]
                        if checkbox_dict[mile_value] == 1:
                            print(f"  Enabled - {milestone.get('cb_title')} {milestone.get('title')}")
                        else:
                            print(f"  Disabled - {milestone.get('cb_title')} {milestone.get('title')}") 
                    ### end Test check

                    dictionaryloop=1
                    if dictionaryloop==1:
                        print("")
                        print("")
                        for name, milestone in milestones.items():
                            # Retrieve the value associated with the key 'mile'
                            mile_value = milestone.get('mile').split('.')[0]
                            if checkbox_dict[mile_value] == 1:
                                _title={}
                                _title["title"]=milestone.get('title')
                                        
                                _description={}
                                _description["description"]=milestone.get("description")
 
                                _required_by_milestone={}
                                _required_by_milestone["required_by_milestone"]=milestone.get("required_by_milestone")
 
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
                                r = datetime.datetime.strptime(d + '-1', "%Y%W-%w")
                                x = r - timedelta(weeks = int(milestone.get("ETA_WW")))
                                year = str(x.isocalendar()[0])
                                week = str(x.isocalendar()[1]).zfill(2)
                                milestoneww = (year + "-" + week)
                                _milestone_eta["milestone_eta"]=milestoneww

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
 
                                print("Creating - milestone ww = "+ str(milestoneww) + " " + milestone.get("title"))
 
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
                        
                        print("Using - " + str(url))
                        print("Using - " + str(linkUrl))
                        
                        for fields in fieldlist:
                            Button_Create["state"] = "disabled"
                            app.ticketInterval = app.ticketInterval +1
                            #----------------------------------------------------------
                            response = postnewHSD(fields) #Uncomment to create tickets
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
 
        # MS_1_selected = tk.StringVar(root)
        # MS_2_selected = tk.StringVar(root)
        # MS_3_selected = tk.StringVar(root)
 
        # entry_1 = tk.Entry(root)
        # entry_2 = tk.Entry(root)
        # entry_3 = tk.Entry(root)
 
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
        Button_info.place(x=417,y=132,width=65,height=30)
        Button_info["command"] = display_info
 
        Button_advance=ttk.Button(root)
        Button_advance["text"] = "->"
        Button_advance["tooltip"] = "Advance options."
        Button_advance.place(x=417,y=162,width=65,height=30)
        Button_advance["command"] = advance_window


    def CheckBox_SelectAll_command(self):
        variables
        if self.varSelectAll.get() == 1:
             for checkbox in variables.values():
                 checkbox['widget'].select()
        else:
             for checkbox in variables.values():
                 checkbox['widget'].deselect()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    canvas = Canvas()
    root.mainloop()