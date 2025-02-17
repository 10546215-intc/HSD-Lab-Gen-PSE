import tkinter as tk
from tkinter import ttk
from tkinter import *
import tkinter.messagebox
import tkinter.font as tkFont
from ttkwidgets import tooltips
import datetime
from datetime import timedelta
import csv
import json
import requests
from requests_kerberos import HTTPKerberosAuth
import urllib3
import os
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
milestone_vals_open = ""
pull_down_mode = "dynamic" # "dynamic" for HSD based "static" for Config file
milestones = {}
checkboxes = {}
variables = {}
mode_var = {}
src_var = {}
program_options = ()
site_options = ()
sites_options=[]
notify_dict={}
customer_options={}
backup_options={}
notify_options={}
lab_dict={}
lead_dict={}
static_vals = {}
#config_open = ""
checkbox_dict={}
static_vals_open = "" #static version number
linkCollection = ""
hsd_source = ""

#url_Prod = "https://hsdes-api.intel.com/"
#url_Pre = "https://hsdes-api-pre.intel.com/"

vernum = "0.98" #application version number
configver = "X" #config version number
dynamicver = "X" #dynamic version number
staticver = "X"
 
#menu_color = "#FF0000"
menu_color = "#e4e5ee"
option_color = "#D4D6E4"
 
def get_milestones():
    try:
        print('\nLoading milestone_vals:')
        with open("dependencies/milestone_vals.csv", encoding="utf8") as data_file:
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
        milestone_vals_open=True
    except:
        print('Failed to load Minestones!')
        milestone_vals_open=False

def mk_checkboxes():
    #########  Create milestone check boxes
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
                checkbox.place(x=10, y=160 + cb_num * 30, width=490, height=25)  # Adjust y position for each checkbox
                checkbox["offvalue"] = "0"
                checkbox["onvalue"] = "1"
                checkbox["variable"] = var
                checkbox.select()
                variables[key] = {'widget': checkbox, 'variable': var, 'text': checkbox["text"], 'mile': milestones[key]["mile"] }
                # Store the checkbox in the dictionary
                checkboxes[cb_num] = checkbox
                cb_num += 1
    except:
        print("Failed to create Milestone Checkboxes!")

def get_opt_menu_list(field, name):
    if name == 'program':
        source = prog_source.get()
    if name == 'Site':    
        source = ste_source.get()   

    tmp_dict=[]
    data_lower=[]
    sites_options.clear()
    customer_options.clear()
    backup_options.clear()
    notify_dict.clear()
    lab_dict.clear()
    lead_dict.clear()
    
    src_file = f"dependencies/{name}.csv"
    print("")
    print(f"Loading {name} menu selections:")
    print("")
    update_hsd_url()
    try:
        headers = {'Content-type': 'application/json'}
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
    if source == 'CSV':     
        try:
            with open(src_file, encoding="utf8") as f:    
                csv_config = csv.DictReader(f)
                
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
                        print(f" - {line[name]} - Not Found In HSD - {hsd_source.get()}uction , Check Name and update {src_file}" )
                        FileOpenMessage=f"Can not find {src_file}:\n"
                        #FileOpenMessage = FileOpenMessage + f'\ndependencies/{name}.csv'
            v_list = tmp_dict
            v_list.sort()
            #config_open=True
            
            # Function to update the options in the OptionMenu
            if name == "program":
                program_options = v_list
                menu = program_menu['menu']
                menu.delete(0, 'end')
                options_name = program_options
                opt_selected = project_option_selected
            elif name == "Site":
                site_options = v_list
                menu = site_menu['menu']
                menu.delete(0, 'end')
                options_name = site_options
                opt_selected = site_option_selected
            else:
                print(f"Menu not identified")
                
            for option in v_list:
                menu.add_command(label=option, command=tk._setit(opt_selected, option))
            #Optionally, set the default value to the first item in the new options
            if options_name:
                opt_selected.set('')
#                opt_selected.set(options_name[0])
            else:
                print("Options list unassigned") 
        except Exception as e:
            print(e)

### Load from HSD query and CSV File        
    if source == 'HSD':
        try:
            v_list = [(item.get('' + field_type + '', None)) for item in data if item.get('' + field_type + '', None)]   
            if name == 'Site': 
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
                            print(f" - {line[name]} - Not Found In HSD" )
                            FileOpenMessage = f"{line[name]} - Not Found In {hsd_source.get()}"

            # Function to update the options in the OptionMenu
            if name == "program":
                program_options = v_list
                menu = program_menu['menu']
                menu.delete(0, 'end')
                options_name = program_options
                opt_selected = project_option_selected
            elif name == "Site":
                site_options = v_list
                menu = site_menu['menu']
                menu.delete(0, 'end')
                options_name = site_options
                opt_selected = site_option_selected
            else:
                print(f"Menu not identified")
                
            for option in v_list:
                menu.add_command(label=option, command=tk._setit(opt_selected, option))
            #Optionally, set the default value to the first item in the new options
            if options_name:
                opt_selected.set('')
            else:
                print("Options list unassigned") 
        except:
            FileOpenMessage="Can not find file(s):\n"
            FileOpenMessage = FileOpenMessage + f'\ndependencies/{name}.csv'

def load_static_vals():
   ### Load Static_vals happens each time menu list are updated
    try:
        with open("dependencies/static_vals.csv", encoding="utf8") as f:
            print('Loading static_vals:')
            csv_static_vals = csv.DictReader(f)
            for line in csv_static_vals:
                static_vals[line['Item']] = line['Value']
        static_vals_open=True
    except:
        static_vals_open=False
        FileOpenMessage="Can not find file(s):\n"
        FileOpenMessage = FileOpenMessage + '\ndependencies/static_vals.csv'

def CheckBox_SelectAll_command():
    variables
    if varSelectAll.get() == 1:
            for checkbox in variables.values():
                checkbox['widget'].select()
    else:
            for checkbox in variables.values():
                checkbox['widget'].deselect()

def update_hsd_url(): 
    source = hsd_source.get()
    global url, linkUrl, icon, auto_c
    global auto_c
    #print(f"update_hsd_url called with source: {source}")
    #print("")
    mode_label.config(text=f"(HSD) {hsd_source.get()}uction")
    if source == "Pre-Prod":
        url = 'https://hsdes-api-pre.intel.com/rest/article'
        linkUrl = 'https://hsdes-pre.intel.com/appstore/article/#/'
        auto_c = 'https://hsdes-api-pre.intel.com/rest/query/autocomplete/support/services_sys_val/'
        icon = PhotoImage(file='dependencies/Y.png')
    if source == "Prod":
        url = 'https://hsdes-api.intel.com/rest/article'
        linkUrl = 'https://hsdes.intel.com/appstore/article/#/'
        auto_c = 'https://hsdes-api.intel.com/rest/query/autocomplete/support/services_sys_val/'
        icon = PhotoImage(file='dependencies/B.png') 

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
    milestone_eta = fields["milestone_eta"]
    required_by_milestone = fields["required_by_milestone"]
    survey_comment = fields["survey_comment"]
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
            "send_mail": varSend_email.get()
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

    # if notify == '':
    #     print('no notify')
    #     readyMessage = readyMessage + "No Notify chosen.\n"
    #     exitFunction=TRUE

    if lab == '':
        print('no Lab')
        readyMessage = readyMessage + "No Lab chosen.\n"
        exitFunction=TRUE

    readyMessage = readyMessage + "\nCheck config.csv file!"
    if exitFunction==TRUE:
        ticketInterval = ticketInterval
        total_tickets = total_tickets
        
        if ticketInterval == total_tickets:
            print(readyMessage)
            tkinter.messagebox.showinfo('Config Error',readyMessage)
        return None
    else:
        data = json.dumps(payload)
        response = requests.post(url, verify=False,auth=HTTPKerberosAuth(), headers = headers, data = data)
        return response.json()

#--------------------------------------------------------------------------------------------------------------------
#### Start Here------------------------------------------------------
def build_ticket_details():
    #global icon
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

            print("")
            print("################################")
            print("## HSD Ticket Creation Report ##")
            print("################################")
            print("")
            print(f"{hsd_source.get()}uction Mode:")
            print("  - HSD Mode Using - " + str(url))
            print("  - HSD Mode  - " + str(linkUrl))
            print("")
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

            ### Test check for tickets that will be created based on check boxes enabled / disabled  Comment out for deployment
            # print("")
            # print("##########################")
            # print("## HSD Milestone Status ##")
            # print("##########################")
            #for name, milestone in milestones.items():
                # mile_value = milestone.get('mile').split('.')[0]
                # if checkbox_dict[mile_value] == 1:
                #     print(f"  Enabled - {milestone.get('cb_title')} {milestone.get('title')}")
                # else:
                #     print(f"  Disabled - {milestone.get('cb_title')} {milestone.get('title')}") 
                ### end Test check

            dictionaryloop=1
            if dictionaryloop==1:
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
                total_tickets = total_tickets

            ticketInterval = ticketInterval

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
                    ticketInterval = ticketInterval +1
                    #----------------------------------------------------------
                    response = postnewHSD(fields) #Uncomment to create tickets
                    #----------------------------------------------------------
                    #print('ticket_number',ticket_number)

                    try:                         
                        tickets.append(response['new_id'])
                        print(f"Created {hsd_source.get()}uction ticket:",response['new_id'])
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
                    Label_ProgressSuccess["text"]= f"Created: {hsd_source.get()}uction " + str(ticket_number) + " / " + str(total_tickets) + " tickets."
                    root.after(1,ProgressBar.update(),Label_ProgressSuccess.update())

                display_tickets='\n\n'

                for line in HSD_ID:
                    display_tickets=display_tickets + str(line) + '\n'

                final_message=""
                final_message=f"Successfully created {hsd_source.get()}uction " + str(ticket_number) + " of " + str(total_tickets) + " tickets. " + str(display_tickets)

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


def saveClipboard(my_string):
    if clipboard == 'TRUE':
        text_file = open(r'dependencies/Clipboard.txt', 'w')
        text_file.write(my_string)
        text_file.close()
        command = 'clip < dependencies/Clipboard.txt'
        os.system(command)

#Close GUI and Python
def close():
    exec
    root.destroy()
    
#########################################################################################################
root = tk.Tk()
root.title("HSD-Lab Gen - Version " + str(vernum))
root.eval("tk::PlaceWindow . center")
# Setting window size
width = 492
height = 740

# Get the screen width and height
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()

# Calculate the x and y coordinates for the window
x = (screenwidth - width) // 2 - 400
y = (screenheight - height) // 8 # Adjust this value to move the window higher up


alignstr = f'{width}x{height}+{x}+{y}'
root.geometry(alignstr)

root.geometry('494x740')
root.resizable(width=False, height=False)

# Create a frame to hold the canvas
frame = tk.Frame(root, width=width, height=height)
frame.pack(fill="both", expand=True)

# Create the canvas
cv = tk.Canvas(frame, width=width, height=height)
cv.pack(fill="both", expand=True)

# Rectangle dimensions
rect_height = 88  # Adjusted to fit within the canvas height

# Calculate the top-left and bottom-right coordinates of the rectangle
rect_x1 = 2  # Start from the left edge of the canvas
rect_y1 = 38  # Fixed y-coordinate for the top of the rectangle
rect_x2 = width -2  # End at the right edge of the canvas
rect_y2 = rect_y1 + rect_height

cv.create_rectangle(rect_x1, rect_y1, rect_x2, rect_y2, fill=menu_color)

# Adjust the lines to be within the rectangle
line1_x = rect_x1 + (width / 3)
line2_x = rect_x1 + (2 * width / 3)
cv.create_line(line1_x, rect_y1, line1_x, rect_y2)
cv.create_line(line2_x, rect_y1, line2_x, rect_y2)

#Add Menu for Options
menu = tk.Menu(root)

project_option_selected = tk.StringVar(root)
program_menu=ttk.OptionMenu(root,project_option_selected,*program_options)
program_menu['tooltip'] = 'Choose program / project.'
program_menu.place(x=10,y=69,width=150,height=25)

site_option_selected = tk.StringVar(root)
site_menu=ttk.OptionMenu(root,site_option_selected,*site_options)
site_menu['tooltip'] = 'Choose Site.'
site_menu.place(x=168,y=69,width=150,height=25)

def advance_window(tog=[0]):
    tog[0] = not tog[0]
    if tog[0]:
        print("Advance Window Open")
        root.geometry("1292x740")
        Button_advance["text"] = "<-"
    else:
        print("Advance Window Close")
        root.geometry("494x740")
        Button_advance["text"] = "->"

hsd_menu = tk.Menu(menu, tearoff=False)
menu.add_cascade(label='HSD Sources', menu=hsd_menu)
# Use a single StringVar to hold the value of the selected radio button
hsd_source = tk.StringVar(value='Pre-Prod')
# Add radio buttons to the menu
hsd_menu.add_radiobutton(label='HSD-Prod', value='Prod', variable=hsd_source,command=update_hsd_url)
hsd_menu.add_radiobutton(label='HSD-Pre', value='Pre-Prod', variable=hsd_source,command=update_hsd_url)

# Create a label to display the current mode
mode_label = tk.Label(root, text=f"{hsd_source.get()}uction", font=("Times", 16), fg="blue")
mode_label.place(x=150, y=4)

prog_menu = tk.Menu(menu, tearoff=False)
menu.add_cascade(label='Program Sources', menu=prog_menu)
# Use a single StringVar to hold the value of the selected radio button
prog_source = tk.StringVar(value = 'CSV')
# Add radio buttons to the menu
prog_menu.add_radiobutton(label='HSD', value='HSD', variable=prog_source, command=lambda: get_opt_menu_list('services_sys_val.support.program','program'))
prog_menu.add_radiobutton(label='CSV', value='CSV', variable=prog_source, command=lambda: get_opt_menu_list('services_sys_val.support.program','program'))

ste_menu = tk.Menu(menu, tearoff=False)
menu.add_cascade(label='Site Sources', menu=ste_menu)
# Use a single StringVar to hold the value of the selected radio button
ste_source = tk.StringVar(value = 'CSV')
# Add radio buttons to the menu
ste_menu.add_radiobutton(label='HSD', value='HSD', variable=ste_source, command=lambda: get_opt_menu_list('support.site','Site'))
ste_menu.add_radiobutton(label='CSV', value='CSV', variable=ste_source, command=lambda: get_opt_menu_list('support.site','Site'))
ste_menu = tk.Menu(menu, tearoff=False)

root.configure(menu = menu)

Button_advance=ttk.Button(root)
Button_advance["text"] = "->"
Button_advance["tooltip"] = "Advance options."
Button_advance.place(x=417,y=4,width=65,height=30)
Button_advance["command"] = advance_window   

#ProgressSuccess Label
Label_ProgressSuccess=tk.Label(root)
ft = tkFont.Font(family='Times',size=14)
Label_ProgressSuccess["font"] = ft
Label_ProgressSuccess["fg"] = "#333333"
Label_ProgressSuccess["justify"] = "left"
Label_ProgressSuccess["text"] = ""
Label_ProgressSuccess.place(x=0,y=340,width=492,height=20)


get_opt_menu_list('services_sys_val.support.program','program') 
get_opt_menu_list('support.site','Site') 
get_milestones()
load_static_vals()
#UpdateIcon()

customer_option_selected = tk.StringVar(root)
notify_option_selected = tk.StringVar(root)

clipboard = static_vals['createClipboard']

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
Label_PODate = tk.Label(root, text="Power On Date", font=("Times", 14), fg="#333333", bg=menu_color, anchor="c")
Label_PODate.place(x=335, y=41, width=150, height=25)

# WorkWeek Label
Label_WorkWeek = tk.Label(root, text="WW#", font=("Times", 14), fg="#333333", bg=menu_color, anchor="e")
Label_WorkWeek.place(x=336, y=69, width=60, height=25)

# Year Label
Label_Year = tk.Label(root, text="Year", font=("Times", 14), fg="#333333", bg=menu_color, anchor="e")
Label_Year.place(x=336, y=99, width=60, height=25)

# Create Select All checkbox
varSelectAll=tk.IntVar(root)
CheckBox_SelectAll=tk.Checkbutton(root)
CheckBox_SelectAll["anchor"] = "w"
#ft = tkFont.Font(family='Times',size=10)
CheckBox_SelectAll["font"] = tkFont.Font(family='Times', size=10)
CheckBox_SelectAll["fg"] = "#333333"
CheckBox_SelectAll["justify"] = "left"
CheckBox_SelectAll["text"] = "Select All"
CheckBox_SelectAll.place(x=10,y=160,width=110,height=25)
CheckBox_SelectAll["offvalue"] = "0"
CheckBox_SelectAll["onvalue"] = "1"
CheckBox_SelectAll["command"] = CheckBox_SelectAll_command
CheckBox_SelectAll["variable"] = varSelectAll
CheckBox_SelectAll.select()

#Create Send Emails checkbox
varSend_email=tk.StringVar(root)
CheckBox_send_email=tk.Checkbutton(root)
CheckBox_send_email["anchor"] = "w"
CheckBox_send_email["font"] = tkFont.Font(family='Times', size=10)
CheckBox_send_email["fg"] = "#333333"
CheckBox_send_email["justify"] = "left"
CheckBox_send_email["text"] = "Send Notification Email"
CheckBox_send_email.place(x=340,y=132,width=160,height=25)
CheckBox_send_email["offvalue"] = "false"
CheckBox_send_email["onvalue"] = "true"
CheckBox_send_email["variable"] = varSend_email
CheckBox_send_email.select()

# Create milestone check boxes
mk_checkboxes()

# Milestones Label
Label_MS=tk.Label(root)
Label_MS["anchor"] = "w"
ft = tkFont.Font(family='Times',size=14, weight="bold")
Label_MS["font"] = ft
Label_MS["fg"] = "#333333"
Label_MS["justify"] = "left"
Label_MS["text"] = "Milestones:"
Label_MS["relief"] = "flat"
Label_MS.place(x=10,y=131,width=186,height=30)

# Progress Label
Label_ProgressSuccess.place(x=30,y=635,width=246,height=20)
Label_ProgressSuccess["text"] = "Ready"
root.after(1,Label_ProgressSuccess.update())

total_tickets = 0
ticket_number = 0

ProgressBar=ttk.Progressbar(root)
ProgressBar.place(x=6,y=665,width=480,height=10)

Button_Create=ttk.Button(root)
Button_Create["text"] = "Create"
Button_Create["tooltip"] = "Create selected milestone tickets and copies links to them in the clipboard cache."
Button_Create.place(x=320,y=630,width=158,height=30)
Button_Create["command"] = build_ticket_details

# if __name__ == "__main__":

root.mainloop()