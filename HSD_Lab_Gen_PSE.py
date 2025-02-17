import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkwidgets import tooltips
import datetime
from datetime import timedelta
import csv
import json
import requests
from requests_kerberos import HTTPKerberosAuth
import urllib3
import os
import webbrowser

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
os.environ["no_proxy"] = "127.0.0.1,localhost,intel.com"

class HSDLabGenApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HSD-Lab Gen - Version 0.98")
        self.root.geometry('494x740')
        self.root.geometry('494x740+100+10')  # Adjust the x_offset and y_offset as needed
        self.root.resizable(width=False, height=False)

        self.url = ""
        self.linkUrl = ""
        self.auto_c = ""
        self.icon = ""
        self.milestone_vals_open = ""
        self.pull_down_mode = "dynamic"
        self.milestones = {}
        self.checkboxes = {}
        self.variables = {}
        self.mode_var = {}
        self.src_var = {}
        self.program_options = ()
        self.site_options = ()
        self.sites_options = []
        self.notify_dict = {}
        self.customer_options = {}
        self.backup_options = {}
        self.notify_options = {}
        self.lab_dict = {}
        self.lead_dict = {}
        self.static_vals = {}
        self.checkbox_dict = {}
        self.static_vals_open = ""
        self.linkCollection = ""
        self.hsd_source = ""
        self.vernum = "0.98"
        self.configver = "X"
        self.dynamicver = "X"
        self.staticver = "X"

        self.menu_color = "#e4e5ee"
        self.option_color = "#D4D6E4"

        self.setup_ui()

    def setup_ui(self):
        # Create a frame to hold the canvas
        frame = tk.Frame(self.root, width=492, height=740)
        frame.pack(fill="both", expand=True)

        # Create the canvas
        cv = tk.Canvas(frame, width=492, height=740)
        cv.pack(fill="both", expand=True)

        # Rectangle dimensions
        rect_height = 88

        # Calculate the top-left and bottom-right coordinates of the rectangle
        rect_x1 = 2
        rect_y1 = 38
        rect_x2 = 492 - 2
        rect_y2 = rect_y1 + rect_height

        cv.create_rectangle(rect_x1, rect_y1, rect_x2, rect_y2, fill=self.menu_color)

        # Adjust the lines to be within the rectangle
        line1_x = rect_x1 + (492 / 3)
        line2_x = rect_x1 + (2 * 492 / 3)
        cv.create_line(line1_x, rect_y1, line1_x, rect_y2)
        cv.create_line(line2_x, rect_y1, line2_x, rect_y2)

        # Add Menu for Options
        menu = tk.Menu(self.root)

        self.project_option_selected = tk.StringVar(self.root)
        self.program_menu = ttk.OptionMenu(self.root, self.project_option_selected, *self.program_options)
        self.program_menu['tooltip'] = 'Choose program / project.'
        self.program_menu.place(x=10, y=69, width=150, height=25)

        self.site_option_selected = tk.StringVar(self.root)
        self.site_menu = ttk.OptionMenu(self.root, self.site_option_selected, *self.site_options)
        self.site_menu['tooltip'] = 'Choose Site.'
        self.site_menu.place(x=168, y=69, width=150, height=25)

        self.hsd_source = tk.StringVar(value='Pre-Prod')
        hsd_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='HSD Sources', menu=hsd_menu)
        hsd_menu.add_radiobutton(label='HSD-Prod', value='Prod', variable=self.hsd_source, command=self.update_hsd_url)
        hsd_menu.add_radiobutton(label='HSD-Pre', value='Pre-Prod', variable=self.hsd_source, command=self.update_hsd_url)

        self.mode_label = tk.Label(self.root, text=f"{self.hsd_source.get()}uction", font=("Times", 16), fg="blue")
        self.mode_label.place(x=150, y=4)

        self.prog_source = tk.StringVar(value='CSV')
        prog_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='Program Sources', menu=prog_menu)
        prog_menu.add_radiobutton(label='HSD', value='HSD', variable=self.prog_source, command=lambda: self.get_opt_menu_list('services_sys_val.support.program', 'program'))
        prog_menu.add_radiobutton(label='CSV', value='CSV', variable=self.prog_source, command=lambda: self.get_opt_menu_list('services_sys_val.support.program', 'program'))

        self.ste_source = tk.StringVar(value='CSV')
        ste_menu = tk.Menu(menu, tearoff=False)
        menu.add_cascade(label='Site Sources', menu=ste_menu)
        ste_menu.add_radiobutton(label='HSD', value='HSD', variable=self.ste_source, command=lambda: self.get_opt_menu_list('support.site', 'Site'))
        ste_menu.add_radiobutton(label='CSV', value='CSV', variable=self.ste_source, command=lambda: self.get_opt_menu_list('support.site', 'Site'))

        self.root.configure(menu=menu)

        self.Button_advance = ttk.Button(self.root, text="->", command=self.advance_window)
        self.Button_advance["tooltip"] = "Advance options."
        self.Button_advance.place(x=417, y=4, width=65, height=30)

        self.Label_ProgressSuccess = tk.Label(self.root, font=tk.font.Font(family='Times', size=14), fg="#333333", justify="left", text="")
        self.Label_ProgressSuccess.place(x=0, y=340, width=492, height=20)

        self.get_opt_menu_list('services_sys_val.support.program', 'program')
        self.get_opt_menu_list('support.site', 'Site')
        self.get_milestones()
        self.load_static_vals()

        self.customer_option_selected = tk.StringVar(self.root)
        self.notify_option_selected = tk.StringVar(self.root)

        self.clipboard = self.static_vals.get('createClipboard', 'FALSE')

        # Create List for WorkWeek OptionMenu
        self.WorkWeekList = list(range(53))

        # Get the Date
        today = datetime.date.today()
        iso_calendar = today.isocalendar()
        self.WorkWeek = iso_calendar[1]
        self.year = iso_calendar[0]
        self.WorkWeekValue_Inside = tk.StringVar(self.root)
        self.YearValue_Inside = tk.StringVar(self.root)

        # Create List for Year OptionMenu
        self.YearList = list(range(2021, 2051))

        self.Option_101 = ttk.OptionMenu(self.root, self.WorkWeekValue_Inside, *self.WorkWeekList)
        self.Option_101['tooltip'] = 'Choose todays known work week for power on.'
        self.Option_101.place(x=400, y=69, width=80, height=25)
        self.WorkWeekValue_Inside.set(self.WorkWeek)

        self.Option_403 = ttk.OptionMenu(self.root, self.YearValue_Inside, *self.YearList)
        self.Option_403['tooltip'] = 'Choose todays known year for power on.'
        self.Option_403.place(x=400, y=99, width=80, height=25)
        self.YearValue_Inside.set(self.year)

        # Create labels for pulldown Boxes
        self.Label_Program = tk.Label(self.root, text="Program", font=("Times", 14), fg="#333333", bg=self.menu_color, anchor="c", justify="left")
        self.Label_Program.place(x=10, y=41, width=150, height=25)

        self.Label_Site = tk.Label(self.root, text="Site", font=("Times", 14), fg="#333333", bg=self.menu_color, anchor="c", justify="center")
        self.Label_Site.place(x=168, y=41, width=150, height=25)

        self.Label_PODate = tk.Label(self.root, text="Power On Date", font=("Times", 14), fg="#333333", bg=self.menu_color, anchor="c")
        self.Label_PODate.place(x=335, y=41, width=150, height=25)

        self.Label_WorkWeek = tk.Label(self.root, text="WW#", font=("Times", 14), fg="#333333", bg=self.menu_color, anchor="e")
        self.Label_WorkWeek.place(x=336, y=69, width=60, height=25)

        self.Label_Year = tk.Label(self.root, text="Year", font=("Times", 14), fg="#333333", bg=self.menu_color, anchor="e")
        self.Label_Year.place(x=336, y=99, width=60, height=25)

        # Create Select All checkbox
        self.varSelectAll = tk.IntVar(self.root)
        self.CheckBox_SelectAll = tk.Checkbutton(self.root, anchor="w", font=tk.font.Font(family='Times', size=10), fg="#333333", justify="left", text="Select All", offvalue="0", onvalue="1", command=self.CheckBox_SelectAll_command, variable=self.varSelectAll)
        self.CheckBox_SelectAll.place(x=10, y=160, width=110, height=25)
        self.CheckBox_SelectAll.select()

        # Create Send Emails checkbox
        self.varSend_email = tk.StringVar(self.root)
        self.CheckBox_send_email = tk.Checkbutton(self.root, anchor="w", font=tk.font.Font(family='Times', size=10), fg="#333333", justify="left", text="Send Notification Email", offvalue="false", onvalue="true", variable=self.varSend_email)
        self.CheckBox_send_email.place(x=340, y=132, width=160, height=25)
        self.CheckBox_send_email.select()

        # Create milestone check boxes
        self.mk_checkboxes()

        # Milestones Label
        self.Label_MS = tk.Label(self.root, anchor="w", font=tk.font.Font(family='Times', size=14, weight="bold"), fg="#333333", justify="left", text="Milestones:", relief="flat")
        self.Label_MS.place(x=10, y=131, width=186, height=30)

        # Progress Label
        self.Label_ProgressSuccess.place(x=30, y=635, width=246, height=20)
        self.Label_ProgressSuccess["text"] = "Ready"
        self.root.after(1, self.Label_ProgressSuccess.update())

        self.total_tickets = 0
        self.ticket_number = 0

        self.ProgressBar = ttk.Progressbar(self.root)
        self.ProgressBar.place(x=6, y=665, width=480, height=10)

        self.Button_Create = ttk.Button(self.root, text="Create", command=self.build_ticket_details)
        self.Button_Create["tooltip"] = "Create selected milestone tickets and copies links to them in the clipboard cache."
        self.Button_Create.place(x=320, y=630, width=158, height=30)

    def get_milestones(self):
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
                    self.milestones[name] = temp_dict
            self.milestone_vals_open = True
        except:
            print('Failed to load Milestones!')
            self.milestone_vals_open = False

    def mk_checkboxes(self):
        print('')
        cb_num = 1
        start_y = 196
        keystones = 0

        for key, item in self.milestones.items():
            if item.get("keystone") == "1":
                keystones += 1
        try:
            for key, item in self.milestones.items():
                if item.get("keystone") == "1":
                    var = tk.IntVar(self.root)
                    checkbox = tk.Checkbutton(self.root)
                    checkbox["anchor"] = "w"
                    checkbox["font"] = tk.font.Font(family='Times', size=10)
                    checkbox["fg"] = "#333333"
                    checkbox["justify"] = "left"
                    checkbox["text"] = self.milestones[key]["cb_title"]
                    checkbox.place(x=10, y=160 + cb_num * 30, width=490, height=25)
                    checkbox["offvalue"] = "0"
                    checkbox["onvalue"] = "1"
                    checkbox["variable"] = var
                    checkbox.select()
                    self.variables[key] = {'widget': checkbox, 'variable': var, 'text': checkbox["text"], 'mile': self.milestones[key]["mile"]}
                    self.checkboxes[cb_num] = checkbox
                    cb_num += 1
        except:
            print("Failed to create Milestone Checkboxes!")

    def get_opt_menu_list(self, field, name):
        if name == 'program':
            source = self.prog_source.get()
        if name == 'Site':
            source = self.ste_source.get()

        tmp_dict = []
        data_lower = []
        self.sites_options.clear()
        self.customer_options.clear()
        self.backup_options.clear()
        self.notify_dict.clear()
        self.lab_dict.clear()
        self.lead_dict.clear()

        src_file = f"dependencies/{name}.csv"
        print("")
        print(f"Loading {name} menu selections:")
        print("")
        self.update_hsd_url()
        try:
            headers = {'Content-type': 'application/json'}
            url_validate = f'{self.auto_c}/{field}'
            field_type = field
            response = requests.get(url_validate, verify=False, auth=HTTPKerberosAuth(), headers=headers)

            if response.status_code != 200:
                return []

            data = response.json().get("data", [])
            data_lower = [(item.get('' + field_type + '', None)) for item in data if item.get('' + field_type + '', None)]

            if not isinstance(data, list):
                return []

        except requests.RequestException as e:
            tk.messagebox.showwarning('Connection Error', 'Not connected to HSDE-ES DB.')
            return []
        except Exception as e:
            tk.messagebox.showwarning('Connection Error', 'Not connected to HSDE-ES DB.')
            return []

        if source == 'CSV':
            try:
                with open(src_file, encoding="utf8") as f:
                    csv_config = csv.DictReader(f)

                    for line in csv_config:
                        if (line[name]) and line.get(name).lower() in data_lower:
                            tmp_dict.append(line[name])
                            if name == 'Site':
                                self.sites_options.append(line['Site'])
                                self.customer_options.update({line['Site']: line['Customer']})
                                self.backup_options.update({line['Site']: line['Backup']})
                                self.notify_dict.update({line['Site']: line['Notify']})
                                self.lab_dict.update({line['Site']: line['Lab']})
                                self.lead_dict.update({line['Site']: line['Customer']})
                        else:
                            print(f" - {line[name]} - Not Found In HSD - {self.hsd_source.get()}uction , Check Name and update {src_file}")
                v_list = tmp_dict
                v_list.sort()

                if name == "program":
                    self.program_options = v_list
                    menu = self.program_menu['menu']
                    menu.delete(0, 'end')
                    options_name = self.program_options
                    opt_selected = self.project_option_selected
                elif name == "Site":
                    self.site_options = v_list
                    menu = self.site_menu['menu']
                    menu.delete(0, 'end')
                    options_name = self.site_options
                    opt_selected = self.site_option_selected
                else:
                    print(f"Menu not identified")

                for option in v_list:
                    menu.add_command(label=option, command=tk._setit(opt_selected, option))
                if options_name:
                    opt_selected.set('')
                else:
                    print("Options list unassigned")
            except Exception as e:
                print(e)

        if source == 'HSD':
            try:
                v_list = [(item.get('' + field_type + '', None)) for item in data if item.get('' + field_type + '', None)]
                if name == 'Site':
                    with open(src_file, encoding="utf8") as f:
                        csv_config = csv.DictReader(f)
                        print('\nloading config and verifying options')
                        for line in csv_config:
                            if (line[name]) and line.get(name).lower() in data_lower:
                                tmp_dict.append(line[name])
                                if name == 'Site':
                                    self.sites_options.append(line['Site'])
                                    self.customer_options.update({line['Site']: line['Customer']})
                                    self.backup_options.update({line['Site']: line['Backup']})
                                    self.notify_dict.update({line['Site']: line['Notify']})
                                    self.lab_dict.update({line['Site']: line['Lab']})
                                    self.lead_dict.update({line['Site']: line['Customer']})
                            else:
                                print(f" - {line[name]} - Not Found In HSD")
                                FileOpenMessage = f"{line[name]} - Not Found In {self.hsd_source.get()}"

                if name == "program":
                    self.program_options = v_list
                    menu = self.program_menu['menu']
                    menu.delete(0, 'end')
                    options_name = self.program_options
                    opt_selected = self.project_option_selected
                elif name == "Site":
                    self.site_options = v_list
                    menu = self.site_menu['menu']
                    menu.delete(0, 'end')
                    options_name = self.site_options
                    opt_selected = self.site_option_selected
                else:
                    print(f"Menu not identified")

                for option in v_list:
                    menu.add_command(label=option, command=tk._setit(opt_selected, option))
                if options_name:
                    opt_selected.set('')
                else:
                    print("Options list unassigned")
            except:
                FileOpenMessage = "Can not find file(s):\n"
                FileOpenMessage = FileOpenMessage + f'\ndependencies/{name}.csv'

    def load_static_vals(self):
        try:
            with open("dependencies/static_vals.csv", encoding="utf8") as f:
                print('Loading static_vals:')
                csv_static_vals = csv.DictReader(f)
                for line in csv_static_vals:
                    self.static_vals[line['Item']] = line['Value']
            self.static_vals_open = True
        except:
            self.static_vals_open = False
            FileOpenMessage = "Can not find file(s):\n"
            FileOpenMessage = FileOpenMessage + '\ndependencies/static_vals.csv'

    def CheckBox_SelectAll_command(self):
        if self.varSelectAll.get() == 1:
            for checkbox in self.variables.values():
                checkbox['widget'].select()
        else:
            for checkbox in self.variables.values():
                checkbox['widget'].deselect()

    def update_hsd_url(self):
        source = self.hsd_source.get()
        if source == "Pre-Prod":
            self.url = 'https://hsdes-api-pre.intel.com/rest/article'
            self.linkUrl = 'https://hsdes-pre.intel.com/appstore/article/#/'
            self.auto_c = 'https://hsdes-api-pre.intel.com/rest/query/autocomplete/support/services_sys_val/'
            self.icon = tk.PhotoImage(file='dependencies/Y.png')
        if source == "Prod":
            self.url = 'https://hsdes-api.intel.com/rest/article'
            self.linkUrl = 'https://hsdes.intel.com/appstore/article/#/'
            self.auto_c = 'https://hsdes-api.intel.com/rest/query/autocomplete/support/services_sys_val/'
            self.icon = tk.PhotoImage(file='dependencies/B.png')

    def postnewHSD(self, fields):
        headers = {'Content-type': 'application/json'}
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
                {"title": title},
                {"description": description},
                {"services_sys_val.support.service_type": service_type},
                {"services_sys_val.support.service_sub_type": service_sub_type},
                {"services_sys_val.support.lab_org": lab_org},
                {"services_sys_val.support.category": category},
                {"component": component},
                {"priority": priority},
                {"support.customer_contact": customer_contact},
                {"support.site": site},
                {"services_sys_val.support.lab": lab},
                {"notify": notify},
                {"services_sys_val.support.org_unit": org_unit},
                {"services_sys_val.support.program": program},
                {"services_sys_val.support.milestone_eta": milestone_eta},
                {"send_mail": self.varSend_email.get()},
                {"services_sys_val.support.required_by_milestone": required_by_milestone},
                {"services_sys_val.support.survey_comment": survey_comment}
            ]
        }

        readyMessage = ''
        total_tickets = 0
        exitFunction = False
        if customer_contact == '':
            print('no customer')
            readyMessage = "No Customer chosen.\n"
            exitFunction = True

        if lab == '':
            print('no Lab')
            readyMessage = readyMessage + "No Lab chosen.\n"
            exitFunction = True

        readyMessage = readyMessage + "\nCheck config.csv file!"
        if exitFunction:
            ticketInterval = ticketInterval
            total_tickets = total_tickets

            if ticketInterval == total_tickets:
                print(readyMessage)
                tk.messagebox.showinfo('Config Error', readyMessage)
            return None
        else:
            data = json.dumps(payload)
            response = requests.post(self.url, verify=False, auth=HTTPKerberosAuth(), headers=headers, data=data)
            return response.json()

    def build_ticket_details(self):
        not_ready = ''
        not_ready1 = ''
        not_ready2 = ''

        if (len(self.project_option_selected.get()) == 0):
            not_ready1 = 'Select Program\n'

        if (len(self.site_option_selected.get()) == 0):
            not_ready2 = 'Select Site\n'

        not_ready = (not_ready1 + not_ready2)
        print(not_ready)

        if (len(not_ready) == 0):
            result = tk.messagebox.askquestion('Create Tickets?', 'Are you sure you want to create tickets?')

            if result == 'yes':
                self.Label_ProgressSuccess["text"] = "Connecting to HSD-ES DB"
                self.root.after(1, self.Label_ProgressSuccess.update())

                selected_site = self.site_option_selected.get()
                selected_lab = self.lab_dict[selected_site]
                selected_notify = self.notify_dict[selected_site]
                selected_lead = self.lead_dict[selected_site]

                print("")
                print("################################")
                print("## HSD Ticket Creation Report ##")
                print("################################")
                print("")
                print(f"{self.hsd_source.get()}uction Mode:")
                print("  - HSD Mode Using - " + str(self.url))
                print("  - HSD Mode  - " + str(self.linkUrl))
                print("")
                print("Selected Program: {}".format(self.project_option_selected.get()))
                print("Selected Site: " + selected_site)
                print("Selected WW: {}".format(self.WorkWeekValue_Inside.get()))
                print("Selected Year: {}".format(self.YearValue_Inside.get()))
                print("Selected Customer: " + selected_lead)
                print("Selected Notify PDL: " + selected_notify)
                print("Selected Lab: " + selected_lab)

                for key, value in self.variables.items():
                    self.checkbox_dict.update({value['mile'].split('.')[0]: int(value['variable'].get())})

                fieldlist = []

                dictionaryloop = 1
                if dictionaryloop == 1:
                    print("")
                    for name, milestone in self.milestones.items():
                        mile_value = milestone.get('mile').split('.')[0]
                        if self.checkbox_dict[mile_value] == 1:
                            _title = {"title": milestone.get('title')}
                            _description = {"description": milestone.get("description")}
                            _required_by_milestone = {"required_by_milestone": milestone.get("required_by_milestone")}
                            _lab_org = {"lab_org": self.static_vals.get("lab_org")}
                            _org_unit = {"org_unit": self.static_vals.get("org_unit")}
                            _category = {"category": self.static_vals.get("category")}
                            _component = {"component": self.static_vals.get("component")}
                            _priority = {"priority": self.static_vals.get("priority")}
                            _status = {"status": self.static_vals.get("status")}
                            _reason = {"reason": self.static_vals.get("reason")}
                            _customer_contact = {"customer_contact": selected_lead}
                            _site = {"site": selected_site}
                            _program = {"program": self.project_option_selected.get()}

                            d = self.YearValue_Inside.get() + self.WorkWeekValue_Inside.get()
                            r = datetime.datetime.strptime(d + '-1', "%Y%W-%w")
                            x = r - timedelta(weeks=int(milestone.get("ETA_WW")))
                            year = str(x.isocalendar()[0])
                            week = str(x.isocalendar()[1]).zfill(2)
                            milestoneww = (year + "-" + week)
                            _milestone_eta = {"milestone_eta": milestoneww}

                            _service_type = {"service_type": self.static_vals.get("service_type")}
                            _service_sub_type = {"service_sub_type": self.static_vals.get("service_sub_type")}
                            _survey_comment = {"survey_comment": self.static_vals.get("survey_comment")}
                            _lab = {'lab': selected_lab}
                            _notify = {'notify': selected_notify}

                            line_dict = {}
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

                            print("Creating - milestone ww = " + str(milestoneww) + " " + milestone.get("title"))

                print("")
                total_tickets = 0
                ticketInterval = 0

                for i in fieldlist:
                    total_tickets = total_tickets + 1
                    total_tickets = total_tickets

                ticketInterval = ticketInterval

                if total_tickets == 0:
                    print('No Tickets')
                    tk.messagebox.showinfo('No Milestones Selected', 'Select at least one Milestone.')
                else:
                    self.ProgressBar['value'] = 0
                    ticket_number = 0
                    error_number = 0
                    tickets = []
                    errors = []
                    HSD_ID = []

                    for fields in fieldlist:
                        self.Button_Create["state"] = "disabled"
                        ticketInterval = ticketInterval + 1
                        response = self.postnewHSD(fields)

                        try:
                            tickets.append(response['new_id'])
                            print(f"Created {self.hsd_source.get()}uction ticket:", response['new_id'])
                            HSD_ID.append(response['new_id'])
                            ticket_number = ticket_number + 1
                        except:
                            try:
                                errors.append(response['message'])
                                print('error:', response['message'])
                                HSD_ID.append(response['message'])
                                error_number = error_number + 1
                            except:
                                print('error')

                        self.ProgressBar['value'] += 100 / total_tickets
                        self.Label_ProgressSuccess["text"] = f"Created: {self.hsd_source.get()}uction " + str(ticket_number) + " / " + str(total_tickets) + " tickets."
                        self.root.after(1, self.ProgressBar.update(), self.Label_ProgressSuccess.update())

                    display_tickets = '\n\n'

                    for line in HSD_ID:
                        display_tickets = display_tickets + str(line) + '\n'

                    final_message = ""
                    final_message = f"Successfully created {self.hsd_source.get()}uction " + str(ticket_number) + " of " + str(total_tickets) + " tickets. " + str(display_tickets)

                    if error_number > 1:
                        final_message = final_message + str(error_number) + " tickets had errors"

                    tk.messagebox.showinfo('Tickets Created?', final_message)

                    self.linkCollection = ""

                    if error_number == 0:
                        for line in HSD_ID:
                            link = ""
                            link = self.linkUrl + str(line)
                            self.linkCollection = self.linkCollection + link + " \n"
                        self.linkCollection = self.linkCollection[0:(len(self.linkCollection) - 2)]
                        self.saveClipboard(self.linkCollection)

                self.Label_ProgressSuccess["text"] = "Ready"
                self.root.after(1, self.Label_ProgressSuccess.update())
                self.Button_Create["state"] = "enabled"

                return
            else:
                return None

        tk.messagebox.showerror('Not Ready!', not_ready)

    def saveClipboard(self, my_string):
        if self.clipboard == 'TRUE':
            text_file = open(r'dependencies/Clipboard.txt', 'w')
            text_file.write(my_string)
            text_file.close()
            command = 'clip < dependencies/Clipboard.txt'
            os.system(command)

    def advance_window(self, tog=[0]):
        tog[0] = not tog[0]
        if tog[0]:
            print("Advance Window Open")
            self.root.geometry("1292x740")
            self.Button_advance["text"] = "<-"

            # Create a frame to hold the notebook and add a border
            notebook_frame = ttk.Frame(self.root, borderwidth=2, relief="solid")
            notebook_frame.place(x=500, y=0, width=792, height=740)

            # Create a notebook for tabs inside the frame
            self.notebook = ttk.Notebook(notebook_frame)
            self.notebook.pack(fill='both', expand=True)

            # Add the Milestone Preview tab
            self.milestone_preview_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.milestone_preview_tab, text="Milestone Preview")
            MilestonePreviewApp(self.milestone_preview_tab)

            # Add the Site Config tab
            self.site_config_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.site_config_tab, text="Site Config")
            SiteDetailsApp(self.site_config_tab, tab_type="site_config")

            # Add the HSD Fields tab
            self.hsd_fields_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.hsd_fields_tab, text="HSD Fields")
            SiteDetailsApp(self.hsd_fields_tab, tab_type="hsd_fields")

            # Bind the tab change event
            self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        else:
            print("Advance Window Close")
            self.root.geometry("494x740")
            self.Button_advance["text"] = "->"
            if hasattr(self, 'notebook'):
                self.notebook.destroy()

    def on_tab_change(self, event):
        # Get the currently selected tab
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")

        # If the HSD Fields tab is selected, set focus to the notebook itself
        if tab_text == "HSD Fields":
            self.notebook.focus_set()

class MilestonePreviewApp:
    def __init__(self, root):
        # Create a frame for the Text widget and scrollbars
        text_frame = ttk.Frame(root)
        text_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # HTML Input Textbox with Scrollbars
        self.html_input = tk.Text(text_frame, wrap='none')
        self.html_input.pack(side='left', fill='both', expand=True)

        # Vertical Scrollbar
        self.v_scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.html_input.yview)
        self.v_scrollbar.pack(side='right', fill='y')
        self.html_input['yscrollcommand'] = self.v_scrollbar.set

        # Horizontal Scrollbar
        self.h_scrollbar = ttk.Scrollbar(root, orient='horizontal', command=self.html_input.xview)
        self.h_scrollbar.pack(fill='x')
        self.html_input['xscrollcommand'] = self.h_scrollbar.set

        # Buttons Frame
        button_frame = ttk.Frame(root)
        button_frame.pack(fill='x', padx=5, pady=5)

        # Preview Button
        self.preview_button = ttk.Button(button_frame, text="Preview", command=self.preview_html)
        self.preview_button.pack(side='left', expand=True, padx=5)

        # Load File Button
        self.load_button = ttk.Button(button_frame, text="Load File", command=self.load_html_file)
        self.load_button.pack(side='left', expand=True, padx=5)

        # Save Milestone Button
        self.save_button = ttk.Button(button_frame, text="Save Milestone", command=self.save_milestone)
        self.save_button.pack(side='left', expand=True, padx=5)

        # Clear Button
        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_entry)
        self.clear_button.pack(side='right', expand=True, padx=5)

        # Convert to CSV Button
        self.convert_button = ttk.Button(button_frame, text="Convert to CSV", command=self.convert_to_csv)
        self.convert_button.pack(side='right', expand=True, padx=5)

        # CSV Frame
        csv_frame = ttk.Frame(root)
        csv_frame.pack(fill='x', padx=5, pady=5)

        # CSV Label
        csv_label = ttk.Label(csv_frame, text="CSV")
        csv_label.pack(side='left', padx=5)

        # CSV Output Textbox
        self.csv_output = tk.Text(csv_frame, height=1, wrap='none')
        self.csv_output.pack(side='left', fill='x', expand=True)

        # Copy Button
        self.copy_button = ttk.Button(csv_frame, text="Copy", command=self.copy_csv_to_clipboard)
        self.copy_button.pack(side='right', padx=5)

        # Variable to store HTML content
        self.html_string = ""

        # Bind right-click to show context menu
        self.html_input.bind("<Button-3>", self.show_context_menu)

        # Create context menu
        self.context_menu = tk.Menu(self.html_input, tearoff=0)
        self.context_menu.add_command(label="Copy", command=self.copy)
        self.context_menu.add_command(label="Paste", command=self.paste)

        # Bind keyboard shortcuts
        self.html_input.bind("<Control-c>", self.copy)
        self.html_input.bind("<Control-v>", self.paste)

        # Home Directory Frame
        home_dir_frame = ttk.Frame(root)
        home_dir_frame.pack(fill='x', padx=5, pady=5)

        # Home Directory Label
        home_dir_label = ttk.Label(home_dir_frame, text="Milestone Home Directory")
        home_dir_label.pack(side='left', padx=5)

        # Home Directory Entry
        self.home_dir_var = tk.StringVar()
        self.home_dir_entry = ttk.Entry(home_dir_frame, textvariable=self.home_dir_var, width=50)
        self.home_dir_entry.pack(side='left', padx=5, fill='x', expand=True)

        # Update/Add Button
        self.update_button = ttk.Button(home_dir_frame, text="Add", command=self.update_home_dir)
        self.update_button.pack(side='left', padx=5)

        # Load settings.json
        self.load_settings()

        # Bind changes in the entry box to update the button label
        self.home_dir_var.trace_add("write", self.update_button_label)

        # Initialize the PyQt5 editor window
        self.editor = None

    def get_settings_path(self):
        """Get the path to the settings.json file in the script's directory."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, "settings.json")

    def load_settings(self):
        """Load settings from settings.json and update the home directory entry."""
        settings_path = self.get_settings_path()
        try:
            with open(settings_path, "r", encoding="utf-8") as file:
                settings = json.load(file)
                milestone_home = settings.get("milestone_home", "")
                self.home_dir_var.set(milestone_home)
                self.update_button_label()
        except FileNotFoundError:
            self.update_button.config(text="Add")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Error reading settings.json. Please check the file format.")

    def update_button_label(self, *args):
        """Update the button label based on the home directory entry."""
        if self.home_dir_var.get().strip():
            self.update_button.config(text="Change")
        else:
            self.update_button.config(text="Add")

    def update_home_dir(self):
        """Update the milestone home directory."""
        initial_dir = self.home_dir_var.get().strip() if self.home_dir_var.get().strip() else None
        selected_dir = filedialog.askdirectory(initialdir=initial_dir)
        if selected_dir:
            if os.access(selected_dir, os.R_OK | os.W_OK):
                self.home_dir_var.set(selected_dir)
                self.update_button.config(text="Change")
                settings = {"milestone_home": selected_dir}
                settings_path = self.get_settings_path()
                try:
                    with open(settings_path, "w", encoding="utf-8") as file:
                        json.dump(settings, file, indent=4)
                    print(f"Updated settings.json with milestone_home: {selected_dir}")
                except Exception as e:
                    messagebox.showerror("Error", f"Error writing to settings.json: {e}")
            else:
                messagebox.showerror("Error", "Verify directory read/write and try again.")
        else:
            self.update_button.config(text="Add")

    def preview_html(self):
        """Create a temporary HTML file and open it in the browser."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            return  # Do nothing if input is empty
        
        # Write to a temporary file
        temp_file = "temp_preview.html"
        with open(temp_file, "w", encoding="utf-8") as file:
            file.write(html_content)

        # Open in web browser
        webbrowser.open(f"file://{os.path.abspath(temp_file)}")

    def load_html_file(self):
        """Open an HTML or text file, clear the input box, and display its contents."""
        initial_dir = self.home_dir_var.get().strip() if self.home_dir_var.get().strip() else None
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("HTML and Text Files", "*.html;*.htm;*.txt")])
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    self.html_string = file.read()
            except UnicodeDecodeError:
                with open(file_path, "r", encoding="latin-1") as file:
                    self.html_string = file.read()
            
            self.clear_entry()  # Clear the entry field before inserting new content
            self.html_input.insert("1.0", self.html_string)  # Insert file content
            self.update_csv_output(self.html_string)

    def save_milestone(self):
        """Save the milestone content to a text file."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            return  # Do nothing if input is empty

        initial_dir = self.home_dir_var.get().strip() if self.home_dir_var.get().strip() else None
        file_path = filedialog.asksaveasfilename(initialdir=initial_dir, defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(html_content)

    def clear_entry(self):
        """Clear the entry field and the CSV output."""
        self.html_input.delete("1.0", tk.END)
        self.csv_output.delete("1.0", tk.END)
        self.csv_output.update_idletasks()  # Ensure the UI updates immediately

    def show_context_menu(self, event):
        """Show the context menu on right-click."""
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def copy(self, event=None):
        """Copy selected text to the clipboard."""
        self.html_input.event_generate("<<Copy>>")

    def paste(self, event=None):
        """Paste text from the clipboard."""
        self.html_input.event_generate("<<Paste>>")

    def parse_geometry(self, geometry):
        """Parse the geometry string to extract width, height, x, and y."""
        dimensions, x_y = geometry.split('+', 1)
        width, height = map(int, dimensions.split('x'))
        x, y = map(int, x_y.split('+'))
        return width, height, x, y

    def on_convert_callback(self, html_content):
        """Callback function to handle converted HTML content."""
        # Insert the converted HTML into the Tkinter HTML input box
        self.clear_entry()
        self.html_input.insert("1.0", html_content)
        self.update_csv_output(html_content)

    def update_csv_output(self, html_content):
        """Convert HTML content to a single line CSV-compatible format."""
        csv_content = ' '.join(html_content.split()).replace(',', ';')  # Replace commas to avoid CSV issues
        self.csv_output.delete("1.0", tk.END)
        self.csv_output.insert("1.0", csv_content)

    def copy_csv_to_clipboard(self):
        """Copy the CSV content to the clipboard."""
        csv_content = self.csv_output.get("1.0", tk.END).strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(csv_content)
        self.root.update()  # Now it stays on the clipboard after the window is closed

    def convert_to_csv(self):
        """Convert HTML content to a CSV-compatible string and paste it into the CSV textbox."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            messagebox.showinfo("Nothing to Convert", "The preview window is empty. Please load or enter HTML content to convert.")
            return

        # Convert HTML content to a single line CSV-compatible format
        csv_content = ' '.join(html_content.split()).replace(',', ';')  # Replace commas to avoid CSV issues
        self.csv_output.delete("1.0", tk.END)
        self.csv_output.insert("1.0", csv_content)

class SiteDetailsApp:
    def __init__(self, root, tab_type):
        self.root = root
        self.sites_data = self.load_sites_data()
        self.static_vals_data = self.load_static_vals_data()
        self.selected_site_index = None

        if tab_type == "site_config":
            self.create_sites_config_section(root)
        elif tab_type == "hsd_fields":
            self.create_hsd_fields_section(root)

    def create_sites_config_section(self, parent):
        # Create a canvas to draw the rectangle with a border
        canvas_width = 600
        canvas_height = 480
        canvas = tk.Canvas(
            parent,
            width=canvas_width,
            height=canvas_height,
            #highlightbackground='black', # Set the border color
            #highlightcolor='black',
            #highlightthickness=2,  # Set the border thickness
        )
        canvas.place(x=10, y=10)

        # Define the rectangle dimensions
        rect_width = 586
        rect_height = 400

        # Calculate the center of the canvas
        center_x = canvas_width // 2
        center_y = canvas_height // 2

        # Calculate the rectangle's top-left and bottom-right coordinates
        rect_x1 = center_x - rect_width // 2
        rect_y1 = center_y - rect_height // 2
        rect_x2 = center_x + rect_width // 2
        rect_y2 = center_y + rect_height // 2

        # Draw a rectangle on the canvas
        canvas.create_rectangle(rect_x1, rect_y1, rect_x2, rect_y2, outline="black", width=2)

        # Add a label for "Sites Config"
        label = ttk.Label(parent, text="Sites Config", font=("Arial", 14, "bold"))
        label.place(x=center_x - 260, y=rect_y1 - 20)

        # Create a frame to hold the listbox, input fields, and buttons
        input_frame = ttk.Frame(parent)

        # Calculate the width and height of the input_frame
        input_frame_width = 500  # Adjust this value based on the actual width of the frame
        input_frame_height = 180  # Adjust this value based on the actual height of the frame

        # Calculate the position to center the input_frame within the rectangle
        input_frame_x = rect_x1 + (rect_width - input_frame_width) // 4
        input_frame_y = rect_y1 + (rect_height - input_frame_height) // 8

        # Place the input_frame at the calculated position
        input_frame.place(x=input_frame_x, y=input_frame_y)

        # Listbox for sites with a specified width
        self.site_listbox = tk.Listbox(input_frame, width=20, height=22)
        self.site_listbox.grid(row=0, column=0, rowspan=4, padx=5, pady=5, sticky="ns")

        # Vertical Scrollbar
        self.v_scrollbar = ttk.Scrollbar(input_frame, orient='vertical', command=self.site_listbox.yview)
        self.v_scrollbar.grid(row=0, column=1, rowspan=4, sticky="ns")
        self.site_listbox['yscrollcommand'] = self.v_scrollbar.set

        # Populate the listbox with site names
        for site in self.sites_data:
            self.site_listbox.insert(tk.END, site['Site'])

        # Bind the listbox selection event
        self.site_listbox.bind('<<ListboxSelect>>', self.on_site_select)

        # Input fields for site details
        self.customer_var = tk.StringVar()
        self.backup_var = tk.StringVar()
        self.notify_var = tk.StringVar()
        self.lab_var = tk.StringVar()

        # Create input fields and buttons inside the frame
        self.create_input_field(input_frame, "Customer:", self.customer_var, 0, 2, self.update_customer)
        self.create_input_field(input_frame, "Backup:", self.backup_var, 1, 2, self.update_backup)
        self.create_input_field(input_frame, "Notify:", self.notify_var, 2, 2, self.update_notify)
        self.create_input_field(input_frame, "Lab:", self.lab_var, 3, 2, self.update_lab)

    def create_hsd_fields_section(self, parent):
        # Create a canvas to draw the rectangle with a border
        canvas_width = 600
        canvas_height = 560
        canvas = tk.Canvas(
            parent,
            width=canvas_width,
            height=canvas_height,
            #highlightbackground='black', # Set the border color
            #highlightcolor='black',
            #highlightthickness=2,  # Set the border thickness
        )
        canvas.place(x=10, y=10)

        # Define the rectangle dimensions
        rect_width = 586
        rect_height = 480

        # Calculate the center of the canvas
        center_x = canvas_width // 2
        center_y = canvas_height // 2

        # Calculate the rectangle's top-left and bottom-right coordinates
        rect_x1 = center_x - rect_width // 2
        rect_y1 = center_y - rect_height // 2
        rect_x2 = center_x + rect_width // 2
        rect_y2 = center_y + rect_height // 2

        # Draw a rectangle on the canvas
        canvas.create_rectangle(rect_x1, rect_y1, rect_x2, rect_y2, outline="black", width=2)

        # Add a label for "HSD Fields"
        label = ttk.Label(parent, text="HSD Fields", font=("Arial", 14, "bold"))
        label.place(x=center_x - 260, y=rect_y1 - 20)

        # Create a frame to hold the input fields and buttons
        input_frame = ttk.Frame(parent)
        input_frame.place(x=rect_x1 + 20, y=rect_y1 + 20)

        # Create input fields for each item in static_vals_data
        self.hsd_vars = {}
        row = 0
        for item, value in self.static_vals_data.items():
            var = tk.StringVar(value=value)
            self.hsd_vars[item] = var
            self.create_input_field(input_frame, f"{item}:", var, row, 0, lambda i=item: self.update_hsd_field(i))
            row += 1

        # Set focus to the parent frame to ensure no input box is selected by default
        parent.focus_set()

    def create_input_field(self, parent, label_text, text_var, row, col_start, update_command):
        """Create a label, entry field, and update button for site details."""
        label = tk.Label(
            parent,
            text=label_text,
            anchor="e",
            #relief="solid",  # Use 'solid' for a solid border
            #borderwidth=1,   # Set the border width
        )
        label.grid(row=row, column=col_start, padx=5, pady=5, sticky="e")

        entry = ttk.Entry(parent, textvariable=text_var, width=30)
        entry.grid(row=row, column=col_start + 1, padx=5, pady=5)

        update_button = ttk.Button(parent, text="Update", command=update_command)
        update_button.grid(row=row, column=col_start + 2, padx=5, pady=5)

    def load_sites_data(self):
        """Load site data from the Site.csv file."""
        sites_data = []
        try:
            with open("dependencies/Site.csv", encoding="utf8") as f:
                csv_reader = csv.DictReader(f)
                for row in csv_reader:
                    sites_data.append(row)
        except FileNotFoundError:
            messagebox.showerror("Error", "Site.csv file not found in dependencies folder.")
        return sites_data

    def load_static_vals_data(self):
        """Load static values from the static_vals.csv file."""
        static_vals_data = {}
        try:
            with open("dependencies/static_vals.csv", encoding="utf8") as f:
                csv_reader = csv.DictReader(f)
                for row in csv_reader:
                    static_vals_data[row['Item']] = row['Value']
        except FileNotFoundError:
            messagebox.showerror("Error", "static_vals.csv file not found in dependencies folder.")
        return static_vals_data

    def on_site_select(self, event):
        """Populate input fields with data from the selected site."""
        selected_index = self.site_listbox.curselection()
        if selected_index:
            self.selected_site_index = selected_index[0]
            selected_site = self.sites_data[self.selected_site_index]
            self.customer_var.set(selected_site.get("Customer", ""))
            self.backup_var.set(selected_site.get("Backup", ""))
            self.notify_var.set(selected_site.get("Notify", ""))
            self.lab_var.set(selected_site.get("Lab", ""))

    def update_customer(self):
        """Update the Customer field for the selected site."""
        self.update_site_field("Customer", self.customer_var.get())

    def update_backup(self):
        """Update the Backup field for the selected site."""
        self.update_site_field("Backup", self.backup_var.get())

    def update_notify(self):
        """Update the Notify field for the selected site."""
        self.update_site_field("Notify", self.notify_var.get())

    def update_lab(self):
        """Update the Lab field for the selected site."""
        self.update_site_field("Lab", self.lab_var.get())

    def update_site_field(self, field_name, new_value):
        """Update a specific field in the Site.csv file for the selected site."""
        if self.selected_site_index is not None:
            self.sites_data[self.selected_site_index][field_name] = new_value
            self.save_sites_data()
            messagebox.showinfo("Update Successful", f"{field_name} updated successfully.")

    def update_hsd_field(self, item):
        """Update a specific field in the static_vals.csv file."""
        new_value = self.hsd_vars[item].get()
        self.static_vals_data[item] = new_value
        self.save_static_vals_data()
        messagebox.showinfo("Update Successful", f"{item} updated successfully.")

    def save_sites_data(self):
        """Save the updated site data back to the Site.csv file."""
        try:
            with open("dependencies/Site.csv", "w", newline='', encoding="utf8") as f:
                fieldnames = ["Site", "Customer", "Backup", "Notify", "Lab"]
                csv_writer = csv.DictWriter(f, fieldnames=fieldnames)
                csv_writer.writeheader()
                csv_writer.writerows(self.sites_data)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Site.csv: {e}")

    def save_static_vals_data(self):
        """Save the updated static values back to the static_vals.csv file."""
        try:
            with open("dependencies/static_vals.csv", "w", newline='', encoding="utf8") as f:
                fieldnames = ["Item", "Value"]
                csv_writer = csv.DictWriter(f, fieldnames=fieldnames)
                csv_writer.writeheader()
                for item, value in self.static_vals_data.items():
                    csv_writer.writerow({"Item": item, "Value": value})
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save static_vals.csv: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = HSDLabGenApp(root)
    root.mainloop()