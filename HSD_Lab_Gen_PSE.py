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
import tempfile

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
        self.milestones_open = ""
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
            print('\nLoading milestones:')
            with open("dependencies/milestones.csv", encoding="utf8") as data_file:
                data = csv.reader(data_file)
                dynamic_headers = next(data)[0:]
                dynamic_headers.append("descriptions")  # Add the "descriptions" header

                for row in data:
                    temp_dict = {}
                    keystone = row[0]
                    Milestone = row[1]
                    name = row[2]
                    values = []

                    for x in row[0:]:
                        values.append(x)
                        
                    for i in range(len(values)):
                        if values[i]:
                            temp_dict[dynamic_headers[i]] = values[i]

                            # Add the "dictionary" field by reading the milestone file
                            milestone_file_path = os.path.join('milestones/', Milestone + '.txt')
                            if os.path.exists(milestone_file_path):
                                with open(milestone_file_path, 'r', encoding='utf8') as milestone_file:
                                    milestone_content = milestone_file.read()
                                    temp_dict['description'] = milestone_content
                            else:
                                print(f"Milestone file '{milestone_file_path}' does not exist.")

                    self.milestones[name] = temp_dict

            self.milestones_open = True
        except Exception as e:
            print(f'Failed to load Milestones! Error: {e}')
            self.milestones_open = False

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
                    self.variables[key] = {'widget': checkbox, 'variable': var, 'text': checkbox["text"], 'Milestone': self.milestones[key]["Milestone"]}
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
        not_ready3 = 0

        for name, milestone in self.milestones.items():
            milestone_file_path = os.path.join('milestones/', str(milestone['Milestone']) + '.txt')
            if os.path.exists(milestone_file_path):
                pass
            else:
                print(f"Milestone file '{milestone_file_path}' does not exist.")
                not_ready3 += 1
                
        if not_ready3 == 0: 
            not_ready3 = ''
        else: not_ready3 = 'Missing Milestone files\n'

        if (len(self.project_option_selected.get()) == 0):
            not_ready1 = 'Select Program\n'

        if (len(self.site_option_selected.get()) == 0):
            not_ready2 = 'Select Site\n'

        not_ready = (not_ready1 + not_ready2 + not_ready3)
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
                    self.checkbox_dict.update({value['Milestone'].split('.')[0]: int(value['variable'].get())})

                fieldlist = []

                dictionaryloop = 1
                if dictionaryloop == 1:
                    print("")
                    for name, milestone in self.milestones.items():
                        mile_value = milestone.get('Milestone').split('.')[0]
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
                            milestoneww = (year + week)
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
        self.root = root
        self.milestones_data = self.load_milestones_data()
        self.selected_milestone_index = None

        # Create a frame to hold the canvas and other components
        main_frame = ttk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Create a canvas to draw the rectangle with a border
        canvas_width = 740
        canvas_height = 400  # Increased height
        canvas = tk.Canvas(
            main_frame,
            width=canvas_width,
            height=canvas_height,
            highlightbackground='black',  # Set the border color
            highlightthickness=2,  # Set the border thickness
        )
        canvas.grid(row=0, column=0, columnspan=2, sticky='nw', pady=10)

        # Define the rectangle dimensions
        rect_width = 720
        rect_height = 340  # Increased height

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

        # Add a label for "Milestones"
        label = ttk.Label(main_frame, text="Milestones", font=("Arial", 14, "bold"))
        label.place(x=center_x - 340, y=rect_y1 - 20)

        # Create a frame to hold the Treeview and scrollbars
        tree_frame = ttk.Frame(canvas)

        # Place the tree_frame inside the canvas
        canvas.create_window((rect_x1, rect_y1), window=tree_frame, anchor='nw', width=rect_width, height=rect_height)

        # Create the Treeview widget with only the 'Milestone' and 'cb_title' columns
        self.tree = ttk.Treeview(tree_frame, columns=('Milestone', 'cb_title'), show='headings')
        self.tree.grid(row=0, column=0, sticky='nsew')

        # Configure the grid to expand the Treeview
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Vertical Scrollbar
        self.v_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        self.tree['yscrollcommand'] = self.v_scrollbar.set

        # Horizontal Scrollbar
        self.h_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')
        self.tree['xscrollcommand'] = self.h_scrollbar.set

        # Add column headings
        self.tree.heading('Milestone', text='Milestone')
        self.tree.heading('cb_title', text='cb_title', anchor=tk.W)

        # Measure the width of the 'Milestone' header text
        header_font = tk.font.Font()
        milestone_header_width = header_font.measure('Milestone')

        # Configure columns
        self.tree.column('Milestone', anchor='w', width=milestone_header_width, stretch=False)  # Set width to header text width
        self.tree.column('cb_title', anchor='w', width=300)  # Increased width
        self.tree.column('Milestone', anchor='center')

        # Populate the Treeview with milestone data
        self.populate_treeview()

        # Bind the Treeview selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_milestone_select)

        # Create input fields for each column in the CSV file
        self.input_vars = {}
        self.create_input_fields(main_frame)

        # Create a frame for buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        # Add a Preview button
        self.preview_button = ttk.Button(button_frame, text="Preview Milestone Description", command=self.preview_html)
        self.preview_button.pack(side='left', padx=5)

        # Add a button to open the HTML editor window
        self.open_html_editor_button = ttk.Button(button_frame, text="Open Milestone Description (HTML) Editor", command=self.open_html_editor)
        self.open_html_editor_button.pack(side='left', padx=5)

    def load_milestones_data(self):
        """Load milestone data from the milestones.csv file."""
        milestones_data = []
        try:
            with open("dependencies/milestones.csv", encoding="utf8") as f:
                csv_reader = csv.DictReader(f)
                for row in csv_reader:
                    milestones_data.append(row)
        except FileNotFoundError:
            messagebox.showerror("Error", "milestones.csv file not found in dependencies folder.")
        return milestones_data

    def populate_treeview(self):
        """Populate the Treeview with milestone data."""
        for milestone in self.milestones_data:
            self.tree.insert('', tk.END, values=(milestone['Milestone'], milestone['cb_title']))

    def create_input_fields(self, parent):
        """Create input fields for each column in the CSV file."""
        if not self.milestones_data:
            return

        # Create a main frame to hold all input field frames
        main_input_frame = ttk.Frame(parent)
        main_input_frame.grid(row=1, column=0, sticky='nw', padx=10, pady=10)

        # Create input fields for each column
        headers = self.milestones_data[0].keys()
        for i, header in enumerate(headers):
            # Create a frame for each input field
            input_frame = ttk.Frame(main_input_frame)
            input_frame.grid(row=i, column=0, sticky='w', padx=5, pady=5)

            # Create label and entry within the frame
            label = ttk.Label(input_frame, text=header, width=15, anchor='w', font=("Arial", 10, "bold"))  # Set a fixed width for labels
            label.grid(row=0, column=0, sticky='w', padx=5)

            var = tk.StringVar()
            entry = ttk.Entry(input_frame, textvariable=var, width=80)
            entry.grid(row=0, column=1, sticky='w', padx=5)

            # Store the variable for later use
            self.input_vars[header] = var

        # Adjust column weights to ensure proper alignment
        main_input_frame.grid_columnconfigure(0, weight=1)

    def on_milestone_select(self, event):
        """Populate input fields with data from the selected milestone."""
        selected_item = self.tree.selection()
        if selected_item:
            selected_milestone = self.tree.item(selected_item, 'values')[0]  # Get the Milestone value
            milestone_data = self.find_milestone_data(selected_milestone)
            if milestone_data:
                for header, value in milestone_data.items():
                    if header in self.input_vars:
                        self.input_vars[header].set(value)

    def find_milestone_data(self, milestone_value):
        """Find and return the milestone data for the given milestone value."""
        for milestone in self.milestones_data:
            if milestone['Milestone'] == milestone_value:
                return milestone
        return None

    def open_html_editor(self):
        """Open the HTML editor window."""
        selected_item = self.tree.selection()
        if selected_item:
            selected_milestone = self.tree.item(selected_item, 'values')[0]  # Get the Milestone value
            HtmlEditorWindow(self.root, selected_milestone)
        else:
            messagebox.showinfo("No Selection", "Please select a milestone to edit.")

    def preview_html(self):
        """Preview the HTML content of the selected milestone."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showinfo("Preview", "Please select a milestone to preview.")
            return

        selected_milestone = self.tree.item(selected_item, 'values')[0]  # Get the Milestone value
        milestone_file_path = os.path.join("milestones", f"{selected_milestone}.txt")

        if not os.path.exists(milestone_file_path):
            messagebox.showinfo("No Milestone File found!", "No Milestone Found Please create and save one.")
            return

        with open(milestone_file_path, "r", encoding="utf-8") as file:
            html_content = file.read()

        if not html_content.strip():
            messagebox.showinfo("Preview", "No content to preview.")
            return

        # Create a temporary file to hold the HTML content
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as temp_file:
            temp_file.write(html_content.encode('utf-8'))
            temp_file_path = temp_file.name

        # Open the temporary file in the default web browser
        webbrowser.open(f"file://{temp_file_path}")

class HtmlEditorWindow:
    def __init__(self, parent, milestone):
        self.window = tk.Toplevel(parent)
        self.window.title("HTML Editor")
        self.milestone = milestone

        # Create a label to show the milestone being edited
        milestone_label = ttk.Label(self.window, text=f"Editing Milestone: {self.milestone}", font=("Arial", 12, "bold"))
        milestone_label.pack(pady=5)

        # Create a frame for the Text widget and scrollbars
        text_frame = ttk.Frame(self.window)
        text_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # HTML Input Textbox with Scrollbars
        self.html_input = tk.Text(text_frame, wrap='none', height=20, width=80)  # Set desired height and width
        self.html_input.pack(side='left', fill='both', expand=True)

        # Vertical Scrollbar for Textbox
        self.text_v_scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.html_input.yview)
        self.text_v_scrollbar.pack(side='right', fill='y')
        self.html_input['yscrollcommand'] = self.text_v_scrollbar.set

        # Horizontal Scrollbar for Textbox
        self.text_h_scrollbar = ttk.Scrollbar(self.window, orient='horizontal', command=self.html_input.xview)
        self.text_h_scrollbar.pack(fill='x')
        self.html_input['xscrollcommand'] = self.text_h_scrollbar.set

        # Buttons Frame
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill='x', pady=10)

        # Preview Button
        self.preview_button = ttk.Button(button_frame, text="Preview", command=self.preview_html)
        self.preview_button.pack(side='left', expand=True, padx=5)

        # Save Milestone Button
        self.save_button = ttk.Button(button_frame, text="Save Milestone", command=self.save_milestone)
        self.save_button.pack(side='left', expand=True, padx=5)

        # Clear Button
        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_entry)
        self.clear_button.pack(side='right', expand=True, padx=5)

        # Load the milestone file if it exists
        self.load_milestone_file()

    def load_milestone_file(self):
        """Load the milestone file and display its contents."""
        milestone_file_path = os.path.join("milestones", f"{self.milestone}.txt")
        if os.path.exists(milestone_file_path):
            with open(milestone_file_path, "r", encoding="utf-8") as file:
                content = file.read()
                self.html_input.delete("1.0", tk.END)
                self.html_input.insert("1.0", content)
        else:
            messagebox.showinfo("No Milestone File found!", "No Milestone Found Please create and save one.")

    def preview_html(self):
        """Preview the HTML content."""
        html_content = self.html_input.get("1.0", tk.END).strip()
        if not html_content:
            messagebox.showinfo("Preview", "No content to preview.")
            return

        # Create a temporary file to hold the HTML content
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as temp_file:
            temp_file.write(html_content.encode('utf-8'))
            temp_file_path = temp_file.name

        # Open the temporary file in the default web browser
        webbrowser.open(f"file://{temp_file_path}")

    def save_milestone(self):
        """Save the milestone content to a file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=self.milestone, filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as file:
                content = self.html_input.get("1.0", tk.END).strip()
                file.write(content)
            # Close the window after saving
            self.window.destroy()

    def clear_entry(self):
        """Clear the entry field and the HTML input."""
        self.html_input.delete("1.0", tk.END)
        
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