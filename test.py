import tkinter
from tkinter import ttk
from tkinter import filedialog
import re
import pandas as pd
from datetime import datetime
import tkinter.messagebox
from PIL import Image, ImageTk

import test2 as customizing

def quit_window (window: tkinter.Tk):
    #quit_message = tkinter.messagebox.askyesnocancel(title='Warning', message='Are you sure you want to quit?')
    #if quit_message == tkinter.TRUE:
        window.destroy()

def home ():
    for widget in mainframe.frame.winfo_children():
        if type(widget) != tkinter.Menu:
            widget.destroy()
    MainRoot (root = mainroot.root, frame = mainframe.frame)

def start (program: str, frame: tkinter.Frame):
    for widget in frame.winfo_children():
        if type(widget) != tkinter.Menu:
            widget.destroy()
    if program == 'ebs_mt940':
        Ebs (frame = frame)
    elif program == 'iban':
        Iban (frame = frame)
    else:
        MigrationFile (frame = frame)

def tab_creation (frame):
    tab = ttk.Notebook(frame)
    tab.pack(fill = tkinter.BOTH, expand = tkinter.TRUE, side = 'left')
    return tab

def browse_file(entry_path: customizing.Entry):
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    entry_path.entry.delete(0, tkinter.END)  # Clear the entry widget
    entry_path.entry.insert(0, file_path)  # Insert the selected file path into the entry widget

def browse_file_xlsx(entry_path: customizing.Entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_path.entry.config(state = 'normal')
    entry_path.entry.delete(0, tkinter.END)  # Clear the entry widget
    entry_path.entry.insert(0, file_path)  # Insert the selected file path into the entry widget

def read_file(entry_path: customizing.Entry, text: tkinter.Text):
    file_path = entry_path.entry.get()
    try:
        with open(file_path, 'r') as file:
            content = file.read()
            text.delete(1.0, tkinter.END)  # Clear the text widget
            text.insert(tkinter.END, content)  # Insert the file content into the text widget
    except FileNotFoundError:
        text.delete(1.0, tkinter.END)
        text.insert(tkinter.END, "File not found.")

def read_file_xlsx(entry_path_str: str):
    file_path = entry_path_str
    with open(file_path, 'r') as file:
        content = pd.read_excel(file_path)

def create_excel_file(sheet_names_list: list, present_fields: list):
    # Ask the user to choose where to save the file
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        # Create a Pandas Excel writer using xlsxwriter as the engine
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            # Write each DataFrame to a different sheet
            for a in sheet_names_list:
                sheet_present_fields = []
                for j in present_fields:
                    if j[0] == a[0]:
                        sheet_present_fields.append(j[1].replace("*", "").replace("+", ""))

                # Create a DataFrame with the first row filled with consecutive numbers
                data = {}

                series_data = pd.Series (data, index=[1])
                df_sheet = pd.DataFrame([series_data], columns=sheet_present_fields)

                df_sheet.to_excel(writer, sheet_name=a[0], index=False)

                # Adjust column size based on the length of the column name
                worksheet = writer.sheets[a[0]]
                for col_num, value in enumerate(sheet_present_fields, 0):
                    max_len = max(df_sheet[value].astype(str).apply(len).max(), len(value))
                    col_width = (max_len + 2) * 1.2  # Adjust the multiplier as needed
                    worksheet.set_column(col_num, col_num, col_width)

        print(f"Excel file saved to: {file_path}")

def migration_forget (sheet_list: list, sheet: str):
    for a in sheet_list:
        if sheet != 'No':
            a.sheet.checkbox.grid_forget() 
        a.main_frame.frame.pack_forget()
        for b in a.field_list:
            b.label.label.grid_forget()
            b.radiobutton_1.grid_forget()
            b.radiobutton_2.grid_forget()
            b.radiobutton_3.grid_forget() 

def info (type: str):
    Info (type = type)

class Info:
    def __init__(
        self,
        type: str
        ):
        self.root = customizing.Root (root_title = 'HELP')
        self.root.root.state('normal')

        self.frame = customizing.Frame (
            root = self.root.root,
            pack_or_grid = 'P'
        )

        self.type_label = customizing.Label (
            frame = self.frame.frame,
            text = type,
            dimension = 24,
            padx = 20,
            pady = 10,
            sticky = tkinter.NW
        )

        self.main_label = customizing.Label (
            frame = self.frame.frame,
            text = '',
            row = 1,
            padx = 10,
            pady = 10,
            sticky = tkinter.NW,
            justify = tkinter.LEFT
        )

        if type == 'EBS MT940 info':
            self.main_label.label.config (text = customizing.ebs_mt940_text)
        elif type == 'IBAN info':
            self.main_label.label.config (text = customizing.iban_text)
        elif type == 'Migration File info':
            self.main_label.label.config (text = customizing.mf_text)

class MainRoot:
    def __init__(
        self,
        root: tkinter.Tk,
        frame: tkinter.Frame
        ):
        self.root = root

        self.frame = frame

        customizing.Label (
             frame = self.frame,
             text = 'SAP HELPER',
             dimension = 40,
             sticky = tkinter.NS,
             foreground = '#229F22',
             pady = 30,
             padx = 650
        )
        customizing.Button (
            frame = self.frame,
            text = 'EBS MT940',
            command = lambda: start('ebs_mt940', self.frame),
            width = 15,
            height = 3,
            row = 1,
            pady = 10,
            padx = 10
        )
        customizing.Button (
            frame = self.frame,
            text = 'IBAN',
            command = lambda: start('iban', self.frame),
            width = 15,
            height = 3,  
            row = 2,
            pady = 10,
            padx = 10
        )
        customizing.Button (
            frame = self.frame,
            text = 'Migration File',
            command = lambda: start('migration_file', self.frame),
            width = 15,
            height = 3,                       
            row = 3,
            pady = 10,
            padx = 10
        )
        
class Ebs:
    def __init__(
        self,
        frame: tkinter.Frame
        ):
        def ebs_analysis ():
            content = self.text.get("1.0", tkinter.END)
            for widget in frame.winfo_children():
                widget.destroy()

            lines = content.split('\n')
            swift = ''
            bank_account_number = ''
            start_date = ''
            end_date = ''
            currency = ''
            opening_balance = 0
            closing_balance = 0

            value_date = ''
            amount = ''
            bank_external_transaction = ''
            bank_ext_tr_description = ''
            position_lst = []
            total_debit = 0
            total_credit = 0

            #header and position value into variables are set here, depending on file content
            for line in lines:
                if line[:3] == '{1:':
                    swift = line[6:14]
                elif line[:4] == ':25:':
                    bank_account_number = line[4:]
                elif line[:5] == ':60F:':
                    start_date = line[6:12]
                    yy = start_date[:2]
                    mm = start_date[2:4]
                    dd = start_date[4:6]
                    start_date = dd + "/" + mm + "/20" + yy
                    currency = line[12:15]
                    opening_balance = line[15:].replace(",", ".")
                    if line[5] == 'C':
                        opening_balance = '-' + line[15:].replace(",", ".")
                elif line[:5] == ':62F:':
                    end_date = line[6:12]
                    yy = end_date[:2]
                    mm = end_date[2:4]
                    dd = end_date[4:6]
                    end_date = dd + "/" + mm + "/20" + yy
                    closing_balance = line[15:].replace(",", ".")
                    if line[5] == 'C':
                        closing_balance = '-' + line[15:].replace(",", ".")
                elif line[:4] == ':61:':
                    if line[14] == 'C' or line[14:16] == 'RD':
                        amount = "-" + re.search(r'(\d+,\d+)', line[4:]).group(1)
                        amount = "{:.2f}".format(float(amount.replace(",", ".")))
                        total_credit += float(amount)
                    else:
                        amount = re.search(r'(\d+,\d+)', line[4:]).group(1)
                        amount = "{:.2f}".format(float(amount.replace(",", ".")))
                        total_debit += float(amount)
                    value_date = line[4:10]
                    yy = value_date[:2]
                    mm = value_date[2:4]
                    dd = value_date[4:6]
                    value_date = dd + "/" + mm + "/20" + yy
                    bank_external_transaction = line[line.find(',') + 3:line.find(',') + 7]
                    if line[line.find(',') + 2].isalpha():
                        bank_external_transaction = line[line.find(',') + 2:line.find(',') + 6]
                    if bank_external_transaction in customizing.ebs_mt940_dict:
                        bank_ext_tr_description = customizing.ebs_mt940_dict[bank_external_transaction]
                    else:
                        bank_ext_tr_description = ''
                    position_lst.append((value_date, amount, bank_external_transaction, bank_ext_tr_description))
                
                
            lst = [('SWIFT', 'BANK ACCOUNT N°', 'START DATE', 'END DATE', 'CURRENCY', 'OPENING BALANCE', 'CLOSING BALANCE'), 
                    (swift, bank_account_number, start_date, end_date, currency, opening_balance, closing_balance),
                    ('', '', '', '', '', '', ''),
                        ('VALUE DATE', 'AMOUNT', 'BANK EXTERNAL TRANSACTION', 'BANK EXT TR DESCRIPTION')]
            
            
            for a in range(len(position_lst)):
                lst.append(position_lst[a])
            lst.append ('')
            lst.append(('OPENING BALANCE', opening_balance))
            lst.append(('TOTAL CREDIT', "{:.2f}".format(total_credit)))
            lst.append(('TOTAL DEBIT', "{:.2f}".format(total_debit)))
            lst.append(('CLOSING BALANCE', closing_balance))
            lst.append(('CHECK', "{:.2f}".format(float(opening_balance) + total_credit + total_debit - float(closing_balance))))
        
            tree = customizing.Treeview (
                frame = self.frame,
                col_text = ['', '', '', '', '', '', ''],
                width_list = [120, 120, 120, 120, 120, 120, 120],
                lst = lst
            )

            check = float(opening_balance) + total_credit + total_debit - float(closing_balance)
            image_path = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\green_tick.png'
            if round(check, 2) != 0:
                image_path = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\red_cross.png'

            # Open the image using Pillow
            image = Image.open(image_path)
            # Convert the image to a format Tkinter supports
            icon_check = ImageTk.PhotoImage(image)

            label_with_icon = tkinter.Label(tree.frame_1.frame, image=icon_check, text="CHECK: ", compound=tkinter.RIGHT, font = ('Calibri', 14, 'bold'), background = '#F0F8FF')
            label_with_icon.grid(column=1, row=0, padx=200, pady=10)

            # Keep a reference to the image to prevent it from being garbage collected
            label_with_icon.image = icon_check


        self.frame = frame

        entry_path = customizing.Entry(
            frame = self.frame,
            width = 80,
            column = 1,
            entry_pady = 10
            )

        #button to upload the .txt file
        customizing.Button(
            frame = self.frame, 
            text = "Upload .txt file", 
            command = lambda: (browse_file (entry_path=entry_path), read_file (entry_path=entry_path, text=self.text)),
            pady = 10
            )

        #text field. It is automatically filled uploading the file. It's also possible to paste here the .txt file content directly
        self.text = tkinter.Text(self.frame, height=44, width=100)
        self.text.grid(row=2, column=0, padx=10, pady=10, columnspan=3)

        self.y_scrollbar = ttk.Scrollbar(self.frame, orient='vertical', command=self.text.yview)
        self.y_scrollbar.grid(row = 2, column = 3, sticky = 'NS')
        self.text.configure(yscrollcommand=self.y_scrollbar.set)

        #button to analyze the file content (a Trevieew will be opened)
        customizing.Button(
            frame = self.frame, 
            text = "Analysis", 
            command = ebs_analysis,
            column = 3
            )

class Iban:
    def __init__(
        self,
        frame: tkinter.Frame
        ):
        def iban_analysis ():
            content = self.text.get("1.0", tkinter.END)
            for widget in frame.winfo_children():
                widget.destroy()
        
            lines = content.split('\n')

            iban = ''
            bank_country = ''
            bank_key = ''
            bank_account_number = ''
            bank_control_key = ''
            swift = ''
            notes = ''
            position_lst = []

            for line in lines:
                iban = ''
                bank_country = ''
                bank_key = ''
                bank_account_number = ''
                bank_control_key = ''
                swift = ''
                notes = ''
                if line != '':
                    iban = line
                    bank_country = line[:2]
                    if bank_country == 'IT':
                        bank_key = line[5:15]
                        bank_account_number = line[15:27]
                        bank_control_key = line[4]
                    elif bank_country == 'ES':
                        bank_key = line[4:12]
                        bank_account_number = line[14:24]
                        bank_control_key = line[12:14]
                    elif bank_country == 'BE':
                        bank_key = line[4:7]
                        bank_account_number = line[4:7] + '-' + line[7:14] + '-' + line[14:16]
                        notes = 'For belgian banks, it needs to enter in bank acount number:- "BANK_KEY-BANK_ACCOUNT_NUMBER-BANK_CONTROL_KEY"'
                    elif bank_country == 'FR':
                        bank_key = line[4:14]
                        bank_account_number = line[14:25]
                        bank_control_key = line[25:27]
                    elif bank_country == 'NL':
                        swift = line[4:8]
                        bank_account_number = line[8:18]
                        notes = 'For Netherland banks, SAP bank key is not relevant to calculate IBAN, SAP extract it from SWIFT'
                    elif bank_country == 'FI':
                        bank_key = line[4:10]
                        bank_account_number = line[10:17]
                        bank_control_key = line[17]
                    elif bank_country == 'LU':
                        bank_key = line[4:7]
                        bank_account_number = line[7:20]
                    elif bank_country == 'CH':
                        bank_key = line[4:9]
                        bank_account_number = line[9:21]
                    elif bank_country == 'GB' or bank_country == 'IE':
                        swift = line[4:8]
                        bank_key = line[8:14]
                        bank_account_number = line[14:22]
                        notes = 'For UK and Ireland, it is used the SWIFT code in the IBAN. It needs to establish if it needs to include it in the SAP bank key'
                    elif bank_country == 'DE':
                        bank_key = line[4:12]
                        bank_account_number = line[12:22]
                    else:
                        notes = 'Error, bank country not recognized'

                    position_lst.append((iban, bank_country, bank_key, bank_account_number, bank_control_key, swift, notes))


            lst = [('IBAN', 'BANK COUNTRY', 'BANK KEY', 'BANK ACCOUNT N°', 'BANK CONTROL KEY', 'SWIFT', 'NOTES')]
            for a in range(len(position_lst)):
                lst.append(position_lst[a])

        
            customizing.Treeview (
                frame = self.frame,
                col_text = ['', '', '', '', '', '', ''],
                width_list = [1, 1, 1, 1, 1, 1, 1000],
                lst = lst
            )

        self.frame = frame

        #IBAN view
        self.text = tkinter.Text(self.frame, height=44, width=100)
        self.text.grid(row=2, column=0, padx=10, pady=10, columnspan=3)

        self.y_scrollbar = ttk.Scrollbar(self.frame, orient='vertical', command=self.text.yview)
        self.y_scrollbar.grid(row = 2, column = 3, sticky = 'NS')
        self.text.configure(yscrollcommand=self.y_scrollbar.set)

        #button to analyze the file content (a Trevieew will be opened)
        customizing.Button(
            frame = self.frame, 
            text = "Analysis", 
            command = iban_analysis,
            column = 3,
            pady = 10
            )

class MigrationFile:
    def __init__(
        self,
        frame: tkinter.Frame
        ):
        def go_back (screen: str):
            if screen == 'sheets':
                for a in sheet_list:
                    a.sheet.checkbox.grid_forget()
                self.entry_path.entry.config(state = 'normal')
                self.entry_path.entry.delete(0, tkinter.END)
                self.go_ahead.config(state = 'disabled')
                self.go_back.config(command = home)
            elif screen == 'fields':
                start('migration_file', mainframe.frame)
            elif screen == 'input':
                for b in self.sheet_names_list:
                    df = pd.read_excel(self.entry_path.entry.get(), b[0])

                    # Get the column names
                    column_list = df.iloc[6, :].tolist()
                    column = 0
                    new_line = 0
                    for c in range(len(column_list)):
                        if c == 10 or c == 20 or c == 30 or c == 40 or c == 50:
                            new_line += 4
                            column = 0
                        sheet_list[b[1]].field_list[c].text_input.grid_forget()
                        sheet_list[b[1]].field_list[c].label.label.grid_forget()
                        sheet_list[b[1]].field_list[c].label.label.grid(row = new_line + 0, column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b[1]].field_list[c].radiobutton_1.grid(row = new_line + 1,column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b[1]].field_list[c].radiobutton_2.grid(row = new_line + 2,column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b[1]].field_list[c].radiobutton_3.grid(row = new_line + 3,column = column, sticky = tkinter.W, padx = 10)
                        column += 1
                self.download_input_button.button.grid_forget()
                self.upload_input_button.button.grid_forget()
                self.go_ahead.config(command = migration_input)
                self.go_back.config(command = lambda: go_back('fields'))

        def sheet_checkboxes (): #this function is called when the file is uploaded

            def mode_command (on_off : str): #a function called when you press on mode radiobutton (if you want to use a mode, the main sheet must be present). Using a mode the mandatory sheets must be in the analysis
                if on_off == 'on':
                    if not any(element in self.sheet_names for element in customizing.migration_file_main_sheet):
                        tkinter.messagebox.showerror(title="ERROR", message='The main sheet for this template is missing in the Excel file. Choose "Generic" or check the uploaded file')
                        self.mode.variable.set('Generic')
                        return
                    for c in range(maximum_sheet):
                        for d in rows:
                            if self.sheet_names[c] in d:
                                if '(mandatory)' in d:
                                    sheet_list[c].sheet.variable.set(1)
                                    sheet_list[c].sheet.checkbox.config (state = 'disabled')
                else:
                    for c in range(maximum_sheet):
                        for d in rows:
                            if self.sheet_names[c] in d:
                                if '(mandatory)' in d:
                                    sheet_list[c].sheet.checkbox.config (state = 'normal')

            migration_forget(sheet_list, 'Yes')
                
            xl = pd.ExcelFile(self.entry_path.entry.get())

            df = pd.read_excel(self.entry_path.entry.get(), 'Field List')

            column_list = df.columns.tolist()

            all_rows = df.iloc[:, 1].tolist()
            rows = []
            for k in all_rows:
                if isinstance(k, str):
                    rows.append(k)

            # Get the sheet names
            self.sheet_names = xl.sheet_names

            #only the first 30 sheets are considered (due to program heaviness)
            maximum_sheet = len(self.sheet_names)
            if len(self.sheet_names) > 30:
                maximum_sheet = 30

            for c in range(maximum_sheet):
                if self.sheet_names[c] != 'Introduction' and self.sheet_names[c] != 'Field List':
                    sheet_list[c].sheet.checkbox.grid(row = c + 1, sticky = tkinter.W)
                    sheet_list[c].sheet.checkbox.config(text = self.sheet_names[c])
            
                for d in rows:
                    if self.sheet_names[c] in d:
                        if '(mandatory)' in d:
                            sheet_list[c].sheet.variable.set(1) #if a sheet is set as mandatory in the Field List sheet, so by default the checkbox will be ticked
                        else:
                            sheet_list[c].sheet.variable.set(0)
            
            for e in customizing.migration_file_modes:
                if e in column_list[0]:
                    self.mode.radiobutton_1.config(command = lambda: mode_command('off'))
                    self.mode.radiobutton_2.config(text = e, value = e, command = lambda: mode_command('on'))
                    self.mode_frame.frame.grid(column = 1, row = 1, rowspan = 100, columnspan = 5, sticky = tkinter.NW, padx = 50)
                    
            self.mode.variable.set('Generic')
            
            self.go_ahead.config(state = 'normal')
            self.go_back.config(command = lambda: go_back('sheets'))
            self.entry_path.entry.config(state = 'disabled')

        def migration_fields ():
            migration_forget(sheet_list, 'No')
            self.mode_frame.frame.grid_forget()
            self.sheet_names_list = []
            
            for b in range(len(sheet_list)):
                if sheet_list[b].sheet.variable.get() == 1: #only sheets ticked will be considered in the analysis
                    self.tab.grid(row = 1, column = 1, rowspan = 100, columnspan = 100, sticky = tkinter.NW)
                    self.tab.add(sheet_list[b].main_frame.frame, text = self.sheet_names[b])

                    df = pd.read_excel(self.entry_path.entry.get(), self.sheet_names[b])

                    # Get the column names
                    column_tech_names = df.iloc[3, :].tolist()
                    column_list = df.iloc[6, :].tolist()
                    column_names = []
                    for column in range(len(column_list)): #if a specific mode is activated, then some fields will be excluded from analysis
                        if self.mode.variable.get() == 'Customer':
                            if column_tech_names[column] not in customizing.mf_customer_general_data and column_tech_names[column] not in customizing.mf_customer_company_data:
                                column_names.append(column_list[column].split('\n')[0])
                        elif self.mode.variable.get() == 'Supplier':
                            if column_tech_names[column] not in customizing.mf_supplier_general_data:
                                column_names.append(column_list[column].split('\n')[0])
                        elif self.mode.variable.get() == 'FI - Accounts receivable open item' or self.mode.variable.get() == 'FI - Accounts payable open item':
                            if column_tech_names[column] not in customizing.mf_bp_open_items:
                                column_names.append(column_list[column].split('\n')[0])
                        elif self.mode.variable.get() == 'FI - G/L account balance and open/line item':
                            if column_tech_names[column] not in customizing.mf_gl_open_items:
                                column_names.append(column_list[column].split('\n')[0])
                        else:
                            column_names.append(column_list[column].split('\n')[0])

                    #only the first 50 fields will be included in the analysis (due to a program heaviness); this is the reason for which for specific mode some fileds will be excluded
                    maximum_field = len(column_names)
                    if len(column_names) > 50:
                        maximum_field = 50

                    column = 0
                    new_line = 0
                    for a in range(maximum_field):
                        if a == 8 or a == 16 or a == 24 or a == 32 or a == 40 or a == 48:
                            new_line += 4
                            column = 0
                        sheet_list[b].field_list[a].label.label.config(text = column_names[a].replace("*", "").replace("+", ""))
                        sheet_list[b].field_list[a].label_text = column_names[a]
                        sheet_list[b].field_list[a].label.label.grid(row = new_line + 0, column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b].field_list[a].radiobutton_1.grid(row = new_line + 1,column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b].field_list[a].radiobutton_2.grid(row = new_line + 2,column = column, sticky = tkinter.W, padx = 10)
                        sheet_list[b].field_list[a].radiobutton_3.grid(row = new_line + 3,column = column, sticky = tkinter.W, padx = 10)
                        if "*" in column_names[a]:
                            sheet_list[b].field_list[a].variable.set('Mandatory') #to have a default for mandatory fields
                        elif "+" in column_names[a]:
                            sheet_list[b].field_list[a].variable.set('Optional') #to have a default for optional fields
                        else:
                            sheet_list[b].field_list[a].variable.set('Not Required')
                        column += 1
                    
                    self.sheet_names_list.append((self.sheet_names[b], b)) #a list with only the sheets to be considered in the analysis
                
                sheet_list[b].sheet.checkbox.config (state = 'disabled')
            
            self.upload_button.button.config(state = 'disabled')
            self.go_back.config(command = lambda: go_back('fields'))
            self.go_ahead.config(command = migration_input)
        
        def upload_input_file (): #this funtion is called when you want to upload template for input values
            for a in self.sheet_names_list:
                df = pd.read_excel(self.upload_input_entry.entry.get(), a[0])

                column_list = df.columns.tolist()

                sheet_present_fields = []
                for j in self.present_fields:
                    if j[0] == a[0]:
                        sheet_present_fields.append((j[1], j[3]))
                for b in sheet_present_fields:
                    sheet_list[a[1]].field_list[b[1]].text_input.delete(1.0, tkinter.END)

                    for d in range(len(column_list)):
                        if column_list[d].replace("*", "").replace("+", "") == b[0].replace("*", "").replace("+", ""): #the program doesn't consider "*" and "+" in reading
                            rows = df.iloc[:, d].tolist()
                            for row in rows:
                                if str(row) != 'nan':
                                    sheet_list[a[1]].field_list[b[1]].text_input.insert(tkinter.END, str(row) + '\n')

        def migration_input ():
            migration_forget(sheet_list, 'No')
            self.present_fields = []
            self.error_list = []

            for a in self.sheet_names_list:
                
                df = pd.read_excel(self.entry_path.entry.get(), a[0])

                # Get the column names
                column_tech_names = df.iloc[3, :].tolist()
                column_list = df.iloc[6, :].tolist()
                column_names = []
                for column in range(len(column_list)): #similar to the one done in migration fields, but in this case also the column number is appended
                    if self.mode.variable.get() == 'Customer':
                        if column_tech_names[column] not in customizing.mf_customer_general_data and column_tech_names[column] not in customizing.mf_customer_company_data:
                            column_names.append((column_list[column].split('\n')[0], column))
                    elif self.mode.variable.get() == 'Supplier':
                        if column_tech_names[column] not in customizing.mf_supplier_general_data:
                            column_names.append((column_list[column].split('\n')[0], column))
                    elif self.mode.variable.get() == 'FI - Accounts receivable open item' or self.mode.variable.get() == 'FI - Accounts payable open item':
                        if column_tech_names[column] not in customizing.mf_bp_open_items:
                            column_names.append((column_list[column].split('\n')[0], column))
                    elif self.mode.variable.get() == 'FI - G/L account balance and open/line item':
                        if column_tech_names[column] not in customizing.mf_gl_open_items:
                            column_names.append((column_list[column].split('\n')[0], column))
                    else:
                        column_names.append((column_list[column].split('\n')[0], column))

                maximum_field = len(column_names)
                if len(column_names) > 50:
                    maximum_field = 50
                
                #for specific modes, there is a check that all the fields excluded from the analysis are blank
                if self.mode.variable.get() == 'Customer' and a[0] == 'General Data' or a[0] == 'Company Data':
                    for n in range(len(column_tech_names)):
                        all_rows = df.iloc[:, n].tolist()
                        if column_tech_names[n] in customizing.mf_customer_general_data or column_tech_names[n] in customizing.mf_customer_company_data:
                            for row in range(len(all_rows)):
                                if row >= 7:
                                    if str(all_rows[row]) != 'nan' or type(all_rows[row]) != float:
                                        self.error_list.append((a[0], 'W001', row + 2 , column_list[n].split('\n')[0].replace("*", "").replace("+", ""), 'This field is not blank, but is in a column not considered in this analysis'))

                elif self.mode.variable.get() == 'Supplier' and a[0] == 'General Data':
                    for n in range(len(column_tech_names)):
                        all_rows = df.iloc[:, n].tolist()
                        if column_tech_names[n] in customizing.mf_supplier_general_data:
                            for row in range(len(all_rows)):
                                if row >= 7:
                                    if str(all_rows[row]) != 'nan' or type(all_rows[row]) != float:
                                        self.error_list.append((a[0], 'W001', row + 2 , column_list[n].split('\n')[0].replace("*", "").replace("+", ""), 'This field is not blank, but is in a column not considered in this analysis'))

                elif (self.mode.variable.get() == 'FI - Accounts receivable open item' and a[0] == 'Customer Open Items') or (self.mode.variable.get() == 'FI - Accounts payable open item' and a[0] == 'Vendor Open Items'):
                    for n in range(len(column_tech_names)):
                        all_rows = df.iloc[:, n].tolist()
                        if column_tech_names[n] in customizing.mf_bp_open_items:
                            for row in range(len(all_rows)):
                                if row >= 7:
                                    if str(all_rows[row]) != 'nan' or type(all_rows[row]) != float:
                                        self.error_list.append((a[0], 'W001', row + 2 , column_list[n].split('\n')[0].replace("*", "").replace("+", ""), 'This field is not blank, but is in a column not considered in this analysis'))

                elif self.mode.variable.get() == 'FI - G/L account balance and open/line item' and a[0] == 'GL Balance':
                    for n in range(len(column_tech_names)):
                        all_rows = df.iloc[:, n].tolist()
                        if column_tech_names[n] in customizing.mf_gl_open_items:
                            for row in range(len(all_rows)):
                                if row >= 7:
                                    if str(all_rows[row]) != 'nan' or type(all_rows[row]) != float:
                                        self.error_list.append((a[0], 'W001', row + 2 , column_list[n].split('\n')[0].replace("*", "").replace("+", ""), 'This field is not blank, but is in a column not considered in this analysis'))

                #checks for mandatory fields (are they filled?) and for not required fields (are they all blank?)
                for b in range(maximum_field):
                    rows = df.iloc[:, column_names[b][1]].tolist()
                    if sheet_list[a[1]].field_list[b].variable.get() == 'Mandatory':
                        self.present_fields.append((a[0], sheet_list[a[1]].field_list[b].label_text, column_names[b][1], b)) #in this list will be included only fields mandatory and optional
                        sheet_list[a[1]].field_list[b].label.label.grid(row = 0, column = b, sticky = tkinter.W, padx = 10)
                        sheet_list[a[1]].field_list[b].text_input.grid(row = 1, column = b, padx = 10)
                        for row in range(len(rows)):
                            if row >= 7:
                                if str(rows[row]) == 'nan' and type(rows[row]) == float:
                                    self.error_list.append((a[0], 'E001', row + 2 , sheet_list[a[1]].field_list[b].label_text.replace("*", "").replace("+", ""), 'This mandatory field is blank'))
                    
                    elif sheet_list[a[1]].field_list[b].variable.get() == 'Optional':
                        self.present_fields.append((a[0], sheet_list[a[1]].field_list[b].label_text, column_names[b][1], b))
                        sheet_list[a[1]].field_list[b].label.label.grid(row = 0, column = b, sticky = tkinter.W, padx = 10)
                        sheet_list[a[1]].field_list[b].text_input.grid(row = 1, column = b, padx = 10)
                    
                    else:
                        for row in range(len(rows)):
                            if row >= 7:
                                if str(rows[row]) != 'nan' or type(rows[row]) != float:
                                    self.error_list.append((a[0], 'E002', row + 2 , sheet_list[a[1]].field_list[b].label_text.replace("*", "").replace("+", ""), 'This not required field is filled'))

            self.tab.select(0)

            self.download_input_button.button.config (command = lambda: create_excel_file (self.sheet_names_list, self.present_fields))
            self.upload_input_button.button.config (command = lambda: (browse_file_xlsx(self.upload_input_entry), read_file_xlsx(self.upload_input_entry.entry.get()), upload_input_file()))
            self.download_input_button.button.grid(row = 0, column = 5, padx = 10)
            self.upload_input_button.button.grid(row = 0, column = 6, padx = 10)

            self.go_back.config(command = lambda: go_back('input'))
            self.go_ahead.config(command = migration_analysis)

        def migration_analysis ():
            mode_counter = 0
            #lists useful to track speficic fields and do specific transversal checks
            bp_code = []
            postal_code = []
            country = []
            bp_code_tax = []
            tax_type = []
            tax_number = []
            bank_country = []
            bank_key = []
            bank_acc_number = []
            bank_cont_key = []
            iban = []
            for a in self.sheet_names_list: #analysis sheets loop
                key_field_list_1 =[]
                key_field_list_2 =[]
                key_field_list_3 =[]

                sheet_present_field = []

                for j in self.present_fields:
                    if j[0] == a[0]:
                        sheet_present_field.append((j[2], j[3])) #it identifies the name and the position of the analysis fields in the sheet
                
                df = pd.read_excel(self.entry_path.entry.get(), a[0])

                # Get the column names
                column_list = df.iloc[6, :].tolist()
                column_names = []
                for column in column_list:
                    column_names.append(column.split('\n')[0])
                
                # Get the column technical names
                column_tech_names = df.iloc[3, :].tolist()
                
                # Get the column details
                column_details = df.iloc[4, :].tolist()
                column_formats = []
                column_int = []
                column_dec = []
                for col_form in range(len(column_details)):
                    if not isinstance(column_details[col_form], str): #check to interrupt analysis if a column has the row 6 (technical information) blank
                        tkinter.messagebox.showerror(title="ERROR", message=f'Wrong Format in Sheet: {a[0]}, column: {column_names[col_form].replace("*", "").replace("+", "")}, row: 6')
                        return
                    column_formats.append(column_details[col_form].split(';')[3]) # column formats
                for col_numb in range(len(column_details)): #correction to be made to template maximum length for some fields
                    if column_tech_names[col_numb] in customizing.migration_file_1_max_digits:
                        column_int.append('1')
                    elif column_tech_names[col_numb] in customizing.migration_file_2_max_digits:
                        column_int.append('2')
                    elif column_tech_names[col_numb] in customizing.migration_file_3_max_digits:
                        column_int.append('3')
                    elif column_tech_names[col_numb] in customizing.migration_file_4_max_digits:
                        column_int.append('4')
                    elif column_tech_names[col_numb] in customizing.migration_file_5_max_digits:
                        column_int.append('5')
                    elif column_tech_names[col_numb] in customizing.migration_file_6_max_digits:
                        column_int.append('6')
                    elif column_tech_names[col_numb] in customizing.migration_file_7_max_digits:
                        column_int.append('7')
                    elif column_tech_names[col_numb] in customizing.migration_file_8_max_digits:
                        column_int.append('8')
                    elif column_tech_names[col_numb] in customizing.migration_file_10_max_digits:
                        column_int.append('10')
                    elif column_tech_names[col_numb] in customizing.migration_file_28_max_digits:
                        column_int.append('28')
                    else:                    
                        column_int.append(column_details[col_numb].split(';')[1]) # column max integer digits
                    column_dec.append(column_details[col_numb].split(';')[2]) # column max decimal digits
                
                # Get key columns number
                column_status = df.iloc[5, :].tolist()
                key_counter = 0
                while column_status[key_counter] == 'Key' or str(column_status[key_counter]) == 'nan':
                    key_counter += 1
                    if key_counter == len(column_status):
                        break

                for b in sheet_present_field:
                    rows = df.iloc[:, b[0]].tolist()
                    input_content = sheet_list[a[1]].field_list[b[1]].text_input.get("1.0", tkinter.END).split('\n')
                    
                    for row in range(len(rows)):
                        if row >= 7 and str(rows[row]) != 'nan': #only for rows with template data and fields not blank
                            #format and length controls
                            if column_formats[b[0]] == 'D': #date
                                if not isinstance(rows[row], datetime): # It recognizes both the SAP custom date format and the Excel date format
                                    self.error_list.append((a[0], 'E003', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'This date is in a wrong format'))

                            elif column_formats[b[0]] == 'N': #integer number
                                if (not isinstance(rows[row], int) or rows[row] > 10**int(column_int[b[0]])):
                                    self.error_list.append((a[0], 'E004', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'This number is in a wrong format'))

                            elif column_formats[b[0]] == 'P': #integer or rational number
                                if not isinstance(rows[row], int) and not isinstance(rows[row], float):
                                    self.error_list.append((a[0], 'E004', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'This number is in a wrong format'))
                                elif isinstance(rows[row], int):
                                    if rows[row] > 10**int(column_int[b[0]]):
                                        self.error_list.append((a[0], 'E005', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'The maximum length of the field is exceeded'))
                                else:
                                    if len(str(rows[row]).split('.')[0]) > int(column_int[b[0]]) or len(str(rows[row]).split('.')[1]) > int(column_dec[b[0]]):
                                        self.error_list.append((a[0], 'E005', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'The maximum length of the field is exceeded'))

                            else: #char type
                                if len(str(rows[row])) > int(column_int[b[0]]): #check for maximum length exceeded
                                    self.error_list.append((a[0], 'E005', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'The maximum length of the field is exceeded'))

                                if rows[3] in customizing.migration_file_space_forbidden_fields and ' ' in str(rows[row]): #check for balnk space present in some fields
                                    self.error_list.append((a[0], 'E010', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'The space in this field is forbidden'))
                                
                                if rows[3] == 'SMTP_ADDR' and '@' not in rows[row]: #check for "@" present in email fields
                                    self.error_list.append((a[0], 'E011', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'In this field the "@" is mandatory'))

                                #specific checks for modes
                                if self.mode.variable.get() == 'Customer':
                                    if rows[3] == 'PARVW' and str(rows[row]) not in customizing.mf_partner_function_customer:
                                        self.error_list.append((a[0], 'E012', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'In this field the standard acceptable values are "AG", "RE", "RG", "WE" and "ZM"'))
                                
                                elif self.mode.variable.get() == 'Supplier':
                                    if rows[3] == 'PARVW' and str(rows[row]) not in customizing.mf_partner_function_supplier:
                                        self.error_list.append((a[0], 'E012', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'In this field the standard acceptable values are "LF", "WL", "BA", "RS" and "ZM"'))
                                    if (rows[3] == 'WT_SUBJCT' or rows[3] == 'WEBRE') and str(rows[row]) != 'X':
                                        self.error_list.append((a[0], 'E012', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'In this field the standard acceptable value is "X"'))

                                elif self.mode.variable.get() == 'FI - Accounts receivable open item' or self.mode.variable.get() == 'FI - Accounts payable open item':
                                    if rows[3] == 'ZLSPR'  and str(rows[row]) != 'X':
                                        self.error_list.append((a[0], 'E012', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'In this field the standard acceptable value is "X"'))

                            #to track all the key fields information
                            if b[1] < key_counter:
                                key_field_list_1.append(rows[row])

                            if not all(element == '' for element in input_content): #check for input fields (blank is always considered)
                                if str(rows[row]) not in input_content:
                                    self.error_list.append((a[0], 'E007', row + 2 , sheet_list[a[1]].field_list[b[1]].label_text.replace("*", "").replace("+", ""), 'Field filled with a value not foreseen among input fields'))                         
                        if row >= 7:
                            if a[0] == 'General Data':
                                if self.mode.variable.get() == 'Customer':
                                    if rows[3] == 'KUNNR':
                                        bp_code.append(str(rows[row]))
                                elif self.mode.variable.get() == 'Supplier':
                                    if rows[3] == 'LIFNR':
                                        bp_code.append(str(rows[row]))

                                if rows[3] == 'POST_CODE1': 
                                    postal_code.append(str(rows[row])) #postal code list (transversal check with country)

                                elif rows[3] == 'COUNTRY':
                                    country.append(str(rows[row])) #country list
                            
                            elif a[0] == 'Bank Master' or a[0] == 'Bank Details':
                                if rows[3] == 'BANKS':
                                    bank_country.append(str(rows[row]))
                                elif rows[3] == 'BANKL':
                                    bank_key.append(str(rows[row]))
                                elif rows[3] == 'BANKN':
                                    bank_acc_number.append(str(rows[row]))
                                elif rows[3] == 'IBAN':
                                    iban.append(str(rows[row]))
                                elif rows[3] == 'BKONT':
                                    bank_cont_key.append(str(rows[row]))
                            
                            elif a[0] == 'Tax Numbers' and (self.mode.variable.get() == 'Customer' or self.mode.variable.get() == 'Supplier'):
                                if rows[3] == 'KUNNR' or rows[3] == 'LIFNR':
                                    bp_code_tax.append(str(rows[row]))

                                elif rows[3] == 'TAXTYPE':
                                    tax_type.append(str(rows[row]))
                                
                                elif rows[3] == 'TAXNUM':
                                    tax_number.append(str(rows[row]))

                if a[0] == 'General Data':
                    if postal_code != []:
                        for w in range(len(country)):
                            if country[w] in customizing.mf_postal_code_country: #postal code checks
                                error_postal_code = customizing.postal_code_check(postal_code[w], country[w])
                                if error_postal_code != '':
                                    self.error_list.append((a[0], 'W002', w + 9 , 'Postal Code', f'The postal code is mandatory by standard for {country[w]}; check OY17 t-code. ' + error_postal_code))
                            else:
                                self.error_list.append((a[0], 'W003', w + 9 , 'Postal Code', f'The postal code for {country[w]} is not analyzed in this program; check OY17 t-code'))
                
                elif a[0] == 'Tax Numbers':
                    if self.mode.variable.get() == 'Customer' or self.mode.variable.get() == 'Supplier':
                        for w in range(len(bp_code)):
                            for x in range(len(bp_code_tax)):
                                if bp_code[w] == bp_code_tax[x]: #vat checks
                                    if country[w] in customizing.mf_vat_country:
                                        error_vat = customizing.vat_check(tax_type[x], tax_number[x], country[w])
                                        if error_vat != '':
                                            self.error_list.append((a[0], 'W004', x + 9 , 'Tax Number', error_vat))
                                    else:
                                        self.error_list.append((a[0], 'W005', x + 9 , 'Tax Number', f'The VAT number for {country[w]} is not analyzed in this program'))

                elif a[0] == 'Bank Master' or a[0] == 'Bank Details':
                    for y in range(len(bank_country)):
                        if bank_country[y] in customizing.mf_bank_country:
                            if a[0] == 'Bank Master': #bank data checks
                                error_bank = customizing.bank_check(a[0], bank_country[y], bank_key[y])
                            else:
                                error_bank = customizing.bank_check(a[0], bank_country[y], bank_key[y], bank_acc_number[y], bank_cont_key[y], iban[y])
                            if error_bank != '':
                                self.error_list.append((a[0], 'W006', y + 9 , '', error_bank + '; check 0Y17 t-code'))
                        else:
                            self.error_list.append((a[0], 'W007', y + 9 , '', f'The bank data for {bank_country[y]} is not analyzed in this program; check OY17 t-code'))

                for c in range(int(len(key_field_list_1)/key_counter)): #check to verify that key fields are not doubled in the same sheet
                    counter = 0
                    for d in range(key_counter):
                        key_field_list_2.append(key_field_list_1[counter + c]) #to sort the key field values
                        counter += int(len(key_field_list_1)/key_counter)
                    if key_field_list_2 in key_field_list_3: #check if key field value is already present in the sheet
                        self.error_list.append((a[0], 'E006', c + 9 , '', 'These key field values are already present in the sheet'))
                    key_field_list_3.append(key_field_list_2)
                    key_field_list_2 = []
                
                #this check is done only if a specific mode is chosen (it verify that in other sheets, the key values of the main sheet are used)
                main_key_field = []
                main_key_fields_list = []
                if self.mode.variable.get() in customizing.migration_file_modes:
                    if a[0] in customizing.migration_file_main_sheet: 
                        self.mode_key_fields = key_field_list_3 #to track all the key field values for the main sheet
                        mode_counter = len(key_field_list_3[0])
                    else:
                        for k in range(len(key_field_list_3)):
                            main_key_field = key_field_list_3[k][:mode_counter]
                            if main_key_field not in self.mode_key_fields:
                                self.error_list.append((a[0], 'E008', k + 9 , '', 'These key field values are not present in the main sheet'))                         
                            main_key_fields_list.append(main_key_field)

                        if a[0] in customizing.migration_file_secondary_sheets: #for secondary sheets, all the key values of the main sheets must be present
                            for n in range(len(self.mode_key_fields)):
                                if self.mode_key_fields[n] not in main_key_fields_list:
                                    self.error_list.append((a[0], 'E009', '' , '', f'The {self.mode_key_fields[n]} key field values are not in this sheet'))                         


            for widget in frame.winfo_children():
                widget.destroy()   

            lst = [('SHEET', 'MESSAGE ID', 'ROW', 'FIELD', 'ERROR MESSAGE')]
            errors_number = 0
            warnings_number = 0
            for a in range(len(self.error_list)):
                lst.append(self.error_list[a])
                if 'E' in self.error_list[a][1]:
                    errors_number += 1
                elif 'W' in self.error_list[a][1]:
                    warnings_number += 1

            tree = customizing.Treeview (
                frame = frame,
                col_text = ['', '', '', '', ''],
                width_list = [120, 30, 30, 120, 540],
                lst = lst
            )

            image_error_path = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\red_cross.png'
            image_warning_path = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\yellow_triangle.png'

            # Open the image using Pillow
            image_error_img = Image.open(image_error_path)
            image_warning_img = Image.open(image_warning_path)
            # Convert the image to a format Tkinter supports
            image_error = ImageTk.PhotoImage(image_error_img)
            image_warning = ImageTk.PhotoImage(image_warning_img)

            label_error = tkinter.Label(tree.frame_1.frame, image=image_error, text = f'Errors: {errors_number}', compound=tkinter.LEFT, font = ('Calibri', 14, 'bold'), background = '#F0F8FF')
            label_error.grid(column=1, row=0, padx=200, pady=10)

            label_warning = tkinter.Label(tree.frame_1.frame, image=image_warning, text = f'Warnings: {warnings_number}', compound=tkinter.LEFT, font = ('Calibri', 14, 'bold'), background = '#F0F8FF')
            label_warning.grid(column=2, row=0, padx=20, pady=10)

            # Keep a reference to the image to prevent it from being garbage collected
            label_error.image = image_error
            label_warning.image = image_warning

        self.frame = frame

        self.tab = ttk.Notebook (self.frame, height=680, width=1320)

        self.sheet_names = []

        self.sheet_names_list = []

        self.mode_key_fields = []

        self.present_fields = []

        self.error_list = []

        self.entry_path = customizing.Entry(
            frame = self.frame,
            width = 80,
            column = 1,
            entry_pady = 10
            )
        
        image_path_back = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\above_thearrow_1550 (1).png'
        self.button_icon_back = tkinter.PhotoImage(file=image_path_back)

        # Create a button with the resized image
        self.go_back = tkinter.Button(self.frame, text="",command = home, image=self.button_icon_back, compound=tkinter.LEFT, background = '#F0F8FF')
        self.go_back.grid(row = 0, column = 3, padx = 5)

        #button to make fields appear
        # Load an image for the button icon
        image_path = 'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\Next_arrow_1559 (1).png'
        self.button_icon = tkinter.PhotoImage(file=image_path)

        # Create a button with the resized image
        self.go_ahead = tkinter.Button(self.frame, text="",command = migration_fields, image=self.button_icon, compound=tkinter.LEFT, background = '#F0F8FF')
        self.go_ahead.grid(row = 0, column = 4, padx = 5)
        self.go_ahead.config(state = 'disabled')

        #button to upload the .xlsx file
        self.upload_button = customizing.Button(
            frame = self.frame, 
            text = "Upload .xlsx file", 
            command = lambda: (browse_file_xlsx (entry_path=self.entry_path), read_file_xlsx (entry_path_str=self.entry_path.entry.get()), sheet_checkboxes ()),
            pady = 10
            )

        self.sheet_1 = customizing.Sheet (frame = frame, tab = self.tab)
        self.sheet_2 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_3 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_4 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_5 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_6 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_7 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_8 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_9 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_10 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_11 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_12 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_13 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_14 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_15 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_16 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_17 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_18 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_19 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_20 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_21 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_22 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_23 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_24 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_25 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_26 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_27 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_28 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_29 = customizing.Sheet (frame = frame, tab = self.tab) 
        self.sheet_30 = customizing.Sheet (frame = frame, tab = self.tab)

        self.download_input_button = customizing.Button (
            frame = self.frame,
            text = 'Download Input Template',
            width = 22
        )
        self.download_input_button.button.grid_forget()

        self.upload_input_button = customizing.Button (
            frame = self.frame,
            text = 'Upload Input Template',
            width = 22
        )
        self.upload_input_button.button.grid_forget()

        self.upload_input_entry = customizing.Entry (
            frame = frame
        )
        self.upload_input_entry.entry.grid_forget()
        self.upload_input_entry.label.label.grid_forget()

        self.mode_frame = customizing.Frame (
            root = self.frame,
            column = 1,
            row = 1,
            row_span = 100
        )

        self.mode = customizing.RadioButton_2 (
            frame = self.mode_frame.frame,
            label_text = 'Analysis Mode',
            text_1 = 'Generic',
            dimension = 15
        )
        self.mode_frame.frame.grid_forget()
        

        sheet_list = [self.sheet_1, self.sheet_2, self.sheet_3, self.sheet_4, self.sheet_5, self.sheet_6, self.sheet_7, self.sheet_8, self.sheet_9, self.sheet_10, self.sheet_11, self.sheet_12, self.sheet_13, self.sheet_14, self.sheet_15, self.sheet_16, self.sheet_17, self.sheet_18, self.sheet_19, self.sheet_20, self.sheet_21, self.sheet_22, self.sheet_23, self.sheet_24, self.sheet_25, self.sheet_26, self.sheet_27, self.sheet_28, self.sheet_29, self.sheet_30]
        
        migration_forget (sheet_list, 'Yes')
       
mainroot = customizing.Root (root_title = 'SAP HELPER')

mainframe = customizing.Frame (
            root = mainroot.root,
            pack_or_grid = 'P'
        )

main_menubar = customizing.MenuBar (
    root = mainroot.root,
    first_label = 'File',
    second_label = 'Go To',
    third_label = 'Help'
)

main_menubar.main_menu_1.add_command (label = 'Quit', command = lambda: quit_window (mainroot.root), font = ('Calibri', 12))
main_menubar.main_menu_2.add_command (label = 'Home', command = home, font = ('Calibri', 12))
main_menubar.main_menu_2.add_command (label = 'Ebs MT940', command = lambda: start('ebs_mt940', mainframe.frame), font = ('Calibri', 12))
main_menubar.main_menu_2.add_command (label = 'IBAN', command = lambda: start('iban', mainframe.frame), font = ('Calibri', 12))
main_menubar.main_menu_2.add_command (label = 'Migration File', command = lambda: start('migration_file', mainframe.frame), font = ('Calibri', 12))
main_menubar.main_menu_3.add_command (label = 'EBS MT940 info', command = lambda: info('EBS MT940 info'), font = ('Calibri', 12))
main_menubar.main_menu_3.add_command (label = 'IBAN info', command = lambda: info('IBAN info'), font = ('Calibri', 12))
main_menubar.main_menu_3.add_command (label = 'Migration File info', command = lambda: info('Migration File info'), font = ('Calibri', 12))


main = MainRoot (root = mainroot.root, frame = mainframe.frame)



mainroot.root.mainloop()


