import tkinter
from tkinter import ttk
import tkinter.messagebox
import pyperclip
import re

ebs_mt940_dict = {
    'FCHG': 'Charges and other expenses',
    'FCHK': 'Cheques',
    'FINT': 'Interest',
    'FRTI': 'Returned item',
    'NBOE': 'Bill of exchange',
    'NCHG': 'Charges and other expenses',
    'NCHK': 'Cheques',
    'NCOL': 'Collections',
    'NCOM': 'Commisions',
    'NDCR': 'Documentary credit',
    'NDDT': 'Direct Debit Item',
    'NDIV': 'Dividends-Warrants',
    'NECK': 'Eurocheques',
    'NEQA': 'Equivalent amount',
    'NFEX': 'Foreign exchange',
    'NINT': 'Interest',
    'NLDP': 'Loan deposit',
    'NMSC': 'Miscellaneous',
    'NRTI': 'Returned item',
    'NSEC': 'Securities',
    'NSTO': 'Standing order',
    'NTCK': 'Travellers cheques',
    'NTRF': 'Transfer',
    'S100': 'SWIFT message 100',
    'S101': 'SWIFT message 101',
    'S103': 'SWIFT message 103',
    'S190': 'SWIFT message 190',
    'S191': 'SWIFT message 191',
    'S200': 'SWIFT message 200',
    'S201': 'SWIFT message 201',
    'S202': 'SWIFT message 202',
    'S203': 'SWIFT message 203',
    'S205': 'SWIFT message 205',
    'S300': 'SWIFT message 300',
    'S320': 'SWIFT message 320',
    'S400': 'SWIFT message 400',
    'S554': 'SWIFT message 554',
    'S556': 'SWIFT message 556',

}

migration_file_modes = ['Fixed asset', 'Customer', 'Supplier']

migration_file_main_sheet = ['Master Details', 'General Data']

#a list of sheets for which is mandatory to have all the key value of the main sheet
migration_file_secondary_sheets = ['Posting Information', 'Time-Dependent Data', 'Depreciation Areas', 'Cumulative Values', #ASSET
                                   'BP Roles', 'Company Data', 'Sales Data', 'Sales Partner', 'Output Tax', #CUSTOMER
                                   'Purchasing Organization Data'] #VENDOR

#a list of technical name fields for which is forbidden to have spaces
migration_file_space_forbidden_fields = ['BUKRS', 'ANLN1', 'ANLN2', 'ANLKL', 'GSBER', 'KOSTL', 'WERKS', 'AFABE', 'ASSETTRTYP', 'CURRENCY', #ASSET
                                         'KUNNR', 'BU_GROUP', 'BPEXT', 'COUNTRY', 'REGION', 'LANGU_CORR', 'SMTP_ADDR', #CUSTOMER - GENERAL DATA
                                         'BP_ROLE', 'MAHNA', 'ZTERM', 'ZWELS_01', 'ZWELS_02', 'ZWELS_03', 'ZWELS_04', 'HBKID', 'AKONT', 'WITHT', 'WT_WITHCD', #CUSTOMER - BP ROLES/COMPANY DATA/WHT
                                         'VKORG', 'VTWEG', 'SPART', 'KDGRP', 'BZIRK', 'VKBUR', 'WAERS', 'KONDA', 'KALKS', 'LPRIO', 'VSBED', 'INCO1', 'KTGRD', 'PARVW', 'KUNN2', 'ALAND', 'TATYP', 'TAXKD', #CUSTOMER - SALES DATA/SALES PARTNER/OUTPUT TAX
                                         'BANKS', 'BANKL', 'BANKN', 'IBAN', 'BKONT', 'TAXTYPE', 'TAXNUM', #CUSTOMER - BANK DATA/TAX NUMBER 
                                         'LIFNR', 'FRGRP', 'ZTERM1', 'REPRF', 'WT_SUBJCT', #VENDOR - COMPANY DATA
                                         'EKORG', 'EKGRP', 'WEBRE', 'BSTAE', 'LIFN2'] #VENDOR - PURCHASING ORGANIZATION DATA/PARTNER FUNCTION

#fields not to be considered in the related sheet for customer template (there is only a check about these fields are blank)
mf_customer_general_data = ['LEGAL_ENTY', 'LEGAL_ORG', 'FOUND_DAT', 'LIQUID_DAT', 'LOCATION_1', 'LOCATION_2', 'LOCATION_3', 'BAHNE', 'BAHNS', 'COUNC', 'CITYC', 'DTAMS', 'DTAWS', 'KNRZA', 'NIELS', 'RPMKR', 'KUKLA', 'HZUOR', 'BRAN1', 'BRAN2', 'BRAN3', 'BRAN4', 'BRAN5', 'KATR1', 'KATR2', 'KATR3', 'KATR4', 'KATR5', 'KATR6', 'KATR7', 'KATR8', 'KATR9', 'KATR10', 'SUFRAMA', 'RG', 'EXP', 'UF', 'RGDATE', 'RIC', 'RNE', 'RNEDATE', 'CNAE', 'LEGALNAT', 'CRTN', 'ICMSTAXPAY', 'INDTYP', 'TDT', 'COMSIZE', 'DECREGPC', 'ECC_NO', 'EXC_REG_NO', 'EXC_RANGE', 'EXC_DIV', 'EXC_COMM', 'EXC_TAX_IND', 'CST_NO', 'LST_NO', 'SERV_TAX_NO', 'PAN_NO', 'PAN_REF_NO', 'BON_AREA_CONF', 'DON_MARK', 'CONSOLIDATE_INVOICE', 'ALLOWANCE_TYPE', 'EINVOICE_MODE', 'J_1KFTBUS', 'J_1KFTIND', 'J_1KFREPRE', 'PH_BIZ_STYLE', 'CITY2', 'HOME_CITY', 'TIME_ZONE', 'LZONE', 'BUILDING', 'ROOM', 'FLOOR', 'CO_NAME', 'HOUSE_NO2', 'STR_SUPPL3', 'LOCATION', 'TXJCD', 'NOTE_TELNR', 'TELNR_LONG_2', 'NOTE_TELNR_2', 'TELNR_LONG_3', 'NOTE_TELNR_3', 'NOTE_MOBILE', 'MOBILE_LONG_2', 'NOTE_MOBILE_2', 'MOBILE_LONG_3', 'NOTE_MOBILE_3', 'NOTE_FAXNR', 'FAXNR_LONG_2', 'NOTE_FAXNR_2', 'FAXNR_LONG_3', 'NOTE_FAXNR_3', 'NOTE_SMTP', 'SMTP_ADDR_2', 'NOTE_SMTP_2', 'SMTP_ADDR_3', 'NOTE_SMTP_3', 'URI_TYP', 'URI_ADDR', 'NOTE_URI', 'SPERR', 'COLLMAN']
mf_customer_company_data = ['TLFNS', 'TLFXS', 'INTAD']

#fields not to be considered in the related sheet for supplier template (there is only a check about these fields are blank)
mf_supplier_general_data = ['LEGAL_ENTY', 'LEGAL_ORG', 'FOUND_DAT', 'LIQUID_DAT', 'LOCATION_1', 'LOCATION_2', 'LOCATION_3', 'DTAMS', 'DTAWS', 'LNRZA', 'ESRNR', 'TERM_LI', 'MIN_COMP', 'COMSIZE', 'DECREGPC', 'CRC_NUM', 'RG', 'EXP', 'UF', 'RGDATE', 'RIC', 'RNE', 'RNEDATE', 'CNAE', 'LEGALNAT', 'CRTN', 'ICMSTAXPAY', 'INDTYP', 'TDT', 'J_1IEXCD', 'J_1IEXRN', 'J_1IEXRG', 'J_1IEXDI', 'J_1IEXCO', 'J_1IVTYP', 'J_1I_CUSTOMS', 'J_1IEXCIVE', 'J_1ISSIST', 'J_1IVENCRE', 'J_1ICSTNO', 'J_1ILSTNO', 'J_1ISERN', 'J_1IPANNO', 'J_1IPANREF', 'J_1IPANVALDT', 'J_1IDEDREF', 'VEN_CLASS', 'J_1KFTBUS', 'J_1KFTIND', 'J_1KFREPRE', 'CATEG', 'STATUS', 'VFNUM', 'VFNID', 'PARTNER_NAME', 'PARTNER_UTR', 'CRN', 'ALLOWANCE_TYPE', 'AU_CARRYING_ENT', 'AU_IND_UNDER_18', 'AU_PAYMENT_NOT_EXCEED_75', 'AU_WHOLLY_INP_TAXED', 'AU_PARTNER_WITHOUT_GAIN', 'AU_NOT_ENTITLED_ABN', 'AU_PAYMENT_EXEMPT', 'AU_PRIVATE_HOBBY', 'AU_DOMESTIC_NATURE', 'SC_CAPITAL', 'SC_CURRENCY', 'CITY2', 'HOME_CITY', 'TIME_ZONE', 'LZONE', 'BUILDING', 'ROOM', 'FLOOR', 'CO_NAME', 'HOUSE_NO2', 'STR_SUPPL3', 'LOCATION', 'TXJCD', 'NOTE_TELNR', 'TELNR_LONG_2', 'NOTE_TELNR_2', 'TELNR_LONG_3', 'NOTE_TELNR_3', 'NOTE_MOBILE', 'MOBILE_LONG_2', 'NOTE_MOBILE_2', 'MOBILE_LONG_3', 'NOTE_MOBILE_3', 'NOTE_FAXNR', 'FAXNR_LONG_2', 'NOTE_FAXNR_2', 'FAXNR_LONG_3', 'NOTE_FAXNR_3', 'NOTE_SMTP', 'SMTP_ADDR_2', 'NOTE_SMTP_2', 'SMTP_ADDR_3', 'NOTE_SMTP_3', 'URI_TYP', 'URI_ADDR', 'NOTE_URI', 'SPERR', 'SPERM']

print(len(mf_supplier_general_data))

migration_file_1_max_digits = ['TAXKD', #CUSTOMER
                               'REPRF', 'WT_SUBJCT', 'WEBRE'] #VENDOR

migration_file_2_max_digits = ['AFABE', #ASSET
                               'WITHT', 'WT_WITHCD', 'VTWEG', 'SPART', 'KDGRP', 'KONDA', 'KALKS', 'LPRIO', 'VSBED', 'KTGRD', 'PARVW'] #CUSTOMER

migration_file_3_max_digits = ['MEINS', 'ASSETTRTYP', #ASSET
                               'COUNTRY', 'REGION', 'INCO1', 'ALAND', 'BANKS', #CUSTOMER
                               'EKGRP'] #VENDOR 

migration_file_4_max_digits = ['BUKRS', 'WERKS', 'GSBER', 'EVALGROUP1', 'EVALGROUP2', 'EVALGROUP3', 'EVALGROUP4', 'AFASL', #ASSET
                               'BU_GROUP', 'KTOKD', 'FRGRP', 'ZTERM', 'MAHNA', 'VKORG', 'VKBUR', 'TATYP', 'TAXTYPE', #CUSTOMER
                               'FRGRP', 'ZTERM1', 'EKORG', 'BSTAE'] #VENDOR

migration_file_5_max_digits = ['CURRENCY', #ASSET
                               'HBKID', 'WAERS'] #CUSTOMER

migration_file_6_max_digits = ['BZIRK'], #CUSTOMER

migration_file_7_max_digits = ['BP_ROLE'] #CUSTOMER

migration_file_8_max_digits = ['ANLKL', 'EVALGROUP5'] #ASSET

migration_file_10_max_digits = ['KOSTL', 'VENDOR_NO', #ASSET
                                'KUNNR', 'AKONT', 'FDGRV', 'ZWELS_01', 'ZWELS_02', 'ZWELS_03', 'ZWELS_04', 'KUNN2', #CUSTOMER
                                'LIFNR'] #VENDOR

migration_file_28_max_digits = ['INCO2'] #CUSTOMER

mf_partner_function_customer = ['AG', 'RE', 'RG', 'WE', 'ZM']

mf_partner_function_supplier = ['LF', 'WL', 'BA', 'RS', 'ZM']

class Root ():
    def __init__(
        self,
        root_title:str,
        tk_or_toplevel: str = 'TK',
        root_geometry:str = '600x400',
        background: str = 'white'
    ):
        if tk_or_toplevel == 'TK':
            self.root = tkinter.Tk()
        else:
            self.root = tkinter.Toplevel()
        self.root.title(root_title)
        self.root.configure(bg = '#F0F8FF')
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        icon_path = r'C:\\Users\\scham\\OneDrive\\Desktop\\SAP HELPER\\Icon\\communication_assistance_help_support_service_information_icon_230472.ico'
        self.root.iconbitmap(icon_path)

        self.root.state('zoomed')

        #self.root.geometry(root_geometry)
        #self.root.minsize(1600, 1200)
        #self.root.attributes('-fullscreen',True)

class Frame ():
    def __init__(
        self,
        root: None,
        pack_or_grid: str = 'G',
        left_or_right: str = 'O',
        background: str = '#F0F8FF',
        column: int = 0,
        row: int = 0,
        sticky: str = '',
        col_span: int = 1,
        row_span: int = 1
    ):
        self.frame = tkinter.Frame(root, background = background)
        if pack_or_grid.upper() == 'P':
            if left_or_right.upper() == 'R':
                self.frame.pack(fill = tkinter.BOTH, expand = tkinter.TRUE, side = 'right')
            elif left_or_right.upper() == 'L':
                self.frame.pack(fill = tkinter.BOTH, expand = tkinter.TRUE, side = 'left')
            else:
                self.frame.pack(fill = tkinter.BOTH)
        else:
            self.frame.grid(column = column, row = row, sticky = sticky, columnspan = col_span, rowspan = row_span)
            self.frame.columnconfigure(0, weight=1)  # Ensure column 0 expands

class Button ():
    def __init__ (
        self,
        frame: tkinter.Frame,
        text: str,
        command = '',
        width: int = 0,
        height:int = 0,
        column:int = 0,
        row:int = 0,
        padx:int = 0,
        pady:int = 0,
        sticky:str = tkinter.W,
        dimension: int = 12
    ):
        self.button = tkinter.Button(frame, text = text, command = command, width = width, height = height, background = '#D8E6EC', font = ('Calibri', dimension, 'bold'))
        self.button.grid(column = column, row = row, padx = padx, pady = pady, sticky = sticky)

class Label ():
    def __init__ (
        self,
        frame: None,
        text: str,
        dimension: int = 13,
        column: int = 0,
        row:int = 0,
        padx: int = 0,
        pady:int = 0,
        sticky: str = '',
        justify = tkinter.CENTER,
        foreground: str = 'black',
        columnspan: int = 1
    ):
        self.label = tkinter.Label(frame, text = text, font = ('Calibri', dimension), justify = justify, background = '#F0F8FF', foreground = foreground)
        self.label.grid(column = column, row = row, padx = padx, pady = pady, sticky = sticky, columnspan = columnspan)

class Entry ():
    def __init__ (
        self,
        frame: None,
        text: str = '',
        column: int = 0,
        row:int = 0,
        label_padx:int = 0,
        label_pady:int = 0,
        width:int = 0,
        entry_padx:int = 0,
        entry_pady:int = 0,
        columnspan: int = 1
    ):
        self.label = Label (
            frame = frame, 
            text = text,
            column = column,
            row = row,
            padx = label_padx,
            pady = label_pady
            )
        self.stringvar = tkinter.StringVar()
        self.entry = tkinter.Entry(frame, text = self.stringvar, width = width)
        self.entry.grid(column = column+1, row = row, padx = entry_padx, pady = entry_pady, columnspan = columnspan)

class Combobox ():
    def __init__ (
        self,
        frame: None,
        command = '',
        values = '',
        width: int = 12,
        column:int = 1,
        row:int = 0,
        padx:int = 10,
        pady:int = 5,
        columnspan: int = 1,
        sticky: str = ''
    ):
        self.text = tkinter.StringVar()
        self.combobox = ttk.Combobox(frame, textvariable = self.text, state = 'readonly', values = values, width = width)
        self.combobox.grid(column = column, row = row, padx = padx, pady = pady, sticky = sticky, columnspan = columnspan)
        self.combobox.bind("<<ComboboxSelected>>", command)

class Checkbox ():
    def __init__(
        self,
        frame: None,
        text:str = '',
        command = '',
        column:int = 0,
        row:int = 0,
        sticky: str = '',
        columnspan: int = 1
    ):
        self.text = text
        self.variable = tkinter.IntVar()
        self.checkbox = tkinter.Checkbutton(frame, text = self.text, variable = self.variable, command = command, background = '#F0F8FF')
        self.checkbox.grid(column = column, row = row, sticky = sticky, columnspan = columnspan)

class MenuBar ():
    def __init__ (
        self,
        root: tkinter.Tk,
        first_label: str,
        second_label: str
    ): 
        self.menubar = tkinter.Menu (root)
        root.config (menu = self.menubar)
        self.main_menu_1 = tkinter.Menu (self.menubar, tearoff = 0)
        self.main_menu_2 = tkinter.Menu (self.menubar, tearoff = 0)
        self.menubar.add_cascade (label = first_label, menu = self.main_menu_1)
        self.menubar.add_cascade (label = second_label, menu = self.main_menu_2)

class Treeview():
    def __init__(
        self,
        frame: tkinter.Frame,
        col_text: list,
        width_list : list,
        lst: list
    ):
        def select_all():
            # Get all item IDs in the Treeview
            item_ids = self.tree.get_children()

            # Select all items
            self.tree.selection_set(item_ids)
            
            on_copy(event = True)

        def on_copy(event):
            selected_items = self.tree.selection()
            if selected_items:
                copied_data = []
                for item_id in selected_items:
                    item_values = self.tree.item(item_id, 'values')
                    copied_data.append('\t'.join(str(value) for value in item_values))

                copied_text = '\n'.join(copied_data)
                pyperclip.copy(copied_text)

        self.frame = frame

        self.frame_1 = Frame (
            root = self.frame,
            pack_or_grid = 'P'
        )
        self.frame_2 = Frame (
            root = self.frame,
            pack_or_grid = 'P'
        )

        self.tree = ttk.Treeview(self.frame_2.frame, columns=col_text, show='headings', height=len(lst))
        for a in range(len(col_text)):
            self.tree.heading(col_text[a], text=col_text[a])
            self.tree.column(col_text[a], width=width_list[a])
        for b in lst:
            self.tree.insert('', tkinter.END, values=b)
        self.tree.pack(side="left", fill="both", expand=True)

        # Add a vertical scrollbar
        self.y_scrollbar = ttk.Scrollbar(self.frame_2.frame, orient='vertical', command=self.tree.yview)
        self.y_scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.y_scrollbar.set)

        self.tree.bind("<Control-c>", on_copy)

        first_item = self.tree.get_children()[0]
        self.tree.selection_set(first_item)

        self.select_all_button = tkinter.Button (self.frame_1.frame, text = 'Select All and Copy', command = select_all, background = '#D8E6EC', font = ('Calibri', 12, 'bold'))
        self.select_all_button.pack(anchor = 'e', pady=10)

class RadioButton_2 ():
    def __init__ (
        self,
        frame: tkinter.Frame,
        label_text: str = '',
        text_1: str = '',
        text_2: str = '',
        command = '',
        dimension: int = 8,
        row: int = 0,
        column: int = 0
    ):
        self.label_text = label_text
        self.text_1 = text_1
        self.text_2 = text_2
        self.label = Label (frame, text = label_text, row = row, column = column, dimension = dimension + 2)
        self.label.label.config(foreground='#191970')
        self.variable = tkinter.StringVar()
        self.radiobutton_1 = tkinter.Radiobutton (frame, text = text_1, variable = self.variable, value = text_1, command = command, font = ('Calibri', dimension), background = '#F0F8FF')
        self.radiobutton_1.grid(row = row + 1, column = column, sticky = tkinter.W)
        self.radiobutton_2 = tkinter.Radiobutton (frame, text = text_2, variable = self.variable, value = text_2, command = command, font = ('Calibri', dimension), background = '#F0F8FF')
        self.radiobutton_2.grid(row = row + 2, column = column, sticky = tkinter.W)

class RadioButton_3 (RadioButton_2):
    def __init__ (
        self,
        frame: tkinter.Frame,
        label_text: str = '',
        text_1: str = '',
        text_2: str = '',
        text_3: str = '',
        command = '',
        dimension: int = 9,
        row: int = 0,
        column: int = 0
    ):
        super ().__init__(
            frame,
            label_text,
            text_1,
            text_2,
            command,
            dimension,
            row,
            column
        )
        self.text_3 = text_3
        self.radiobutton_3 = tkinter.Radiobutton (frame, text = text_3, variable = self.variable, value = text_3, command = command,font = ('Calibri', dimension), background = '#F0F8FF')
        self.radiobutton_3.grid(row = row + 3, column = column)
        self.text_input = tkinter.Text(frame, height = 37, width = 15)

class ScrollableFrame(tkinter.Frame):
    def __init__(self, master=None, **kwargs):
        tkinter.Frame.__init__(self, master, **kwargs)

        # Create a canvas and add it to the frame
        self.canvas = tkinter.Canvas(self, borderwidth=0, highlightthickness=0, background = '#F0F8FF')

        # Create a frame inside the canvas to hold the widgets
        self.inner_frame = tkinter.Frame(self.canvas, background = '#F0F8FF')

        # Add a horizontal scrollbar and link it to the canvas
        self.h_scrollbar = tkinter.Scrollbar(self, orient="horizontal", command=self.canvas.xview, background = '#F0F8FF')
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        # Pack the canvas and scrollbar into the frame
        self.canvas.pack(side="top", fill = tkinter.BOTH, expand = tkinter.TRUE)
        self.h_scrollbar.pack(side="bottom", fill="x")

        # Add the inner frame to the canvas
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        # Configure the canvas to update the scroll region when the frame size changes
        self.inner_frame.bind("<Configure>", self.on_inner_frame_configure)

        # Bind the canvas to respond to mousewheel events for scrolling
        #self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

    def on_inner_frame_configure(self, event):
        # Update the scroll region to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    #def on_mousewheel(self, event):
        # Handle mousewheel scrolling
    #    if event.delta:
    #        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_widget(self, widget, **kwargs):
        # Add a widget to the inner frame
        widget.grid(**kwargs)

    def add_widgets(self, *widgets):
        # Add multiple widgets to the inner frame
        for widget in widgets:
            widget.grid()

class Sheet ():
    def __init__ (
        self,
        frame: tkinter.Frame,
        tab: ttk.Notebook
    ):
        self.sheet = Checkbox (frame)
        self.main_frame = Frame (tab, 'P')
        self.scrollable_frame = ScrollableFrame(self.main_frame.frame)
        self.scrollable_frame.pack(side = "top", fill = tkinter.BOTH, expand = tkinter.TRUE)
        self.frame_tab = self.scrollable_frame.inner_frame
        self.field_1 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_2 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_3 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_4 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        self.field_5 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_6 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_7 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_8 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_9 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_10 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_11 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_12 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_13 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_14 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_15 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_16 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_17 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_18 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_19 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_20 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_21 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_22 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_23 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_24 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_25 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_26 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_27 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_28 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_29 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        self.field_30 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_31 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_32 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_33 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_34 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_35 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_36 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_37 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_38 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_39 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_40 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_41 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_42 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_43 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_44 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_45 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_46 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_47 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_48 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_49 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_50 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_51 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_52 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_53 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_54 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_55 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_56 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_57 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_58 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_59 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_60 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_61 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_62 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_63 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_64 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_65 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_66 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_67 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_68 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_69 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_70 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_71 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_72 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_73 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_74 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_75 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_76 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_77 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_78 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_79 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_80 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_81 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_82 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_83 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_84 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_85 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_86 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_87 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_88 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_89 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_90 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_91 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_92 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_93 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_94 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_95 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_96 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_97 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_98 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_99 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_100 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_101 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_102 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_103 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_104 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_105 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_106 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_107 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_108 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_109 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_110 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_111 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_112 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_113 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_114 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_115 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_116 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_117 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_118 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_119 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_120 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_121 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_122 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_123 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_124 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_125 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_126 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_127 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_128 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_129 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_130 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_131 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_132 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_133 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_134 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_135 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_136 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_137 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_138 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_139 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_140 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_141 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_142 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_143 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_144 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_145 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_146 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_147 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_148 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_149 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_150 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_151 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_152 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_153 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_154 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_155 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_156 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_157 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_158 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_159 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_160 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_161 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_162 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_163 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_164 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_165 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_166 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_167 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_168 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_169 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_170 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_171 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_172 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_173 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_174 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_175 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_176 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_177 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_178 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_179 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required') 
        #self.field_180 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_181 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_182 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_183 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_184 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_185 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_186 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_187 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_188 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_189 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_190 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_191 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_192 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_193 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_194 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_195 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_196 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_197 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_198 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_199 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        #self.field_200 = RadioButton_3 (self.frame_tab, text_1 = 'Mandatory', text_2 = 'Optional', text_3 = 'Not Required')
        self.field_list = [self.field_1, self.field_2, self.field_3, self.field_4, self.field_5, self.field_6, self.field_7, self.field_8, self.field_9, self.field_10,
                           self.field_11, self.field_12, self.field_13, self.field_14, self.field_15, self.field_16, self.field_17, self.field_18, self.field_19, self.field_20,
                           self.field_21, self.field_22, self.field_23, self.field_24, self.field_25, self.field_26, self.field_27, self.field_28, self.field_29, self.field_30,
                           self.field_31, self.field_32, self.field_33, self.field_34, self.field_35, self.field_36, self.field_37, self.field_38, self.field_39, self.field_40,
                           self.field_41, self.field_42, self.field_43, self.field_44, self.field_45, self.field_46, self.field_47, self.field_48, self.field_49, self.field_50]
                           #self.field_51, self.field_52, self.field_53, self.field_54, self.field_55, self.field_56, self.field_57, self.field_58, self.field_59, self.field_60,
                           #self.field_61, self.field_62, self.field_63, self.field_64, self.field_65, self.field_66, self.field_67, self.field_68, self.field_69, self.field_70,
                           #self.field_71, self.field_72, self.field_73, self.field_74, self.field_75, self.field_76, self.field_77, self.field_78, self.field_79, self.field_80,
                           #self.field_81, self.field_82, self.field_83, self.field_84, self.field_85, self.field_86, self.field_87, self.field_88, self.field_89, self.field_90,
                           #self.field_91, self.field_92, self.field_93, self.field_94, self.field_95, self.field_96, self.field_97, self.field_98, self.field_99, self.field_100

def postal_code_check (postal_code: str, country: str):
    error = ''
    if country == 'AD' and not re.compile(r'^AD\d{3}$').match(postal_code):
        error = 'Should be in format "ANA NAN"'
    elif country == 'CA' and not re.compile(r'^[A-Z]{1}\d{1}[A-Z]{1} \d{1}[A-Z]{1}\d{1}$').match(postal_code):
        error = 'Should be in format "ANA NAN"'
    elif (country == 'CZ' or country == 'GR' or country == 'SE' or country == 'SK') and not re.compile(r'^\d{3} \d{2}$').match(postal_code):
        error = 'Should be in format "NNN NN"'
    elif (country == 'DE' or country == 'IT' or country == 'FR' or country == 'ES') and not re.compile(r'^\d{5}$').match(postal_code):
        error = 'Should be in format "NNNNN"'
    elif country == 'GB' and len(postal_code) > 9:
        error = 'Should not have a length greater than 9'
    elif country == 'MT' and not re.compile(r'^[A-Z]{3} \d{4}$').match(postal_code):
        error = 'Should be in format "AAA NNNN"'
    elif country == 'NL' and not re.compile(r'^\d{4} [A-Z]{2}$').match(postal_code):
        error = 'Should be in format "NNNN AA"'
    elif country == 'PL' and not re.compile(r'^\d{2}-\d{3}$').match(postal_code):
        error = 'Should be in format "NN-NNN"'
    elif country == 'PT' and not re.compile(r'^\d{4}-\d{3}$').match(postal_code):
        error = 'Should be in format "NNNN-NNN"'

    return error


def vat_check (tax_type: str, tax_number:str, country: str):
    error = ''
    if country == 'IT':
        if tax_type == 'IT0' and not re.compile(r'^IT\d{11}$').match(tax_number):
            error = f'For country {country} and tax type {tax_type}, the tax number should be in format "ITNNNNNNNNNNN"'
        elif tax_type == 'IT1' and not re.compile(r'^[A-Z]{6}\d{2}[A-Z]{1}\d{2}[A-Z]{1}\d{3}[A-Z]{1}$').match(tax_number):
            error = f'For country {country} and tax type {tax_type}, the tax number should be in format "AAAAAANNANNANNNA"'
        elif tax_type == 'IT2' and not re.compile(r'^\d{11}$').match(tax_number):
            error = f'For country {country} and tax type {tax_type}, the tax number should be in format "NNNNNNNNNNN"'

    return error