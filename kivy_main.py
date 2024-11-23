from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.widget import Widget
from kivy.graphics import Color, Line
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.dropdown import DropDown
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from datetime import datetime
from kivymd.app import MDApp
import ctypes

import kivy_customizing
import kivy_engine


class MyApp(MDApp):
    def build(self):

        self.error = ''

        self.sm = ScreenManager()

        # Create main screen
        main_screen = Screen(name='Main')

        # Set the background color of the entire window
        self.clearcolor = (0.941, 0.973, 1, 1)

        Window.maximize()

        self.title = "SAP Helper"

        # Create a self.layout1
        self.layout1 = FloatLayout()

        main_label = Label(
            text='SAP HELPER',
            font_size='35sp',
            color=(0, 0, 0, 1),  # color for text
            pos_hint={'center_x': 0.5, 'y': 0.85},
            size_hint=(0.1, 0.1)
        )

        ebs = kivy_customizing.KivyButton (
            layout=self.layout1,
            text = 'EBS MT940',
            x = 0.13,
            y = 0.7
        )
        ebs.button.bind(on_press=self.go_to_ebs)
        iban = kivy_customizing.KivyButton (
            layout=self.layout1,
            text = 'IBAN',
            x = 0.13,
            y = 0.55
        )
        iban.button.bind(on_press=self.go_to_iban)
        migration = kivy_customizing.KivyButton (
            layout=self.layout1,
            text = 'Migration Files',
            x = 0.13,
            y = 0.4
        )
        migration.button.bind(on_press=self.go_to_migration)

        # Set layout as content of main screen
        main_screen.add_widget(self.layout1)
        self.layout1.add_widget(main_label)


        # ebs screen
        ebs_screen = Screen(name='EBS MT940')
        self.layout2 = FloatLayout()

        self.ebs_main = DragAndDrop (sm = self.sm, layout=self.layout2)
        self.ebs_main.bind_dropfile()  # Bind the drag-and-drop event

        ebs_screen.add_widget(self.layout2)


        # iban screen
        iban_screen = Screen(name='IBAN')
        self.layout3 = FloatLayout()

        self.iban_main = DragAndDrop (sm = self.sm, layout=self.layout3)
        self.iban_main.bind_dropfile()  # Bind the drag-and-drop event

        iban_screen.add_widget(self.layout3)



        # migration screen
        migration_screen = Screen(name='MIGRATION')
        self.layout4 = FloatLayout()

        self.migration_main = Migration (sm = self.sm, layout=self.layout4)
        self.migration_main.bind_dropfile()  # Bind the drag-and-drop event

        migration_screen.add_widget(self.layout3)

        
        # Add screens to screen manager
        self.sm.add_widget(main_screen)
        self.sm.add_widget(ebs_screen)
        self.sm.add_widget(iban_screen)


        return self.sm

    def go_to_ebs(self, instance):
        # Switch to ebs screen
        self.sm.current = 'EBS MT940'
    
    def go_to_iban(self, instance):
        # Switch to iban screen
        self.sm.current = 'IBAN'
    
    def go_to_migration(self, instance):
        # Switch to migration screen
        self.sm.current = 'MIGRATION'
    

class DragAndDrop ():
    def __init__ (
            self,
            sm: ScreenManager,
            layout: FloatLayout
            ):
        self.sm = sm
        self.layout = layout
        self.file_path = ""

        self.go_back_button = kivy_customizing.KivyButton(
                layout = self.layout,
                size_hint = (None, None),
                background_normal = 'go_back.png',
                background_down = 'go_back.png',
                background_disabled_normal = 'go_back.png',
                x = 0.07,
                y = 0.9
            )
        self.go_back_button.button.bind(on_press=self.go_back)

        self.home = kivy_customizing.KivyButton(
                layout = self.layout,
                size_hint = (None, None),
                background_normal = 'home.png',
                background_down = 'home.png',
                background_disabled_normal = 'home.png',
                x = 0.12,
                y = 0.9
            )
        self.home.button.bind(on_press=self.go_to_home)

        self.reset = kivy_customizing.KivyButton(
                layout = self.layout,
                size_hint = (None, None),
                background_normal = 'reset.png',
                background_down = 'reset.png',
                background_disabled_normal = 'reset.png',
                x = 0.17,
                y = 0.9
            )
        self.reset.button.bind(on_press=self.reset_content)

        self.analysis = kivy_customizing.KivyButton(
                layout = self.layout,
                size_hint = (None, None),
                background_normal = 'analysis.png',
                background_down = 'analysis.png',
                background_disabled_normal = 'analysis.png',
                x = 0.22,
                y = 0.9
            )
        self.analysis.button.bind(on_press=self.start_analysis)

        # Text area to display file content
        self.file_content_input = TextInput(
            hint_text="",
            readonly=True,
            font_size=18,  # Change this value to adjust font size
            size_hint=(0.9, 0.8),  # Width: 90% of the screen, Height: 80% of the screen
            pos_hint={'center_x': 0.5, 'center_y': 0.45}
        )
        self.drop_label = Label(
            text="Drag and Drop the file here",
            font_size='50sp',
            color=(0, 0, 0, 0.5),  # color for text
            pos_hint={'center_x': 0.5, 'y': 0.5},
            size_hint=(0.1, 0.1)
        )

        # Add widgets to layout
        self.layout.add_widget(self.file_content_input)
        self.layout.add_widget(self.drop_label)

    def go_to_home(self, instance):
        if self.sm.current not in ['EBS MT940', 'IBAN']:
            self.sm.remove_widget(self.analysis_screen)
        # Switch to main screen
        self.sm.current = 'Main'
    
    def go_back(self, instance):
        # Switch to previous screen
        if self.sm.current in ['EBS MT940', 'IBAN']:
            self.sm.current = 'Main'
        elif self.sm.current == 'Analysis_ebs':
            self.sm.remove_widget(self.analysis_screen)
            self.sm.current = 'EBS MT940'
        else:
            self.sm.remove_widget(self.analysis_screen)
            self.sm.current = 'IBAN'

    def bind_dropfile(self):
        """Bind the drag-and-drop event to the current window."""
        Window.bind(on_dropfile=self.on_file_drop)

    def reset_content(self, instance):
        """Clears the content of the file content text area."""
        self.file_content_input.text = ""
        if not self.drop_label.parent:
            self.layout.add_widget(self.drop_label)
    
    def on_file_drop(self, window, file_path):
        self.layout.remove_widget(self.drop_label)

        # Decode the file path from bytes to string
        self.file_path = file_path.decode("utf-8")

        # Check if the file is a .txt file
        if self.file_path.endswith(".txt"):
            # Read and display the content of the file
            try:
                with open(file_path, 'r') as file:
                    self.file_content_input.text = file.read()
                    self.file_content_input.cursor = (0, 0)  # Set cursor to the start of the text
            except Exception as e:
                self.file_content_input.text = f"Error reading file: {e}"
        else:
            self.file_content_input.text = "Please drop a valid .txt file!"
    
    def start_analysis(self, instance):
        # analysis screen
        if self.sm.current == 'EBS MT940':
            self.analysis_screen = Screen(name='Analysis_ebs')
        else:
            self.analysis_screen = Screen(name='Analysis_iban')

        self.layout_analysis = FloatLayout()

        self.go_back_button = kivy_customizing.KivyButton(
                layout = self.layout_analysis,
                size_hint = (None, None),
                background_normal = 'go_back.png',
                background_down = 'go_back.png',
                background_disabled_normal = 'go_back.png',
                x = 0.07,
                y = 0.95
            )
        self.go_back_button.button.bind(on_press=self.go_back)

        self.home = kivy_customizing.KivyButton(
                layout = self.layout_analysis,
                size_hint = (None, None),
                background_normal = 'home.png',
                background_down = 'home.png',
                background_disabled_normal = 'home.png',
                x = 0.12,
                y = 0.95
            )
        self.home.button.bind(on_press=self.go_to_home)

        self.export = kivy_customizing.KivyButton(
                layout = self.layout_analysis,
                size_hint = (None, None),
                background_normal = 'excel.png',
                background_down = 'excel.png',
                background_disabled_normal = 'excel.png',
                x = 0.17,
                y = 0.95
            )
        

        # Create a ScrollView for the table content
        self.scroll_view = ScrollView(size_hint=(1, None), size=(800, 800), pos=(0, 60))

        if self.sm.current == 'EBS MT940':
            header_list = ['', '', '', '', '', '', '']
            
            self.export.button.bind(on_press = lambda instance: kivy_customizing.export_to_excel(layout = self.layout_analysis, table = self.table, file_path = self.file_path, file_name = "EBS MT940 export", ebs = "YES"))
            
            ebs_check = kivy_engine.EbsEngine (content = self.file_content_input.text)

            self.positions = [('SWIFT', 'BANK ACCOUNT N°', 'START DATE', 'END DATE', 'CURRENCY', 'OPENING BALANCE', 'CLOSING BALANCE'), 
                    (ebs_check.swift, ebs_check.bank_account_number, ebs_check.start_date, ebs_check.end_date, ebs_check.currency, ebs_check.opening_balance, ebs_check.closing_balance),
                    ('', '', '', '', '', '', ''),
                    ('VALUE DATE', 'AMOUNT', 'BANK EXTERNAL TRANSACTION', 'BANK EXT TR DESCRIPTION', '', '', '')]

            for a in ebs_check.position_lst:
                self.positions.append(a)

            self.positions.append (('', '', '', '', '', '', ''))
            self.positions.append(('OPENING BALANCE', ebs_check.opening_balance, '', '', '', '', ''))
            self.positions.append(('TOTAL CREDIT', "{:.2f}".format(ebs_check.total_credit), '', '', '', '', ''))
            self.positions.append(('TOTAL DEBIT', "{:.2f}".format(ebs_check.total_debit), '', '', '', '', ''))
            self.positions.append(('CLOSING BALANCE', ebs_check.closing_balance, '', '', '', '', ''))
            self.positions.append(('CHECK', "{:.2f}".format(float(ebs_check.opening_balance) + ebs_check.total_credit + ebs_check.total_debit - float(ebs_check.closing_balance)), '', '', '', '', ''))

            
            widths = [285, 285, 285, 285, 285, 285, 285]

        elif self.sm.current == 'IBAN':
            header_list = ['IBAN', 'BANK COUNTRY', 'BANK KEY', 'BANK ACCOUNT N°', 'BANK CONTROL KEY', 'SWIFT', 'NOTES']

            self.export.button.bind(on_press = lambda instance: kivy_customizing.export_to_excel(layout = self.layout_analysis, table = self.table, file_path = self.file_path, file_name = "IBAN"))

            iban_check = kivy_engine.IbanEngine (content = self.file_content_input.text)

            self.positions = iban_check.position_lst

            widths = [350, 200, 200, 200, 200, 150, 500]


        self.table = kivy_customizing.KivyTable (
            scroll_view = self.scroll_view,
            layout = self.layout_analysis,
            headers = header_list,
            rows = self.positions, # it should be a list of lists
            widths = widths # the sum of the components should be 1, It identify the relative position of the columns
        )

        self.layout_analysis.add_widget(self.scroll_view)
        self.analysis_screen.add_widget(self.layout_analysis)

        # Add screens to screen manager
        self.sm.add_widget(self.analysis_screen)

        if self.sm.current == 'EBS MT940':
            self.sm.current = 'Analysis_ebs'
        else:
            self.sm.current = 'Analysis_iban'

class Migration ():
    def __init__ (
            self,
            sm: ScreenManager,
            layout: FloatLayout
            ):
        self.sm = sm
        self.layout = layout
        self.file_path = ""
    


    def bind_dropfile(self):
        """Bind the drag-and-drop event to the current window."""
        Window.bind(on_dropfile=self.on_file_drop)

    
    def on_file_drop(self, window, file_path):
        self.layout.remove_widget(self.drop_label)

        # Decode the file path from bytes to string
        self.file_path = file_path.decode("utf-8")

        # Check if the file is a .txt file
        if self.file_path.endswith(".txt"):
            # Read and display the content of the file
            try:
                with open(file_path, 'r') as file:
                    self.file_content_input.text = file.read()
                    self.file_content_input.cursor = (0, 0)  # Set cursor to the start of the text
            except Exception as e:
                self.file_content_input.text = f"Error reading file: {e}"
        else:
            self.file_content_input.text = "Please drop a valid .txt file!"
        

        
        
        
        
        


if __name__ == '__main__':
    MyApp().run()