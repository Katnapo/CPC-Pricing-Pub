import tkinter as tk
from tkinter import ttk
import re
import constants
from constants import Constants
from tkinter import filedialog
from tkcalendar import DateEntry
import datetime
class PricingApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.bottom_content_exists = 0
        self.title("Excel Sheet Generator For BC Price Uploads")
        self.geometry("600x400")
        self.style = ttk.Style(self)
        self.style.theme_use("alt")

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True)

        self.output_filename = None
        self.filename = None

        # Create the Simple Settings tab
        self.simple_settings_frame = ttk.Frame(self.tabs)
        self.tabs.add(self.simple_settings_frame, text="Simple Settings")
        self.create_simple_settings()

        # # Create the Advanced Settings tab
        self.advanced_settings_frame = ttk.Frame(self.tabs)
        self.advanced_settings_frame.pack(expand=True)
        self.tabs.add(self.advanced_settings_frame, text="Advanced Settings")
        self.create_advanced_settings()

        self.exchange_rate_settings = ttk.Frame(self.tabs)
        self.tabs.add(self.exchange_rate_settings, text="Exchange Rate Settings")
        self.create_exchange_rate_settings()

        self.validation_settings = ttk.Frame(self.tabs)
        self.tabs.add(self.validation_settings, text="Validation")
        self.create_validation_settings()

        # Add Upload and Submit buttons outside the tabs
        self.create_bottom_content()
        self.set_config_values()

    def create_bottom_content(self):

            # Add Upload and Submit buttons outside the tabs
        if self.bottom_content_exists == 0:

            self.progress_bar_label = ttk.Label(self, text="")
            self.progress_bar_label.pack(side="top", fill="x")

            self.progress_bar = ttk.Progressbar(self, orient="horizontal", length=200, mode="determinate")
            self.progress_bar.pack(side="top", fill="x")

            self.upload_button = ttk.Button(self, text="Upload Input", command=self.upload_action_input)
            self.upload_button.pack(side="left", ipadx=10)

            self.submit_button = ttk.Button(self, text="Submit", command=self.submit_action)
            self.submit_button.pack(side="right", ipadx=10)

            self.information_text = ttk.Label(self, text="No file selected")
            self.information_text.pack(side="bottom", fill="both")

            self.bottom_content_exists = 1

    def set_progressbar_label(self, text):
        self.progress_bar_label.config(text=text)
    def progressbar_callback(self, value, text, reset=False):

        if reset:
            self.progress_bar["value"] = 0
            self.progress_bar_label.config(text="")
            return None

        self.set_progressbar_label(text)

        if self.progress_bar["value"] > 100:
            return None

        self.progress_bar["value"] +=value
        self.update_idletasks()
    def create_simple_settings(self):
        # Create the Starting Date field
        self.create_bottom_content()

        self.starting_date_calendar = DateEntry(self.simple_settings_frame, selectmode="day", state="disabled")
        self.starting_date_calendar.grid(row=0, column=2)
        # Set starting date format to dd/mm/yyyy
        self.starting_date_calendar.config(date_pattern="dd/mm/yyyy")

        self.starting_date_label = ttk.Label(self.simple_settings_frame, text="Starting Date")
        self.starting_date_label.grid(row=0, column=0, padx=(5))

        self.starting_date_options = Constants.startDateConstants
        self.starting_date_dropdown = ttk.Combobox(self.simple_settings_frame, values=self.starting_date_options)
        self.starting_date_dropdown.grid(row=0, column=1, pady=(5))
        self.starting_date_dropdown.bind("<<ComboboxSelected>>", self.starting_date_dropdown_action)

        # Create the Ending Date field
        self.ending_date_label = ttk.Label(self.simple_settings_frame, text="Ending Date")
        self.ending_date_label.grid(row=1, column=0, padx=(5))

        self.ending_date_checkbox_var = tk.IntVar()
        self.ending_date_checkbox = ttk.Checkbutton(self.simple_settings_frame, variable=self.ending_date_checkbox_var, command=self.toggle_ending_date)
        self.ending_date_checkbox.grid(row=1, column=1)

        self.ending_date_calendar = DateEntry(self.simple_settings_frame, selectmode="day", state="disabled")
        self.ending_date_calendar.grid(row=1, column=2, pady=(5))
        # Set starting date format to dd/mm/yyyy
        self.ending_date_calendar.config(date_pattern="dd/mm/yyyy")

        # Create the Custom Price Type field
        self.custom_price_label = ttk.Label(self.simple_settings_frame, text="Custom Price Type")
        self.custom_price_label.grid(row=2, column=0, pady=(5), padx=(5))

        self.custom_price_options = Constants.possiblePriceTypes
        self.custom_price_dropdown = ttk.Combobox(self.simple_settings_frame, values=self.custom_price_options)
        self.custom_price_dropdown.grid(row=2, column=1)

        self.api_use_tick_var = tk.IntVar()
        self.api_use_label = ttk.Label(self.simple_settings_frame, text="Use BC API Prices")
        self.api_use_label.grid(row=4, column=0, pady=(5))
        self.api_use_tick =ttk.Checkbutton(self.simple_settings_frame, variable=self.api_use_tick_var, command=self.toggle_use_api)
        self.api_use_tick.grid(row=4, column=1)

        self.close_off_tick_var = tk.IntVar()
        self.close_off_label = ttk.Label(self.simple_settings_frame, text="Close Off Prices")
        self.close_off_label.grid(row=5, column=0, pady=(2))
        self.close_off_tick = ttk.Checkbutton(self.simple_settings_frame, variable=self.close_off_tick_var, state="disabled")
        self.close_off_tick.grid(row=5, column=1)

    def starting_date_dropdown_action(self, event):

        import datetime

        selected_option = self.starting_date_dropdown.get()
        self.starting_date_calendar.config(state="enabled")

        if selected_option == "Use Custom Date":
            self.starting_date_calendar.drop_down()
            constants.Constants.startDateCustomization = 2

        elif selected_option == "Use Today's Date":
            self.starting_date_calendar.set_date(datetime.date.today())
            self.starting_date_calendar.config(state="disabled")
            constants.Constants.startDateCustomization = 1

        elif selected_option == "Use Date in Input Sheet":
            self.starting_date_calendar.delete(0, tk.END)
            self.starting_date_calendar.selection_clear()
            self.starting_date_calendar.config(state="disabled")
            constants.Constants.startDateCustomization = 3

    def toggle_ending_date(self):

        if self.ending_date_checkbox_var.get() == 1:
            self.ending_date_calendar.config(state="enabled")
            constants.Constants.customEndDateBool= True
        else:
            self.ending_date_calendar.delete(0, tk.END)
            self.ending_date_calendar.selection_clear()
            constants.Constants.customEndDateBool = False
            self.ending_date_calendar.config(state="disabled")

    def toggle_use_api(self):

        if self.api_use_tick_var.get() == 1:
            self.close_off_tick.config(state="enabled")
            self.input_column_label_price.config(text="Multipliter Column Name")

        else:
            self.close_off_tick_var.set(0)
            self.close_off_tick.config(state="disabled")
            self.input_column_label_price.config(text="Price Column Name")
    def create_advanced_settings(self):

        # Create the Input Column Names list editor
        self.input_column_label_identifier = ttk.Label(self.advanced_settings_frame, text="Identifier Column Name")
        self.input_column_label_identifier.grid(row=0, column=0)
        self.input_column_textbox_identifier = ttk.Entry(self.advanced_settings_frame, width=20)
        self.input_column_textbox_identifier.grid(row=0, column=1)

        self.input_column_label_price = ttk.Label(self.advanced_settings_frame, text="Price Column Name")
        self.input_column_label_price.grid(row=1, column=0)
        self.input_column_textbox_price = ttk.Entry(self.advanced_settings_frame, width=20)
        self.input_column_textbox_price.grid(row=1, column=1)

        self.input_column_label_date = ttk.Label(self.advanced_settings_frame, text="Date Column Name")
        self.input_column_label_date.grid(row=2, column=0)
        self.input_column_textbox_date = ttk.Entry(self.advanced_settings_frame, width=20)
        self.input_column_textbox_date.grid(row=2, column=1)

        self.by_search_radius_label = ttk.Label(self.advanced_settings_frame, text="Column Search Radius")
        self.by_search_radius_label.grid(row=3, column=0)
        self.by_search_radius_textbox = ttk.Entry(self.advanced_settings_frame)
        self.by_search_radius_textbox.grid(row=3, column=1)

        # Create the None Tolerance section

        self.none_tolerance_label = ttk.Label(self.advanced_settings_frame, text="None Tolerance")
        self.none_tolerance_label.grid(row=4, column=0, pady=(10))

        self.none_tolerance_var = tk.IntVar()
        self.none_tolerance_spinbox = ttk.Spinbox(self.advanced_settings_frame, from_=0, to=100, width=5, textvariable=self.none_tolerance_var)
        self.none_tolerance_spinbox.grid(row=4, column=1)

        self.none_tolerance_increment_button = tk.Button(self.advanced_settings_frame, text="▲", command=self.increment_value)
        self.none_tolerance_decrement_button = tk.Button(self.advanced_settings_frame, text="▼", command=self.decrement_value)


        # Create the Output Suffix, Identifier Column, and Price Column fields
        self.output_suffix_label = ttk.Label(self.advanced_settings_frame, text="Output Suffix")
        self.output_suffix_label.grid(row=5, column=0)

        self.output_suffix_textbox = ttk.Entry(self.advanced_settings_frame, width=20)
        self.output_suffix_textbox.grid(row=5, column=1)

        self.customer_number_prefix_label = ttk.Label(self.advanced_settings_frame, text="Customer Number Prefix")
        self.customer_number_prefix_label.grid(row=6, column=0)

        self.customer_number_prefix_textbox = ttk.Entry(self.advanced_settings_frame, width=20)
        self.customer_number_prefix_textbox.grid(row=6, column=1)

    def increment_value(self):
        self.none_tolerance_var.set(self.none_tolerance_var.get() + 1)

    def decrement_value(self):
        self.none_tolerance_var.set(self.none_tolerance_var.get() - 1)

    def create_exchange_rate_settings(self):

        self.exchange_table = ttk.Treeview(self.exchange_rate_settings, columns=("c1", "c2"), show='headings')
        self.exchange_table.column("# 1", anchor=tk.CENTER, width=100)
        self.exchange_table.heading("# 1", text="Currency")
        self.exchange_table.column("# 2", anchor=tk.CENTER, width=100)
        self.exchange_table.heading("# 2", text="Exchange Rate")

        self.exchange_table.grid(row=0, column=0)

        self.exchange_table_add_button = ttk.Button(self.exchange_rate_settings, text="Add", command=self.add_exchange_rate)
        self.exchange_table_add_button.grid(row=0, column=1, sticky="n", padx=10, pady=10)

        self.exchange_table_add_label_currency = ttk.Label(self.exchange_rate_settings, text="Currency:")
        self.exchange_table_add_label_currency.grid(row=0, column=2, sticky="n", padx=10, pady=10)
        self.exchange_table_add_textbox_currency = ttk.Entry(self.exchange_rate_settings, width=20)
        self.exchange_table_add_textbox_currency.grid(row=0, column=3, sticky="n", padx=10, pady=10)

        self.exchange_table_add_label_exchange_rate = ttk.Label(self.exchange_rate_settings, text="Exchange Rate:")
        self.exchange_table_add_label_exchange_rate.grid(row=0, column=2, sticky="n", padx=10, pady=40)
        self.exchange_table_add_textbox_exchange_rate = ttk.Entry(self.exchange_rate_settings, width=20)
        self.exchange_table_add_textbox_exchange_rate.grid(row=0, column=3, sticky="n", padx=10, pady=40)

        self.exchange_table_delete_button = ttk.Button(self.exchange_rate_settings, text="Delete", command=self.delete_exchange_rate)
        self.exchange_table_delete_button.grid(row=0, column=1, sticky="n", padx=10, pady=60)
        pass

    def add_exchange_rate(self):

        self.exchange_table.insert("", tk.END, values=(self.exchange_table_add_textbox_currency.get(),
                                                       self.exchange_table_add_textbox_exchange_rate.get()))
        pass

    def delete_exchange_rate(self):

        selected_item = self.exchange_table.selection()[0]
        self.exchange_table.delete(selected_item)
        pass

    def create_validation_settings(self):

        self.validation_settings_output_upload_label = ttk.Label(self.validation_settings, text="Upload output file for validation:")
        self.validation_settings_output_upload_label.grid(row=0, column=0, sticky="n", padx=10, pady=10)
        self.validation_settings_output_upload_button = ttk.Button(self.validation_settings, text="Upload", command=self.upload_action_output)
        self.validation_settings_output_upload_button.grid(row=0, column=1, sticky="n", padx=10, pady=10)
        self.validation_settings_output_upload_file_label = ttk.Label(self.validation_settings, text="No file selected")
        self.validation_settings_output_upload_file_label.grid(row=0, column=2, sticky="n", padx=10, pady=10)

        self.validation_settings_test_coverage_label = ttk.Label(self.validation_settings, text="Test Coverage:")
        self.validation_settings_test_coverage_label.grid(row=1, column=0, sticky="n", padx=10, pady=10)
        self.validation_settings_test_coverage_textbox = ttk.Entry(self.validation_settings, width=20)
        self.validation_settings_test_coverage_textbox.grid(row=1, column=1, sticky="n", padx=10, pady=10)

        self.validation_random_conversions_label = ttk.Label(self.validation_settings, text="Validate Random Conversions")
        self.validation_random_conversions_label.grid(row=2, column=0, sticky="n", padx=10, pady=10)
        self.validation_random_conversions_button = ttk.Button(self.validation_settings, text="Run", command=self.run_validation_tests_price_conversion)
        self.validation_random_conversions_button.grid(row=2, column=1, sticky="n", padx=10, pady=10)
        self.validation_random_conversions_help_button = ttk.Button(self.validation_settings, text="Help", command=self.validate_random_conversions_help)
        self.validation_random_conversions_help_button.grid(row=2, column=2, sticky="n", padx=10, pady=10)

        self.validation_input_output_size_match_label = ttk.Label(self.validation_settings, text="Validate Input and Output Size Match")
        self.validation_input_output_size_match_label.grid(row=3, column=0, sticky="n", padx=10, pady=10)
        self.validation_input_output_size_match_button = ttk.Button(self.validation_settings, text="Run", command=self.run_validation_tests_input_output_size_match)
        self.validation_input_output_size_match_button.grid(row=3, column=1, sticky="n", padx=10, pady=10)
        self.validation_input_output_size_match_help_button = ttk.Button(self.validation_settings, text="Help", command=self.validate_size_match_help)
        self.validation_input_output_size_match_help_button.grid(row=3, column=2, sticky="n", padx=10, pady=10)

        pass

    def run_validation_tests_price_conversion(self):

        if self.output_filename:

            print("Running validation tests")

            import validation_tests
            testPriceCalc = validation_tests.TestPriceCalc()

            try:
                response = testPriceCalc.validateRandomConversionCalcs(self.progressbar_callback)

                if response:
                    self.information_text.config(text="Random Conversion calc test was a success.")
                else:
                    self.information_text.config(text="Random Conversion calc test failed. See error.csv for failed lines.")

                self.progressbar_callback(0, "", reset=True)

            except Exception as e:
                self.information_text.config(text="Random Conversion calc test failed due to error. Error is: " + str(e))
                self.progressbar_callback(0, "", reset=True)

    def run_validation_tests_input_output_size_match(self):

        if self.output_filename:

            print("Running validation tests")

            import validation_tests
            testPriceCalc = validation_tests.TestPriceCalc()

            try:
                response = testPriceCalc.checkInputAndOutputSizeMatch(self.progressbar_callback)

                if response:
                    self.information_text.config(text="Input and Output size match test was a success.")
                else:
                    self.information_text.config(text="Input and Output size match test failed. See error.txt for failed lines.")

            except Exception as e:

                self.information_text.config(text="Input and Output size match test failed due to error. Error is: " + str(e))
                self.progressbar_callback(0, "", reset=True)
    def validate_random_conversions_help(self):

        top = tk.Toplevel()
        top.geometry("750x250")
        top.title("What is Random Conversion Calcs?")
        ttk.Label(top, text=Constants.randomConversionCalcHelp, font=('Mistral 12 bold')).pack(side="top", fill="both", expand=True)

    def validate_size_match_help(self):

        top = tk.Toplevel()
        top.geometry("750x250")
        top.title("What is Input and Output Size Match?")
        ttk.Label(top, text=Constants.inputOutputSizeMatchHelp, font=('Mistral 12 bold')).pack(side="top", fill="both", expand=True)

    def set_config_values(self):

        import json

        self.config_file = open("config.json", "r")
        self.config_data = json.load(self.config_file)
        self.config_file.close()

        for currency in self.config_data["exchange rates and countries"]:
            self.exchange_table.insert("", tk.END, values=(currency["country"], currency["exchange_rate"]))

        self.input_column_textbox_identifier.insert(0, self.config_data["identifier column"])
        self.input_column_textbox_price.insert(0, self.config_data["price column"])
        self.input_column_textbox_date.insert(0, self.config_data["date column"])

        self.by_search_radius_textbox.insert(0, self.config_data["column search radius"])
        self.none_tolerance_var.set(self.config_data["none tolerance"])
        self.output_suffix_textbox.insert(0, self.config_data["output suffix"])
        self.customer_number_prefix_textbox.insert(0, self.config_data["customer number prefix"])

        self.validation_settings_test_coverage_textbox.insert(0, self.config_data["test coverage"])

    def get_todays_date(self):
        import datetime
        return datetime.datetime.today().strftime('%d/%m/%Y')

    def upload_action_input(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File for Import")

        if filename:

            self.information_text.config(text="File selected: " + filename)
            self.filename = filename
            constants.Constants.originalBookLocation = filename
            constants.Constants.originalBookName = filename.split("/")[-1]

    def upload_action_output(self):

        filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File for Import")

        if filename:

            self.output_filename = filename
            self.validation_settings_output_upload_file_label.config(text="File selected: " + filename)
            constants.Constants.outputBookLocation = filename
            constants.Constants.outputBookName = filename.split("/")[-1]

            self.create_output_file()

    def create_output_file(self):

        import os, shutil

        current_dir = os.getcwd()
        destination_dir = os.path.join(current_dir, constants.Constants.outputBookName)

        try:
            shutil.copyfile(self.output_filename, destination_dir)
        except shutil.SameFileError:
            pass

    def setBasicConstants(self):

        #TODO: Move price type to here rather than having its constant set in the simple settings tab

        constants.Constants.priceType = self.custom_price_dropdown.get()

        if constants.Constants.priceType is None or constants.Constants.priceType == "":
            constants.Constants.priceType = constants.Constants.possiblePriceTypes[0]

        try:
            constants.Constants.customStartDate = datetime.datetime.strftime(self.starting_date_calendar.get_date(),
                                                                             "%d/%m/%Y")
        except Exception as e:

            constants.Constants.customStartDate = None

        try:
            constants.Constants.customEndDate = datetime.datetime.strftime(self.ending_date_calendar.get_date(),
                                                                           "%d/%m/%Y")
        except Exception as e:

            # This error occurs when the user does not select an end date and therefore the calendar is empty,
            # meaning TK cannot convert it to a datetime object. This is fine as we can just set the end date to
            # None and the program will know to use the current date as the end date.

            constants.Constants.customEndDate = None

        constants.Constants.useAPI = bool(self.api_use_tick_var.get())
        constants.Constants.closeOff = bool(self.close_off_tick_var.get())

    def setAdvancedConstants(self):

        identifier = self.input_column_textbox_identifier.get()
        if identifier:
            constants.Constants.inputIdentifierColName = identifier

        priceColumn = self.input_column_textbox_price.get()
        if priceColumn:
            constants.Constants.inputPriceColName = priceColumn

        dateColumn = self.input_column_textbox_date.get()
        if dateColumn:
            constants.Constants.inputStartDateColName = dateColumn


        # If API usage is set to 1, but we are generating markdowns or promotions and opening prices
        # we need to use the price column as a container for the multiplier column. Any other scenario using the
        # API will use just the identifier column - this logic is captured below.

        if self.api_use_tick_var.get() == 1 and (self.custom_price_dropdown.get() == "Full Price" or self.close_off_tick_var.get() == 1):
            constants.Constants.inputColumnIdentifiers = [constants.Constants.inputIdentifierColName]
        else:
            constants.Constants.inputColumnIdentifiers = [constants.Constants.inputIdentifierColName,
                                                              constants.Constants.inputPriceColName]

        if constants.Constants.startDateCustomization == 3:
            constants.Constants.inputColumnIdentifiers.append(constants.Constants.inputStartDateColName)

        try:
            searchRadius = int(self.by_search_radius_textbox.get())
            constants.Constants.searchRadius = searchRadius

        except:
            self.information_text.config(text="Please enter a valid search radius")
            return False

        constants.Constants.noneCountTolerance = self.none_tolerance_var.get()
        constants.Constants.outputBookSuffix = self.output_suffix_textbox.get()
        constants.Constants.customerNumberPrefix = self.customer_number_prefix_textbox.get()

        return True

    def setExchangeRateConstants(self):

        # Goes through exchange rate table and sets the exchange rate constants

        self.exchange_table_data = self.exchange_table.get_children()
        tempArray= []
        for row in self.exchange_table_data:

            currency = self.exchange_table.item(row)["values"]

            if currency[0] == "" or currency[1] == "":
                continue

            try:
                currency[1] = float(currency[1])

            except:
                self.information_text.config(text="Please enter a valid exchange rate")
                return False

            tempArray.append(currency)

        constants.Constants.countryCodeExchange = tempArray

        return True
    def submit_action(self):

        import os
        import shutil
        if self.filename:

            current_dir = os.getcwd()
            destination_dir = os.path.join(current_dir, constants.Constants.inputBookName)

            try:
                shutil.copyfile(self.filename, destination_dir)
            except shutil.SameFileError:
                pass

            self.setBasicConstants()
            if not self.setAdvancedConstants():
                return

            if not self.setExchangeRateConstants():
                return

            from classes import Controller
            controller = Controller()
            response = controller.run(self.progressbar_callback)

            if response is None:
                self.information_text.config(text="File created successfully. Please look for " +
                                                  constants.Constants.outputBookName + " in the same directory as this program.")
            else:
                self.information_text.config(text="Error: " + str(response))
            self.progress_bar["value"] = 0

        else:
            self.information_text.config(text="Please select a file to upload")
            self.progress_bar["value"] = 0