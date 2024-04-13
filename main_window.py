import os
import os.path
from idlelib.tooltip import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import sys
import filename_methods as fm
import misc_methods as mm
from constants import *
from session_log import SessionLog
from session_process_xlsx import SessionProcessXLSX
from user_entry import UserEntry


# Class to create GUI
class MainWindow:
    # Dependencies: MainWindow communicate with classes that are related to GUI contents and buttons
    def __init__(self):  # instantiation function. Use root for GUI and refers to main window

        root = Tk()
        root.title("BOM Merger")
        self.root_frame = root
        self.user_entry = UserEntry()
        self.session_log = SessionLog(self.user_entry)
        self.session_process_xlsx = None

        self.rows_to_peak = [0] * self.user_entry.n_rows_to_peak

        # GUI Frames
        self.frame_root_title = Frame(root, highlightthickness=0)
        self.frame_root_session = LabelFrame(root, width=200, height=390, padx=5, pady=5, text="Session")
        self.frame_root_commands = LabelFrame(root, width=200, height=80, padx=5, pady=5, text="")

        # Disable resizing the window
        # root.resizable(False, False)

        # Grids
        self.frame_root_title.grid(row=0, column=0, padx=10, pady=5, ipadx=5, ipady=5)
        self.frame_root_session.grid(row=1, column=0, sticky="W", padx=10, pady=(5, 5), ipadx=5, ipady=2)
        self.frame_root_commands.grid(row=2, column=0, sticky="W", padx=10, pady=(10, 5), ipadx=5, ipady=2)
        # self.frame_root_session.grid_propagate(False)  # Prevents automatic resizing of frame
        # self.frame_root_commands.grid_propagate(False)  # Prevents automatic resizing of frame

        entry_validation_positive_numbers = root.register(mm.only_positive_numbers)
        entry_validation_numbers = root.register(mm.only_digits)
        entry_validation_numbers_space = root.register(mm.digits_or_space)
        entry_validation_positive_numbers_comma = root.register(mm.positive_numbers_or_comma)

        ######################################################################
        # Frame Session
        # Labels
        label_file_name = Label(self.frame_root_session, text="File Name")
        label_rows_to_peak = Label(self.frame_root_session, text="Rows to Peak")
        label_rows = Label(self.frame_root_session, text="Rows")
        label_component_index = Label(self.frame_root_session, text="Component Index")
        label_quantity_index = Label(self.frame_root_session, text="Quantity Index")

        # Entries
        self.entry_file_name_entry = StringVar()
        self.entry_file_name_entry.trace("w", lambda name, index, mode, entry_file_name_entry=self.entry_file_name_entry: self.entry_update_file_name_and_suffix())
        self.entry_file_name = Entry(self.frame_root_session, width=81, textvariable=self.entry_file_name_entry)
        self.entry_file_name.insert(END, (FILE_NAME if FILE_NAME else FILE_NAME_INIT))

        self.entry_file_location_entry = StringVar()
        self.entry_file_location_entry.trace("w", lambda name, index, mode, entry_file_location_entry=self.entry_file_location_entry: self.entry_update_file_location())
        self.entry_file_location = Entry(self.frame_root_session, width=80, textvariable=self.entry_file_location_entry)
        self.entry_file_location.insert(END, os.path.normcase(FILE_LOCATION))

        self.entry_rows_to_peak_entry = IntVar()
        self.entry_rows_to_peak_entry.trace("w", lambda name, index, mode, entry_rows_to_peak_entry=self.entry_rows_to_peak_entry: self.entry_update_rows_to_peak())
        self.entry_rows_to_peak = Entry(self.frame_root_session, width=6, textvariable=self.entry_rows_to_peak_entry, validate="key", validatecommand=(entry_validation_positive_numbers, '%P'))
        self.entry_rows_to_peak.insert(END, str(N_ROWS_TO_PEAK_DEFAULT))

        self.entry_component_index_entry = StringVar()
        self.entry_component_index_entry.trace("w", lambda name, index, mode, entry_component_index_entry=self.entry_component_index_entry: self.entry_update_component_index())
        self.entry_component_index = Entry(self.frame_root_session, width=10, textvariable=self.entry_component_index_entry, validate="key", validatecommand=(entry_validation_positive_numbers, '%P'))
        self.entry_component_index.insert(END, str(COMPONENT_INDEX_DEFAULT))

        self.entry_quantity_index_entry = StringVar()
        self.entry_quantity_index_entry.trace("w", lambda name, index, mode, entry_quantity_index_entry=self.entry_quantity_index_entry: self.entry_update_quantity_index())
        self.entry_quantity_index = Entry(self.frame_root_session, width=10, textvariable=self.entry_quantity_index_entry, validate="key", validatecommand=(entry_validation_positive_numbers, '%P'))
        self.entry_quantity_index.insert(END, str(QUANTITY_INDEX_DEFAULT))

        # Textbox
        self.textbox_rows = Text(self.frame_root_session, height=10, width=50)
        self.textbox_rows_vscroll_bar = Scrollbar(self.frame_root_session, orient="vertical")
        self.textbox_rows_vscroll_bar.config(command=self.textbox_rows.yview)
        self.textbox_rows.config(yscrollcommand=self.textbox_rows_vscroll_bar.set)
        self.textbox_row_clear()

        # Buttons
        self.button_choose_single_file = Button(self.frame_root_session, text="Choose File", command=self.choose_file, pady=0, width=10, fg='brown')
        self.button_file_peak = Button(self.frame_root_session, text="File Peak", command=self.button_file_peak, pady=0, width=10)

        # Grids
        label_file_name.grid(row=5, column=0, sticky=E)
        self.entry_file_name.grid(row=5, column=1, sticky=W)
        self.button_file_peak.grid(row=7, column=0, sticky=NE)
        self.show_button_choose_single_file()
        self.enable_button_file_peak()
        self.entry_file_location.grid(row=6, column=1, sticky=W, padx=(5, 0))
        label_rows_to_peak.grid(row=6, column=1, sticky=W, padx=(500, 0))
        self.entry_rows_to_peak.grid(row=6, column=1, sticky=W, padx=(580, 0))
        label_rows.grid(row=7, column=0, sticky=E)
        self.textbox_rows.grid(row=7, column=1, sticky=W, padx=(20, 0))
        self.textbox_rows_vscroll_bar.grid(row=7, column=1, sticky=W)
        label_component_index.grid(row=8, column=0, sticky=W)
        self.entry_component_index.grid(row=8, column=1, sticky=W)
        label_quantity_index.grid(row=9, column=0, sticky=W)
        self.entry_quantity_index.grid(row=9, column=1, sticky=W)
        # END OF FRAME #######################################################

        ######################################################################
        # Frame Commands
        self.button_open_folder = Button(self.frame_root_commands, text="Open Folder", command=lambda: self.open_folder(self.user_entry.file_location), pady=3, width=20)
        self.button_process_file = Button(self.frame_root_commands, text="Process File", pady=3, width=20, fg="green", command=self.process_file)
        self.button_exit = Button(self.frame_root_commands, text="Exit", fg='red', command=self.quit_program, pady=3, width=20)

        # Grids
        self.button_open_folder.grid(row=1, column=1, padx=(2, 0))
        self.button_process_file.grid(row=1, column=3)
        self.button_exit.grid(row=1, column=4)
        # END OF FRAME #######################################################

        self.root_frame.mainloop()

    ######################################################################

    def set_state(self, widget, state):
        print(type(widget))
        try:
            widget.configure(state=state)
        except:
            pass
        for child in widget.winfo_children():
            self.set_state(child, state=state)

    def gui_entry_unlock(self):
        self.set_state(self.frame_root_session, state='normal')
        self.set_state(self.frame_root_commands, state='normal')

    # Disable user entries in offline mode
    def gui_entry_lock(self):
        self.set_state(self.frame_root_session, state='disabled')
        self.set_state(self.frame_root_commands, state='disabled')

    def entry_update_component_index(self):
        try:
            entry_component_index = self.entry_component_index_entry.get()
            self.user_entry.component_index = int(entry_component_index)
            print("::user_entry.component_index: ", self.user_entry.component_index)
        except:
            self.user_entry.component_index = COMPONENT_INDEX_DEFAULT
            print("::user_entry.component_index: ", self.user_entry.component_index)

    def entry_update_quantity_index(self):
        try:
            entry_quantity_index = self.entry_quantity_index_entry.get()
            self.user_entry.quantity_index = int(entry_quantity_index)
            print("::user_entry.quantity_index: ", self.user_entry.quantity_index)
        except:
            self.user_entry.quantity_index = QUANTITY_INDEX_DEFAULT
            print("::user_entry.quantity_index: ", self.user_entry.quantity_index)

    def entry_update_rows_to_peak(self):
        try:
            entry_rows_to_peak = self.entry_rows_to_peak_entry.get()
            self.user_entry.n_rows_to_peak = int(entry_rows_to_peak)
            print("::user_entry.n_rows_to_peak: ", self.user_entry.n_rows_to_peak)
        except:
            self.user_entry.n_rows_to_peak = N_ROWS_TO_PEAK_DEFAULT
            print("::user_entry.erows_to_peak: ", self.user_entry.n_rows_to_peak)
        self.rows_to_peak = [0] * self.user_entry.n_rows_to_peak

    def entry_update_file_location(self):
        file_location = self.entry_file_location_entry.get()
        if fm.FileNameMethods.check_file_location_valid(file_location):
            self.user_entry.file_location = file_location
        else:
            self.user_entry.file_location = FILE_LOCATION
        print("::user_entry.file_location: ", self.user_entry.file_location)

    def entry_update_file_name_and_suffix(self):
        file_name_suffix = self.entry_file_name_entry.get()
        file_name = os.path.splitext(file_name_suffix)[0]
        file_suffix = os.path.splitext(file_name_suffix)[1]

        if fm.FileNameMethods.check_filename_components_exists(self.user_entry.file_location, file_name, file_suffix):
            self.user_entry.file_name = file_name
            self.user_entry.file_suffix = file_suffix
        else:
            self.user_entry.file_name = FILE_NAME
            self.user_entry.file_suffix = FILE_SUFFIX
        print("::user_entry.file_name: ", self.user_entry.file_name)
        print("::user_entry.file_suffix: ", self.user_entry.file_suffix)

    @staticmethod
    def quit_program():
        sys.exit()

    def peak_file(self):
        try:
            if fm.FileNameMethods.check_filename_components_exists(self.user_entry.file_location, self.user_entry.file_name, self.user_entry.file_suffix) is not True:
                message = 'Invalid file\n'
                message_colour = 'red'
                self.session_log.write_textbox(message, message_colour)
                messagebox.showerror('Error', message)
                return
        except Exception as e:
            e = 'Invalid file\n'
            message_colour = 'red'
            self.session_log.write_textbox(e, message_colour)
            messagebox.showerror('Error', e)
            return

        message = 'Peak File: ' + fm.FileNameMethods.build_file_name_full(self.user_entry.file_location, self.user_entry.file_name, self.user_entry.file_suffix) + '\n'
        message_colour = 'brown'
        self.session_log.write_textbox(message, message_colour)
        self.session_process_xlsx = SessionProcessXLSX(self.user_entry, self.session_log, self.textbox_rows)

    def process_file(self):
        # Check File exists
        try:
            if fm.FileNameMethods.check_filename_components_exists(self.user_entry.file_location, self.user_entry.file_name, self.user_entry.file_suffix) is not True:
                message = 'Invalid file\n'
                message_colour = 'red'
                self.session_log.write_textbox(message, message_colour)
                messagebox.showerror('Error', message)
                return
        except Exception as e:
            e = 'Invalid file\n'
            message_colour = 'red'
            self.session_log.write_textbox(e, message_colour)
            messagebox.showerror('Error', e)
            return
        message = 'Start Session\n'
        message_colour = 'brown'
        self.session_log.write_textbox(message, message_colour)
        self.display_session_settings()
        # self.gui_entry_lock()
        self.session_process_xlsx = SessionProcessXLSX(self.user_entry, self.session_log, self.textbox_rows)
        # self.gui_entry_unlock()
        message = 'Finish Session\n'
        message_colour = 'brown'
        self.session_log.write_textbox(message, message_colour)

    ######################################################################

    ######################################################################
    def hide_button_choose_single_file(self):
        self.button_choose_single_file.grid_forget()

    def show_button_choose_single_file(self):
        self.button_choose_single_file.grid(row=6, column=0, sticky=E)

    def enable_button_file_peak(self):
        self.button_file_peak["state"] = ACTIVE

    def disable_button_file_peak(self):
        self.button_file_peak["state"] = DISABLED

    # Methods for data files
    def choose_file(self):
        # Clear entry fields before selecting file, else the entry will be cleared after file is selected
        # and the user_entry fields will be cleared as well
        file_location_current = self.user_entry.file_location
        self.entry_file_name.delete(0, END)
        self.entry_file_location.delete(0, END)

        file_type = [('xlsx', '*.xlsx'), ('xls', '*.xls')]
        file_full_name = filedialog.askopenfilename(initialdir=file_location_current, title="Select File", filetypes=file_type, defaultextension=file_type)
        self.user_entry.file_name = os.path.splitext(os.path.basename(file_full_name))[0]
        self.user_entry.file_suffix = os.path.splitext(os.path.basename(file_full_name))[1]
        self.user_entry.file_location = os.path.dirname(file_full_name)

        self.entry_file_name.insert(0, (self.user_entry.file_name + self.user_entry.file_suffix))
        self.entry_file_location.insert(0, os.path.normcase(self.user_entry.file_location))
        message = 'File Loaded: ' + str(file_full_name) + '\n'
        message_colour = 'black'
        self.session_log.write_textbox(message, message_colour)
        print("::file name: ", self.user_entry.file_name)
        print("::file suffix: ", self.user_entry.file_suffix)
        print("::file location: ", self.user_entry.file_location)

    def button_file_peak(self):
        self.peak_file()

    def open_folder(self, folder_path):
        temp_path = os.path.realpath(folder_path)
        try:
            os.startfile(temp_path)
        except:
            try:
                os.mkdir(FILE_LOCATION)
                self.user_entry.file_location = FILE_LOCATION
                self.session_log.write_textbox("Folder Created", "blue")
                print("Folder Created")
            except OSError as e:
                print("Failed to Create Folder")
                e = Exception("Failed to Create Folder")
                self.session_log.write_textbox(str(e), "red")
                raise e

    def _save_data_pandas(self, folder, file_name, data):
        file_address = (folder + "/" + file_name + '.csv')
        if os.path.isfile(file_address):
            self.session_log.write_textbox("File Exists and Will Be Overwritten: ", "red")
        pd.DataFrame(data).to_csv(file_address, index=False, header=False)

    ######################################################################

    @staticmethod
    def message_box(title, data):
        try:
            messagebox.showinfo(title=title, message=data)
        except:
            data = 'Invalid data'

    def display_session_settings(self):
        message = 'File To Process: ' + fm.FileNameMethods.build_file_name_full(
            self.user_entry.file_location, self.user_entry.file_name, self.user_entry.file_suffix) + '\n'
        colour = "blue"
        self.session_log.write_textbox(message, colour)

    def textbox_row_clear(self):
        self.textbox_rows.configure(state='normal')
        self.textbox_rows.delete('1.0', 'end')
        self.textbox_rows.insert('end', 'No Files Loaded')
        self.textbox_rows.configure(state='disabled')

    def textbox_update(self, data):
        self.textbox_rows.configure(state='normal')
        self.textbox_rows.delete('1.0', 'end')
        self.textbox_rows.insert('end', data)
        self.textbox_rows.configure(state='disabled')
