__author__ = 'paulj@paulj.net'

import csv
import glob
import json
import os
import re
import sqlite3
import tkinter as tk
from datetime import datetime
from shutil import copy2 as copy
from sqlite3 import Error
from tkinter import filedialog, messagebox
from tkinter import ttk


class MainApplication:
    def __init__(self):
        # set ui variables
        self.db_path = ""
        self.conn = None
        self.cursor = None
        self.parent = tk.Tk()
        self.parent.minsize(400, 0)
        self.parent.resizable(False, False)
        self.parent.title("Bluebeam Studio Markup Database Tool")
        # top menu
        self.menubar = tk.Menu(self.parent)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="New", command=self.new_db)
        self.filemenu.add_command(label="Open", command=self.open_db)
        self.filemenu.add_command(label="Close", command=self.close_db, state='disabled')
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=self.parent.quit)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        self.editmenu = tk.Menu(self.menubar, tearoff=0)
        self.editmenu.add_command(label="Load New Date", command=self.load_new_date)
        self.editmenu.add_command(label="Edit Past Dates", command=self.edit_past_dates)
        self.menubar.add_cascade(label="Edit", menu=self.editmenu, state='disabled')
        self.toolmenu = tk.Menu(self.menubar, tearoff=0)
        self.toolmenu.add_command(label="Export Latest Date", command=self.export_latest_date)
        self.toolmenu.add_command(label="Custom Export", command=self.custom_export)
        self.menubar.add_cascade(label="Tools", menu=self.toolmenu, state='disabled')
        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="Guide", command=self.guide)
        self.helpmenu.add_command(label="About...", command=self.about)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)
        self.parent.config(menu=self.menubar)

        # initialize window
        self.parent.option_add('*Font', '24')
        self.frame1 = tk.Frame()
        self.frame1.pack(padx=5, pady=5, anchor='w')
        self.db_label1 = tk.Label(self.frame1, text="Database:")
        self.db_label1.grid(column=0, row=0)
        self.db_label2 = tk.Label(self.frame1, text="No database loaded")
        self.db_label2.grid(column=1, row=0, columnspan=10)
        self.break_label2 = tk.Label(self.frame1, text="")
        self.break_label2.grid(column=0, row=1, columnspan=3)
        self.title_label1 = tk.Label(self.frame1, text="Previous")
        self.title_label1.grid(column=1, row=2, sticky='w')
        self.title_label2 = tk.Label(self.frame1, text="Total")
        self.title_label2.grid(column=2, row=2, sticky='w')
        self.runs_label1 = tk.Label(self.frame1, text="Runs:")
        self.runs_label1.grid(column=0, row=3, sticky='e')
        self.runs_label2 = tk.Label(self.frame1, text="N/A")
        self.runs_label2.grid(column=1, row=3, sticky='w')
        self.runs_label3 = tk.Label(self.frame1, text="N/A")
        self.runs_label3.grid(column=2, row=3, sticky='w')
        self.files_label1 = tk.Label(self.frame1, text="Files:")
        self.files_label1.grid(column=0, row=4, sticky='e')
        self.files_label2 = tk.Label(self.frame1, text="N/A")
        self.files_label2.grid(column=1, row=4, sticky='w')
        self.files_label3 = tk.Label(self.frame1, text="N/A")
        self.files_label3.grid(column=2, row=4, sticky='w')
        self.markups_label1 = tk.Label(self.frame1, text="Markups:")
        self.markups_label1.grid(column=0, row=5, sticky='e')
        self.markups_label2 = tk.Label(self.frame1, text="N/A")
        self.markups_label2.grid(column=1, row=5, sticky='w')
        self.markups_label3 = tk.Label(self.frame1, text="N/A")
        self.markups_label3.grid(column=2, row=5, sticky='w')
        self.parent.mainloop()

    def init_window(self):
        self.update_window()
        self.filemenu.entryconfig("Close", state='normal')
        self.menubar.entryconfig("Edit", state='normal')
        self.menubar.entryconfig("Tools", state='normal')

    def update_window(self):
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.cursor.execute('''SELECT update_time FROM files ORDER BY update_time DESC LIMIT 1''')
        previous_run = self.cursor.fetchone()
        if previous_run is not None:
            previous_run = previous_run[0]
            self.cursor.execute('''SELECT COUNT(DISTINCT update_time) FROM files''')
            total_runs = self.cursor.fetchone()
            self.cursor.execute('''SELECT COUNT(file_id) FROM files WHERE update_time = ?''', [previous_run])
            previous_files = self.cursor.fetchone()
            self.cursor.execute('''SELECT COUNT(file_id) FROM files''')
            total_files = self.cursor.fetchone()
            self.cursor.execute('''SELECT COUNT(markup_id) FROM markups
                                               INNER JOIN files on files.file_id = markups.file_id
                                               WHERE files.update_time = ?''', [previous_run])
            previous_markups = self.cursor.fetchone()
            self.cursor.execute('''SELECT COUNT(markup_id) FROM markups''')
            total_markups = self.cursor.fetchone()
            previous_run = previous_run[:-9]
        else:
            previous_run = "N/A"
            total_runs = 0
            previous_files = 0
            total_files = 0
            previous_markups = 0
            total_markups = 0

        self.db_label2.configure(text=self.db_path)
        self.runs_label2.configure(text=previous_run)
        self.runs_label3.configure(text=total_runs)
        self.files_label2.configure(text=previous_files)
        self.files_label3.configure(text=total_files)
        self.markups_label2.configure(text=previous_markups)
        self.markups_label3.configure(text=total_markups)

    def new_db(self):
        # wait for user to choose db filename
        self.db_path = filedialog.asksaveasfilename(defaultextension=".sqlite")
        if len(self.db_path) != 0:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.cursor.execute('''CREATE TABLE files(
                                   file_id INTEGER PRIMARY KEY AUTOINCREMENT,
                                   filename TEXT,
                                   update_time DATETIME,
                                   friendly_name TEXT)''')
            self.cursor.execute('''CREATE TABLE markups(
                                   file_id TEXT,
                                   markup_id TEXT,
                                   type TEXT,
                                   page TEXT,
                                   author TEXT,
                                   subject TEXT,
                                   comment TEXT,
                                   color TEXT,
                                   colorfill TEXT,
                                   opacity TEXT,
                                   opacityfill TEXT,
                                   rotation TEXT,
                                   status TEXT,
                                   checked TEXT,
                                   locked TEXT,
                                   datecreated DATETIME,
                                   datemodified DATETIME,
                                   linewidth TEXT,
                                   x TEXT,
                                   y TEXT,
                                   width TEXT,
                                   height TEXT,
                                   linestyle TEXT,
                                   space TEXT,
                                   ExtendedProperties JSON,
                                   colortext TEXT,
                                   layer TEXT,
                                   parent TEXT,
                                   grouped TEXT,
                                   PRIMARY KEY (file_id, markup_id),
                                   FOREIGN KEY (file_id) REFERENCES files(file_id))''')
            self.conn.commit()
            self.init_window()

    def open_db(self):
        # wait for user to select update folder
        self.db_path = filedialog.askopenfilename(title='Select a markups database',
                                                  initialdir=os.getcwd(),
                                                  filetypes=(('sqlite files', 'sqlite'),))
        if len(self.db_path) != 0:
            self.init_window()

    def close_db(self):
        self.conn = None
        self.cursor = None
        self.db_path = ""
        self.db_label2.configure(text="No database loaded")
        self.runs_label2.configure(text="N/A")
        self.runs_label3.configure(text="N/A")
        self.files_label2.configure(text="N/A")
        self.files_label3.configure(text="N/A")
        self.markups_label2.configure(text="N/A")
        self.markups_label3.configure(text="N/A")

        self.filemenu.entryconfig("Close", state='disabled')
        self.menubar.entryconfig("Edit", state='disabled')
        self.menubar.entryconfig("Tools", state='disabled')

    def load_new_date(self):
        # wait for user to select update folder
        dir_choice = filedialog.askdirectory(initialdir=os.getcwd(), title='Please select a directory')
        if len(dir_choice) == 0:
            return True
        else:
            dirname = os.path.normpath(dir_choice)
        filenames = glob.glob(r'%s\**\*.json' % dirname, recursive=True)

        # see if any previous updates happened today and prompt for overwrite or cancel
        self.cursor.execute('''SELECT file_id FROM files WHERE date(update_time) = date('now')''')
        previous_files = self.cursor.fetchall()
        if len(previous_files) != 0:
            # yes/no to overwrite or cancel
            if not messagebox.askokcancel("Duplicate Date Warning!",
                                          "The database already contains entries with today's date.\n"
                                          "Press OK to overwrite previous data from today or CANCEL to exit."):

                return
            else:
                self.cursor.executemany('''DELETE FROM files WHERE file_id = ?''', previous_files)
                self.cursor.executemany('''DELETE FROM markups WHERE file_id = ?''', previous_files)

        # set start time of update and initialize console window
        now_time = datetime.now().replace(microsecond=0)
        now_time_string = now_time.strftime("%Y-%m-%d %H:%M:%S")
        print("Start time: %s\n" % now_time_string)

        # read available columns
        # noinspection SqlResolve
        self.cursor.execute('''SELECT * FROM pragma_table_info('markups')''')
        column_pos = 0
        columns = {}
        sql_markup = ''
        new_insert = []
        for column in self.cursor.fetchall():
            columns[column[1]] = column_pos
            column_pos += 1
            sql_markup += '?,'
            new_insert.append(None)
        sql_markup = sql_markup[:-1]

        # backup db before making any changes
        print("Backing up db before update.\n")
        copy(self.db_path, self.db_path + '.backup')
        total_markups = 0

        for filename in filenames:
            # a lot of work needs to be done to clean up this json data
            with open(filename, "r") as json_file:
                json_data = json_file.read()
                if json_data == '[""]':
                    print("No Markups! Skipping %s" % filename)
                    continue
                else:
                    print("Importing %s" % filename)
                    # modify this to strip most of the path, only include the items inside the folder not the full path

                    self.cursor.execute('''INSERT INTO files VALUES (null,?,?,?)''',
                                        (filename, now_time, os.path.basename(filename)[11:-5]))
                    file_id = self.cursor.lastrowid
                    # strip square brackets and quotes around entire line
                    json_data = json_data[2:-2]
                    # clean up first level
                    json_data = re.sub(r"(?<!\|)'", r'"', json_data)
                    json_data = re.sub(r'":"{', r'":{', json_data)
                    json_data = re.sub(r'\|\'}"', r"|'}", json_data)
                    # clean up second level
                    json_data = re.sub(r"(?<!\|)\|'", r'"', json_data)
                    # clean up second level escaped characters
                    json_data = re.sub(r'\|\|\|\|\|\',"', r'|", "', json_data)
                    json_data = re.sub(r'(?<!\|)\|\|\|\\"', r'\"', json_data)
                    json_data = re.sub(r'(?<!\|)\|\|\|\|(?![|r])', r'|', json_data)
                    json_data = re.sub(r'(?<!\|)\|\|r', r'\\n', json_data)
                    try:
                        markups = json.loads(json_data)
                    except ValueError:
                        messagebox.showwarning("Error!", "Error loading main JSON data from file '%s'." % filename)
                        return

                    markups_inserts = []

                    for markup_item in markups.items():
                        insert = new_insert.copy()
                        insert[columns['file_id']] = file_id
                        insert[columns['markup_id']] = markup_item[0]
                        for key, value in markup_item[1].items():
                            if key == 'ExtendedProperties':
                                extended_json = value
                                # cleanup extended data
                                extended_json = re.sub(r"\\", r"\\\\", extended_json)
                                # clean up first level
                                extended_json = re.sub(r"(?<!\|)\|\|\|'", r'"', extended_json)
                                # clean up first level escaped characters
                                extended_json = re.sub(r'(?<!\|)\|\|\|\|\|\|\|\\"', r'\"', extended_json)
                                extended_json = re.sub(r'(?<!\|)\|\|\|\|\|\|\|"', r'\"', extended_json)
                                extended_json = re.sub(r"(?<!\|)\|\|\|\|\|\|\|'", r"'", extended_json)
                                extended_json = re.sub(r'(?<!\|)\|\|\|\|r', r'\\n', extended_json)
                                extended_json = re.sub(r'(?<!\|)\|\|\|\|\|\|\|\|', r'|', extended_json)
                                # validate json
                                try:
                                    extended = json.loads(extended_json)
                                except ValueError:
                                    messagebox.showwarning(
                                        "Error!", "Error loading extended JSON data from file '%s', markup '%s'"
                                                  % (filename, markup_item[0]))
                                    return
                                # gather extended data
                                # noinspection PyTypeChecker
                                insert[columns['ExtendedProperties']] = extended_json
                            else:
                                try:
                                    # noinspection PyTypeChecker
                                    insert[columns[key]] = re.sub(r"(?<!\|)\|\|\|'", "'", value)
                                except KeyError:
                                    messagebox.showwarning("Error!", "New key (%s) found in file '%s'."
                                                           % (key, filename))
                                    return
                        markups_inserts.append(insert)
            # noinspection SqlInsertValues
            self.cursor.executemany('''INSERT INTO markups VALUES (%s)''' % sql_markup, markups_inserts)
            total_markups += len(markups_inserts)

        self.conn.commit()
        self.cursor.execute('''VACUUM''')

        print("Successfully imported %s files containing %s markups in %s seconds."
              % (str(len(filenames)), total_markups, (datetime.now() - now_time).total_seconds()))
        self.update_window()
        return

    def edit_past_dates(self):
        root_epd = tk.Tk()
        root_epd.title('Edit Past Date')
        root_epd.geometry("360x180")
        root_epd.resizable(False, False)

        self.cursor.execute('''SELECT DISTINCT update_time FROM files''')
        all_runs = self.cursor.fetchall()
        run_list = [run[0] for run in all_runs]

        def delete_record():
            # Backup db before making any changes
            print("Backing up database before update......")
            copy(self.db_path, self.db_path + '.backup')

            self.cursor.execute('''SELECT file_id FROM files WHERE (update_time) = ?''', (clicked.get(),))
            chosen_files = self.cursor.fetchall()
            self.cursor.executemany('''DELETE FROM files WHERE file_id = ?''', chosen_files)
            self.cursor.executemany('''DELETE FROM markups WHERE file_id = ?''', chosen_files)

            self.conn.commit()
            self.cursor.execute('''VACUUM''')
            print(f'Successfully Deleted {clicked.get()} Data From Database')
            self.update_window()
            root_epd.destroy()

        def set_end_date(e):  # <- parameter 'e' value is not used
            start_date = clicked_start.get()
            delete_multiple.configure(state="disabled")
            if start_date != "Start Date":
                copy_list = end_list.copy()
                for index in range(0, start_list.index(start_date)):
                    copy_list.remove(end_list[index])
                copy_list.insert(0, "End Date")
                drop_end.configure(values=copy_list, state="readonly")
                drop_end.set(copy_list[0])
            else:
                drop_end.set("End Date")
                drop_end.configure(state="disabled")

        def set_button(e):
            end_date = clicked_end.get()
            if end_date != "End Date":
                delete_multiple.configure(state="normal")
            else:
                delete_multiple.configure(state="disabled")

        def delete_multiple_records():
            start = run_list.index(clicked_start.get())
            end = run_list.index(clicked_end.get()) + 1
            # Table import
            print("Backing up database before update......")
            copy(self.db_path, self.db_path + '.backup')
            for index in range(start, end):
                self.cursor.execute('''SELECT file_id FROM files WHERE (update_time) = ?''', (run_list[index],))
                chosen_files = self.cursor.fetchall()
                self.cursor.executemany('''DELETE FROM files WHERE file_id = ?''', chosen_files)
                self.cursor.executemany('''DELETE FROM markups WHERE file_id = ?''', chosen_files)
                print(f'Deleted {run_list[index]} Data From Database')
            self.conn.commit()
            print("Committing changes...")
            self.cursor.execute('''VACUUM''')
            self.update_window()
            root_epd.destroy()
            print(f'Successfully Deleted Data Between {clicked_start.get()} and {clicked_end.get()} From Database')

        frame1 = tk.Frame(root_epd)
        frame1.pack(anchor="n", side="left")

        clicked = tk.StringVar(frame1)
        clicked.set(run_list[0])
        drop = ttk.Combobox(frame1, textvariable=clicked, values=run_list, state="readonly")
        drop.grid(row=1, column=0, padx=20, pady=(5, 0), ipady=5, sticky="n")
        tk.Button(frame1, text="Delete Data", command=delete_record).grid(row=3, column=0, ipady=2)
        tk.Label(frame1, text="Delete by Date").grid(row=0, column=0, padx=20, pady=(5, 10))

        start_list = run_list.copy()
        start_list.insert(0, "Start Date")
        clicked_start = tk.StringVar(root_epd)
        clicked_start.set(start_list[0])
        drop_start = ttk.Combobox(frame1, textvariable=clicked_start, values=start_list, state="readonly")
        drop_start.grid(row=1, column=1, padx=20, pady=(5, 0), sticky="n")
        drop_start.bind("<<ComboboxSelected>>", set_end_date)

        end_list = run_list.copy()
        clicked_end = tk.StringVar(root_epd)
        clicked_end.set("End Date")
        drop_end = ttk.Combobox(frame1, textvariable=clicked_end, values=end_list, state="disabled")
        drop_end.grid(row=2, column=1, padx=20, sticky="n")
        drop_end.bind("<<ComboboxSelected>>", set_button)

        delete_multiple = tk.Button(frame1, text="Delete Multiple Dates", command=delete_multiple_records,
                                    state="disabled")
        delete_multiple.grid(row=3, column=1, padx=20, pady=15, ipady=2)
        tk.Label(frame1, text="Delete by Date Range").grid(row=0, column=1, padx=20, pady=(10, 5))

        root_epd.mainloop()

    def export_latest_date(self):
        try:
            # Fetch latest date in table
            self.cursor.execute('''SELECT update_time FROM Files ORDER BY update_time DESC LIMIT 1''')
            previous_run = self.cursor.fetchone()
            previous_run = previous_run[0]
            print(f"Latest date is: {previous_run} ")
            print(f"Extracting {previous_run} data from sqlite...")
            self.cursor.execute('''SELECT files.file_id, filename, update_time, friendly_name, markup_id, type, page, author, subject, comment, color, 
            colorfill, opacity, opacityfill, rotation, status, checked, locked, datecreated, datemodified, linewidth, x, y, width, height, 
            linestyle, space, extendedproperties, colortext, layer, parent, grouped FROM markups
                                   INNER JOIN files on files.file_id = markups.file_id
                                   WHERE update_time = ?''', [previous_run])
            title = f"latest markups data - {previous_run[:10]}.csv"
            with open(title, "w", newline="") as csv_file:
                csv_writer = csv.writer(csv_file, delimiter=",")
                csv_writer.writerow([i[0] for i in self.cursor.description])
                csv_writer.writerows(self.cursor)
            print("Data Exported Successfully Into {}".format(os.getcwd() + r"\{}".format(title)))
        except Error as e:
            print(e)

    def custom_export(self):
        root_custom_export = tk.Tk()
        root_custom_export.title('Custom Export')
        root_custom_export.geometry('450x200')
        root_custom_export.resizable(False, False)

        self.cursor.execute('''SELECT DISTINCT update_time FROM files''')
        all_runs = self.cursor.fetchall()
        run_list = [run[0] for run in all_runs]

        def export_all():
            try:
                # Table import
                print("Extracting data from sqlite...")
                self.cursor.execute('''SELECT files.file_id, filename, update_time, friendly_name, markup_id, type, page, author, subject, comment, color,
                colorfill, opacity, opacityfill, rotation, status, checked, locked, datecreated, datemodified, linewidth, x, y, width, height,
                linestyle, space, extendedproperties, colortext, layer, parent, grouped FROM markups INNER JOIN files on files.file_id = markups.file_id''')
                with open("markups data.csv", "w", newline="") as csv_file:
                    csv_writer = csv.writer(csv_file, delimiter=",")
                    csv_writer.writerow([i[0] for i in self.cursor.description])
                    csv_writer.writerows(self.cursor)
                print("Data Exported Successfully Into {}".format(os.getcwd() + r"\markups_data.csv"))
            except Error as e:
                print(e)

        def export_date():
            try:
                # Table import
                print(f"Extracting {clicked.get()[:10]} data from sqlite...")
                self.cursor.execute('''SELECT files.file_id, filename, update_time, friendly_name, markup_id, type, page, author, subject, comment, color,
                colorfill, opacity, opacityfill, rotation, status, checked, locked, datecreated, datemodified, linewidth, x, y, width, height,
                linestyle, space, extendedproperties, colortext, layer, parent, grouped FROM markups
                INNER JOIN files on files.file_id = markups.file_id
                WHERE update_time = ?''', [clicked.get()])
                title = f"markups {clicked.get()[:10]}.csv"
                with open(title, "w", newline="") as csv_file:
                    csv_writer = csv.writer(csv_file, delimiter=",")
                    csv_writer.writerow([i[0] for i in self.cursor.description])
                    csv_writer.writerows(self.cursor)
                print("Data Exported Successfully Into {}".format(os.getcwd() + r"\{}".format(title)))
            except Error as e:
                print(e)

        def set_end_date(e):
            start_date = clicked_start.get()
            export_multiple.configure(state="disabled")
            if start_date != "Start Date":
                copy_list = end_list.copy()
                for index in range(0, start_list.index(start_date)):
                    copy_list.remove(end_list[index])
                copy_list.insert(0, "End Date")
                drop_end.configure(values=copy_list, state="readonly")
                drop_end.set(copy_list[0])
            else:
                drop_end.set("End Date")
                drop_end.configure(state="disabled")

        def set_button(e):
            end_date = clicked_end.get()
            if end_date != "End Date":
                export_multiple.configure(state="normal")
            else:
                export_multiple.configure(state="disabled")

        def export_date_range():
            try:
                start = run_list.index(clicked_start.get())
                end = run_list.index(clicked_end.get()) + 1
                # Table import
                print(
                    f"Extracting data between {clicked_start.get()[:10]} and {clicked_end.get()[:10]} from sqlite...")
                title = f"markups {clicked_start.get()[:10]} to {clicked_end.get()[:10]}.csv"
                with open(title, "w", newline="") as csv_file:
                    for index in range(start, end):
                        self.cursor.execute('''SELECT files.file_id, filename, update_time, friendly_name, markup_id, type, page, author, subject, comment, color,
                                       colorfill, opacity, opacityfill, rotation, status, checked, locked, datecreated, datemodified, linewidth, x, y, width, height,
                                       linestyle, space, extendedproperties, colortext, layer, parent, grouped FROM markups
                                       INNER JOIN files on files.file_id = markups.file_id
                                       WHERE update_time = ?''', [run_list[index]])
                        csv_writer = csv.writer(csv_file, delimiter=",")
                        csv_writer.writerow([i[0] for i in self.cursor.description])
                        csv_writer.writerows(self.cursor)
                print("Data Exported Successfully Into {}".format(os.getcwd() + r"\{}".format(title)))
            except Error as e:
                print(e)

        frame1 = tk.Frame(root_custom_export, pady=10)
        frame1.pack(anchor="n", side="top")
        frame2 = tk.Frame(root_custom_export)
        frame2.pack(anchor="n", side="left")
        frame3 = tk.Frame(root_custom_export)
        frame3.pack(anchor="n", side="right")

        tk.Button(frame1, text="Export All Runs", command=export_all).grid(row=0, column=0, padx=20, sticky="w")

        clicked = tk.StringVar(root_custom_export)
        clicked.set(run_list[0])
        drop = ttk.Combobox(frame2, textvariable=clicked, values=run_list, state="readonly")
        drop.grid(row=1, column=0, padx=50, pady=(5, 5), ipady=5, sticky="n")
        tk.Button(frame2, text="Export Date", command=export_date).grid(row=3, column=0, padx=50, pady=10, ipadx=2)
        tk.Label(frame2, text="Export by Date").grid(row=0, column=0, padx=50, pady=3)

        start_list = run_list.copy()
        start_list.insert(0, "Start Date")
        clicked_start = tk.StringVar(root_custom_export)
        clicked_start.set(start_list[0])
        drop_start = ttk.Combobox(frame3, textvariable=clicked_start, values=start_list, state="readonly")
        drop_start.grid(row=1, column=1, padx=20, sticky="n")
        drop_start.bind("<<ComboboxSelected>>", set_end_date)

        end_list = run_list.copy()
        clicked_end = tk.StringVar(root_custom_export)
        clicked_end.set("End Date")
        drop_end = ttk.Combobox(frame3, textvariable=clicked_end, values=end_list, state="disabled")
        drop_end.grid(row=2, column=1, padx=20, sticky="n")
        drop_end.bind("<<ComboboxSelected>>", set_button)

        export_multiple = tk.Button(frame3, text="Export Date Range", command=export_date_range, state="disabled")
        export_multiple.grid(row=3, column=1, padx=20, pady=5, ipady=2, sticky="n")
        tk.Label(frame3, text="Export by Date Range").grid(row=0, column=1, padx=50, pady=3)

        root_custom_export.mainloop()

    # noinspection PyMethodMayBeStatic
    def guide(self):
        root_guide = tk.Tk()
        root_guide.title('Guide')
        root_guide.geometry("800x700")

        body_text = f"""Welcome to Bluebeam Studio Markup Database Help\n    
Get started 
{"This program is a tool created by Paul Jeffrey to convert the JSON files in Bluebeam to an SQLite database that " +
 "can be imported into Power BI. You need Python downloaded to start running the program. Once you open the program, " +
 "a window will pop up. This program includes four primary tabs: File, Edit, Tools, and Help."}\n
File
{"Under File, you can choose the following commands: New, Open, Close, and Exit. " +
 "The New tool creates a new file in this program. You will not need to use it often since we are only " +
 "using one SQLite file that was previously opened. The Open tool opens an SQLite file on your computer. " +
 "Note that the file type must be SQLite. After opening the file, you need to follow the procedure from " +
 "the student hand-over document to successfully run an update on Power BI. Use Close command to close the " +
 "currently opening file. This command is unavailable when you initialize the program since no database is loaded. " +
 "To close the program, simply click the X icon at the top right corner or go to File > Exit."}\n
Edit
{"You cannot Edit (Load New Date & Edit Past Date) because no database is loaded. You will be able to use " +
 "those functions once the SQLite database is open. Click Load New Date to modify the SQLite file by adding data " +
 "from a date. A window pops up asking for a folder to load data from a new date. Edit Past Dates " +
 "allows you to see the previous updates of the SQLite file and modify these versions in a pop-up window. " +
 "There is a dropdown window with a list of run dates to delete from the SQLite database. " +
 "Changes committed from those functions will be shown on the main application window."}\n
Tools
{"You cannot Edit (Export Latest Date & Custom Export) because no database is loaded. " +
 "You can use those functions once the SQLite database is open. Both functions transform " +
 "data into an Excel CSV file named accordingly. Export Last Date enables you to export the SQLite database " +
 "from the latest run. Custom Export Date allows you to export data sorted by run date from the SQLite database. " +
 "You can export the enormous whole SQLite database to store all run dates in a single spreadsheet. " +
 "Alternatively, you can pick a date from the dropdown menu and export the specific date you desire."}\n
Help
{"Click Guide to see a brief introduction of the tools and their functions in this program. " +
 "Click About… to see the program’s copyright information and contributors."}
"""

        guide_text = tk.Text(root_guide, wrap="word", width=int("%d" % root_guide.winfo_screenwidth()),
                             height=int("%d" % root_guide.winfo_screenheight()), font=('Calibri', 12), padx=10, pady=10)
        guide_text.insert("insert", body_text)
        guide_text.configure(state="disabled")
        guide_text.pack()
        root_guide.mainloop()

    # noinspection PyMethodMayBeStatic
    def about(self):
        root_about = tk.Tk()
        root_about.title('About...')
        root_about.geometry("440x200")
        root_about.resizable(False, False)

        welcome = tk.Label(root_about, text="""
    Welcome to Bluebeam Studio Markup Database Help
    """, fg='blue', font=16)
        welcome.grid(sticky="w")
        tk.Label(root_about, text="""
        Product: JSON to Database Markup Converting Tool\n
        Author: Paul Jeffrey - paulj@paulj.net\n
        Assistant: Duc Huy Ngo
        """, font=('Calibri', 12), justify="left").grid(row=1, sticky="w")

        root_about.mainloop()


if __name__ == "__main__":
    app = MainApplication()
