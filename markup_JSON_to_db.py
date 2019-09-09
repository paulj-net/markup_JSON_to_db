__author__ = 'paulj@paulj.net'

from datetime import datetime
import glob
import json
import os
import re
import sqlite3
import sys
import tkinter
from tkinter import filedialog
from shutil import copy2 as copy


def create_db(path):
    conn = sqlite3.connect(path)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE files(
                        file_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        filename TEXT,
                        update_time DATETIME,
                        friendly_name TEXT)''')
    cursor.execute('''CREATE TABLE markups(
                        file_id TEXT REFERENCES files,
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
                        PRIMARY KEY (file_id, markup_id))''')
    conn.commit()
    conn.close()


if __name__ == "__main__":
    # wait for user to select update folder
    root = tkinter.Tk()
    root.withdraw()
    dir_choice = filedialog.askdirectory(initialdir=os.getcwd(), title='Please select a directory')
    root.destroy()
    if len(dir_choice) == 0:
        sys.exit()
    else:
        dirname = os.path.normpath(dir_choice)
    filenames = glob.glob(r'%s\**\*.json' % dirname, recursive=True)

    # set start time of update
    now_time = datetime.now().replace(microsecond=0)
    now_time_string = now_time.strftime("%Y-%m-%dT%H:%M:%S")
    print("Start time: %s" % now_time.strftime("%H:%M:%S"))
    # create new or backup existing db
    db_path = r'.\markups.sqlite'
    if os.path.exists(db_path):
        print("Backing up db before update.")
        copy(db_path, db_path + '.backup')
    else:
        print("Creating fresh db (no existing db found!).")
        create_db(db_path)
    # connect to db and read available columns
    mconn = sqlite3.connect(db_path)
    mcursor = mconn.cursor()
    # noinspection SqlResolve
    mcursor.execute('''SELECT * FROM pragma_table_info('markups')''')
    column_pos = 0
    columns = {}
    sql_markup = ''
    new_insert = []
    for column in mcursor.fetchall():
        columns[column[1]] = column_pos
        column_pos += 1
        sql_markup += '?,'
        new_insert.append(None)
    sql_markup = sql_markup[:-1]
    total_markups = 0

    for filename in filenames:
        # a lot of work needs to be done to clean up this json data
        with open(filename, "r") as json_file:
            print("Importing %s" % filename)
            mcursor.execute('''INSERT INTO files VALUES (null,?,?,?)''', (filename, now_time,
                                                                          os.path.basename(filename)[11:-5]))
            file_id = mcursor.lastrowid
            json_data = json_file.read()
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
            print("Error loading main JSON data from file '%s'.\n"
                  "Exiting script." % filename)
            sys.exit()

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
                        print("Error loading extended JSON data from file '%s', markup '%s'.\n"
                              "Exiting script." % (filename, markup_item[0]))
                        sys.exit()
                    # gather extended data
                    # noinspection PyTypeChecker
                    insert[columns['ExtendedProperties']] = extended_json
                else:
                    try:
                        # noinspection PyTypeChecker
                        insert[columns[key]] = re.sub(r"(?<!\|)\|\|\|'", "'", value)
                    except KeyError:
                        print("New key (%s) found in file '%s'.\n"
                              "Exiting script." % (key, filename))
                        sys.exit()
            markups_inserts.append(insert)
        # noinspection SqlInsertValues
        mcursor.executemany('''INSERT INTO markups VALUES (%s)''' % sql_markup, markups_inserts)
        total_markups += len(markups_inserts)

    mconn.commit()
    mcursor.execute('''VACUUM''')
    mconn.close()

    print("Successfully imported %s files containing %s markups in %s seconds."
          % (str(len(filenames)), total_markups, (datetime.now() - now_time).total_seconds()))
