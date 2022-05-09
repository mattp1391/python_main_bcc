import sqlite3
import time
import sys
import os

import tkinter as tk
from tkinter import filedialog as tf
from tkinter import messagebox as tmb
from tkinter import constants as tc


class TfAimsunSqliteTool(tk.Frame):

    def __init__(self, root):
        tk.Frame.__init__(self, root)
        self.colNames = None
        self.colType = None
        self.AvailableReps = None
        self.AvailableTables = []
        # options for buttons
        button_opt = {'fill': tc.BOTH, 'padx': 5, 'pady': 5}

        # self.msg = tk.Message(self, text='').grid(row=6, column=0)
        tk.Button(self, text='Input File', command=self.askopenfilename).grid(row=0, column=0, sticky='E')
        tk.Button(self, text='Output File', command=self.ask_output_file).grid(row=1, column=0, sticky='E')
        tk.Button(self, text='ImportTables', command=self.get_list_box_selection).grid(row=6, column=1, sticky='E')
        self.var = self.user_file_input("SQLITE InputFile required")  # .grid(row=0,column=1)
        self.var2 = self.user_file_input2("Output file required")
        tk.Label(self, textvariable=self.var).grid(row=0, column=1, sticky='W')
        tk.Label(self, textvariable=self.var2).grid(row=1, column=1, sticky='W')
        tk.Label(self, text='SQLITE Tables').grid(row=4, column=0, sticky='W')
        self.listbox = tk.Listbox(self, selectmode='EXTENDED', exportselection=0)
        self.listbox.grid(row=5, column=0, sticky='W')
        tk.Label(self, text='Replications').grid(row=4, column=1, sticky='W')
        self.listboxReps = tk.Listbox(self, selectmode='EXTENDED', exportselection=0)
        self.listboxReps.grid(row=5, column=1, sticky='W')
        tk.Label(self, text='Replications').grid(row=4, column=1, sticky='W')
        self.file_opt = options = {}
        options['defaultextension'] = '.sqlite'
        options['filetypes'] = [('sqlite files', '.sqlite'), ('all files', '.*')]
        options['initialdir'] = 'C:\\'
        options['initialfile'] = 'myfile.sqlite'
        options['parent'] = root
        options['title'] = 'Please Choose SQLITE File'
        self.dir_opt = options = {}
        options['initialdir'] = 'C:\\'
        options['mustexist'] = False
        options['parent'] = root
        options['title'] = 'This is a title'

    def get_tables(self):

        if tk.Entry(self, textvariable=self.var).get()[-6:].lower() == "sqlite":
            f = tk.Entry(self, textvariable=self.var).get()
            conn = sqlite3.connect(tk.Entry(self, textvariable=self.var).get())
            c = conn.cursor()
            table_list = []
            for row in c.execute('SELECT name FROM sqlite_master WHERE type = "table" ORDER BY name').fetchall():
                table_list = table_list + [str(list(row)[0])]
            item = 0
            if table_list is not None:
                self.listbox.delete(0, tk.END)
                for i in table_list:
                    item += 1
                    self.listbox.insert(tk.END, i)
            c.close()
            self.AvailableTables = table_list
            # return tableList

    def askopenfile(self):

        """Returns an opened file in read mode."""
        # var.set(fd.askopenfilename())
        return tf.askopenfile(mode='r', **self.file_opt)

    def add_to_list(self, list_elements, list_box_to_adjust):
        item = 0
        list_box_to_adjust.delete(0, tk.END)
        if list_elements is not None:
            for i in list_elements:
                item += 1
                list_box_to_adjust.insert(tk.END, i)

        self.AvailableReps = list_elements

    def askopenfilename(self):

        """Returns an opened file in read mode.
        This time the dialog just returns a filename and the file is opened by your own code.
        """
        filename = tf.askopenfilename(**self.file_opt)
        if filename:
            self.var.set(filename)
            self.get_tables()
            sql = self.select_data('SIM_INFO', None, columns='SIM_INFO.did')
            conn = sqlite3.connect(tk.Entry(self, textvariable=self.var).get())
            c = conn.cursor()
            try:
                c.execute(sql)
            except:
                sql = ''
            table_data = c.fetchall()
            # tableData=(str(tableData
            c.close()

            rep_list = []
            for i in table_data:
                string = str(i)
                string = string.replace('(', '')
                string = string.replace(',)', '')
                rep_list.append(string)

            self.add_to_list(rep_list, self.listboxReps)

    def ask_save_as_file(self):

        """Returns an opened file in write mode."""

        f = tf.asksaveasfile(mode='w', **self.file_opt)
        # print f
        if f is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            print('None')
            return
        else:
            self.var.set(f)

    def ask_output_file(self):

        """Returns an opened file in write mode."""

        f = tf.asksaveasfilename(**self.file_opt)

        # print f
        if f is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            print('None')
            return
        else:
            self.var2.set(f)

    def asksaveasfilename(self):

        """Returns an opened file in write mode.
        This time the dialog just returns a filename and the file is opened by your own code.
        """

        # get filename
        f = tf.asksaveasfilename(**self.file_opt)

        # open file on your own
        if f:
            self.var.set(f)

    def askdirectory(self):

        """Returns a selected directoryname."""

        return tf.askdirectory(**self.dir_opt)

    def user_file_input(self, text):
        var = tk.StringVar(root)
        var.set(text)
        return var

    def user_file_input2(self, text):
        var2 = tk.StringVar(root)
        var2.set(text)
        return var2

    def get_list_box_selection(self):

        tables_to_import = self.listbox.curselection()

        new_sqlite_tables = ['SIM_INFO', 'META_INFO', 'META_SUB_INFO']
        for t in tables_to_import:
            if t not in new_sqlite_tables:
                new_sqlite_tables.append(self.AvailableTables[t])
        print(new_sqlite_tables)

        reps_to_import = self.listboxReps.curselection()
        new_reps = []
        print(self.AvailableReps)
        for r in reps_to_import:
            # print self.AvailableReps[r]
            new_reps.append(self.AvailableReps[r])
        print(new_reps)
        conn = sqlite3.connect(tk.Entry(self, textvariable=self.var).get())
        c = conn.cursor()
        conn2 = sqlite3.connect(tk.Entry(self, textvariable=self.var2).get())
        c2 = conn2.cursor()
        if '.sqlite' not in tk.Entry(self, textvariable=self.var).get():
            tmb.showerror("Error", "Please choose an sqlite output file")
            c.close()
            c2.close()
            sys.exit()

        if '.sqlite' not in tk.Entry(self, textvariable=self.var2).get():
            tmb.showerror("Error", "Please choose an sqlite output file")
            c.close()
            c2.close()
            sys.exit()
        self.colNames = []
        self.colType = []

        for t in self.AvailableTables:
            if t in new_sqlite_tables:
                c.execute('PRAGMA TABLE_INFO({})'.format(t))
                all_info = c.fetchall()

                col_info = [tup[1] for tup in all_info]
                col_type = [tup[2] for tup in all_info]

                sql = self.create_sqlite_table(t, col_info, col_type)
                c2.execute(sql)

                # if type(replicationIds)== int:
                #        replicationIds=newReps

                for replicationId in new_reps:

                    sql = self.select_data(t, int(replicationId))
                    # print SQL
                    c.execute(sql)
                    table_data = c.fetchall()
                    print('rows to be inserted for replication ' + str(replicationId) + ': ' + str(len(table_data)))
                    if len(table_data) > 0:
                        sql = self.delete_data(t, replicationId)
                        print(sql)
                        c2.execute(sql)
                        for r in table_data:

                            sql_row_data = self.insert_data(t, r)
                            if sql_row_data is not None:
                                c2.execute(sql_row_data)

        conn2.commit()
        c.close()
        c2.close()
        print('done')

    def create_sqlite_table(self, table_name, column_name_list, column_type_list):

        col = 0
        if len(column_name_list) == len(column_type_list):
            sql = ' CREATE TABLE If NOT EXISTS ' + table_name + "("
            for c in column_name_list:

                sql = sql + c + " " + column_type_list[col]
                if col + 1 != len(column_type_list):
                    sql = sql + ", "
                    col += 1
                else:
                    sql = sql + ");"

        return sql

    def delete_data(self, table_name, replication_id=None):
        print(table_name)
        if replication_id is None:
            sql = ' DELETE FROM ' + table_name
        else:
            sql = 'DELETE FROM ' + table_name + ' WHERE did=' + str(replication_id)

        return sql

    def insert_data(self, table_name, values):
        sql = ""
        if len(values) == 0:
            return None
        v_no = 1
        for v in values:
            if v_no == 1:
                sql = sql + 'INSERT INTO ' + table_name + ' VALUES (' + "'" + str(v) + "'"
            else:
                sql = sql + ', ' + "'" + str(v) + "'"
            v_no = v_no + 1
        sql = sql + ")"
        sql = sql.replace("'None'", "null")

        return sql

    def select_data(self, table_name, replication_id=None, columns='*'):
        sql = 'SELECT ' + columns + ' FROM ' + table_name
        if replication_id is not None:
            sql = sql + ' WHERE  ' + table_name + '.did = ' + str(replication_id)
        return sql


if __name__ == '__main__':
    root = tk.Tk()
    root.title("GTA SQLITE Filter")
    # tableList=['abc', '123']
    TfAimsunSqliteTool(root).pack()
    root.mainloop()
