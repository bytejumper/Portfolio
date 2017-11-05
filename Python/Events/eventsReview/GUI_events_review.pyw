#! /usr/bin/python3

""" eventsReview_GUI

Graphical User Interface for eventsReview.py
"""

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as mb
import tkinter.filedialog as fd
import os
import numpy as np
import events_review as er


class EventReview(tk.Tk):
    def __init__(self, parent):
        """ Initialize GUI

        Sets options for file selections
        """
        tk.Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
        self.prev_path = ''
        self.reset_selections()

        self.open_opt = options = {}
        options['filetypes'] = [('csv files', '.csv')]
        # if directory doesn't exist, initialdir should use default (current working directory)
        options['initialdir'] = 'U:\\COE Advancement\\Work Requests\\Temp Files'

        self.save_opt = options = {}
        options['filetypes'] = [('All Files', '.*'), ('Excel files', '.xlsx')]
        options['initialdir'] = 'Z:\\03 Events\\Event Review Reports'

    def initialize(self):
        """ Define GUI window frames and widgets.

        Frame 1 contains drop-down for output format (self.drop), and buttons for
        source file selection (source_button and prev_button) and actions
        (reset and okay).  okay only reads the file enough to
        provide summary information in frame 2.

        Frame 2 contains summary information for the file, and a button to initiate
        processing the file (self.save).  All buttons and checkboxes in this frame
        are disabled until a source file has been read.
        """
        self.notebook = ttk.Notebook(self)

        # First tab: shows file selection options
        frame1 = ttk.Frame(self.notebook)

        label = ttk.Label(frame1, text='Type')
        label.grid(column=0, row=0)
        source_button = ttk.Button(frame1, text='Source File', command=self.click_source)
        source_button.grid(column=0, row=1)
        self.prev_button = ttk.Button(frame1, text='Previous File', command=self.click_prev)
        self.prev_button.grid(column=0, row=2)
        self.mgo_button = ttk.Button(frame1, text='MGO File', command=self.click_mgo)
        self.mgo_button.grid(column=0, row=3)

        format_types = ('UIF', 'ENG', 'Event Flag')
        self.drop = ttk.Combobox(frame1, state='readonly', values=format_types)
        self.drop.bind('<<ComboboxSelected>>', self.drop_select)
        self.drop.grid(column=1, row=0)
        lookup = ttk.Label(frame1, text='LookupID')
        lookup.grid(column=2, row=0)
        self.lookup_entry = ttk.Entry(frame1)
        self.lookup_entry.grid(column=3, row=0)
        self.source_file = tk.StringVar()
        label = ttk.Label(frame1, textvariable=self.source_file)
        label.grid(column=1, row=1, columnspan=3)
        self.prev_file = tk.StringVar()
        label = ttk.Label(frame1, textvariable=self.prev_file)
        label.grid(column=1, row=2, columnspan=3)
        self.mgo_file = tk.StringVar()
        label = ttk.Label(frame1, textvariable=self.mgo_file)
        label.grid(column=1, row=3)

        reset = ttk.Button(frame1, text='Clear All', command=self.reset_selections)
        reset.grid(column=2, row=4)

        okay = ttk.Button(frame1, text='GO', command=self.run_script)
        okay.grid(column=3, row=4)

        # Second tab: shows summary information
        frame2 = ttk.Frame(self.notebook)

        label = ttk.Label(frame2, text='# Individuals')
        label.grid(column=0, row=0)
        label = ttk.Label(frame2, text='Individuals w/ email')
        label.grid(column=0, row=1)
        label = ttk.Label(frame2, text='# Households')
        label.grid(column=0, row=2)
        self.rsvps = ttk.Label(frame2, text='New RSVPs')
        self.rsvps.grid(column=0, row=3)
        self.del_source = tk.IntVar()
        self.c1 = ttk.Checkbutton(frame2, text="Delete source file on save",
                                  variable=self.del_source, onvalue=1, offvalue=0)
        self.c1.grid(column=0, row=4, sticky=tk.W)
        self.del_prev = tk.IntVar()
        self.c2 = ttk.Checkbutton(frame2, text="Delete previous file on save",
                                  variable=self.del_prev, onvalue=1, offvalue=0)
        self.c2.grid(column=0, row=5, sticky=tk.W)

        self.indv = tk.IntVar()
        indv = ttk.Label(frame2, textvariable=self.indv)
        indv.grid(column=1, row=0, sticky=tk.W)
        self.eml = tk.IntVar()
        eml = ttk.Label(frame2, textvariable=self.eml)
        eml.grid(column=1, row=1, sticky=tk.W)
        self.hh = tk.IntVar()
        hh = ttk.Label(frame2, textvariable=self.hh)
        hh.grid(column=1, row=2, sticky=tk.W)
        self.new = tk.IntVar()
        self.rsvp_val = ttk.Label(frame2, textvariable=self.new)
        self.rsvp_val.grid(column=1, row=3, sticky=tk.W)
        self.save = ttk.Button(frame2, text='Save', command=self.click_save)
        self.save.grid(column=1, row=4, rowspan=2)

        self.notebook.add(frame1, text='File Selection')
        self.notebook.add(frame2, text='Summary Info')
        self.notebook.pack()
        self.notebook.bind('<<NotebookTabChanged>>', self.tab_change)

    def drop_select(self, event=None):
        """ Event handler for self.drop

        Changes frame widgets depending on which drop-box option is selected.
            prev_button is deactivated when 'Event Flag' is selected.
            prev_button selects MGO list when 'ENG' is selected.
            RSVP summary information in frame 2 is only shown when 'UIF' is selected.
        """
        if self.drop.get() != 'UIF':
            self.prev_button.state(['disabled'])
            self.prev_file.set('')
            self.rsvps.config(foreground=self['bg'])
            self.rsvp_val.config(foreground=self['bg'])
            if self.drop.get() == 'Event Flag':
                self.mgo_button.state(['disabled'])
                self.mgo_file.set('')
        else:
            self.rsvps.config(foreground='black')
            self.rsvp_val.config(foreground='black')

    def click_source(self):
        """ Event handler for source_button

        Opens file selection dialogue.  When a file has been selected
        (self.source_path), shows file name.
        """
        self.source_path = fd.askopenfilename(**self.open_opt)
        # set source_file to just the file name
        self.source_file.set(os.path.basename(self.source_path))

    def click_prev(self):
        """ Event handler for prev_button

        Opens file selection dialogue.  When a file has been selected
        (self.prev_path), shows file name.
        When format_type == 'ENG', changes allowed filetype from csv to txt
        """
        self.prev_path = fd.askopenfilename(**self.open_opt)
        # set prev_file to just the file name
        self.prev_file.set(os.path.basename(self.prev_path))

    def click_mgo(self):
        """ Event handler for mgo_button

        Opens file selection dialogue.  When a file has been selected
        (self.mgo_path), shows file name.
        """
        self.mgo_path = fd.askopenfilename(
            initialdir='Z:\\03 Information Specialist\\Scripts\\Events\\dependencies',
            filetypes=[('txt files', '.txt')])
        # set mgo_path to just the file name
        self.mgo_file.set(os.path.basename(self.mgo_path))

    def tab_change(self, event):
        active = self.notebook.index(self.notebook.select())
        if active == 0:
            self.reset_summary()

    def reset_selections(self):
        """ Event handler for reset

        Clears self.source_file, self.prev_file, self.mgo_file, and self.lookup_entry
        Resets drop-down to show 'UIF'
        Calls drop_select to reset widget states.
        """
        self.drop.set('UIF')
        for i in [self.source_file, self.prev_file, self.mgo_file]:
            i.set('')
        self.lookup_entry.delete(0, 'end')
        self.drop_select()

    def reset_summary(self):
        """ Event handler for self.notebook

        Clears all summary values
        Disables self.c1, self.c2, and self.save
        Disables summary tab
        """
        for i in [self.indv, self.eml, self.hh, self.new]:
            i.set(0)
        for i in [self.c1, self.c2, self.save]:
            i.state(['disabled'])
        self.notebook.tab(1, state='disabled')

    def run_script(self):
        """ Event handler for okay

        Checks for selected file.  If file was selected, activates widgets in
        frame 2.  Reads file and displays summary information.
        """
        if ((self.source_file.get() == '') & (self.mgo_file.get() == '')) | \
                ((self.source_file.get() == '') & (self.drop.get() == 'Event Flag')):
            mb.showerror('File Error', 'You did not select a file.')
        else:
            self.df = er.basic(self.source_path)
            if (len(self.df.columns) != 45) & (self.df.columns[0] != 'HOUSEHOLDLOOKUPID'):
                mb.showerror('Source File Error', 'You did not select a summary output file.')
                self.source_file.set('')

            else:
                self.notebook.tab(1, state='normal')
                self.notebook.select(1)
                self.save.state(['!disabled'])
                self.c1.state(['!disabled'])
                if self.drop.get() == 'UIF':
                    self.c2.state(['!disabled'])

                self.indv.set(len(self.df['PROSPECTLOOKUPID']))
                self.eml.set(sum(self.df['Email']))
                self.hh.set(len(np.unique(self.df['HOUSEHOLDLOOKUPID'])))

                # Compare to previous file, if applicable
                if self.drop.get() == 'UIF':
                    if self.prev_file.get() != '':
                            past = er.basic(self.prev_path)
                            if (len(past.columns) != 46) & (past.columns[0] != 'HOUSEHOLDLOOKUPID'):
                                mb.showerror('Previous File Error', 'You did not select a summary output file.')
                                self.prev_file.set('')
                            else:
                                self.df = self.df.merge(past[['PROSPECTLOOKUPID', 'PROSPECTNAME']],
                                                        how='left', on='PROSPECTLOOKUPID')
                                self.df.rename(columns={'PROSPECTNAME_x': 'PROSPECTNAME',
                                                        'PROSPECTNAME_y': 'Old'}, inplace=True)
                                self.new.set(len(self.df[self.df['Old'].astype(str) == 'nan']))

    def click_save(self):
        """ Event handler for self.save

        Opens directory selection dialogue.  If a directory was
        selected (save), the file is formatted and the output files are saved.
        """
        self.df = er.management(self.df)
        self.df = er.delete(self.df, self.drop.get())
        save = fd.asksaveasfilename(
            initialfile=self.source_file.get()[:-4],
            **self.save_opt)
        try:
            er.format_file(self.df, save, self.drop.get(), self.mgo_path, self.lookup_entry.get())
            if self.del_source.get() == 1:
                os.remove(self.source_path)
            if self.del_prev.get() == 1:
                os.remove(self.prev_path)
            mb.showinfo('Save', 'Completed!')
        except Exception as e:
            mb.showerror('Error', 'Could not format as desired.' +
                         '\n' + str(e))
            pass

if __name__ == '__main__':
    app = EventReview(None)
    app.title('Format Event Files')
    app.mainloop()
