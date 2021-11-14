import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd

from datetime import date, timedelta

import csv


def open_hha_file():
    filetypes = (
        ('csv files', '*.csv'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open HHA Invoice .csv file',
        initialdir='/',
        filetypes=filetypes)

    print("Loading HHA .csv file: %r" % filename)

    with open(filename, 'r') as csvfile:
        reader = csv.DictReader(csvfile)

        headings = reader.fieldnames
        data = []
        for row in reader:
            data.append(row)

    clear_and_populate_tree_from_list(input_tree, headings, data)

    processed_headings, processed_data = convert_hha_to_quickbooks(headings, data)
    clear_and_populate_tree_from_list(output_tree, processed_headings, processed_data)

    global input_filename
    input_filename = filename


def save_quickbooks_file(headings, data):
    filename = fd.asksaveasfilename(
        title='Save Quickbooks Invoice .csv file',
        initialdir='/',
        defaultextension='.csv')

    print("Saving Quickbooks .csv file: %r" % filename)

    with open(filename, 'w') as csvfile:
        writer = csv.DictWriter(csvfile, headings)
        writer.writeheader()
        for row in data:
            writer.writerow(row)

    global output_filename
    output_filename = filename


def create_table(master):
    table_margin = Frame(master, width=width)
    table_margin.pack(side=TOP, fill=BOTH, expand=1)
    scrollbar_x = Scrollbar(table_margin, orient=HORIZONTAL)
    scrollbar_y = Scrollbar(table_margin, orient=VERTICAL)

    tree = ttk.Treeview(table_margin, selectmode="extended", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    scrollbar_y.config(command=tree.yview)
    scrollbar_y.pack(side=RIGHT, fill=Y)
    scrollbar_x.config(command=tree.xview)
    scrollbar_x.pack(side=BOTTOM, fill=X)
    tree.pack(fill=BOTH, expand=1)

    return table_margin, scrollbar_x, scrollbar_y, tree


def clear_and_populate_tree_from_list(tree: ttk.Treeview, headings, data):
    tree.delete(*tree.get_children())

    tree['columns'] = headings

    tree.column('#0', stretch=NO, minwidth=0, width=0)

    for i, heading in enumerate(headings):
        tree.heading(heading, text=heading, anchor=W)
        tree.column('#%d' % (i+1), stretch=NO, minwidth=0, width=200)

    for row in data:
        tree.insert('', 0, values=list(row.values()))

    root.update()

def convert_hha_to_quickbooks(headings, data):
    new_headings = [
        "InvoiceNo",
        "Customer",
        "InvoiceDate",
        "DueDate",
        "Terms",
        "ItemAmount",
        "ExternalInvoiceId",
        "Service Date",
        "ItemDescription"
    ]

    new_data = []

    for row in data:
        new_data.append({
            "InvoiceNo": data["hha_invoice_id"],
            "Customer": "EverCare Choice, Inc.",
            "InvoiceDate": date.fromisoformat(data["invoice_date"]).strftime("%m/%d/%Y"),
            "DueDate": (date.fromisoformat(data["invoice_date"]) + timedelta(days=30)).strftime("%m/%d/%Y"),
            "Terms": "Net 30",
            "ItemAmount": data["total_amount"],
            "ExternalInvoiceId": data["hha_invoice_id"],
            "Service Date": data["visit_date"],
            "ItemDescription": "CDPAS"
        })


# construct GUI
root = tk.Tk()
root.title('Preprocess HHA Invoice .csv for QuickBooks Import')

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

width = screen_width / 2
height = screen_height / 2

x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)
root.geometry("%dx%d+%d+%d" % (width, height, x, y))

ActionBar = Frame(root, width=width)
ActionBar.pack(side=TOP, fill=X)

input_filename = ''
output_filename = ''

open_button = ttk.Button(
    ActionBar,
    text='Open HHA Invoice .csv file',
    command=open_hha_file
)

open_button.pack(side=LEFT, expand=True)

save_button = ttk.Button(
    ActionBar,
    text='Save Quickbooks Invoice .csv file',
    command=save_quickbooks_file
)

save_button.pack(side=LEFT, expand=True)

input_table_margin, input_scrollbar_x, input_scrollbar_y, input_tree = create_table(root)
output_table_margin, output_scrollbar_x, output_scrollbar_y, output_tree = create_table(root)



# run the application
root.mainloop()

