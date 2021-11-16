import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd

from datetime import date, timedelta
import re
import functools
import os

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
            # hha puts a single quote in the total amount field for some reason ¯\_(ツ)_/¯
            row["total_amount"] = re.sub(r'[^\d|.]', '', row["total_amount"])
            data.append(row)
            # print(row["invoice_id"] + ' ' + row['service_date'])

    clear_and_populate_tree_from_list(input_tree, headings, data)

    processed_headings, processed_data = convert_hha_to_quickbooks(headings, data)
    clear_and_populate_tree_from_list(output_tree, processed_headings, processed_data)

    save_button_widget = ActionBar.children.get("save_button")

    if save_button_widget:
        save_button_widget.command = functools.partial(save_quickbooks_file, processed_headings, processed_data)
    else:
        save_button = ttk.Button(
            ActionBar,
            text='Save Quickbooks Invoice .csv file',
            command=functools.partial(save_quickbooks_file, processed_headings, processed_data),
            name="save_button"
        )

        save_button.pack(side=LEFT, expand=True)

    root.update()

    global output_filename
    output_filename = filename


def save_quickbooks_file(headings, data):
    save_directory = os.path.dirname(input_filepath)
    input_filename = os.path.basename(input_filepath)
    default_output_filename = "qb_import_from_hha_invoices_" + data[0]["InvoiceNo"] + '-' + data[-1]["InvoiceNo"]

    filename = fd.asksaveasfilename(
        title='Save Quickbooks Invoice .csv file',
        initialdir=save_directory,
        initialfile=default_output_filename,
        defaultextension='.csv'
    )

    print("Saving Quickbooks .csv file: %r" % filename)

    with open(filename, 'w', newline='') as csvfile:
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
    root.update()

    tree['columns'] = headings

    tree.column('#0', stretch=NO, minwidth=0, width=0)

    for i, heading in enumerate(headings):
        tree.heading(heading, text=heading, anchor=W)
        tree.column('#%d' % (i+1), stretch=NO, minwidth=0, width=200)

    for row in data:
        tree.insert('', END, values=list(row.values()))

    root.update()

def convert_hha_to_quickbooks(headings, data):
    new_headings = [
        "InvoiceNo",
        "Customer",
        "Service Date",
        "InvoiceDate",
        "DueDate",
        "Terms",
        "ItemAmount",
        "ExternalInvoiceId",
        "ItemDescription"
    ]

    new_data = []

    for row in data:
        # check to make sure we only have evercare invoices
        # # this would need to be extended when we want to support multiple payers
        if int(row["contract_id"]) == 45210:
            invoice_date_re = re.search(r'(\d+)/(\d+)/(\d+)', row["invoice_date"])
            from pprint import pprint
            pprint(row["invoice_id"] + " " + row["service_date"])
            invoice_month = invoice_date_re.group(1)
            invoice_day = invoice_date_re.group(2)
            invoice_year = invoice_date_re.group(3)
            invoice_date = date(int(invoice_year), int(invoice_month), int(invoice_day))

            service_date_re = re.search(r'(\d*)/(\d*)/(\d*)', row["service_date"])
            service_month = service_date_re.group(1)
            service_day = service_date_re.group(2)
            service_year = service_date_re.group(3)
            service_date = date(int(service_year), int(service_month), int(service_day))

            new_data.append({
                "InvoiceNo": row["invoice_id"],
                "Customer": "EverCare Choice, Inc.",
                "Service Date": service_date.strftime("%m/%d/%Y"),
                "InvoiceDate": invoice_date.strftime("%m/%d/%Y"),
                "DueDate": (invoice_date + timedelta(days=30)).strftime("%m/%d/%Y"),
                "Terms": "Net 30",
                "ItemAmount": row["total_amount"],
                "ExternalInvoiceId": row["invoice_id"],
                "ItemDescription": "CDPAS"
            })

    return new_headings, new_data


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

input_filepath = ''
output_filepath = ''

open_button = ttk.Button(
    ActionBar,
    text='Open HHA Invoice .csv file',
    command=open_hha_file,
    name="open_button"
)

open_button.pack(side=LEFT, expand=True)

input_table_margin, input_scrollbar_x, input_scrollbar_y, input_tree = create_table(root)
output_table_margin, output_scrollbar_x, output_scrollbar_y, output_tree = create_table(root)

# run the application
root.mainloop()

