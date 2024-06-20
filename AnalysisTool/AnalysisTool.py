import tkinter as tk
from tkinter import Listbox, END
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from functools import partial
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import calendar
import threading
import time
from functools import partial
import os
from browse_file import browse_file
from process import process_data
from Purchase_Analysis import Purchase_Analysis

processing = False

def main():
    def on_closing():
        if tk.messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()

    def update_process_button_command(event=None):
        selected_option = user_choice_combobox2.get()
        if selected_option == 'Purchase_Analysis':
            process_button.config(command=partial(Purchase_Analysis, root, input_file_entry,process_button, progress_bar))
        else:
            process_button.config(command=partial(process_data, root, input_file_entry, user_choice_combobox, user_choice_combobox2, process_button, progress_bar))

    root = tk.Tk()
    root.title("Analysis Tool")
    root.protocol("WM_DELETE_WINDOW", on_closing)

    input_file_label = tk.Label(root, text="Select Excel File:")
    input_file_label.grid(row=1, column=0, padx=10, pady=5)

    input_file_entry = tk.Entry(root, width=50)
    input_file_entry.grid(row=1, column=1, padx=10, pady=5)

    browse_button = tk.Button(root, text="Browse", command=partial(browse_file, input_file_entry))
    browse_button.grid(row=1, column=2, padx=10, pady=5)

    options = ['day', 'month']
    user_choice_combobox = ttk.Combobox(root, values=options, state="readonly")
    user_choice_combobox.grid(row=2, column=1, padx=10, pady=5)
    user_choice_combobox.current(0)

    selected_option = ['STD VS Réel', 'Valorisation', 'Purchase_Analysis']
    user_choice_combobox2 = ttk.Combobox(root, values=selected_option, state="readonly")
    user_choice_combobox2.grid(row=3, column=1, padx=10, pady=5)
    user_choice_combobox2.current(0)
    user_choice_combobox2.bind('<<ComboboxSelected>>', update_process_button_command)

    process_button = tk.Button(root, text="Process Data")
    process_button.grid(row=4, column=1, padx=10, pady=10)

    progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
    progress_bar.grid(row=5, column=1, padx=10, pady=5)

    update_process_button_command()  # Initial setup based on the default selection

    root.mainloop()

if __name__ == "__main__":
    main()

