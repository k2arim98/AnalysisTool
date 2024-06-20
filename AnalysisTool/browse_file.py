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


def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, filename)