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

processing = False

def categorize_date(date, filter_by):
    if filter_by == "day":
        return "Decade1" if date.day <= 10 else ("Decade3" if date.day > 20 else "Decade2")
    elif filter_by == "month":
        month_name = calendar.month_name[date.month]
        return f"{month_name.capitalize()}"
    else:
        return "Invalid filter choice"
