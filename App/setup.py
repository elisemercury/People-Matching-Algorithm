from distutils.core import setup # Need this to handle modules
import py2exe  # We have to import all modules used in our program
import pandas as pd
import scipy.spatial
import random
import xlsxwriter
from datetime import date
import time
import tkinter as ttk
from tkinter import *
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from ttkthemes import ThemedTk
import textwrap
import algorithm
import os

setup(windows=[{"script":'main.py', "icon_resources":[(1,"icon.png")]}]) # Calls setup function to indicate that we're dealing with a single console application