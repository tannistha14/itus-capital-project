# example.py — small wrapper so xlwings can import the module named after the workbook
import os, sys
# ensure folder is on sys.path (usually not needed but harmless)
sys.path.insert(0, os.path.dirname(__file__))

from daily_data_udf import *
