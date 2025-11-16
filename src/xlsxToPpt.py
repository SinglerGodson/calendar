import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os


def excel_to_ppt(excel_file = './2.xlsx', ppt_file = './2.ppt', sheet_name=0, column_name=None, column_index=0):
   print('test ')