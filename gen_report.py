import pandas as pd
from pandas.plotting import table
from matplotlib import pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT

document = Document()
document.add_heading('Sample Report', 0)
