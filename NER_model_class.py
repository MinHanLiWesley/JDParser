from os import linesep
import spacy
import re
from ner_train import test_model
import tika
from tika import parser
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import PurePath, PurePosixPath





