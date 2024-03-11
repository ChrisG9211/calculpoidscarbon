import plotly.graph_objs as go
from plotly.subplots import make_subplots
import urllib.request as ur
import json
import pickle
from collections import defaultdict, OrderedDict
import copy
import clr
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import BuiltInParameter, FilteredElementCollector
import datetime
import openpyxl
from pypac import get_pac, PACSession