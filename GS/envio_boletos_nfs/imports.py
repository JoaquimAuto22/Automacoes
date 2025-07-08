from email.mime.multipart import MIMEMultipart
from typing import List, Optional, Dict
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from tkinter import filedialog
from PyPDF2 import PdfReader
from email import encoders
from tqdm import tqdm
from PIL import Image
import tkinter as tk
import pandas as pd
import pytesseract
import numpy as np
import cv2 as cv
import openpyxl 
import smtplib
import shutil
import fitz
import time
import os
import re