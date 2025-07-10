from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from typing import List, Dict
from tqdm import tqdm
from PIL import Image
from sys import exit
import pytesseract
import cv2 as cv
import fitz
import time
import os
import re  

path = "ECO PARK  NF -20778  SINGULAR.pdf"

def pdf_to_img(path: str, page: int = 0) -> None:
    pdf_document = fitz.open(path)
    page = pdf_document.load_page(page)
    image = page.get_pixmap()
    
    img_path = 'temp_img.jpg'
    image.save(img_path)
    pdf_document.close()

    image = Image.open(img_path)

    cnpj_cliente = image.crop((77, 230, 180, 250))

    cnpj_cliente.save("cnpj.jpg")
   

def extract_text(img_path: str, config: str = '--psm 9') -> str:
    img = cv.imread(img_path)
    if img is None:
        raise FileNotFoundError(f"Imagem não encontrada: {img_path}")
    
    scale_percent = 300
    new_width = int(img.shape[1] * scale_percent / 150)
    new_height = int(img.shape[0] * scale_percent / 150)
    img = cv.resize(img, (new_width, new_height), interpolation=cv.INTER_LANCZOS4)

    img = cv.cvtColor(img, cv.COLOR_BGR2GRAY)

    text = pytesseract.image_to_string(img, config=config)
    return text.strip()

def extrair_somente_numeros(texto: str) -> str:
    """Remove tudo que não for número."""
    return re.sub(r'\D', '', texto)

if __name__ == "__main__":
    try:
        pdf_to_img(path)
        texto_cliente = extract_text('cnpj.jpg')
        cnpj_numerico = extrair_somente_numeros(texto_cliente)

        print("CNPJ ", cnpj_numerico)


    except Exception as e:
        print(f"Erro: {e}")
