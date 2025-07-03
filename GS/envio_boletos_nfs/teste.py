import fitz
import os
from PIL import Image
import pytesseract
import re

# Caminho para o Tesseract instalado (ajuste se necessário)


# Regex para encontrar CNPJs com ou sem pontuação
CNPJ_REGEX = r'\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}'

def extrair_cnpjs_do_pdf(path_pdf: str, page: int = 0):
    try:
        # Converter PDF em imagem
        doc = fitz.open(path_pdf)
        pagina = doc.load_page(page)
        pix = pagina.get_pixmap(dpi=300)
        temp_img_path = 'temp_cnpj.jpg'
        pix.save(temp_img_path)
        doc.close()

        # Ler imagem e aplicar OCR
        imagem = Image.open(temp_img_path)
        texto = pytesseract.image_to_string(imagem, lang='por')

        # Buscar CNPJs com regex
        cnpjs = re.findall(CNPJ_REGEX, texto)
        cnpjs = [re.sub(r'\D', '', c) for c in cnpjs]  # limpar pontuação

        # Remover duplicatas mantendo ordem
        vistos = set()
        cnpjs_unicos = []
        for c in cnpjs:
            if c not in vistos:
                vistos.add(c)
                cnpjs_unicos.append(c)

        os.remove(temp_img_path)

        # Mostrar resultados
        if cnpjs_unicos:
            print("CNPJs encontrados:")
            for cnpj in cnpjs_unicos:
                print("-", cnpj)
            return cnpjs_unicos
        else:
            print("Nenhum CNPJ encontrado.")
            return []

    except Exception as e:
        print(f"Erro ao processar PDF: {e}")
        return []

# Caminho do seu arquivo PDF
pdf_path = "GS/envio_boletos_nfs/SINGULAR_FACILITIES_CE/arquivos_organizados/Pastas_Mescladas/nfs sem documento/27-05 ASSOCIACAO RESERVA CAMARA.pdf"
extrair_cnpjs_do_pdf(pdf_path)
