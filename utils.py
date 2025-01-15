from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import pandas as pd
import os
import matplotlib.pyplot as plt
from openpyxl import load_workbook


def couper(image, left, upper, right, lower):
    cropped_image = image.crop((left, upper, right, lower))
    return cropped_image

def mise_dans_excel(image, chiffre_excel, case_excel,dictionnaire_de_renomage,sheet):
    # Chemin vers Tesseract OCR (assurez-vous que Tesseract est installé)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # Extraire le texte de l'image
    extracted_text = pytesseract.image_to_string(image)
    if extracted_text in dictionnaire_de_renomage:
        extracted_text = dictionnaire_de_renomage[extracted_text]
        # Liste des caractères à remplacer
    chars_to_replace = ["'", "‘", "@", "(", ".", ":", "[", "]", ","]

    # Remplacer les caractères indésirables
    for char in chars_to_replace:
        extracted_text = extracted_text.replace(char, "")  

    # Insérer le texte extrait dans la cellule spécifiée
    cell = f"{chiffre_excel}{case_excel}"
    sheet[cell] = extracted_text
    
    
def ameliorer_image(image_path):
    # Charger l'image
    image = Image.open(image_path)
    image = image.convert('L') # noir et blanc
    """image.filter(ImageFilter.MedianFilter(size=4)) # Réduire le bruit"""
    """image.filter(ImageFilter.SHARPEN) # Augmenter la netteté""" #idee d'amélioration
    # Augmenter légèrement le contraste
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2) # 1.0 signifie aucun changement, >1.0 augmente le contraste
    return image

def creer_toutes_les_cases():
    # Listes des pixels des images pour la couper en un tableau 
    lignes = [230, 249, 275, 297, 322, 346, 368, 391, 417, 440, 462, 487, 510, 535, 558, 580, 603, 627, 650, 675, 700, 725, 750, 775, 793, 817, 841, 863, 888, 910, 934, 955, 981, 1004, 1027, 1050, 1100]
    colonnes = [80, 190, 385, 581, 775, 970, 1165, 1360, 1555, 1730]
    i=0
    pixel_case= []
    while i < len(lignes)-1:
        z=0
        while z < len(colonnes)-1:
            pixel_case.append(((lignes[i],colonnes[z]),(lignes[i+1],colonnes[z+1])))
            z+=1
        i+=1
    return pixel_case