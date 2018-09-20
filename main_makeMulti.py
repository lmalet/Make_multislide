"""
Auteur: Ludovic Malet
Date: juillet 2018
Description: Converti des fichiers PPT, PPTX et PDF en panneau multislide
"""
from Make_multi_class import MakeMulti
from PPT_2_PDF import PPT2PDF
import os
import time

start_time = time.time()

FilesPPT = PPT2PDF("Files")
FilesPPT.convert()
FilesPDF = MakeMulti("Files")
FilesPDF.extractPDFPages()

print("--- %s seconds ---" % (time.time() - start_time))
print("Fin du programme")

os.system("pause")