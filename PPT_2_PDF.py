"""
Auteur: Ludovic Malet
Date: juillet 2018
Description: Permet de convertir les fichiers PPT et PPTX en PDF
"""

import win32com.client
import glob
import os

class PPT2PDF:

    def __init__(self, folder):
        self.folder = folder
        self.files = glob.glob(os.path.join(folder, "*.ppt*"))
        self.formatType = 32

    def convert(self):
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = 1
        for file in self.files:
            filename = os.getcwd() + '\\' + file
            newname = os.path.splitext(filename)[0] + ".pdf"
            try:
                deck = powerpoint.Presentations.Open(filename)
                deck.SaveAs(newname, self.formatType)
                deck.Close()
            except Exception as inst:
                print(type(inst))
                print(inst.args)
                print(inst)
        powerpoint.Quit()