"""
Auteur: Ludovic Malet
Date: juillet 2018
Description: Extrait les pages d'un fichier PDF et les assemble en un seul panneau
"""
import glob
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import subprocess as sp
from PIL import Image as Img
import shutil


class MakeMulti:

    def __init__(self, folder):
        self.folder = folder
        self.files = glob.glob(os.path.join(folder, "*.pdf"))


    def folder_path(self, file):
        file_path, file_ext = os.path.splitext(file)
        return file_path


    def createFolder(self, file_path):
        new_path = self.folder_path(file_path)
        if not os.path.exists(new_path):
            os.mkdir(new_path)

    def extractPDFPages(self):
        """
        Extrait les pages du fichier PDF
        Appelle ensuite la fonction makeMulti
        :param files: les fichiers PDF qui sont le dossier "Files"
        :return:
        """
        for file in self.files:
            self.createFolder(file)
            file_path = self.folder_path(file)
            # file_name = file_path.split("\\")[1]
            file_name = os.path.basename(file_path)
            print("\nCréation du fichier " + str(file_name))

            inputpdf = PdfFileReader(open(file, "rb"))

            for i in range(inputpdf.numPages):
                output = PdfFileWriter()
                output.addPage(inputpdf.getPage(i))

                output_filename = str(file_name) + "-page{:0>2d}.pdf".format(i + 1)
                output_filename_path = os.path.join(file_path, output_filename)
                print("\t Extraction de la page " + output_filename)

                with open(output_filename_path, 'wb') as out:
                    output.write(out)
                self.makeImage(output_filename_path)
            self.makeMulti(file_path)
            # shutil.rmtree(file_path)
            print("\t Suppression des fichiers temporaires")


    def makeImage(self, output_filename):
        """
        Convertit les fichiers PDF en JPG à l'aide du logiciel ImageMagick
        :param output_filename: page du PDF
        :return:
        """
        cmd = os.path.join("data", "ImageMagick", "convert.exe")
        cmd += " -density 150 " #ON DIRAIT QU'IL FAUT CHANGER LA RÉSOLUTION LORSQU'IL Y A PLUS DE 20 IMAGES
        file_name, file_ext = os.path.splitext(output_filename)
        cmd += output_filename + " " + file_name + ".jpg"
        print("\t Conversion de la page en JPG")
        proc = sp.run(cmd, shell=True)


    def makeMulti(self, file_path):
        """
        Assemble les fichiers JPG créés avec la fonction extractPDFPages
        :param file_path: Dossier contenant les fichiers JPG d'un multislide
        :return:
        """
        JPG_Files = glob.glob(os.path.join(file_path, "*.jpg"))
        # file_path2 = file_path.split("\\")
        file_path2 = os.path.basename(file_path)
        images_list = []
        for file_JPG in JPG_Files:
            images_list.append(file_JPG)
        if len(JPG_Files) > 42:
            print("Erreur : le fichier " + file_path2)
            print(" possède " + str(len(JPG_Files)) + " diapos, contacter l'auteur")
        else:
            try:
                imgs = [Img.open(i) for i in images_list]
                (width, height) = imgs[0].size

                x, y = self.switch_taille(len(JPG_Files))

                if 1 < len(JPG_Files) < 21:
                    self.createBigImage(x, y, imgs, width, height, JPG_Files, file_path2)

                elif 20 < len(JPG_Files) < 43:
                    choix = int(input("Le multislide dépasse 20 diapos, voulez vous continuer? oui: 1, non: 2" + "\n"))
                    if choix == 1:
                        self.createBigImage(x, y, imgs, width, height, JPG_Files, file_path2)
                    elif choix == 2:
                        print("Le multislide ne sera pas créé")
                    else:
                        print("Le multislide ne sera pas créé")
            except Exception as e:
                print(e)


    def switch_taille(self, taille):
        """
        Détermine (colonnes x rangées) suivant le nombre de diapos
        :param taille: Correspond au nombre de diapos
        :return: tuple(colonnes x rangées)
        """
        switcher = {
            2: (1, 2),
            3: (2, 2),
            4: (2, 2),
            5: (2, 3),
            6: (2, 3),
            7: (3, 3),
            8: (3, 3),
            9: (3, 3),
            10: (3, 4),
            11: (3, 4),
            12: (3, 4),
            13: (3, 5),
            14: (3, 5),
            15: (3, 5),
            16: (4, 4),
            17: (4, 5),
            18: (4, 5),
            19: (4, 5),
            20: (4, 5),
            21: (4, 6),
            22: (4, 6),
            23: (4, 6),
            24: (4, 6),
            25: (5, 5),
            26: (5, 6),
            27: (5, 6),
            28: (5, 6),
            29: (5, 6),
            30: (5, 6),
            31: (5, 7),
            32: (5, 7),
            33: (5, 7),
            34: (5, 7),
            35: (5, 7),
            36: (6, 6),
            37: (6, 7),
            38: (6, 7),
            39: (6, 7),
            40: (6, 7),
            41: (6, 7),
            42: (6, 7)
        }
        return switcher.get(taille)


    def createBigImage(self, x, y, imgs, width, height, files_in_folder, file_name):
        try:
            new_img = Img.new('RGB', (x * width, y * height), color=self.background_color(imgs[1]))
            i = 0
            for j in range(y):
                for k in range(x):
                    if len(files_in_folder) == i:
                        pass
                    else:
                        new_img.paste(im=imgs[i], box=(k * width, j * height))
                        i += 1
            new_img.save(os.path.join('Multislides', file_name + '.pdf'), format='pdf')
            print("\t Multislide créé")
        except Exception as e:
            print(e)

    def background_color(self, image):
        """
        Retourne la couleur majoritaire dans l'image
        :param image:
        :return: valeur_max
        """
        dico = {}
        for value in image.getdata():
            if value in dico.keys():
                dico[value] += 1
            else:
                dico[value] = 1

        valeur_max = max(dico, key=dico.get)
        return valeur_max
