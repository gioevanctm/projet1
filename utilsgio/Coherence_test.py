from tkinter import *
from tkinter import tix
import os.path
from tkinter.filedialog import *
from tkinter.messagebox import *
import os
import shapefile as sf
from shapely.geometry import Polygon
#from interface.Interface_Coherence_test import interface_coherence

class coherence:
    # Création de la Classe
    def __init__(self): #initialisation des attributs
        self.shp_zsro = None
        self.shp_pfsro = None
        self.shp_zpbo = None
        self.shp_pfpbo = None
        self.shp_piq = None
        self.shp_error = None
        self.nb_champ = 0


    def loadShpfile1(self, path1):
        listefield =[]
        self.shp_zsro = sf.Reader(path1)
        for field in self.shp_zsro.fields:
            listefield.append(field[0])
        return listefield[1:]


    def loadShpfile2(self, path2):
        listefield =[]
        self.shp_pfsro = sf.Reader(path2)
        for field in self.shp_pfsro.fields:
            listefield.append(field[0])
        return listefield[1:]

    def loadShpfile3(self, path3):
        listefield =[]
        self.shp_zpbo = sf.Reader(path3)
        for field in self.shp_zpbo.fields:
            listefield.append(field[0])
        return listefield[1:]

    def loadShpfile4(self, path4):
        listefield =[]
        self.shp_pfpbo = sf.Reader(path4)
        for field in self.shp_pfpbo.fields:
            listefield.append(field[0])
        return listefield[1:]

    def loadShpfile5(self, path5):
        listefield =[]
        self.shp_piq = sf.Reader(path5)
        for field in self.shp_piq.fields:
            listefield.append(field[0])
        return listefield[1:]

    # AFFICHAGE DU CONTENU DE LA 1ER COLONNE
    def loadShp(self, path1, typefichier):
        if typefichier == "zsro":
            liste_rec = []
            self.shp_zsro = sf.Reader(path1)
            for record in self.shp_zsro.records():
                liste_rec.append(record[self.nb_champ])
            return liste_rec[:]


    def loadShp2(self, path2, typefichier):

        liste_rec = []
        self.shp_pfsro = sf.Reader(path2)
        for record in self.shp_pfsro.records():
                liste_rec.append(record[self.nb_champ])
        return liste_rec[:]






