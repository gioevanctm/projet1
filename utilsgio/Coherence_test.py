from tkinter import *
from tkinter import tix
import os.path
import os


import shapefile as sf
from shapely.geometry import Polygon, MultiPolygon, Point

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
        self.id_metier_zsro = []
        self.id_metier_pfsro = []
        self.id_metier_zpbo = []
        self.id_metier_pfpbo = []
        self.id_metier_piq = []
        self.nb_logement_piq = []
        self.compt_log = 0
        self.compt_log_total = 0
        self.compt_total = 0
        self.type_log = 0

        self.poly = None
        self.list_cord_poly = []
        self.Multi_poly = None
        self.list_pfsro = None
        self.list_pfpbo = None
        self.list_piq = None

        self.erreur_zsro = 0
        self.erreur_pfsro = 0
        self.erreur_zpbo = 0
        self.erreur_pfpbo = 0
        self.erreur_logement = 0
        self.erreurs_totales = 0



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

    # Récupération des records d'une colonne choisi pour le zsro
    def loadShp(self, path1, typefichier):
        if typefichier == "zsro":
            self.id_metier_zsro.clear() # effacer en cas de reclique
            for record in self.shp_zsro.records():
                self.id_metier_zsro.append(record[self.nb_champ])
            return self.id_metier_zsro[:]

    # Récupération des records d'une colonne choisi pour le pfsro
    def loadShp2(self, path2, typefichier):
        self.id_metier_pfsro.clear() # effacer en cas de reclique
        for record in self.shp_pfsro.records():
            self.id_metier_pfsro.append(record[self.nb_champ])
        return self.id_metier_pfsro[:]

    # Récupération des records d'une colonne choisi pour le zpbo
    def loadShp3(self, path3, typefichier):
        self.id_metier_zpbo.clear()  # effacer en cas de reclique
        for record in self.shp_zpbo.records():
            self.id_metier_zpbo.append(record[self.nb_champ])
        return self.id_metier_zpbo[:]

    # Récupération des records d'une colonne choisi pour le pfpbo
    def loadShp4(self, path4, typefichier):
        self.id_metier_pfpbo.clear()  # effacer en cas de reclique
        for record in self.shp_pfpbo.records():
            self.id_metier_pfpbo.append(record[self.nb_champ])
        return self.id_metier_pfpbo[:]

    # Récupération des records d'une colonne choisi pour le piquetage
    def loadShp5(self, path5):
        self.shp_piq = sf.Reader(path5)
        self.id_metier_piq.clear()  # effacer en cas de reclique
        for record in self.shp_piq.records():
            self.id_metier_piq.append(record[self.nb_champ])
        return self.id_metier_piq[:]

    # Récupération des records d'une colonne choisi pour le piquetage
    def loadShp55(self, path5):
        self.shp_piq = sf.Reader(path5)
        self.nb_logement_piq.clear()  # effacer en cas de reclique
        for record in self.shp_piq.records():
            self.nb_logement_piq.append(record[self.nb_champ])
        return self.nb_logement_piq[:]

    def analyse(self):
        set1 = set(self.id_metier_zsro)
        set2 = set(self.id_metier_pfsro)
        set3 = set(self.id_metier_zpbo)
        set4 = set(self.id_metier_pfpbo)
        self.compt_log = 0
        self.compt_log_total = 0
        self.compt_total = 0
        self.type_log = 0
        # ---------------------------------------------------------------------------------
        print(" ------------------------Vérification des id métier des ZSRO et des SRO------------------------ ")
        for a in set1:
            if a in set2:
                #print("la ZSRO", a, "a un SRO associé")
                ue = 1

            else:
                print("la ZSRO", a, "n'a pas de SRO associé")
                self.erreur_zsro += 1
        for b in set2:
            if b in set1:
                #print("le SRO", b, "a une ZSRO associé")
                ue = 1
            else:
                print("le SRO", b, "n'as pas de ZSRO associé")
                self.erreur_pfsro += 1

        print(" ------------------------Vérification des id métier des ZPBO et des PBO------------------------ ")

        for a in set3:
            if a in set4:
                #print("la ZPBO", a, "a un PBO associé")
                le = 1
            else:
                print("la ZPBO", a, "n'as pas de PBO associé")
                self.erreur_zpbo += 1
        for b in set4:
            if b in set3:
                #print("le PBO", b, "a une ZPBO associé")
                le = 1
            else:
                print("le PBO", b, "n'as pas de ZPBO associé")
                self.erreur_pfpbo += 1

        print(" ------------------------Vérification des du nombre de logements dans les ZSRO------------------------ ")

        for i in range(len(self.shp_zsro.shapes())):
            self.poly = Polygon(self.shp_zsro.shape(i).points)
            for j in range(len(self.shp_piq)):
                self.list_piq = self.shp_piq.shape(j).points
                if Point(*self.list_piq).within(self.poly):
                    #print("Le logement {} est dans le shape {}".format(self.id_metier_piq[j], self.id_metier_zsro[i]))
                    self.compt_log += self.nb_logement_piq[j]
                else:
                    le = 1
                    #print("Ce logement {} ne fait parti d'aucun shape".format(self.id_metier_piq[j]))
            print("il y a {} logement dans le shape{}".format(self.compt_log, self.id_metier_zsro[i]))
            if self.compt_log < 300 or self.compt_log > 400:
                self.erreur_logement += 1
            print("--------------------------------------------------------------------------------------------------------")
            self.compt_log_total += self.compt_log
            self.compt_log = 0
        for j in range(len(self.shp_piq)):
            self.compt_total += self.nb_logement_piq[j]
        self.compt_total = self.compt_total-self.compt_log_total
        print("il y a au total {} logements dans la ZSRO et {} logement hors Zones".format(self.compt_log_total, self.compt_total))

        print(" ------------------------Vérification des du nombre de logements dans les ZPBO------------------------ ")

        self.compt_log = 0
        self.compt_log_total = 0

        #for e in range(len(self.shp_piq)):


        for i in range(len(self.shp_zpbo.shapes())):
            self.poly = Polygon(self.shp_zpbo.shape(i).points)

            for j in range(len(self.shp_piq)):
                self.list_piq = self.shp_piq.shape(j).points
                if Point(*self.list_piq).within(self.poly):
                    #print("Le logement {} est dans le shape {}".format(self.id_metier_piq[j], self.id_metier_zsro[i]))
                    self.compt_log += self.nb_logement_piq[j]
                    if (self.nb_logement_piq[j]) > 4:
                        self.type_log = 1 #immeuble
                        print("C'est une immeuble")


            print("il y a {} logement dans le shape{}".format(self.compt_log, self.id_metier_zpbo[i]))

            if self.type_log == 0:
                if self.compt_log > 5:
                    self.erreur_logement += 1
                    print("il y a plus de 6 logements - erreur")
            self.type_log = 0



            print("--------------------------------------------------------------------------------------------------------")
            self.compt_log_total += self.compt_log
            self.compt_log = 0
        for j in range(len(self.shp_piq)):
            self.compt_total += self.nb_logement_piq[j]
        self.compt_total = self.compt_total - self.compt_log_total
        print("il y a au total {} logements dans la ZSRO et {} logement hors Zones".format(self.compt_log_total,
                                                                                           self.compt_total))

        self.erreurs_totales = self.erreur_zsro + self.erreur_pfsro + self.erreur_zpbo + self.erreur_pfpbo + self.erreur_logement
        print("il y a {} erreurs au niveau des Zones SRO, {} erreurs au niveau des Emplacement des SRO,\n {} erreurs au niveau des Zones PBO, {} erreurs au niveau des PBO.\n Il y {} au niveau des logements.\n Il y a donc {} erreurs au totale.".format(self.erreur_zsro, self.erreur_pfsro, self.erreur_zpbo, self.erreur_pfpbo, self.erreur_logement, self.erreurs_totales))



















