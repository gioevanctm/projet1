from tkinter import *
import os.path
import os
import xlrd
import xlsxwriter
from datetime import date


import shapefile as sf
from shapely.geometry import Polygon, MultiPolygon, Point

class coherence:
    # Création de la Classe
    def __init__(self): #initialisation des attributs
        self.shp_zsro = None    #initialisation des variables
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
        self.list_log = []
        self.compt_log = 0
        self.compt_log_total = 0
        self.compt_total = 0
        self.type_log = 0

        self.h = 0
        self.f = 0

        self.poly = None
        self.list_cord_poly = []
        self.Multi_poly = None
        self.list_pfsro = None
        self.list_pfpbo = None
        self.list_piq = None

        self.ZSROxSRO = []
        self.SROxZSRO = []
        self.ZPBOxPBO = []
        self.PBOxZPBO = []
        self.PIQxZPBO = []

        self.erreur_zsro = 0
        self.erreur_pfsro = 0
        self.erreur_zpbo = 0
        self.erreur_pfpbo = 0
        self.erreur_log_hors_zone = 0
        self.erreur_logement = 0
        self.erreur_logement_tot = 0
        self.list_err_log = []

        self.erreurs_totales = 0
        self.err = ""

        self.worksheet = None

        self.path_xls = "C:\\Users\\g.battery\\Desktop\\rapport\\Rapport_QGIS.xlsx"
        self.workbook = xlsxwriter.Workbook(self.path_xls)

        self.regular_text = self.workbook.add_format({'font_size': 10})
        self.title = self.workbook.add_format({'bold': True, 'font_size': 11, 'text_wrap': True, 'align': 'center'})
        self.error = self.workbook.add_format({'font_size': 10, 'shrink': True})
        self.total =self.workbook.add_format({'bold': True, 'font_size': 10, 'border': True})
        self.red_error = self.workbook.add_format({'font_color': 'red', 'font_size': 10, 'shrink': True})
        self.orange_error = self.workbook.add_format({'font_color': 'orange', 'font_size': 10, 'shrink': True})
        self.violet_error = self.workbook.add_format({'font_color': 'violet', 'font_size': 10, 'shrink': True})
        self.red_error_total = self.workbook.add_format({'font_color': 'red', 'font_size': 10, 'bold': True,'border': True, 'bg_color': '#ffc7ce'})
        self.green = self.workbook.add_format({'font_color': 'green', 'font_size': 10})
        self.orange_error_back = self.workbook.add_format({'font_color': 'orange', 'font_size': 10, 'bold': True, 'border': True, 'bg_color': '#fcd5b4'})
        self.red_error_back = self.workbook.add_format({'font_color': 'red', 'font_size': 10, 'border': True, 'bold': True, 'bg_color': '#ffc7ce'})

        # Fonction création de l'entête de chaque feuille du classeur Excel. La première ligne est remplie avec les titres suivants :
        def entete(wkb, name, col_title):
            worksheet = wkb.add_worksheet(name)
            worksheet.write(0, 0, col_title, self.title)
            return worksheet


        self.worksheet_SRO = entete(self.workbook, "Rapport_SRO", "Cohérence des données Shape : SRO/PIQ_BAL")
        self.title_ = ("SRO", "ZSROxSRO", "SROxZSRO", "Lgt Total", "Lgt Total données attributaires", "Différence", "Erreur\nPiq_Bal")  #liste entete Page 1

        self.worksheet_PBO = entete(self.workbook, "Rapport PBOxZPBO_Bal", "Cohérence des données Shape : PBO/ZPBO")
        self._title_ = ("ZPBOxPBO", "PBOxZPBO")                                                                                         #liste entete page 2

        self.worksheet_PIQ = entete(self.workbook, "Rapport ZPBOxPIQ_Bal", "Cohérence des données Shape : ZPBO/PIQ_BAL")
        self._title_1 = ("PBO", "PIQxPBO", "Lgt données attributaires", "Différences", "Type Logement","µModule\nZPBO", "Erreur\nPiq_Bal", "Erreur différence")                                          #liste entete Page 3

        self.worksheet_SRO.write_row(0, 0, self.title_, self.title)     #écrire toutes les entetes
        self.worksheet_PBO.write_row(0, 0, self._title_, self.title)    #écrire toutes les entetes
        self.worksheet_PIQ.write_row(0, 0, self._title_1, self.title)   #écrire toutes les entetes
        self.worksheet_SRO.freeze_panes(1, 0)       #Freeze la premier ligne
        self.worksheet_PBO.freeze_panes(1, 0)       #Freeze la premier ligne
        self.worksheet_PIQ.freeze_panes(1, 0)       #Freeze la premier ligne

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
        self.erreur_log_hors_zone = 0
        self.compt_log = 0
        self.compt_log_total = 0
        self.compt_total = 0
        self.type_log = 0
        self.h = 0
        self.f = 0
        self.list_err_log = 0
        # ---------------------------------------------------------------------------------
        print(" ------------------------Vérification des id métier des ZSRO et des SRO------------------------ ")
        for a in set1:
            if a in set2:
                self.ZSROxSRO.append("la ZSRO " + a[14:] + " a un SRO associé")
                #print("la ZSRO", a, "a un SRO associé")
                ue = 1

            else:
                print("la ZSRO", a, "n'a pas de SRO associé")
                self.ZSROxSRO.append("la ZSRO " + a[14:] + " n'a pas de SRO associé")
                self.erreur_zsro += 1
        for b in set2:
            if b in set1:
                self.SROxZSRO.append("le SRO " + b[14:] + " a une ZSRO associé")
                print("le SRO", b, "a une ZSRO associé")
                ue = 1
            else:
                self.SROxZSRO.append("le SRO " + b[14:] + " n'as pas de ZSRO associé")
                print("le SRO", b, "n'as pas de ZSRO associé")
                self.erreur_pfsro += 1

        print(" ------------------------Vérification des id métier des ZPBO et des PBO------------------------ ")

        for a in set3:
            if a in set4:
                #self.ZPBOxPBO.append("la ZPBO " + a[14:] + " a un PBO associé")
                #print("la ZPBO", a, "a un PBO associé")
                le = 1
            else:
                self.ZPBOxPBO.append("la ZPBO " + a[14:] + " n'as pas de PBO associé")
                print("la ZPBO", a[13:], "n'as pas de PBO associé")
                self.erreur_zpbo += 1
        for b in set4:
            if b in set3:
                #self.PBOxZPBO.append("le PBO " + b[14:] + " a une ZPBO associé")
                #print("le PBO", b[13:], "a une ZPBO associé")
                le = 1
            else:
                self.PBOxZPBO.append("le PBO " + b[13:] + " n'as pas de ZPBO associé")
                print("le PBO", b[13:], "n'as pas de ZPBO associé")
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
            self.list_log.append(self.compt_log)
            if self.compt_log < 300 or self.compt_log > 400:
                self.erreur_logement += 1
            self.compt_log_total += self.compt_log
            self.compt_log = 0
            print("--------------------------------------------------------------------------------------------------------")



            row = (os.path.basename(self.shp_zsro.record(i)['id_metier_']), self.ZSROxSRO[i], self.SROxZSRO[i],
                   self.list_log[i], self.shp_zsro.record(i)['nb_el'])
            self.worksheet_SRO.set_column(i + 1, 0, 26)
            self.worksheet_SRO.write_row(i + 1, 0, row, self.regular_text)


            self.worksheet_SRO.write_formula(i + 1, 5, "=D" + str(i + 2) + "-E" + str(i + 2), self.red_error)
            self.worksheet_SRO.write_formula(i + 1, 6, "=IF(OR(D" + str(i + 2) + ">449,D" + str(i + 2) + "<300,F" + str(i + 2) + "<>0),1,0)", self.red_error)


            # Write a conditional format over a range donnée compté.
            self.worksheet_SRO.conditional_format('D' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'not between',
                                                    'minimum': 300,
                                                    'maximum': 450,
                                                    'format': self.red_error_back})

            self.worksheet_SRO.conditional_format('D' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'between',
                                                    'minimum': 400,
                                                    'maximum': 450,
                                                    'format': self.orange_error_back})

            # Write a conditional format over a range attributaire.
            self.worksheet_SRO.conditional_format('E' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'not between',
                                                    'minimum': 300,
                                                    'maximum': 450,
                                                    'format': self.red_error_back})

            self.worksheet_SRO.conditional_format('E' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'between',
                                                    'minimum': 400,
                                                    'maximum': 450,
                                                    'format': self.orange_error_back})


            # Write a conditional format over a range.
            self.worksheet_SRO.conditional_format('F' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'not equal to',
                                                    'value': 0,
                                                    'format': self.red_error_back})

            self.worksheet_SRO.conditional_format('G' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'not equal to',
                                                    'value': 0,
                                                    'format': self.red_error_back})


            self.h += 1

        self.worksheet_SRO.write(self.h + 1, 0, 'Total', self.total)
        self.worksheet_SRO.write_formula((self.h + 1), 3, "=SUM(D2:D" + str(self.h + 1) + ")", self.total)
        self.worksheet_SRO.write_formula((self.h + 1), 4, "=SUM(E2:E" + str(self.h + 1) + ")", self.total)
        self.worksheet_SRO.write_formula((self.h + 1), 5, "=SUM(F2:F" + str(self.h + 1) + ")", self.total)
        self.worksheet_SRO.write_formula((self.h + 1), 6, "=SUM(G2:G" + str(self.h + 1) + ")", self.red_error_total)

        for j in range(len(self.shp_piq)):
            self.compt_total += self.nb_logement_piq[j]
        self.compt_total = self.compt_total-self.compt_log_total

        print("il y a au total {} logements dans la ZSRO et {} logement hors Zones".format(self.compt_log_total, self.compt_total))


        print(" ------------------------Vérification des du nombre de logements dans les ZPBO------------------------ ")

        self.compt_log = 0
        self.compt_log_total = 0


        for i in range(len(self.shp_zpbo.shapes())):
            self.poly = Polygon(self.shp_zpbo.shape(i).points)

            for j in range(len(self.shp_piq)):
                self.list_piq = self.shp_piq.shape(j).points
                if Point(*self.list_piq).within(self.poly):
                    #print("Le logement {} est dans le shape {}".format(self.id_metier_piq[j], self.id_metier_zsro[i]))
                    self.compt_log += self.nb_logement_piq[j]

                    if (self.shp_zpbo.record(i)['nb_module']) == 0:
                        self.type_log = -1 #Anomalie
                        print("erreur Aucun module indiquer")

                    if (self.shp_zpbo.record(i)['nb_module']) == 1:
                        self.type_log = 1
                        print("C'est une immeuble avec 1 modules")

                    if (self.shp_zpbo.record(i)['nb_module']) == 2:
                        self.type_log = 2 #immeuble
                        print("C'est une immeuble avec 2 modules")

                    if (self.shp_zpbo.record(i)['nb_module']) >= 3:
                        self.type_log = 3 #immeuble
                        print("C'est une immeuble avec 3 modules")

            print("il y a {} logement dans le shape{}".format(self.compt_log, self.id_metier_zpbo[i]))

            if self.type_log == -1:   #anomalie
                self.list_err_log = 1

            if self.type_log == 1:
                if self.compt_log > 5:
                    self.erreur_logement += 1
                    self.list_err_log = 1
                    print("il y a plus de 6 logements - erreur")
                else:
                    self.list_err_log = 0

            if self.type_log == 2:
                if self.compt_log > 10:
                    self.list_err_log = 1
                else:
                    self.list_err_log = 0

            if self.type_log > 3:
                if self.compt_log > 16:
                    self.list_err_log = 1
                else:
                    self.list_err_log = 0
            self.f += 1



            self.worksheet_PIQ.set_column(i + 1, 0, 15)
            row2 = (os.path.basename(self.id_metier_zpbo[i]), self.compt_log, self.shp_zpbo.record(i)['nb_el'], "", self.type_log, self.shp_zpbo.record(i)['nb_module'])
            self.worksheet_PIQ.write_row(i + 1, 0, row2, self.regular_text)

            self.worksheet_PIQ.write_formula(i + 1, 3, "=B" + str(i + 2) + "-C" + str(i + 2), self.red_error)

            self.worksheet_PIQ.write_formula(i + 1, 7, "=IF(D" + str(i + 2) +"<>0,1,)", self.red_error)
            self.worksheet_PIQ.write(i + 1, 6, self.list_err_log, self.red_error)
            self.list_err_log = 0
            self.type_log = 0

            # Write a conditional format over a range.
            self.worksheet_PIQ.conditional_format('G' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'equal to',
                                                    'value': 1,
                                                    'format': self.red_error_back})

            # Write a conditional format over a range.
            self.worksheet_PIQ.conditional_format('H' + str(i + 2), {'type': 'cell',
                                                    'criteria': 'not equal to',
                                                    'value': 0,
                                                    'format': self.red_error_back})

            self.worksheet_PBO.set_column(i + 1, 0, 26)
            print("--------------------------------------------------------------------------------------------------------")
            self.compt_log_total += self.compt_log
            self.compt_log = 0

        self.worksheet_PIQ.write(self.f + 1, 0, 'Total', self.total)
        self.worksheet_PIQ.write(0, 8, "Le type de logement:", self.regular_text)
        self.worksheet_PIQ.write(1, 8,"-1:il n'y a pas de module(anomalie)",self.regular_text)
        self.worksheet_PIQ.write(2, 8,"1:il n'y a qu'un seul module",self.regular_text)
        self.worksheet_PIQ.write(3, 8,"2:il y a 2 modules", self.regular_text)
        self.worksheet_PIQ.write(4, 8,"3:il y a 3 ou plusieurs modules",self.regular_text)

        self.worksheet_PIQ.write_formula(self.f + 1, 1, "=SUM(B2:B" + str(self.f + 1) + ")", self.total)
        self.worksheet_PIQ.write_formula(self.f + 1, 2, "=SUM(C2:C" + str(self.f + 1) + ")", self.total)
        self.worksheet_PIQ.write_formula(self.f + 1, 3, "=SUM(D2:D" + str(self.f + 1) + ")", self.red_error_total)
        self.worksheet_PIQ.write_formula(self.f + 1, 5, "=SUM(F2:F" + str(self.f + 1) + ")", self.total)
        self.worksheet_PIQ.write_formula(self.f + 1, 6, "=SUM(G2:G" + str(self.f + 1) + ",H2:H" + str(self.f + 1) + ")", self.red_error_total)

        for i in range(len(self.ZPBOxPBO)):
            self.worksheet_PBO.write(i + 1, 0, self.ZPBOxPBO[i], self.red_error)
        for i in range(len(self.PBOxZPBO)):
            self.worksheet_PBO.write(i + 1, 1, self.PBOxZPBO[i], self.red_error)

        self.worksheet_PBO.write(1, 2,"Il y a " + str(self.erreur_zpbo) + " erreurs au niveau des Zones PBO, " + str(self.erreur_pfpbo) + " erreurs au niveau des PBO.", self.regular_text)



        for j in range(len(self.shp_piq)):
            self.compt_total += self.nb_logement_piq[j]
        self.erreur_log_hors_zone = self.compt_total - self.compt_log_total
        print("Il y a au total {} logements dont, {} dans les ZSRO et {} logement hors Zones".format(self.compt_total, self.compt_log_total, self.erreur_log_hors_zone))

        self.erreurs_totales = self.erreur_zsro + self.erreur_pfsro + self.erreur_zpbo + self.erreur_pfpbo + self.erreur_logement
        print("Il y a {} erreurs au niveau des Zones SRO, {} erreurs au niveau des Emplacement des SRO,\n {} erreurs au niveau des Zones PBO, {} erreurs au niveau des PBO.\n Il y {} au niveau des logements.\n Il y a donc {} erreurs au totale.".format(self.erreur_zsro, self.erreur_pfsro, self.erreur_zpbo, self.erreur_pfpbo, self.erreur_logement, self.erreurs_totales))

        self.workbook.close()

















