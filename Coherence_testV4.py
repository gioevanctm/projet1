# -*- coding: latin-1 -*-
from tkinter import *
import os.path
import os
import xlrd
import xlsxwriter
from datetime import date
import shapefile as sf
import shutil
from shapely.geometry import Polygon, MultiPolygon, Point

class coherence:
    # Cr�ation de la Classe
    def __init__(self): #initialisation des attributs
        self.shp_zsro = None    #initialisation des variables
        self.shp_pfsro = None
        self.shp_zpbo = None
        self.shp_pfpbo = None
        self.shp_piq = None

        self.path1 = StringVar()
        self.path2 = StringVar()
        self.path3 = StringVar()
        self.path4 = StringVar()
        self.path5 = StringVar()

        self.path1prj = StringVar()
        self.path2prj = StringVar()
        self.path3prj = StringVar()
        self.path4prj = StringVar()
        self.path5prj = StringVar()

        self.shp_err_zsro = None  # initialisation des variables
        self.shp_err_pfsro = None
        self.shp_err_zpbo = None
        self.shp_err_pfpbo = None
        self.shp_err_piq = None

        self.path_destination = StringVar()

        self.id_metier_zsro = []
        self.id_metier_pfsro = []
        self.id_metier_zpbo = []
        self.id_metier_pfpbo = []
        self.id_metier_piq = []
        self.nb_logement_piq = []

        self.record_zsro = []
        self.record_pfsro = []
        self.record_zpbo = []
        self.record_pfpbo = []
        self.record_piq = []

        self.shape_zsro=[]
        self.shape_pfsro = []
        self.shape_zpbo = []
        self.shape_pfpbo = []
        self.shape_piq = []

        self.list_log = []

        self.compt_log = 0
        self.compt_log_total = 0
        self.compt_total = 0
        self.type_log = 0

        self.h = 0
        self.f = 0

        self.poly = None
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

        self.etat_prog = StringVar()
        self.etat_prog.set("S�lectionner le fichier ZSRO")

    def analyse(self):
        try:
            if str(self.path_destination) == "PY_VAR11" or str(self.path_destination) == "":

                self.shp_err_zsro = sf.Writer("C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_ZE_PMZ", shapeType=5)
                shutil.copy(str(self.path1prj), "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_ZE_PMZ.prj") #r�cup�ration du fichier projection
                self.shp_err_pfsro = sf.Writer("C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_pf_PMZ", shapeType=1)
                shutil.copy(str(self.path2prj), "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_pf_PMZ.prj") #r�cup�ration du fichier projection
                self.shp_err_zpbo = sf.Writer("C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_ZE_PB.shp", shapeType=5)
                shutil.copy(str(self.path3prj), "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_ZE_PB.prj") #r�cup�ration du fichier projection
                self.shp_err_pfpbo = sf.Writer("C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_pf_PB.shp", shapeType=1)
                shutil.copy(str(self.path4prj), "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_pf_PB.prj") #r�cup�ration du fichier projection
                self.shp_err_piq = sf.Writer("C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_logements.shp", shapeType=1)
                shutil.copy(str(self.path5prj), "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\shape_erreurs\\Erreurs_logements.prj") #r�cup�ration du fichier projection
                self.path_xls = "C:\\Users\\" + os.getlogin() + "\\Documents\\Rapport\\Rapport_QGIS" + str( date.today()) + ".xlsx"
                # self.etat_prog.set("Votre rapport est dans mes documents->rapport.")
            else:
                self.shp_err_zsro = sf.Writer(str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_ZE_PMZ", shapeType=5)
                shutil.copy(str(self.path1prj), str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_ZE_PMZ.prj") #r�cup�ration du fichier projection
                self.shp_err_pfsro = sf.Writer(str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_pf_PMZ", shapeType=1)
                shutil.copy(str(self.path2prj), str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_pf_PMZ.prj") #r�cup�ration du fichier projection
                self.shp_err_zpbo = sf.Writer(str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_ZE_PB.shp", shapeType=5)
                shutil.copy(str(self.path3prj), str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_ZE_PB.prj") #r�cup�ration du fichier projection
                self.shp_err_pfpbo = sf.Writer(str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_pf_PB.shp", shapeType=1)
                shutil.copy(str(self.path4prj), str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_pf_PB.prj") #r�cup�ration du fichier projection
                self.shp_err_piq = sf.Writer(str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_logements.shp", shapeType=1)
                shutil.copy(str(self.path5prj), str(self.path_destination) + "\\Rapport\\Shape_erreurs\\Erreurs_logements.prj") #r�cup�ration du fichier projection
                self.path_xls = str(self.path_destination) + "\\Rapport\\Rapport_QGIS" + str(date.today()) + ".xlsx"
                # self.etat_prog.set("Votre rapport est dans la destination choisie.")

            def entete(wkb, name, col_title):
                worksheet = wkb.add_worksheet(name)
                worksheet.write(0, 0, col_title, self.title)
                return worksheet

            self.worksheet = None
            self.workbook = xlsxwriter.Workbook(self.path_xls)

            self.regular_text = self.workbook.add_format({'font_size': 10})
            self.title = self.workbook.add_format({'bold': True, 'font_size': 11, 'text_wrap': True, 'align': 'center'})
            self.error = self.workbook.add_format({'font_size': 10, 'shrink': True})
            self.total = self.workbook.add_format({'bold': True, 'font_size': 10, 'border': True})
            self.red_error = self.workbook.add_format({'font_color': 'red', 'font_size': 10, 'shrink': True})
            self.orange_error = self.workbook.add_format({'font_color': 'orange', 'font_size': 10, 'shrink': True})
            self.violet_error = self.workbook.add_format({'font_color': 'violet', 'font_size': 10, 'shrink': True})
            self.red_error_total = self.workbook.add_format(
                {'font_color': 'red', 'font_size': 10, 'bold': True, 'border': True, 'bg_color': '#ffc7ce'})
            self.green = self.workbook.add_format({'font_color': 'green', 'font_size': 10})
            self.orange_error_back = self.workbook.add_format(
                {'font_color': 'orange', 'font_size': 10, 'bold': True, 'border': True, 'bg_color': '#fcd5b4'})
            self.red_error_back = self.workbook.add_format(
                {'font_color': 'red', 'font_size': 10, 'border': True, 'bold': True, 'bg_color': '#ffc7ce'})

            self.worksheet_SRO = entete(self.workbook, "Rapport_SRO", "Coh�rence des donn�es Shape : SRO/PIQ_BAL")
            self.title_ = ("SRO", "ZSROxSRO", "SROxZSRO", "Lgt Total", "Lgt Total donn�es attributaires", "Diff�rence",
                           "Erreur\nPiq_Bal")  # liste entete Page 1

            self.worksheet_PBO = entete(self.workbook, "Rapport PBOxZPBO_Bal", "Coh�rence des donn�es Shape : PBO/ZPBO")
            self._title_ = ("ZPBOxPBO", "PBOxZPBO")  # liste entete page 2

            self.worksheet_PIQ = entete(self.workbook, "Rapport ZPBOxPIQ_Bal",
                                        "Coh�rence des donn�es Shape : ZPBO/PIQ_BAL")
            self._title_1 = (
            "PBO", "PIQxPBO", "Lgt donn�es attributaires", "Diff�rences", "Type Logement", "�Module\nZPBO",
            "Erreur\nPiq_Bal", "Erreur diff�rence")  # liste entete Page 3

            self.worksheet_SRO.write_row(0, 0, self.title_, self.title)  # �crire toutes les entetes
            self.worksheet_PBO.write_row(0, 0, self._title_, self.title)  # �crire toutes les entetes
            self.worksheet_PIQ.write_row(0, 0, self._title_1, self.title)  # �crire toutes les entetes
            self.worksheet_SRO.freeze_panes(1, 0)  # Freeze la premier ligne
            self.worksheet_PBO.freeze_panes(1, 0)  # Freeze la premier ligne
            self.worksheet_PIQ.freeze_panes(1, 0)  # Freeze la premier ligne

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

            self.shp_err_zsro.fields = self.shp_zsro.fields[1:]  # r�cup�re tous les champs
            self.shp_err_pfsro.fields = self.shp_pfsro.fields[1:]  # r�cup�re tous les champs
            self.shp_err_zpbo.fields = self.shp_zpbo.fields[1:]  # r�cup�re tous les champs
            self.shp_err_pfpbo.fields = self.shp_pfpbo.fields[1:]  # r�cup�re tous les champs
            self.shp_err_piq.fields = self.shp_piq.fields[1:]  # r�cup�re tous les champs

            # Zsro
            for i in range(len(self.record_zsro)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_zsro[i]['date_creat']) == "b''":
                    self.record_zsro[i]['date_creat'] = ""
            for i in range(len(self.record_zsro)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_zsro[i]['date_modif']) == "b''":
                    self.record_zsro[i]['date_modif'] = ""

            # Pfsro
            for i in range(len(self.record_pfsro)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfsro[i]['date_pmpa']) == "b''":
                    self.record_pfsro[i]['date_pmpa'] = ""
            for i in range(len(self.record_pfsro)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfsro[i]['date_creat']) == "b''":
                    self.record_pfsro[i]['date_creat'] = ""
            for i in range(len(self.record_pfsro)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfsro[i]['date_modif']) == "b''":
                    self.record_pfsro[i]['date_modif'] = ""

            # zpbo
            for i in range(len(self.record_zpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_zpbo[i]['date_creat']) == "b''":
                    self.record_zpbo[i]['date_creat'] = ""
            for i in range(len(self.record_zpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_zpbo[i]['date_modif']) == "b''":
                    self.record_zpbo[i]['date_modif'] = ""

            for i in range(len(self.shp_zpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.shp_zpbo.record(i)['date_creat']) == "b''":
                    self.shp_zpbo.record(i)['date_creat'] = ""
            for i in range(len(self.shp_zpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.shp_zpbo.record(i)['date_modif']) == "b''":
                    self.shp_zpbo.record(i)['date_modif'] = ""

            # pfpbo
            for i in range(len(self.record_pfpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfpbo[i]['date_pmpa']) == "b''":
                    self.record_pfpbo[i]['date_pmpa'] = ""
            for i in range(len(self.record_pfpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfpbo[i]['date_creat']) == "b''":
                    self.record_pfpbo[i]['date_creat'] = ""
            for i in range(len(self.record_pfpbo)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_pfpbo[i]['date_modif']) == "b''":
                    self.record_pfpbo[i]['date_modif'] = ""

            # piq
            for i in range(len(self.record_piq)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_piq[i]['date_creat']) == "b''":
                    self.record_piq[i]['date_creat'] = ""
            for i in range(len(self.record_piq)):  # remplacer une valeur NULL au format date par ""
                if str(self.record_piq[i]['date_modif']) == "b''":
                    self.record_piq[i]['date_modif'] = ""

            self.ZSROxSRO = []
            self.SROxZSRO = []
            self.ZPBOxPBO = []
            self.PBOxZPBO = []

            # ---------------------------------------------------------------------------------
            # print(" ------------------------V�rification des id m�tier des ZSRO et des SRO------------------------ ")
            for a in set1:
                if a in set2:
                    self.ZSROxSRO.append("la ZSRO " + a[14:] + " a un SRO associ�")
                    le = 1
                    # print(self.id_metier_zsro.index(a))         #r�cupere l'index de l'id metier
                    # print(self.record_zsro[self.id_metier_zsro.index(a)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_zsro[self.id_metier_zsro.index(a)]) #r�cup�re le shape de l'index de l'id m�tier

                    # self.shp_err_zsro.record(*self.record_zsro[self.id_metier_zsro.index(a)]) #ecris les records dans le nouveau shape
                    # self.shp_err_zsro.shape(self.shape_zsro[self.id_metier_zsro.index(a)])#ecris le shape de l'index de l'id m�tier

                else:
                    self.ZSROxSRO.append("la ZSRO " + a[14:] + " n'a pas de SRO associ�")
                    self.erreur_zsro += 1
                    # print(self.id_metier_zsro.index(a))         #r�cupere l'index de l'id metier
                    # print(self.record_zsro[self.id_metier_zsro.index(a)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_zsro[self.id_metier_zsro.index(a)]) #r�cup�re le shape de l'index de l'id m�tier

                    self.shp_err_zsro.record(
                        *self.record_zsro[self.id_metier_zsro.index(a)])  # ecris les records dans le nouveau shape
                    self.shp_err_zsro.shape(
                        self.shape_zsro[self.id_metier_zsro.index(a)])  # ecris le shape de l'index de l'id m�tier
            for b in set2:
                if b in set1:
                    self.SROxZSRO.append("le SRO " + b[14:] + " a une ZSRO associ�")
                    le = 1
                    # print(self.id_metier_pfsro.index(b))         #r�cupere l'index de l'id metier
                    # print(*self.record_pfsro[self.id_metier_pfsro.index(b)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_pfsro[self.id_metier_pfsro.index(b)]) #r�cup�re le shape de l'index de l'id m�tier

                    # self.shp_err_pfsro.record(*self.record_pfsro[self.id_metier_pfsro.index(b)]) #ecris les records dans le nouveau shape
                    # self.shp_err_pfsro.shape(self.shape_pfsro[self.id_metier_pfsro.index(b)])#ecris le shape de l'index de l'id m�tier
                else:
                    self.SROxZSRO.append("le SRO " + b[14:] + " n'as pas de ZSRO associ�")
                    self.erreur_pfsro += 1
                    # print(self.id_metier_pfsro.index(b))         #r�cupere l'index de l'id metier
                    # print(self.record_pfsro[self.id_metier_pfsro.index(b)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_pfsro[self.id_metier_pfsro.index(b)]) #r�cup�re le shape de l'index de l'id m�tier

                    self.shp_err_pfsro.record(
                        *self.record_pfsro[self.id_metier_pfsro.index(b)])  # ecris les records dans le nouveau shape
                    self.shp_err_pfsro.shape(
                        self.shape_pfsro[self.id_metier_pfsro.index(b)])  # ecris le shape de l'index de l'id m�tier
            # print("-----------------------------------------------------------------------------------------------------------------------------")
            # print(" ------------------------V�rification des id m�tier des ZPBO et des PBO------------------------ ")

            for a in set3:
                if a in set4:
                    # self.ZPBOxPBO.append("la ZPBO " + a[14:] + " a un PBO associ�")
                    le = 1
                    # print(self.id_metier_zpbo.index(a))         #r�cupere l'index de l'id metier
                    # print(self.record_zpbo[self.id_metier_zpbo.index(a)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_zpbo[self.id_metier_zpbo.index(a)]) #r�cup�re le shape de l'index de l'id m�tier
                    # self.shp_err_zpbo.record(*self.record_zpbo[self.id_metier_zpbo.index(a)]) #ecris les records dans le nouveau shape
                    # self.shp_err_zpbo.shape(self.shape_zpbo[self.id_metier_zpbo.index(a)])#ecris le shape de l'index de l'id m�tier
                else:
                    self.ZPBOxPBO.append("la ZPBO " + a[14:] + " n'as pas de PBO associ�")
                    self.erreur_zpbo += 1
                    # print(self.id_metier_zpbo.index(a))         #r�cupere l'index de l'id metier
                    # print(self.record_zpbo[self.id_metier_zpbo.index(a)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_zpbo[self.id_metier_zpbo.index(a)]) #r�cup�re le shape de l'index de l'id m�tier
                    self.shp_err_zpbo.record(
                        *self.record_zpbo[self.id_metier_zpbo.index(a)])  # ecris les records dans le nouveau shape
                    self.shp_err_zpbo.shape(
                        self.shape_zpbo[self.id_metier_zpbo.index(a)])  # ecris le shape de l'index de l'id m�tier
            for b in set4:
                if b in set3:
                    # self.PBOxZPBO.append("le PBO " + b[14:] + " a une ZPBO associ�")
                    le = 1
                    # print(self.id_metier_pfpbo.index(b))         #r�cupere l'index de l'id metier
                    # print(*self.record_pfpbo[self.id_metier_pfpbo.index(b)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(list(self.record_pfpbo[self.id_metier_pfpbo.index(b)]))  # r�cup�re le record de l'index de l'id m�tier
                    # self.shp_err_pfpbo.record(*self.record_pfpbo[self.id_metier_pfpbo.index(b)]) #ecris les records dans le nouveau shape
                    # self.shp_err_pfpbo.shape(self.shape_pfpbo[self.id_metier_pfpbo.index(b)])#ecris le shape de l'index de l'id m�tier
                else:
                    self.PBOxZPBO.append("le PBO " + b[13:] + " n'as pas de ZPBO associ�")
                    self.erreur_pfpbo += 1
                    # print("______________________________erreur PBO____________________________________")
                    # print(self.id_metier_pfpbo.index(b))         #r�cupere l'index de l'id metier
                    # print(*self.record_pfpbo[self.id_metier_pfpbo.index(b)])  # r�cup�re le record de l'index de l'id m�tier
                    # print(self.shape_pfpbo[self.id_metier_pfpbo.index(b)]) #r�cup�re le shape de l'index de l'id m�tier
                    self.shp_err_pfpbo.record(
                        *self.record_pfpbo[self.id_metier_pfpbo.index(b)])  # ecris les records dans le nouveau shape
                    self.shp_err_pfpbo.shape(
                        self.shape_pfpbo[self.id_metier_pfpbo.index(b)])  # ecris le shape de l'index de l'id m�tier

            # print(" ------------------------V�rification des du nombre de logements dans les ZSRO------------------------ ")

            for i in range(len(self.shp_zsro.shapes())):
                self.poly = Polygon(self.shp_zsro.shape(i).points)
                for j in range(len(self.shp_piq)):
                    self.list_piq = self.shp_piq.shape(j).points
                    if Point(*self.list_piq).within(self.poly):
                        self.compt_log += self.nb_logement_piq[j]
                    # else:
                    # le = 1
                    # print("Ce logement {} ne fait parti d'aucun shape".format(self.id_metier_piq[j]))
                    j += 1
                # print("il y a {} logement dans le shape{}".format(self.compt_log, self.id_metier_zsro[i]))
                self.list_log.append(self.compt_log)
                if self.compt_log < 300 or self.compt_log > 400:
                    self.erreur_logement += 1
                self.compt_log_total += self.compt_log
                self.compt_log = 0
                # print("--------------------------------------------------------------------------------------------------------")

                row = (os.path.basename(self.shp_zsro.record(i)['id_metier_']), self.ZSROxSRO[i], self.SROxZSRO[i],
                       self.list_log[i], self.shp_zsro.record(i)['nb_el'])
                self.worksheet_SRO.set_column(i + 1, 0, 26)
                self.worksheet_SRO.write_row(i + 1, 0, row, self.regular_text)

                self.worksheet_SRO.write_formula(i + 1, 5, "=D" + str(i + 2) + "-E" + str(i + 2), self.red_error)
                self.worksheet_SRO.write_formula(i + 1, 6,
                                                 "=IF(OR(D" + str(i + 2) + ">449,D" + str(i + 2) + "<300,F" + str(
                                                     i + 2) + "<>0),1,0)", self.red_error)

                # Write a conditional format over a range donn�e compt�.
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
            self.compt_total = self.compt_total - self.compt_log_total

            # print("il y a au total {} logements dans la ZSRO et {} logement hors Zones".format(self.compt_log_total, self.compt_total))

            # print(" ------------------------V�rification des du nombre de logements dans les ZPBO------------------------ ")

            self.compt_log = 0
            self.compt_log_total = 0

            for i in range(len(self.shp_zpbo.shapes())):
                self.poly = Polygon(self.shp_zpbo.shape(i).points)
                for j in range(len(self.shp_piq)):
                    self.list_piq = self.shp_piq.shape(j).points
                    if Point(*self.list_piq).within(self.poly):
                        # print("Le logement {} est dans le shape {}".format(self.id_metier_piq[j], self.id_metier_zsro[i]))

                        self.compt_log += self.nb_logement_piq[j]

                        if (self.shp_zpbo.record(i)['nb_module']) == 0:
                            self.type_log = -1  # Anomalie
                            # print("erreur Aucun module indiquer")

                        if (self.shp_zpbo.record(i)['nb_module']) == 1:
                            self.type_log = 1
                            # print("C'est une immeuble avec 1 modules")

                        if (self.shp_zpbo.record(i)['nb_module']) == 2:
                            self.type_log = 2  # immeuble
                            # print("C'est une immeuble avec 2 modules")

                        if (self.shp_zpbo.record(i)['nb_module']) >= 3:
                            self.type_log = 3  # immeuble
                            # print("C'est une immeuble avec 3 modules")

                # print("il y a {} logement dans le shape{}".format(self.compt_log, self.id_metier_zpbo[i]))
                if self.type_log == -1:  # anomalie
                    self.list_err_log = 1
                    self.shp_err_zpbo.record(*self.shp_zpbo.record(i))
                    self.shp_err_zpbo.shape(self.shp_zpbo.shape(i))
                    self.shp_err_pfpbo.record(
                        *self.record_pfpbo[self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])])
                    self.shp_err_pfpbo.shape(
                        self.shp_pfpbo.shape(self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])))

                if self.type_log == 1:
                    if self.compt_log > 5:
                        self.erreur_logement += 1
                        self.list_err_log = 1
                        self.shp_err_zpbo.record(*self.shp_zpbo.record(i))
                        self.shp_err_zpbo.shape(self.shp_zpbo.shape(i))
                        self.shp_err_pfpbo.record(
                            *self.record_pfpbo[self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])])
                        self.shp_err_pfpbo.shape(
                            self.shp_pfpbo.shape(self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])))


                    else:
                        self.list_err_log = 0

                if self.type_log == 2:
                    if self.compt_log > 10:
                        self.list_err_log = 1
                        self.shp_err_zpbo.record(*self.shp_zpbo.record(i))
                        self.shp_err_zpbo.shape(self.shp_zpbo.shape(i))
                        self.shp_err_pfpbo.record(
                            *self.record_pfpbo[self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])])
                        self.shp_err_pfpbo.shape(
                            self.shp_pfpbo.shape(self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])))

                    else:
                        self.list_err_log = 0

                if self.type_log > 3:
                    if self.compt_log > 16:
                        self.list_err_log = 1
                        self.shp_err_zpbo.record(*self.shp_zpbo.record(i))
                        self.shp_err_zpbo.shape(self.shp_zpbo.shape(i))
                        self.shp_err_pfpbo.record(
                            *self.record_pfpbo[self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])])
                        self.shp_err_pfpbo.shape(
                            self.shp_pfpbo.shape(self.id_metier_pfpbo.index(self.shp_zpbo.record(i)['id_metier_'])))

                    else:
                        self.list_err_log = 0
                self.f += 1

                self.worksheet_PIQ.set_column(i + 1, 0, 15)
                row2 = (os.path.basename(self.id_metier_zpbo[i]), self.compt_log, self.shp_zpbo.record(i)['nb_el'], "",
                        self.type_log, self.shp_zpbo.record(i)['nb_module'])
                self.worksheet_PIQ.write_row(i + 1, 0, row2, self.regular_text)

                self.worksheet_PIQ.write_formula(i + 1, 3, "=B" + str(i + 2) + "-C" + str(i + 2), self.red_error)

                self.worksheet_PIQ.write_formula(i + 1, 7, "=IF(D" + str(i + 2) + "<>0,1,)", self.red_error)
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
                # print("--------------------------------------------------------------------------------------------------------")
                self.compt_log_total += self.compt_log
                self.compt_log = 0

            self.worksheet_PIQ.write(self.f + 1, 0, 'Total', self.total)
            self.worksheet_PIQ.write(0, 8, "Le type de logement:", self.regular_text)
            self.worksheet_PIQ.write(1, 8, "-1:il n'y a pas de module(anomalie)", self.regular_text)
            self.worksheet_PIQ.write(2, 8, "1:il n'y a qu'un seul module", self.regular_text)
            self.worksheet_PIQ.write(3, 8, "2:il y a 2 modules", self.regular_text)
            self.worksheet_PIQ.write(4, 8, "3:il y a 3 ou plusieurs modules", self.regular_text)

            self.worksheet_PIQ.write_formula(self.f + 1, 1, "=SUM(B2:B" + str(self.f + 1) + ")", self.total)
            self.worksheet_PIQ.write_formula(self.f + 1, 2, "=SUM(C2:C" + str(self.f + 1) + ")", self.total)
            self.worksheet_PIQ.write_formula(self.f + 1, 3, "=SUM(D2:D" + str(self.f + 1) + ")", self.red_error_total)
            self.worksheet_PIQ.write_formula(self.f + 1, 5, "=SUM(F2:F" + str(self.f + 1) + ")", self.total)
            self.worksheet_PIQ.write_formula(self.f + 1, 6,
                                             "=SUM(G2:G" + str(self.f + 1) + ",H2:H" + str(self.f + 1) + ")",
                                             self.red_error_total)

            for i in range(len(self.ZPBOxPBO)):
                self.worksheet_PBO.write(i + 1, 0, self.ZPBOxPBO[i], self.red_error)
            for i in range(len(self.PBOxZPBO)):
                self.worksheet_PBO.write(i + 1, 1, self.PBOxZPBO[i], self.red_error)

            self.worksheet_PBO.write(1, 2,
                                     "Il y a " + str(self.erreur_zpbo) + " erreurs au niveau des Zones PBO, " + str(
                                         self.erreur_pfpbo) + " erreurs au niveau des PBO.", self.regular_text)

            for j in range(len(self.shp_piq)):
                self.compt_total += self.nb_logement_piq[j]
            self.erreur_log_hors_zone = self.compt_total - self.compt_log_total
            # print("Il y a au total {} logements dont, {} dans les ZSRO et {} logement hors Zones".format(self.compt_total, self.compt_log_total, self.erreur_log_hors_zone))

            self.erreurs_totales = self.erreur_zsro + self.erreur_pfsro + self.erreur_zpbo + self.erreur_pfpbo + self.erreur_logement
            # print("Il y a {} erreurs au niveau des Zones SRO, {} erreurs au niveau des Emplacement des SRO,\n {} erreurs au niveau des Zones PBO, {} erreurs au niveau des PBO.\n Il y {} au niveau des logements.\n Il y a donc {} erreurs au totale.".format(self.erreur_zsro, self.erreur_pfsro, self.erreur_zpbo, self.erreur_pfpbo, self.erreur_logement, self.erreurs_totales))
            # faire un bouton fermer
            self.etat_prog.set("Le rapport est dans Documents ou dans la destination choisie.")

            self.shp_err_zsro.close()
            self.shp_err_pfsro.close()
            self.shp_err_zpbo.close()
            self.shp_err_pfpbo.close()
            self.shp_err_piq.close()

            self.workbook.close()
        except AttributeError:
            self.etat_prog.set("Veuillez selectionner tous les shapes!")
        except ValueError:
            self.etat_prog.set("Les shapes ne correspondent pas!")
        except xlsxwriter.exceptions.FileCreateError:
            self.etat_prog.set("Le fichier excel est Ouvert, Fermez le et R�appuyer")



