import shapefile as sf
from shapely.geometry import Point, Polygon
import xlrd
import xlsxwriter
import os
from datetime import date
import tkinter
from tkinter.filedialog import *

filepath_zpmz = askopenfilename(title="Ouverture shape : ZPMZ", filetypes=[('ShapeFile', '.shp'), ('all files', '.*')])

shp_ZPMZ = sf.Reader(filepath_zpmz)

#shp_ZPMZ = sf.Reader('C:\\Users\\m.rosillette\\Documents\\Analyse\\PRO\\Projet_Qgis_Nord Caraïbe\\Shp\\Etude '
                    # 'TRSP\\ZE_PMZ_SAINT-PIERRE.shp')

filepath_pmz = askopenfilename(title="Ouverture shape : PMZ", filetypes=[('ShapeFile', '.shp'), ('all files', '.*')])

# shp_PMZ = sf.Reader('C:\\Users\\m.rosillette\\Documents\\Analyse\\PRO\\Projet_Qgis_Nord Caraïbe\\Shp\\Etude '
                     #'TRSP\\PF_PMZ SAINT PIERRE.shp')
shp_PMZ = sf.Reader(filepath_pmz)

filepath_piq = askopenfilename(title="Ouverture shape : PIQ Bal", filetypes=[('ShapeFile', '.shp'), ('all files', '.*')])

#shp_PIQ = sf.Reader('C:\\Users\\m.rosillette\\Documents\\Analyse\\PRO\\Projet QGIS Case Pilote_APD\\Piquetage\\PIQ_BAL_Case_Pilote.shp')

shp_PIQ = sf.Reader(filepath_piq)

filepath_zpb = askopenfilename(title="Ouverture shape : ZPB", filetypes=[('ShapeFile', '.shp'), ('all files', '.*')])

# shp_ZPB = sf.Reader('C:\\Users\\m.rosillette\\Documents\\Analyse\\PRO\\Projet_Qgis_Nord '
                    #'Caraïbe\\Shp\\Etude D2\\ZE_PB_SAINT-PIERRE.shp')

shp_ZPB = sf.Reader(filepath_zpb)

#filepath = askopenfilename(title="Ouverture shape", filetypes=[('ShapeFile', '.shp'), ('all files', '.*')])

# shp_PB = sf.Reader

# path_test_r = "C:\\Users\\m.rosillette\\Documents\\Check_Bal\\Bal_Nettoye CP.xlsx"
path_test_w = "C:\\ Users\\g.battery\Desktop\\fichier_a_comparer\\test\\Rapport_QGIS_CAS-PRO_" + str(date.today()) + ".xlsx"
# wb = xlrd.open_workbook(path_test_r)
# sheet = wb.sheet_by_index(0)

workbook = xlsxwriter.Workbook(path_test_w)

regular_text = workbook.add_format({'font_size': 10})
title = workbook.add_format({'bold': True, 'font_size': 11, 'text_wrap': True, 'align': 'center'})
error = workbook.add_format({'font_size': 10, 'shrink': True})
red_error = workbook.add_format({'font_color': 'red', 'font_size': 10, 'shrink': True})
orange_error = workbook.add_format({'font_color': 'orange', 'font_size': 10, 'shrink': True})
violet_error = workbook.add_format({'font_color': 'violet', 'font_size': 10, 'shrink': True})
red = workbook.add_format({'font_color': 'red', 'font_size': 10, 'bold': True, 'bg_color': '#ffc7ce'})
green = workbook.add_format({'font_color': 'green', 'font_size': 10})
orange = workbook.add_format({'font_color': 'orange', 'font_size': 10, 'bold': True, 'bg_color': '#fcd5b4'})
violet = workbook.add_format({'font_color': 'violet', 'font_size': 10, 'bold': True, 'bg_color': '#d2caec'})


# Fonction création de l'entête de chaque feuille du classeur Excel. La première ligne est remplie avec les titres suivants :
def entete(wkb, name, col_title):
    worksheet = wkb.add_worksheet(name)
    worksheet.write(0, 0, col_title, title)
    return worksheet


worksheet_pmz = entete(workbook, "Rapport_PMZ", "Cohérence des données Shape : PMZ/PIQ_BAL") # création d'une nouvelle feuille avec la fonction entete

list_point_piq = [] #création d'une liste avec les coordonées des points
for shape in shp_PIQ.shapes():
    list_point_piq.append(shape.points[0])

title_ = ("PMZ", "PF_PMZxZPMZ", "Lgt Res\nZPMZxPiq_Bal", "Lgt Pro\nZPMZxPiq_Bal", "Lgt Tot\nZPMZxPiq_Bal", "Lgt Res\nZPMZ",
          "Lgt Pro\nZPMZ", "Lgt Tot\nZPMZ", "Erreur\nPiq_Bal", "Ligne(s)") #titre 1er feuille
worksheet_pmz.write_row(0, 0, title_, title)
for j in range(len(shp_ZPMZ.shapes())):
    comp_tot = 0 #pour reset a chaque shape (enlever pro et res)
    comp_res = 0
    comp_pro = 0
    comp_err = 0
    comp_coegeo = "NOK"
    err = ""
    poly = Polygon(shp_ZPMZ.shape(j).points)#recupération de la forme sur le plan pour le stocquer dans poly pour vérifié si le point est dans la forme
    for i in range(len(list_point_piq)):
        if Point(list_point_piq[i]).within(poly): # si l'élément i est dans le polygone
            try:
                comp_res += shp_PIQ.record(i)['nb_logemen']# ajouter au nb de logement
                # comp_pro += shp_PIQ.record(i)['NB_PRO']
            except UnicodeDecodeError:
                comp_err += 1 #si echec err+1
                err += str(i) + "\n" # ajouter la valeur et retour a la ligne
    #for i in range(len(shp_PMZ)):
        # if os.path.basename(shp_PMZ.record(i)['id_metier_']) == os.path.basename(shp_ZPMZ.record(j)['id_metier_']) and Point(shp_PMZ.shape(i).points[0]).within(poly):
            #comp_coegeo = "OK"
    row = (os.path.basename(shp_ZPMZ.record(j)['id_metier_']), comp_coegeo, comp_res, comp_pro, comp_res + comp_pro,
           " ", " ", shp_ZPMZ.record(j)['nb_el'], comp_err) #rentrer les valeur dans excel
    worksheet_pmz.write_row(j + 1, 0, row, regular_text) #ecriture sur excel
    worksheet_pmz.write(j + 1, 9, err[:-1], error)
    j += 1

worksheet_pb = entete(workbook, "Rapport PBxPIQ_Bal", "Cohérence des données Shape : PB/PIQ_BAL") #2iem feuille
title_ = ("PB", "Lgt Res\nZPBxPIQ", "Lgt Pro\nZPBxPIQ", "Lgt Tot\nZPBxPIQ", "Lgt Tot\nZPBO", "µModule\nZPBO",
            "Erreur\nPiq_Bal", "Ligne(s)")
worksheet_pb.write_row(4, 0, title_, title) #titre a la ligne 4
comp_o = 0
comp_r = 0
comp_v = 0
points_cent = [] #remplacer par les points pbo
for i in range(len(shp_ZPB)):
    poly = Polygon(shp_ZPB.shape(i).points)
    points_cent.append(Point(poly.centroid.coords[:][0][0], poly.centroid.coords[:][0][1]))

for j in range(len(shp_ZPB.shapes())):
    comp_res = 0
    comp_pro = 0#enlever pro et res
    comp_err = 0
    comp_coegeo = "NOK"
    err = ""
    poly = Polygon(shp_ZPB.shape(j).points)# stocquer le contour des zpb dans poly
    for i in range(len(list_point_piq)):
        if Point(list_point_piq[i]).within(poly): # si le point pic est dans poly on add le nb
            try:
                comp_res += shp_PIQ.record(i)['nb_logemen']
                # comp_pro += shp_PIQ.record(i)['NB_PRO']
            except UnicodeDecodeError:
                comp_err += 1
    row = (os.path.basename(shp_ZPB.record(j)['id_metier_']), comp_res, comp_pro, comp_res + comp_pro,
           shp_ZPB.record(j)['nb_el'], shp_ZPB.record(j)['nb_module'], comp_err)
    if comp_res + comp_pro > shp_ZPB.record(j)['nb_module']*6: #si comp_res+comp_ on regarde si on dépasse le nb de fibre dans un module (il doit y avoir 6 par module max donc on test
        worksheet_pb.write_row(j + 5, 0, row, red)
        comp_r += 1
        if err is not "":
            worksheet_pb.write(j + 5, 7, err[:-1], red_error)
    elif comp_res + comp_pro > int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)) or comp_res + comp_pro !=shp_ZPB.record(j)['nb_el']: #si le total est supérieur
        if comp_res + comp_pro > int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)): # si on dépasse 80% =erreur
            moy = []
            for pt in points_cent: # pour chaque centre de zpb
                if Point(poly.centroid.coords[:][0][0], poly.centroid.coords[:][0][1]).distance(pt) <= 200: #faire la moyenne du nombre de logement qui ya dans tousl es zpb sur 200m
                    for k in range(len(shp_ZPB)):
                        if pt == Polygon(shp_ZPB.shape(k).points).centroid:
                            moy.append(shp_ZPB.record(k)['nb_el']) #on stocque le nb de logements
            if sum(moy) / len(moy) <= int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)):#additionne toutes les valeur de la liste et on divise par le nb d'élément (moyenne)
                worksheet_pb.write_row(j + 5, 0, row, regular_text)# écriture dans la feuille
            else:
                worksheet_pb.write_row(j + 5, 0, row, orange)#sinon on écrire en orange (probleme)
                comp_o += 1
        if comp_res + comp_pro != shp_ZPB.record(j)['nb_el']: # on compare les 2 valeurs
            worksheet_pb.write_row(j + 5, 0, row, regular_text)
        if err is not "":
            worksheet_pb.write(j + 5, 7, err[:-1], orange_error)
    else:
        worksheet_pb.write_row(j + 5, 0, row, regular_text)
        if err is not "":
            worksheet_pb.write(j + 5, 7, err[:-1], error)

    j += 1

row_o = ("PB saturé", comp_o)
worksheet_pb.write_row(1, 0, row_o, regular_text)
row_r = ("PB non conforme", comp_r)
worksheet_pb.write_row(2, 0, row_r, red)

workbook.close()

