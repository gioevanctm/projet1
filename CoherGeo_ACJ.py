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
path_test_w = "C:\\Users\\m.rosillette\\Documents\\Check_Bal\\Rapport_QGIS_CAS-PRO_" + str(date.today()) + ".xlsx"
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


worksheet_pmz = entete(workbook, "Rapport_PMZ", "Cohérence des données Shape : PMZ/PIQ_BAL")

list_point_piq = []
for shape in shp_PIQ.shapes():
    list_point_piq.append(shape.points[0])

title_ = ("PMZ", "PF_PMZxZPMZ", "Lgt Res\nZPMZxPiq_Bal", "Lgt Pro\nZPMZxPiq_Bal", "Lgt Tot\nZPMZxPiq_Bal", "Lgt Res\nZPMZ",
          "Lgt Pro\nZPMZ", "Lgt Tot\nZPMZ", "Erreur\nPiq_Bal", "Ligne(s)")
worksheet_pmz.write_row(0, 0, title_, title)
for j in range(len(shp_ZPMZ.shapes())):
    comp_tot = 0
    comp_res = 0
    comp_pro = 0
    comp_err = 0
    comp_coegeo = "NOK"
    err = ""
    poly = Polygon(shp_ZPMZ.shape(j).points)
    for i in range(len(list_point_piq)):
        if Point(list_point_piq[i]).within(poly):
            try:
                comp_res += shp_PIQ.record(i)['nb_logemen']
                # comp_pro += shp_PIQ.record(i)['NB_PRO']
            except UnicodeDecodeError:
                comp_err += 1
                err += str(i) + "\n"
    #for i in range(len(shp_PMZ)):
        # if os.path.basename(shp_PMZ.record(i)['id_metier_']) == os.path.basename(shp_ZPMZ.record(j)['id_metier_']) and Point(shp_PMZ.shape(i).points[0]).within(poly):
            #comp_coegeo = "OK"
    row = (os.path.basename(shp_ZPMZ.record(j)['id_metier_']), comp_coegeo, comp_res, comp_pro, comp_res + comp_pro,
           " ", " ", shp_ZPMZ.record(j)['nb_el'], comp_err)
    worksheet_pmz.write_row(j + 1, 0, row, regular_text)
    worksheet_pmz.write(j + 1, 9, err[:-1], error)
    j += 1

worksheet_pb = entete(workbook, "Rapport PBxPIQ_Bal", "Cohérence des données Shape : PB/PIQ_BAL")
title_ = ("PB", "Lgt Res\nZPBxPIQ", "Lgt Pro\nZPBxPIQ", "Lgt Tot\nZPBxPIQ", "Lgt Tot\nZPBO", "µModule\nZPBO",
            "Erreur\nPiq_Bal", "Ligne(s)")
worksheet_pb.write_row(4, 0, title_, title)
comp_o = 0
comp_r = 0
comp_v = 0
points_cent = []
for i in range(len(shp_ZPB)):
    poly = Polygon(shp_ZPB.shape(i).points)
    points_cent.append(Point(poly.centroid.coords[:][0][0], poly.centroid.coords[:][0][1]))

for j in range(len(shp_ZPB.shapes())):
    comp_res = 0
    comp_pro = 0
    comp_err = 0
    comp_coegeo = "NOK"
    err = ""
    poly = Polygon(shp_ZPB.shape(j).points)
    for i in range(len(list_point_piq)):
        if Point(list_point_piq[i]).within(poly):
            try:
                comp_res += shp_PIQ.record(i)['nb_logemen']
                # comp_pro += shp_PIQ.record(i)['NB_PRO']
            except UnicodeDecodeError:
                comp_err += 1
    row = (os.path.basename(shp_ZPB.record(j)['id_metier_']), comp_res, comp_pro, comp_res + comp_pro,
           shp_ZPB.record(j)['nb_el'], shp_ZPB.record(j)['nb_module'], comp_err)
    if comp_res + comp_pro > shp_ZPB.record(j)['nb_module']*6:
        worksheet_pb.write_row(j + 5, 0, row, red)
        comp_r += 1
        if err is not "":
            worksheet_pb.write(j + 5, 7, err[:-1], red_error)
    elif comp_res + comp_pro > int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)) or comp_res + comp_pro !=shp_ZPB.record(j)['nb_el']:
        if comp_res + comp_pro > int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)):
            moy = []
            for pt in points_cent:
                if Point(poly.centroid.coords[:][0][0], poly.centroid.coords[:][0][1]).distance(pt) <= 200:
                    for k in range(len(shp_ZPB)):
                        if pt == Polygon(shp_ZPB.shape(k).points).centroid:
                            moy.append(shp_ZPB.record(k)['nb_el'])
            if sum(moy) / len(moy) <= int(round(shp_ZPB.record(j)['nb_module'] * 6 * 8 / 10)):
                worksheet_pb.write_row(j + 5, 0, row, regular_text)
            else:
                worksheet_pb.write_row(j + 5, 0, row, orange)
                comp_o += 1
        if comp_res + comp_pro != shp_ZPB.record(j)['nb_el']:
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

