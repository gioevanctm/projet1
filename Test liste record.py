import xlrd
import shapefile
from tkinter.filedialog import *
from tkinter import *
import os, os.path
from tkinter import *
import tkinter.messagebox

path = "D:\Gio\Bureau\dossier test\gendarmerie.xlsx"
w = shapefile.Writer("D:\Gio\Bureau\dossier test\gendarmerie")

w = shapefile.Writer(path)
w.field('Telephone', "C")
w.field("Departement","N")
w.field("code_commune_insee","N")
w.field("voie","C")
w.field("code_postal", "N")
w.field("Commune", "C")
w.field("geocodage_epsg")
w.field("X", "N")
w.field("Y", "N")

wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0)  # premier feuille
type_cellule=""
list(type_cellule)

for c in sheet.row_values(1):
    if type(c) is str:
        type_cellule +="C "

    else:type_cellule +="N "
    #print(type_cellule)

sa= type_cellule.split( )
print(sa)
#print(type(liste))
#print(liste)
for i in sheet.row_values(0):

    #print(*sheet.row_values(i))
    #(print(type_cellule)
    w.field(i, sa(i))
    #w.field(*sheet.row_values(i), *
    print(type_cellule)



#for i in range(1, sheet.nrows):

    # print(sheet.row_values(i))
    #print(sheet.row_values(i))
    #w.record(*sheet.row_values(i)[:-2])
    #w.point(*sheet.row_values(i)[-4:-2])


w.close()
