import xlrd
import shapefile
import tkinter
from tkinter.filedialog import *
from tkinter import Button
from tkinter import Label
from tkinter import StringVar
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os, os.path

path = askopenfilename(title="Ouverture fichier excel", filetypes=[('Excel files', '*.xlsx; *.xls'), ('all files', '.*')])

path = "D:\\Gio\\Documents\\projet\\gendarmerie.xlsx"

w = shapefile.Writer("D:\\Gio\\Documents\\projet\\gendarmerie")
w.field('Telephone', "C")
w.field("Departement","N")
w.field("code_commune_insee","N")
w.field("voie","C")
w.field("code_postal", "N")
w.field("Commune", "C")#Ajout dans un champ
w.field("geocodage_epsg")
w.field("X", "N")
w.field("Y", "N")

wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0)#premier feuille
for i in range(1, sheet.nrows):
    print(sheet.row_values(i))
    w.record(Telephone=sheet.row_values(i)[0], Departement=sheet.row_values(i)[1], code_commune_insee=sheet.row_values(i)[2],
             voie=sheet.row_values(i)[3], code_postal=sheet.row_values(i)[4], Commune=sheet.row_values(i)[5], geocodage_epsg=sheet.row_values(i)[6],
             X=sheet.row_values(i)[7], Y=sheet.row_values(i)[8])
    w.point(sheet.row_values(i)[7],sheet.row_values(i)[8])

w.close()

#w.record(Telephone=)