import shapefile
import xlrd


class xlstoshape:
    # Création de la Classe
    def __init__(self):#initialistation des attribus
        self.wb = None
        self.sheet = None
        self.w = None
        self.type_cellule = []

    def convert_file(self): #définition d'une fonction
        type_cellule=[]
        for c in self.sheet.row_values(1):
            if type(c) is str:
                type_cellule.append("C")
            else:
                type_cellule.append("N")

        for c in range(self.sheet.ncols): #création de champ
            self.w.field(self.sheet.row_values(0)[c], type_cellule[c])


        for i in range(1, self.sheet.nrows): #for 1 jusqu'a le nb total de ligne
            self.w.record(*self.sheet.row_values(i)) # remplire les données attributaires
            self.w.point(*self.sheet.row_values(i)[-4:-2]) #placer les points # on va devoir demander a l'utilisation où sont les données géographique

        self.w.close() #fermeture du fichier excel

    def load_xls_file(self, path): #définition d'une fonction qui choisie le fichier excel
        self.wb = xlrd.open_workbook(path) #ouverture du fichier excel
        self.sheet = self.wb.sheet_by_index(0) #choisir la premier feuille
        return self.sheet.row_values(0) #retourne la premiere ligne

    def create_shape(self, path_dest, type_shp): #définition d'une fonction qui crée le shape
        if type_shp == "Points": #don teste les type
            self.w = shapefile.Writer(path_dest, ShapeType=1) #shape de type point
        elif type_shp == "Lignes":
            self.w = shapefile.Writer(path_dest, ShapeType=3) #shape de type ligne
        elif type_shp == "Polygones":
            self.w = shapefile.Writer(path_dest, ShapeType=5) #shape de type pointpolygones
        else:
            raise TypeError("Il faut préciser le type de Shape") # lève l'erreur et revoyer le texte
