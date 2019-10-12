import shapefile
import xlrd


class xlstoshape:
    # Création de la Classe
    def __init__(self):
        self.wb = None
        self.sheet = None
        self.w = None

    def convert_file(self):
        for c in self.sheet.row_values(0):
            self.w.field(c, 'C')

        for i in range(1, self.sheet.nrows):
            self.w.record(*self.sheet.row_values(i))
            self.w.point(*self.sheet.row_values(i)[-4:-2])

        self.w.close()

    def load_xls_file(self, path):
        self.wb = xlrd.open_workbook(path)
        self.sheet = self.wb.sheet_by_index(0)
        return self.sheet.row_values(0)

    def create_shape(self, path_dest, type_shp):
        if type_shp == "Points":
            self.w = shapefile.Writer(path_dest, ShapeType=1)
        elif type_shp == "Lignes":
            self.w = shapefile.Writer(path_dest, ShapeType=3)
        elif type_shp == "Polygones":
            self.w = shapefile.Writer(path_dest, ShapeType=5)
        else:
            raise TypeError("Il faut préciser le type de Shape")
