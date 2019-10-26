import shapefile as sf


class piqCrossZA:
    # Création de la Classe
    def __init__(self):
        self.shp_zsro = None
        self.shp_pfsro = None
        self.shp_zpbo = None
        self.shp_pfpbo = None
        self.shp_piq = None
        self.wb = None
        self.sheet = None
        self.shp_error = None

    def loadShpfile(self, path, typedata):
        if typedata == "zsro":
            self.shp_zsro = sf.Reader(path)
            listFields = []
            for field in self.shp_zsro.fields:
                listFields.append(field[0])
            return listFields[1:]
        elif typedata == "pfsro":
            self.shp_pfsro = sf.Reader(path)
            listFields = []
            for field in self.shp_pfsro.fields:
                listFields.append(field[0])
            return listFields[1:]
        elif typedata == "zpbo":
            self.shp_zpbo = sf.Reader(path)
            listFields = []
            for field in self.shp_zpbo.fields:
                listFields.append(field[0])
            return listFields[1:]
        elif typedata == "pfpbo":
            self.shp_pfpbo = sf.Reader(path)
            listFields = []
            for field in self.shp_pfpbo.fields:
                listFields.append(field[0])
            return listFields[1:]
        elif typedata == "piq":
            self.shp_piq = sf.Reader(path)
            listFields = []
            for field in self.shp_piq.fields:
                listFields.append(field[0])
            return listFields[1:]

    def crossAnalyse(self):
        # do something
        return