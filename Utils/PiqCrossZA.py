import os

import shapefile as sf
import xlsxwriter
from shapely.geometry import Polygon


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

    def crossAnalyse(self,path_dest, field_zsro, field_pfsro, field_zpbo, field_pfpbo, field_piq):
        # do something
        print(field_zsro, field_pfsro, field_zpbo, field_pfpbo, field_piq)
        wkb = self.xlsReport(path_dest)
        regular_text = wkb.add_format({'font_size': 10})
        error = wkb.add_format({'font_size': 10, 'shrink': True})
        worksheet_sro = wkb.add_worksheet("Rapport SRP")
        worksheet_sro.write(0, 0, "Cohérence des données Shape : PMZ/PIQ_BAL", title)
        for j in range(len(self.shp_zsro.shapes())):
            comp_tot = 0
            comp_res = 0
            comp_pro = 0
            comp_err = 0
            comp_coegeo = "NOK"
            err = ""
            poly = Polygon(self.shp_zsro.shape(j).points)
            #for i in range(len(list_point_piq)):
            #    if Point(list_point_piq[i]).within(poly):
            #        try:
            #            comp_res += self.shp_piq.record(i)[str(field_piq)]
            #            # comp_pro += shp_PIQ.record(i)['NB_PRO']
            #        except UnicodeDecodeError:
            #            comp_err += 1
            #            err += str(i) + "\n"
            # for i in range(len(shp_PMZ)):
            # if os.path.basename(shp_PMZ.record(i)['id_metier_']) == os.path.basename(shp_ZPMZ.record(j)['id_metier_']) and Point(shp_PMZ.shape(i).points[0]).within(poly):
            # comp_coegeo = "OK"
            row = (
            os.path.basename(self.shp_zsro.record(j)['id_metier_']), comp_coegeo, comp_res, comp_pro, comp_res + comp_pro,
            " ", " ", self.shp_zsro.record(j)[str(field_zsro)], comp_err)
            worksheet_sro.write_row(j + 1, 0, row, regular_text)
            worksheet_sro.write(j + 1, 9, err[:-1], error)
            j += 1

    def xlsReport(self, path_dest):
        workbook = xlsxwriter.Workbook(path_dest)

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

        return workbook