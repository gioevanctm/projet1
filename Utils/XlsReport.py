from xlsxwriter import workbook

class xlsrpt(workbook):
    def __init__(self, path):
        workbook.__init__(path)
        self.regular_text = workbook.add_format({'font_size': 10})
        self.title = workbook.add_format({'bold': True, 'font_size': 11, 'text_wrap': True, 'align': 'center'})
        self.error = workbook.add_format({'font_size': 10, 'shrink': True})
        self.red_error = workbook.add_format({'font_color': 'red', 'font_size': 10, 'shrink': True})
        self.orange_error = workbook.add_format({'font_color': 'orange', 'font_size': 10, 'shrink': True})
        self.violet_error = workbook.add_format({'font_color': 'violet', 'font_size': 10, 'shrink': True})
        self.red = workbook.add_format({'font_color': 'red', 'font_size': 10, 'bold': True, 'bg_color': '#ffc7ce'})
        self.green = workbook.add_format({'font_color': 'green', 'font_size': 10})
        self.orange = workbook.add_format({'font_color': 'orange', 'font_size': 10, 'bold': True, 'bg_color': '#fcd5b4'})
        self.violet = workbook.add_format({'font_color': 'violet', 'font_size': 10, 'bold': True, 'bg_color': '#d2caec'})

    def newSheet(self, sheetname, title):
        wks = self.a