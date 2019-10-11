import os.path
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.ttk import Treeview

from XlstoShape import xlstoshape


class interface_xlstoshape(Frame):
    # Création de la Classe
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=768, height=596, **kwargs)
        self.pack(fill=BOTH, expand=True)
        self.path = None
        self.destination_file = None
        self.wb = None
        self.value_output = StringVar()
        self.value_input = StringVar()
        self.value_input.set("Charger un fichier")
        self.convert_tool = xlstoshape()

        # Création des widgets
        self.frame_home = Frame(self)
        self.frame_home.pack(fill=X)
        self.bouton = Button(self.frame_home, text="Parcourir...", command=self.chose_file)
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.entree = Entry(self.frame_home, textvariable=self.value_input)
        self.entree.pack(fill=X, padx=5, expand=True)

        self.frame_tab = Frame(self, borderwidth=2, relief=GROOVE)
        self.frame_tab.pack(fill=BOTH, expand=True)
        self.tableau = Treeview(self.frame_tab)
        self.tableau['show'] = 'headings'
        self.tableau.pack(side=TOP, fill=BOTH, expand=True)

        self.frame_param = LabelFrame(self, text="Paramètre de conversion")
        self.frame_param.pack(fill=X, expand=True)
        self.listed = ttk.Combobox(self.frame_param, values=["Points", "Lignes", "Polygones"])
        self.listed.pack(padx=5, pady=5)

        self.frame_out = LabelFrame(self, text="Paramètre de sortie")
        self.frame_out.pack(fill=X, side=BOTTOM)
        self.sortie = Entry(self.frame_out, textvariable=self.value_output)
        self.sortie.pack(fill=X, padx=5, expand=True)

        self.ok = Button(self.frame_out, text="Executer", command=self.action)
        self.ok.pack(side=LEFT, padx=5, pady=5)

    def chose_file(self):
        self.path = askopenfilename(title="Ouverture fichier excel",
                              filetypes=[('Excel files', '*.xlsx; *.xls'), ('excel files', '*.xlsx ;*.xls')])
        self.value_input.set(self.path)

        head = self.convert_tool.load_xls_file(self.path)
        self.tableau["columns"] = head
        self.tableau.column("#0", stretch=True)
        for name_col in head:
            self.tableau.column(name_col, stretch=True)
            self.tableau.heading(name_col, text=name_col)

        for i in range(1, 10):
            self.tableau.insert("", "end", "", values=self.convert_tool.sheet.row_values(i))

        self.destination_file, destination_ext = os.path.splitext(self.path)
        self.value_output.set(self.destination_file)

    def action(self):
        try:
            self.convert_tool.create_shape(self.sortie.get(), self.listed.get())
        except TypeError as err:
            askokcancel("Attention", *err.args)
            return

        self.convert_tool.convert_file()
        print("Success")


fenetre = Tk()
interface = interface_xlstoshape(fenetre)

interface.mainloop()
interface.destroy()
