from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *

from Utils.PiqCrossZA import piqCrossZA


class interface_piqCrossZa(Frame):
    # Création de la Classe
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, **kwargs)
        self.pack(fill=BOTH, expand=True)
        self.path = None
        self.destination_file = None
        self.wb = None
        self.value_output = StringVar()
        self.value_input_zsro = StringVar()
        self.value_input_zsro.set("Charger la Zone Arrière SRO")
        self.value_input_pfsro = StringVar()
        self.value_input_pfsro.set("Charger les emplacements SRO")
        self.value_input_zpbo = StringVar()
        self.value_input_zpbo.set("Charger la Zone arrière PBO")
        self.value_input_pfpbo = StringVar()
        self.value_input_pfpbo.set("Charger les emplacements PBO")
        self.value_input_piq = StringVar()
        self.value_input_piq.set("Charger le Piq_BAL")
        self.analyse_tool = piqCrossZA()

        # Création des widgets
        #
        # Frame chargement ZSRO ----------------------------------------------------------------------------------------
        #
        self.frame_zsro = Frame(self)
        self.frame_zsro.pack(fill=X)
        self.entree_zsro = Entry(self.frame_zsro, textvariable=self.value_input_zsro)
        self.entree_zsro.pack(side=LEFT, fill=X, padx=5, pady=5, expand=True)
        self.bouton = Button(self.frame_zsro, text="Parcourir...", command= lambda : self.chose_file("zsro"))
        self.bouton.pack(side=RIGHT, padx=5)
        self.listed_zsro = ttk.Combobox(self.frame_zsro, state="readonly")
        self.listed_zsro.pack(pady=7)
        #
        # Frame chargement PFSRO ---------------------------------------------------------------------------------------
        #
        self.frame_pfsro = Frame(self)
        self.frame_pfsro.pack(fill=X)
        self.entree_pfsro = Entry(self.frame_pfsro, textvariable=self.value_input_pfsro)
        self.entree_pfsro.pack(side=LEFT, fill=X, padx=5, expand=True)
        self.bouton = Button(self.frame_pfsro, text="Parcourir...", command= lambda : self.chose_file("pfsro"))
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.listed_pfsro = ttk.Combobox(self.frame_pfsro, state="readonly", values=["Points", "Lignes", "Polygones"])
        self.listed_pfsro.pack(pady=7)
        #
        # Frame chargement ZPBO ----------------------------------------------------------------------------------------
        #
        self.frame_zpbo = Frame(self)
        self.frame_zpbo.pack(fill=X)
        self.entree_zpbo = Entry(self.frame_zpbo, textvariable=self.value_input_zpbo)
        self.entree_zpbo.pack(side=LEFT, fill=X, padx=5, expand=True)
        self.bouton = Button(self.frame_zpbo, text="Parcourir...", command= lambda : self.chose_file("zpbo"))
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.listed_zpbo = ttk.Combobox(self.frame_zpbo, state="readonly", values=["Points", "Lignes", "Polygones"])
        self.listed_zpbo.pack(pady=7)
        #
        # Frame chargement PFPBO ---------------------------------------------------------------------------------------
        #
        self.frame_pfpbo = Frame(self)
        self.frame_pfpbo.pack(fill=X)
        self.entree_pfpbo = Entry(self.frame_pfpbo, textvariable=self.value_input_pfpbo)
        self.entree_pfpbo.pack(side=LEFT, fill=X, padx=5, expand=True)
        self.bouton = Button(self.frame_pfpbo, text="Parcourir...", command= lambda : self.chose_file("pfpbo"))
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.listed_pfpbo = ttk.Combobox(self.frame_pfpbo, state="readonly", values=["Points", "Lignes", "Polygones"])
        self.listed_pfpbo.pack(pady=7)
        #
        # Frame chargement PIQ -----------------------------------------------------------------------------------------
        #
        self.frame_piq = Frame(self)
        self.frame_piq.pack(fill=X)
        self.entree_piq = Entry(self.frame_piq, textvariable=self.value_input_piq)
        self.entree_piq.pack(side=LEFT, fill=X, padx=5, expand=True)
        self.bouton = Button(self.frame_piq, text="Parcourir...", command= lambda : self.chose_file("piq"))
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.listed_piq = ttk.Combobox(self.frame_piq, state="readonly", values=["Points", "Lignes", "Polygones"])
        self.listed_piq.pack(pady=7)

        self.frame_out = LabelFrame(self, text="Paramètre de sortie")
        self.frame_out.pack(fill=X, side=BOTTOM)
        self.frame_file_output = Frame(self.frame_out)
        self.frame_file_output.pack(fill=X)
        self.sortie = Entry(self.frame_file_output, textvariable=self.value_output)
        self.sortie.pack(side=LEFT, fill=X, padx=5, expand=True)
        self.bouton = Button(self.frame_file_output, text="Parcourir...", command=self.save_file)
        self.bouton.pack(side=RIGHT, padx=5, pady=5)
        self.frame_exe = Frame(self.frame_out)
        self.frame_exe.pack(fill=X, side=BOTTOM)
        self.ok = Button(self.frame_exe, text="Executer", command=self.action)
        self.ok.pack(side=LEFT, padx=5, pady=5)

    def chose_file(self, typeData):
        if typeData == "zsro":
            self.path = askopenfilename(title="Chagement des Zones arrières SRO",
                                  filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])
            self.value_input_zsro.set(self.path)
            self.listed_zsro['values'] = self.analyse_tool.loadShpfile(self.path, "zsro")
        elif typeData == "pfsro":
            self.path = askopenfilename(title="Chagement des emplacements SRO",
                                        filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])
            self.value_input_pfsro.set(self.path)
            self.listed_pfsro['values'] = self.analyse_tool.loadShpfile(self.path, "pfsro")
        elif typeData == "zpbo":
            self.path = askopenfilename(title="Chagement des Zones arrières PBO",
                                        filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])
            self.value_input_zpbo.set(self.path)
            self.listed_zpbo['values'] = self.analyse_tool.loadShpfile(self.path, "zpbo")
        elif typeData == "pfpbo":
            self.path = askopenfilename(title="Chagement des emplacements PBO",
                                        filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])
            self.value_input_pfpbo.set(self.path)
            self.listed_pfpbo['values'] = self.analyse_tool.loadShpfile(self.path, "pfpbo")
        elif typeData == "piq":
            self.path = askopenfilename(title="Chagement du Piq_BAL",
                                        filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])
            self.value_input_piq.set(self.path)
            self.listed_piq['values'] = self.analyse_tool.loadShpfile(self.path, "zsro")


        #self.destination_file, destination_ext = os.path.splitext(self.path)
        #self.value_output.set(self.destination_file)

    def save_file(self):
        self.path_test = asksaveasfilename()
        self.value_output.set(self.path_test)

    def action(self):
        try:
            self.convert_tool.create_shape(self.sortie.get(), self.listed.get())
        except TypeError as err:
            askokcancel("Attention", *err.args)
            return

        self.convert_tool.convert_file()
        print("Success")


fenetre = Tk()
interface = interface_piqCrossZa(fenetre)

interface.mainloop()
interface.destroy()
