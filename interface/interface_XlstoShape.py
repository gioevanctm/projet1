import os.path
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.ttk import Treeview

from XlstoShape import xlstoshape


class interface_xlstoshape(Frame):
    # Création de la Classe avec comme parametre frame(cadre)
    def __init__(self, fenetre, **kwargs): #constructeur de la classe(fonction qui va s'activé quand on va utiliser la classe(parametre de base)) self:contient les variable(attribue et methode) **liste d'argument de base que peut recevoir la frame(transforme
        Frame.__init__(self, fenetre, width=768, height=596, **kwargs) #appel duconstructeur de frame (** kwargs apppel des parametre sous forme de dictionnaire)
        self.pack(fill=BOTH, expand=True) #tout replir (fenetre) et si on modifie la taille va augmenter en fonction d besoin)
        self.path = None #intialisation des attribues
        self.destination_file = None #intialisation des attribues
        self.wb = None #intialisation des attribues
        self.value_output = StringVar()  #intialisation des attribues
        self.value_input = StringVar() #intialisation des attribues
        self.value_input.set("Charger un fichier") #intialisation des attribues
        self.convert_tool = xlstoshape() #intialisation des attribues (instancier la classe qu'on a récuperer) (convert_tool contient la classe)

        # Création des widgets
        self.frame_home = Frame(self)  #
        self.frame_home.pack(fill=X) #afficher sur tkiner
        self.bouton = Button(self.frame_home, text="Parcourir...", command=self.chose_file) #création de bouton
        self.bouton.pack(side=RIGHT, padx=5, pady=5)  #afficher le bouton et le placer
        self.entree = Entry(self.frame_home, textvariable=self.value_input) #création de l'entrer
        self.entree.pack(fill=X, padx=5, expand=True) #afficher le champ et le placer (expend true (si plus grand va s'adapter))

        self.frame_tab = Frame(self, borderwidth=2, relief=GROOVE) #création du contour de la fenetre
        self.frame_tab.pack(fill=BOTH, expand=True) #afficher le champ et le placer (expend true (si plus grand va s'adapter))
        self.tableau = Treeview(self.frame_tab) #treeview documentation
        self.tableau['show'] = 'headings' #titre

        #Ajout de la scrollbar
        self.hsb = Scrollbar(self.frame_tab, orient="horizontal", command=self.tableau.xview) #On la place dans self.frame_tab au sens horizontal
        self.hsb.pack(side=BOTTOM, fill=X) #on place la barre en bas avec "X"(ordonnée) en horizontal
        self.vsb = Scrollbar(self.frame_tab, orient="vertical", command=self.tableau.yview) #On la place dans self.frame_tab au sens vertical
        self.vsb.pack(side=RIGHT, fill=Y) #on place la barre a droite avec "Y"(abscisse) en vertical
        self.tableau.configure(xscrollcommand=self.hsb.set, yscrollcommand=self.vsb.set) #on les configure sur le tableau

        self.tableau.pack(side=TOP, fill=BOTH, expand=True) #afficher le tableau (expend true (si plus grand va s'adapter))
        self.frame_param = LabelFrame(self, text="Paramètre de conversion") #parametre avec un nom
        self.frame_param.pack(fill=X, expand=True) #placer
        self.listed = ttk.Combobox(self.frame_param, values=["Points", "Lignes", "Polygones"]) #liste d'élément
        self.listed.pack(padx=5, pady=5) #placer

        self.frame_out = LabelFrame(self, text="Paramètre de sortie") #cadre avec du txt
        self.frame_out.pack(fill=X, side=BOTTOM) #placer
        self.sortie = Entry(self.frame_out, textvariable=self.value_output) #entry sortie
        self.sortie.pack(fill=X, padx=5, expand=True) #placer

        self.ok = Button(self.frame_out, text="Executer", command=self.action) #bouton executer
        self.ok.pack(side=LEFT, padx=5, pady=5) #placer





    def chose_file(self):
        self.path = askopenfilename(title="Ouverture fichier excel",
                              filetypes=[('Excel files', '*.xlsx; *.xls'), ('excel files', '*.xlsx ;*.xls')]) #ouvrir un fichier
        self.value_input.set(self.path) # value_imput devient le chemin

        head = self.convert_tool.load_xls_file(self.path) #appel a la fonction excel (stocker dans head)
        self.tableau["columns"] = head #création d'un tableau tableau avec les valeurs de la premiere ligne dans les entete
        self.tableau.column("#0", stretch=True) #placer (stretch: largeur de chaque colone
        for name_col in head:   # parcourir la liste des colones et on récupère chaque nom
            self.tableau.column(name_col, stretch=True) #on rempli l'entete
            self.tableau.heading(name_col, text=name_col) #on rempli l'entete titre des colones

        for i in range(1, 25): #afficher les 10er valeur du tableau dans le tableau
            self.tableau.insert("", "end", "", values=self.convert_tool.sheet.row_values(i)) # l'inserer dans le tableau

        self.destination_file, destination_ext = os.path.splitext(self.path) #on met a jour la destination quand on met le chemin de depart et on enleve l'extention
        self.value_output.set(self.destination_file) #mise a jour de la valeur

    def action(self): # executer
        try:
            self.convert_tool.create_shape(self.sortie.get(), self.listed.get()) #créaction du shape avec les parametre entre sortie et type de shape
        except TypeError as err: #si exeption faire:
            askokcancel("Attention", *err.args) # boite de dialogue avec 2 boutons
            return #stop

        self.convert_tool.convert_file() #convertion du fichier
        print("Success")


fenetre = Tk() #on instantie tk
interface = interface_xlstoshape(fenetre) #on instantie l'inerface xlsshap avec le parametre fenetre

interface.mainloop() #mainloop
interface.destroy() #fermer la fenetre apres