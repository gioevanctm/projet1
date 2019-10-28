from tkinter import *
from tkinter import tix
import os.path
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter.ttk import Treeview
import os

fenetre = tix.Tk()
fenetre.title('test')
fenetre.geometry('700x250')
class interface_coherence(Frame):
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=0, height=0, **kwargs)
# Création de la fenêtre principale (main window)


#initialisation
        self.liste1=[1,2,3,4]
        self.liste2=[]
        self.liste3=[]
        self.liste4=[]
        self.liste5=[]
        self.value_input1 = StringVar()
        self.value_input2 = StringVar()
        self.value_input3 = StringVar()
        self.value_input4 = StringVar()
        self.value_input5 = StringVar()
        self.value_input1.set("Charger un fichier")
        self.value_input2.set("Charger un fichier")
        self.value_input3.set("Charger un fichier")
        self.value_input4.set("Charger un fichier")
        self.value_input5.set("Charger un fichier")
        self.varcombo1 = tix.StringVar()
        self.varcombo2 = tix.StringVar()
        self.varcombo3 = tix.StringVar()
        self.varcombo4 = tix.StringVar()
        self.varcombo5 = tix.StringVar()
        self.champs_choisi1 = StringVar()
        self.champs_choisi2 = StringVar()
        self.champs_choisi3 = StringVar()
        self.champs_choisi4 = StringVar()
        self.champs_choisi5 = StringVar()

# création des Frames dans la fenêtre principale
        self.Frame1 = Frame(fenetre, relief=GROOVE)
        self.Frame2 = Frame(fenetre,  relief=GROOVE)
        self.Frame3 = Frame(self.fenetre, relief=GROOVE)
        self.Frame4 = Frame(self.fenetre, relief=GROOVE)
        self.Frame5 = Frame(self.fenetre, relief=GROOVE)

# Placement des frames
        self.Frame1.place(x=20, y=00, width=650, height=40)
        self.Frame2.place(x=20, y=40, width=650, height=40)
        self.Frame3.place(x=20, y=80, width=650, height=40)
        self.Frame4.place(x=20, y=120, width=650, height=40)
        self.Frame5.place(x=20, y=160, width=650, height=40)

        #Frame exécution
        self.Frame6 = Frame(fenetre,relief=GROOVE)
        self.Frame6.place(x=200, y=200, width=100, height=40)

# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME1 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree1 = Entry(fenetre, textvariable=value_input1)  # création de l'entrer
        self.entree1.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))
        self.combo1 = tix.ComboBox(fenetre, editable=1, dropdown=1, variable=varcombo1, command=Affiche1)
        self.combo1.entry.config(state='readonly')  ## met la zone de texte en lecture seule
        for i in range(0, len(self.liste1)):
            self.combo1.insert(i, self.liste1[i])  # Liste est la liste qui récupère les champs
        self.combo1.pack(side=LEFT, padx=5, pady=5)

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME1
#-----------------------------------------------------------------------------------------------------------------------
#Affectation de valeur liste déroulante
    def Affiche1(self, evt):
        self.champs_choisi1 = self.varcombo1.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        print(self.champs_choisi1)

#Définition des chemins
    def chose_file1(self):
        self.path1 = askopenfilename(title="Ouverture fichier excel", filetypes=[('Excel files', '*.xlsx; *.xls')])  # ouvrir un fichier
        self.value_input1.set(self.path1)  # value_imput devient le chemin


#liste déroulante



#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME6
#-----------------------------------------------------------------------------------------------------------------------


#fekfekek




fenetre.mainloop()

