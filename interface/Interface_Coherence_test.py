from tkinter import tix
from tkinter.filedialog import *
import shapefile
import utilsgio.Coherence_test



class interface_coherence(Frame):# Création de la fenêtre principale (main window)
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=100, height=100,**kwargs)
        self.pack(fill=BOTH, expand=True)

#initialisation
        self.listefield1=[]
        self.listefield2=[]
        self.listefield3=[]
        self.listefield4=[]
        self.listefield5=[]
        self.listefield55 = []
        self.value_input_zsro = StringVar()
        self.value_input_pfsro = StringVar()
        self.value_input_zpbo = StringVar()
        self.value_input_pfpbo = StringVar()
        self.value_input_piq = StringVar()
        self.value_input_zsro.set("Charger la Zone Arrière SRO")
        self.value_input_pfsro.set("Charger les emplacements SRO")
        self.value_input_zpbo.set("Charger la Zone arrière PBO")
        self.value_input_pfpbo.set("Charger les emplacements PBO")
        self.value_input_piq.set("Charger le Piq_BAL")
        self.varcombo1 = tix.StringVar()
        self.varcombo2 = tix.StringVar()
        self.varcombo3 = tix.StringVar()
        self.varcombo4 = tix.StringVar()
        self.varcombo5 = tix.StringVar()
        self.varcombo55 = tix.StringVar()
        self.varcombo1.set('choisir "id_métier_"')
        self.varcombo2.set('choisir "id_métier_"')
        self.varcombo3.set('choisir "id_métier_"')
        self.varcombo4.set('choisir "id_métier_"')
        self.varcombo5.set('choisir "id_métier"')
        self.varcombo55.set('choisir "nb_logemen"')

        self.champs_choisi1 = StringVar()
        self.champs_choisi2 = StringVar()
        self.champs_choisi3 = StringVar()
        self.champs_choisi4 = StringVar()
        self.champs_choisi5 = StringVar()
        self.champs_choisi55 = StringVar()
        self.path1 = StringVar()
        self.path2 = StringVar()
        self.path3 = StringVar()
        self.path4 = StringVar()
        self.path5 = StringVar()
        self.coherence = utilsgio.Coherence_test.coherence()
        self.loadShp = None
        self.loadShp2 = None
        self.loadShp3 = None
        self.loadShp4 = None
        self.loadShp5 = None

# création des Frames dans la fenêtre principale
        self.Frame1 = Frame(self, relief=GROOVE)
        self.Frame2 = Frame(self, relief=GROOVE)
        self.Frame3 = Frame(self, relief=GROOVE)
        self.Frame4 = Frame(self, relief=GROOVE)
        self.Frame5 = Frame(self, relief=GROOVE)
# Placement des frames
        self.Frame1.place(x=20, y=00, width=900, height=40)
        self.Frame2.place(x=20, y=40, width=900, height=40)
        self.Frame3.place(x=20, y=80, width=900, height=40)
        self.Frame4.place(x=20, y=120, width=900, height=40)
        self.Frame5.place(x=20, y=160, width=900, height=40)
        #Frame exécution
        self.Frame6 = Frame(self, relief=GROOVE)
        self.Frame6.place(x=180, y=200, width=300, height=40)
# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 1 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree1 = Entry(self.Frame1, textvariable=self.value_input_zsro)  # création de l'entrer
        self.entree1.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))
        self.combo1 = tix.ComboBox(self.Frame1, editable=1, dropdown=1, variable=self.varcombo1, command=self.Affiche1)
        self.combo1.entry.config(state='readonly')  ## met la zone de texte en lecture seule



        # bouton1 parcourir
        self.bouton1 = Button(self.Frame1, text="Parcourir...", command=self.chose_file1)  # création de bouton
        self.bouton1.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 2 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree2 = Entry(self.Frame2, textvariable=self.value_input_pfsro)  # création de l'entrer
        self.entree2.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))
        # liste déroulante
        self.combo2 = tix.ComboBox(self.Frame2, editable=1, dropdown=1, variable=self.varcombo2, command=self.Affiche2)
        self.combo2.entry.config(state='readonly')  ## met la zone de texte en lecture seule

        # bouton parcourir
        self.bouton2 = Button(self.Frame2, text="Parcourir...", command=self.chose_file2)  # création de bouton
        self.bouton2.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 3 constructeur
#-----------------------------------------------------------------------------------------------------------------------
        self.entree3 = Entry(self.Frame3, textvariable=self.value_input_zpbo)  # création de l'entrer
        self.entree3.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))
        # liste déroulante
        self.combo3 = tix.ComboBox(self.Frame3, editable=1, dropdown=1, variable=self.varcombo3, command=self.Affiche3)
        self.combo3.entry.config(state='readonly')  ## met la zone de texte en lecture seule

        # bouton parcourir
        self.bouton3 = Button(self.Frame3, text="Parcourir...", command=self.chose_file3)  # création de bouton
        self.bouton3.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer


# -----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 4 constructeur
# -----------------------------------------------------------------------------------------------------------------------
        self.entree4 = Entry(self.Frame4, textvariable=self.value_input_pfpbo)  # création de l'entrer
        self.entree4.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))
        # liste déroulante
        self.combo4 = tix.ComboBox(self.Frame4, editable=1, dropdown=1, variable=self.varcombo4, command=self.Affiche4)
        self.combo4.entry.config(state='readonly')  ## met la zone de texte en lecture seule

        # bouton parcourir
        self.bouton4 = Button(self.Frame4, text="Parcourir...", command=self.chose_file4)  # création de bouton
        self.bouton4.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer
# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 5 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree5 = Entry(self.Frame5, textvariable=self.value_input_piq)  # création de l'entrer
        self.entree5.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # liste déroulante

        self.combo5 = tix.ComboBox(self.Frame5, editable=1, dropdown=1, variable=self.varcombo5, command=self.Affiche5)
        self.combo5.entry.config(state='readonly')  ## met la zone de texte en lecture seule

        self.combo55 = tix.ComboBox(self.Frame5, editable=1, dropdown=1, variable=self.varcombo55, command=self.Affiche55)
        self.combo55.entry.config(state='readonly')  ## met la zone de texte en lecture seule

        # bouton parcourir
        self.bouton5 = Button(self.Frame5, text="Parcourir...", command=self.chose_file5)  # création de bouton
        self.bouton5.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

# -----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 6 constructeur
# -----------------------------------------------------------------------------------------------------------------------
    # bouton parcourir

        self.bouton6 = Button(self.Frame6, text="Lancer l'analyse", command=self.coherence.analyse)  # création de bouton
        self.bouton6.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 1 fonction
#-----------------------------------------------------------------------------------------------------------------------


#Affectation de valeur liste déroulante
    def Affiche1(self, evt):
        self.champs_choisi1 = self.varcombo1.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield1.index(self.champs_choisi1) #récupération de l'index du champs choisi
        self.loadShp = self.coherence.loadShp(self.path1, "zsro")# récupération des données en fonction de l'index du champs choisi

#Définition des chemins
    def chose_file1(self):
        self.path1 = askopenfilename(title="Chagement des Zones arrières SRO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des Zones arrières SRO
        self.value_input_zsro.set(self.path1)  # value_imput devient le chemin

        self.listefield1 = self.coherence.loadShpfile1(self.path1)
        self.combo1.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(self.listefield1))#effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield1)):
            self.combo1.insert(i, self.listefield1[i])  # Liste est la liste qui récupère les champs
            i+=1

        self.combo1.pack(side=LEFT, padx=5, pady=5, expand=False)



#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 2 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Affectation de valeur liste déroulante
    def Affiche2(self, evt):
        self.champs_choisi2=self.varcombo2.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield2.index(self.champs_choisi2) #récupération de l'index du champs choisi
        self.loadShp2 = self.coherence.loadShp2(self.path2, "zsro")# récupération des données en fonction de l'index du champs choisi

#Définition des chemins
    def chose_file2(self):
        self.path2 = askopenfilename(title="Chagement des emplacements SRO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier emplacement des SRO
        self.value_input_pfsro.set(self.path2)  # value_imput devient le


        self.listefield2 = self.coherence.loadShpfile2(self.path2)
        self.combo2.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(self.listefield2))#effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield2)):
            self.combo2.insert(i, self.listefield2[i])  # Liste est la liste qui récupère les champs
            i +=1
        self.combo2.pack(side=LEFT, padx=5, pady=5)

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 3 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Affectation de valeur liste déroulante
    def Affiche3(self, evt):
        self.champs_choisi3=self.varcombo3.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield3.index(self.champs_choisi3) #récupération de l'index du champs choisi
        self.loadShp3 = self.coherence.loadShp3(self.path3, "zpbo")# récupération des données en fonction de l'index du champs choisi

#Définition des chemins
    def chose_file3(self):
        self.path3 = askopenfilename(title="Chagement des Zones arrières PBO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des Zones arrières PBO
        self.value_input_zpbo.set(self.path3)  # value_imput devient le chemin

        self.listefield3 = self.coherence.loadShpfile3(self.path3)
        self.combo3.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(self.listefield3))  # effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield3)):
            self.combo3.insert(i, self.listefield3[i])  # Liste est la liste qui récupère les champs
            i+=1
        self.combo3.pack(side=LEFT, padx=5, pady=5)

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 4 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Affectation de valeur liste déroulante
    def Affiche4(self, evt):
        self.champs_choisi4=self.varcombo4.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield4.index(self.champs_choisi4) #récupération de l'index du champs choisi
        self.loadShp4 = self.coherence.loadShp4(self.path4, "zpbo")# récupération des données en fonction de l'index du champs choisi


#Définition des chemins
    def chose_file4(self):
        self.path4 = askopenfilename(title="Chagement des emplacements PBO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des emplacement des PBO
        self.value_input_pfpbo.set(self.path4)  # value_imput devient le chemin

        self.listefield4 = self.coherence.loadShpfile4(self.path4)
        self.combo4.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(self.listefield4))  # effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield4)):
            self.combo4.insert(i, self.listefield4[i])  # Liste est la liste qui récupère les champs
            i+=1
        self.combo4.pack(side=LEFT, padx=5, pady=5)

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 5 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Affectation de valeur liste déroulante
    def Affiche5(self, evt):
        self.champs_choisi5=self.varcombo5.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield5.index(self.champs_choisi5)  # récupération de l'index du champs choisi
        self.loadShp5 = self.coherence.loadShp5(self.path5)
        print(self.loadShp5)
    def Affiche55(self, evt):
        self.champs_choisi55 = self.varcombo55.get()  # On affecte le champs a "champs_choisi" et on affiche a l'ecran la valeur selectionnee
        self.coherence.nb_champ = self.listefield55.index(self.champs_choisi55)  # récupération de l'index du champs choisi
        self.loadShp55 = self.coherence.loadShp55(self.path5)
        print(self.loadShp55)

#Définition des chemins
    def chose_file5(self):
        self.path5 = askopenfilename(title="Chagement du Piq_BAL", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des emplacements des logements
        self.value_input_piq.set(self.path5)  # value_imput devient le chemin

        self.listefield5 = self.coherence.loadShpfile5(self.path5)
        self.combo5.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(self.listefield5))  # effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield5)):
            self.combo5.insert(i, self.listefield5[i])  # Liste est la liste qui récupère les champs
            i+=1
        self.combo5.pack(side=RIGHT, padx=5, pady=5)

        self.listefield55 = self.coherence.loadShpfile5(self.path5)
        self.combo55.subwidget_list['slistbox'].subwidget_list['listbox'].delete(0, len(
            self.listefield55))  # effacer le contenu de la combobox si on reclique
        for i in range(0, len(self.listefield55)):
            self.combo55.insert(i, self.listefield55[i])  # Liste est la liste qui récupère les champs
            i+=1
        self.combo55.pack(side=LEFT, padx=5, pady=5)



fenetre = tix.Tk()
interface = interface_coherence(fenetre)
fenetre.title('cohérence géographique')
fenetre.geometry('950x250')
interface.mainloop()