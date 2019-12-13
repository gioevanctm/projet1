from tkinter import tix
from tkinter.filedialog import *
import shapefile
import Coherence_test
import shapefile as sf


class interface_coherence(Frame):# Création de la fenêtre principale (main window)
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=100, height=100,**kwargs)
        self.pack(fill=BOTH, expand=True)

#initialisation
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

        self.path1 = StringVar()
        self.path2 = StringVar()
        self.path3 = StringVar()
        self.path4 = StringVar()
        self.path5 = StringVar()
        self.coherence = Coherence_test.coherence()
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
        self.Frame1.place(x=20, y=00, width=400, height=40)
        self.Frame2.place(x=20, y=40, width=400, height=40)
        self.Frame3.place(x=20, y=80, width=400, height=40)
        self.Frame4.place(x=20, y=120, width=400, height=40)
        self.Frame5.place(x=20, y=160, width=400, height=40)
        #Frame exécution
        self.Frame6 = Frame(self, relief=GROOVE)
        self.Frame6.place(x=5, y=200, width=422, height=40)
# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 1 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree1 = Entry(self.Frame1, textvariable=self.value_input_zsro)  # création de l'entrer
        self.entree1.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # bouton1 parcourir
        self.bouton1 = Button(self.Frame1, text="Parcourir...", command=self.chose_file1)  # création de bouton
        self.bouton1.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 2 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree2 = Entry(self.Frame2, textvariable=self.value_input_pfsro)  # création de l'entrer
        self.entree2.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # bouton parcourir
        self.bouton2 = Button(self.Frame2, text="Parcourir...", command=self.chose_file2)  # création de bouton
        self.bouton2.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 3 constructeur
#-----------------------------------------------------------------------------------------------------------------------
        self.entree3 = Entry(self.Frame3, textvariable=self.value_input_zpbo)  # création de l'entrer
        self.entree3.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # bouton parcourir
        self.bouton3 = Button(self.Frame3, text="Parcourir...", command=self.chose_file3)  # création de bouton
        self.bouton3.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer


# -----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 4 constructeur
# -----------------------------------------------------------------------------------------------------------------------
        self.entree4 = Entry(self.Frame4, textvariable=self.value_input_pfpbo)  # création de l'entrer
        self.entree4.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # bouton parcourir
        self.bouton4 = Button(self.Frame4, text="Parcourir...", command=self.chose_file4)  # création de bouton
        self.bouton4.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer
# ----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 5 constructeur
# ----------------------------------------------------------------------------------------------------------------------
        self.entree5 = Entry(self.Frame5, textvariable=self.value_input_piq)  # création de l'entrer
        self.entree5.pack(side=LEFT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))

        # bouton parcourir
        self.bouton5 = Button(self.Frame5, text="Parcourir...", command=self.chose_file5)  # création de bouton
        self.bouton5.pack(side=RIGHT, padx=5, pady=5)  # afficher le bouton et le placer

# -----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 6 constructeur
# -----------------------------------------------------------------------------------------------------------------------
    # bouton execution

        self.bouton6 = Button(self.Frame6, text="Lancer l'analyse", command=self.coherence.analyse)  # création de bouton

        self.bouton6.pack(side=LEFT, padx=5, pady=5)  # afficher le bouton et le placer

        self.entree66 = Entry(self.Frame6, textvariable=self.coherence.etat_prog, background="#16B84E")  # création de l'entrer
        self.entree66.pack(side=RIGHT, fill=X, padx=5, expand=True)  # afficher le champ et le placer (expend true (si plus grand va s'adapter))


    #def double_fonction(self):
        #self.coherence.etat_prog.set("En cours de création du Rapport excel veuillez patienter")
        #self.coherence.analyse()
#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 1 fonction
#-----------------------------------------------------------------------------------------------------------------------

#Définition des chemins
    def chose_file1(self):
        self.path1 = askopenfilename(title="Chagement des Zones arrières SRO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des Zones arrières SRO
        self.value_input_zsro.set(os.path.splitext(os.path.basename(self.path1))[0])  # découpé le chemin et l'extention pour laisser  que le nom du fichier et mettre dans entry


        self.coherence.shp_zsro = sf.Reader(self.path1)

        self.coherence.id_metier_zsro.clear()  # effacer en cas de reclique
        self.coherence.record_zsro.clear()     # effacer en cas de reclique
        self.coherence.shape_zsro.clear()      # effacer en cas de reclique
        for i in range(len(self.coherence.shp_zsro.records())):
            self.coherence.id_metier_zsro.append(self.coherence.shp_zsro.record(i)['id_metier_'])
            self.coherence.record_zsro.append(self.coherence.shp_zsro.record(i)) #récupère le record zsro
            self.coherence.shape_zsro.append(self.coherence.shp_zsro.shape(i))   #récupère le shape zsro


        self.coherence.etat_prog.set("Sélectionner le fichier PFSRO")
#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 2 fonction
#-----------------------------------------------------------------------------------------------------------------------
    def chose_file2(self):
        self.path2 = askopenfilename(title="Chagement des emplacements SRO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier emplacement des SRO
        self.value_input_pfsro.set(os.path.splitext(os.path.basename(self.path2))[0])  # découpé le chemin et l'extention pour laisser  que le nom du fichier et mettre dans entry

        self.coherence.shp_pfsro = sf.Reader(self.path2)

        self.coherence.id_metier_pfsro.clear()  # effacer en cas de reclique
        self.coherence.record_pfsro.clear()     # effacer en cas de reclique
        self.coherence.shape_pfsro.clear()      # effacer en cas de reclique
        for i in range(len(self.coherence.shp_pfsro.records())):
            self.coherence.id_metier_pfsro.append(self.coherence.shp_pfsro.record(i)['id_metier_'])
            self.coherence.record_pfsro.append(self.coherence.shp_pfsro.record(i)) #récupère le record pfsro
            self.coherence.shape_pfsro.append(self.coherence.shp_pfsro.shape(i))   #récupère le shape pfsro
        #print(self.coherence.id_metier_pfsro)
        self.coherence.etat_prog.set("Sélectionner le fichier ZPBO")
#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 3 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Définition des chemins
    def chose_file3(self):
        self.path3 = askopenfilename(title="Chagement des Zones arrières PBO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des Zones arrières PBO
        self.value_input_zpbo.set(os.path.splitext(os.path.basename(self.path3))[0]) # découpé le chemin et l'extention pour laisser  que le nom du fichier et mettre dans entry

        self.coherence.shp_zpbo = sf.Reader(self.path3)

        self.coherence.id_metier_zpbo.clear()  # effacer en cas de reclique
        self.coherence.record_zpbo.clear()     # effacer en cas de reclique
        self.coherence.shape_zpbo.clear()      # effacer en cas de reclique
        for i in range(len(self.coherence.shp_zpbo.records())):
            self.coherence.id_metier_zpbo.append(self.coherence.shp_zpbo.record(i)['id_metier_'])
            self.coherence.record_zpbo.append(self.coherence.shp_zpbo.record(i)) #récupère le record zpbo
            self.coherence.shape_zpbo.append(self.coherence.shp_zpbo.shape(i))   #récupère le shape zpbo
        #print(self.coherence.id_metier_zpbo)
        self.coherence.etat_prog.set("Sélectionner le fichier PFPBO")
#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 4 fonction
#-----------------------------------------------------------------------------------------------------------------------
#Définition des chemins
    def chose_file4(self):
        self.path4 = askopenfilename(title="Chagement des emplacements PBO", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des emplacement des PBO
        self.value_input_pfpbo.set(os.path.splitext(os.path.basename(self.path4))[0])  # découpé le chemin et l'extention pour laisser  que le nom du fichier et mettre dans entry

        self.coherence.shp_pfpbo = sf.Reader(self.path4)

        self.coherence.id_metier_pfpbo.clear()  # effacer en cas de reclique
        self.coherence.record_pfpbo.clear()     # effacer en cas de reclique
        self.coherence.shape_pfpbo.clear()      # effacer en cas de reclique
        for i in range(len(self.coherence.shp_pfpbo.records())):
            self.coherence.id_metier_pfpbo.append(self.coherence.shp_pfpbo.record(i)['id_metier_'])
            self.coherence.record_pfpbo.append(self.coherence.shp_pfpbo.record(i)) #récupère le record pfpbo
            self.coherence.shape_pfpbo.append(self.coherence.shp_pfpbo.shape(i))   #récupère le shape pfpbo

        #print(self.coherence.id_metier_pfpbo)
        self.coherence.etat_prog.set("Sélectionner le fichier du piquetage")
#-----------------------------------------------------------------------------------------------------------------------
#                                               FRAME 5 fonction
#-----------------------------------------------------------------------------------------------------------------------

#Définition des chemins
    def chose_file5(self):
        self.path5 = askopenfilename(title="Chagement du Piq_BAL", filetypes=[('ShapeFiles', '*.shp'), ('All files', '*.*')])  # ouvrir un fichier des emplacements des logements
        self.value_input_piq.set(self.path5)  # value_imput devient le chemin
        self.value_input_piq.set(os.path.splitext(os.path.basename(self.path5))[0])  # découpé le chemin et l'extention pour laisser  que le nom du fichier et mettre dans entry

        self.coherence.shp_piq = sf.Reader(self.path5)

        self.coherence.id_metier_piq.clear()  # effacer en cas de reclique
        self.coherence.record_piq.clear()     # effacer en cas de reclique
        self.coherence.shape_piq.clear()      # effacer en cas de reclique
        for i in range(len(self.coherence.shp_piq.records())):
            self.coherence.id_metier_piq.append(self.coherence.shp_piq.record(i)['id_metier_'])
            self.coherence.record_piq.append(self.coherence.shp_piq.record(i)) #récupère le record piq
            self.coherence.shape_piq.append(self.coherence.shp_piq.shape(i))   #récupère le shape piq
        #print(self.coherence.id_metier_piq)

        self.coherence.nb_logement_piq.clear()  # effacer en cas de reclique
        for i in range(len(self.coherence.shp_piq.records())):
            self.coherence.nb_logement_piq.append(self.coherence.shp_piq.record(i)['nb_logemen'])
        #print(self.coherence.nb_logement_piq)
        self.coherence.etat_prog.set("Appuyer sur le bouton <<lancer l'annalyse>> et patienter ")


    #def Affiche_Exception(self, MonErreur):
        #self.fenetre = Toplevel()
        #self.fenetre.title('Erreur')
        #self.Label(self.mafenetre, text=MonErreur)


fenetre = Tk()
logoctm=('lib\\logo.ico')
fenetre.iconbitmap(logoctm)
interface = interface_coherence(fenetre)
fenetre.title('cohérence géographique')
fenetre.geometry('450x250')
interface.mainloop()