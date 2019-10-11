from tkinter.filedialog import *



class main_interface(Frame):
    def __init__(self, fenetre, **kwargs):
        Frame.__init__(self, fenetre, width=768, height=596, **kwargs)
        self.pack(fill=BOTH, expand=True)
        self.photoAccueil = PhotoImage(file = r"C:\Users\m.rosillette\Documents\Apprentissage\Gio_Project\projet1\Home_black_50x.png")

        # Création des widgets
        self.frame1 = Frame(self)
        self.frame1.pack(fill=X)
        self.boutonAccueil = Button(self.frame1, image=self.photoAccueil)
        self.boutonAccueil.pack(side=RIGHT)

        self.frame2 = Frame(self)
        self.frame2.pack(fill=X)
        self.bouton = Button(self.frame2, text="Convertir Xls to Shape")
        self.bouton.pack()

fenetre = Tk()
interface = main_interface(fenetre)
interface.mainloop()
interface.destroy()
