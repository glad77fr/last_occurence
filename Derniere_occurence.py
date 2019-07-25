from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
#from tkinter.filedialog import *
import tkinter
import pandas as pd
import json
import tkinter.ttk as ttk
from Last_occurence import loading, last_occurence

class Application(Tk):

    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
        self.title("Dernière occurence")
        self.resizable(width=False, height=False)
        self.load_data()
        self.mainloop()

    def initialize(self):
        self.rep_source_affiche = StringVar()
        self.rep_cible_affiche = StringVar()
        self.liste_champs = Listbox()
        self.liste_val = StringVar()

        self.control_menu = Frame(self)
        self.control_menu.pack()

        self.conf_menu = LabelFrame(self.control_menu, text="Configuation", width=430, height=300, labelanchor='n')
        self.conf_menu.grid(row=1, column=1, pady=10, padx=10)

        self.conf_data = LabelFrame(self.control_menu, text="Configuation données", width=43, height=300, labelanchor='n')
        self.conf_data.grid(row=2, column=1, pady=10, padx=10)

        self.entree_label = Label(self.conf_menu, text="Fichier source")
        self.entree_label.grid(row=1, column=1, padx=20,sticky=W)

        self.entree = Entry(self.conf_menu, textvariable=self.rep_source_affiche, width=50)
        self.entree.grid(row=1, column=2,sticky=W, padx=0)

        self.imp_Button = Button(self.conf_menu, text="Sélection...", command=self.import_file)
        self.imp_Button.grid(row=1, column=3, padx=15, pady=10,sticky=W)

        self.cible_label = Label(self.conf_menu, text="Répertoire cible")
        self.cible_label.grid(row=2,column=1, padx=20,sticky=W)

        self.cible_entree = Entry(self.conf_menu, textvariable=self.rep_cible_affiche, width=50)
        self.cible_entree.grid(row=2, column=2, sticky=W, padx=0)

        self.cible_Button = Button(self.conf_menu, text="Sélection...", command = self.rep_cible)
        self.cible_Button.grid(row=2, column=3, padx=15, pady=10,sticky=W)

        self.ex_menu = LabelFrame(self.control_menu, text="Chargement", width=430, height=200, labelanchor='n')
        self.ex_menu.grid(row=1, column=2,padx=10, pady=10)

        self.ex_Button = Button(self.ex_menu, text="Charger", command = self.chargement,height = 1, width = 18)
        self.ex_Button.grid(row=3, column=1, padx=10, pady=10,sticky=W)

        self.quitButton = Button(self.ex_menu, text="Quitter", command=self.destroy,height = 1, width = 18)
        self.quitButton.grid(row=4, column=1, padx=10, pady=10,sticky=W)

        self.mat = ttk.Combobox(self.conf_data, values=self.liste_champs,width=30, height=1)
        self.mat.grid(row=4, column=2, padx=10, pady=10, sticky=W)

        self.deb = ttk.Combobox(self.conf_data, values=self.liste_champs, width=30, height=1)
        self.deb.grid(row=5, column=2, padx=10, pady=10, sticky=W)

        self.lab_mat = Label(self.conf_data, text="Champ dernière occurence")
        self.lab_mat.grid(row=4,column=1, padx=20,sticky=W)

        self.date_deb = Label(self.conf_data, text="Date occurence")
        self.date_deb.grid(row=5, column=1, padx=20, sticky=W)

        self.ex_menu = LabelFrame(self.control_menu, text="Execution", width=200, height=100, labelanchor='n')
        self.ex_menu.grid(row=2, column=2, padx=20, pady=10)

        self.ex_Button = Button(self.ex_menu, text="Executer", command=self.execution, height=1, width=18)
        self.ex_Button.grid(row=2, column=2, padx=10, pady=10, sticky=W)

    def import_file(self):
        rep = filedialog.askopenfilename(initialdir="/", title="Select file",
                                         filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if rep != "":
            self.rep_source_affiche.set(rep)
            self.source = pd.read_excel(rep)
            try:
                self.save_data["source"] = rep
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)
            except:
                self.save_data.update({"Source": rep})
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)

    def rep_cible(self):

        self.rep_cible = filedialog.askdirectory()
        if self.rep_cible != "":
            self.rep_cible_affiche.set(self.rep_cible)

            try:
                self.save_data["cible"] = self.rep_cible
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)
            except:
                self.save_data.update({"Cible": self.rep_cible })
                with open('data.txt', 'w+') as outfile:
                    json.dump(self.save_data, outfile)

    def chargement(self):

        if self.rep_source_affiche.get() == "":
            messagebox.showerror("Ficher source non sélectionné", "Veuillez sélectionner un fichier source")

        if self.rep_cible_affiche.get() =="":
            messagebox.showerror("Répertoire source non sélectionné", "Veuillez sélectionner un répertoire source")

        chargement = loading(self.rep_source_affiche.get())
        maliste = [x for x in chargement[0]]
        self.mat["values"] = maliste
        self.deb["values"] = maliste

    def execution(self):
        if self.mat.get() != "":
            mat = self.mat.get()
        else:
            messagebox.showerror("Données non sélectionnées", "Veuillez sélectionner un matricule")
            return

        if self.deb.get() != "":
            date_ref = self.deb.get()
        else:
            messagebox.showerror("Données non sélectionnées", "Veuillez sélectionner un champ date")
            return

        try:  # Chargement des colonnes
            chargement = loading(self.rep_source_affiche.get())
        except:
            messagebox.showinfo("Problème de chargement", "Problème de chargement du fichier, assurez vous que le fichier est bien un fichier Excel ou CSV")
            return

        try :
            resultat = last_occurence(chargement[1],mat,date_ref)
            writer = pd.ExcelWriter(self.rep_cible_affiche.get() + "/export.xlsx", engine='xlsxwriter',
                                        date_format='dd/mm/yyyy')
            resultat.to_excel(writer, index=False)
            writer.save()
            messagebox.showinfo("Chargement", "Le fichier export a été créé et déposé dans le répertoire " + self.rep_cible_affiche.get() )
        except :
            messagebox.showerror("Problème de format de données","Le champ date selectionné doit être sous le format dd/mm/yyyy")

    def load_data(self):
        try:
            with open('data.txt') as json_file:
                self.save_data = json.load(json_file)
            try:
                self.rep_cible_affiche.set(self.save_data["cible"])
            except:
                pass
            try:
                self.rep_source_affiche.set(self.save_data["source"])
            except:
                pass
        except:
            self.save_data={}

toto = Application(None)

toto.mainloop()