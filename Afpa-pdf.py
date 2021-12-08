from tkinter import filedialog
from tkinter import *
import openpyxl
import tkinter as tk
import json
from tkinter.messagebox import askyesno

# parametrage de la fenetre
app = Tk()
app.geometry("720x500")
app["bg"] = "#77BA99"

# les tableau
moyenDeComunication= [
    "Aucun contact",
    "Métis",
    "Mail",
    "Téléphone",
    "Autres whatsapp, google drive ...",
    "métis + autres moyens",
    "Mail + autres moyens",
    "Téléphone + autres moyens"
]

competence = [
    
    "Période d'intégration DWWM",
    "Développer la partie front-end d’une application web ou web mobile en intégrant les recommandations de sécurité",
    "Développer la partie back-end d’une application web ou web mobile en intégrant les recommandations de sécurité",
    "Période Entreprise DWWM",
    "Période Certification DWWM",
    "Compétences transversales"
]

sequences = [
    "Séquence d'intégration",
    "Réaliser une interface utilisateur web statique et adaptable en effectuant une veille technologique en langue française ou anglaise",
    "Maquetter une application avec un contenu en langue française ou anglaise",
    "Développer une interface utilisateur web dynamique en effectuant une veille technologique y compris en anglais",
    "Réaliser une interface utilisateur avec une solution de gestion de contenu ou e-commerce en effectuant une veille technologique y compris en anglais",
    "Créer une base de données en effectuant une veille à partir de documentation en langue française ou anglaise",
    "Développer les composants d'accès aux données en recherchant éventuellement en langue anglaise des solutions innovantes",
    "Développer la partie back-end d'une application web ou web mobile en pratiquant une veille technologique y compris en anglais",
    "Développer des composants dynamiques en utilisant les bibliothèques d'une application de gestion de contenu en pratiquant une veille technologique y compris en anglais",
    "Période Entreprise D2WM",
    "Période Certification D2WM",
    "Compétences transversales"
]

# toutes les fonction
def file():

    def write():

        question = askyesno(title="Afpa-Pdf", message="êtes vous sur de vos infos ?")
        sheetParameters = file["1.Paramétrage"]
        columnParameters = sheetParameters["A"]
        count = 0
        for i in range(1227, len(columnParameters)):
            result = columnParameters[i].value
            if result != None:
                count = count + 1
            else:
                count = count

        count = 31 + count
        if question:

            sheetSelected = file[variable.get()]
            
            column = ["I","N", "S", "X", "AC"]
            columnCompetence = ["J", "O", "T", "Y", "AD"]
            columnSequence = ["K","P","U","Z","AE"]
            columnHours = ["M","R","AB"]
            columnwednesday = sheetSelected["w"]
            columnfriday = sheetSelected["AG"]

            
            for a in column:
                sheet = sheetSelected[a]
                for i in range(31,int(count)):
                   sheet[i].value = listcom.get()

            for z in columnCompetence:
                sheetCompetence = sheetSelected[z]
                for e in range(31,int(count)):
                   sheetCompetence[e].value = listcompetence.get()
            
            for r in columnSequence:
                sheetSequence = sheetSelected[r]
                for t in range(31,int(count)):
                   sheetSequence[t].value = listsequences.get()

            for y in columnHours:
                sheetHours = sheetSelected[y]
                for u in range(31,int(count)):
                   sheetHours[u].value = "8"
            
            for o in range(31,int(count)):
                
                columnwednesday[o].value = "7"

            for p in range(31,int(count)):
                columnfriday[p].value = "4"


            file.save(filepath)
            title = Label(app, text="Le fichier a bien était remplis", font=("Verdana", 15, "italic bold"), fg="red", bg="#77BA99").place(x='200', y='400')

            
            

    filepath = filedialog.askopenfilename()
    file = openpyxl.load_workbook(filepath)
    # recupération des feuilles
    onglet = file.sheetnames

    # menu deroulan de chaque listes
    listcom= StringVar()
    listcom.set(moyenDeComunication[0])
    opt = tk.OptionMenu(app, listcom,*moyenDeComunication)
    opt.config(width=10, font=('Helvetica', 8))
    opt.place(x='50', y='150')

    listcompetence= StringVar()
    listcompetence.set(competence[0])
    opt = tk.OptionMenu(app, listcompetence,*competence)
    opt.config(width=20, font=('Helvetica', 8))
    opt.place(x='180', y='150')

    listsequences= StringVar()
    listsequences.set(sequences[0])
    opt = tk.OptionMenu(app, listsequences,*sequences)
    opt.config(width=20, font=('Helvetica', 8))
    opt.place(x='370', y='150')

    variable = StringVar()
    variable.set(onglet[0])
    opt = tk.OptionMenu(app, variable, *onglet)
    opt.config(width=10, font=('Helvetica', 8))
    opt.place(x="580", y='150')
    # bouton valider
    valider = Button(app, text="Valider", bg="#7A28CB", bd="3", relief = "flat", font=("Verdana", 7, "bold"), fg="white", command=write).place(x='320', y='300')

# affichage du titre
title = Label(app, text="Suivi De Réalisation à Distance", font=("Verdana", 15, "italic bold"), fg="red", bg="#77BA99").place(x='200', y='30')

# appel la fonction file
file = Button(app, text="Veuillez choisir votre fichier .xlsx", bg="#7A28CB", bd="3", relief = "flat", font=("Verdana", 7, "bold"), fg="white", command=file).place(x='270', y='80')

app.mainloop()

