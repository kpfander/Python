import os
import uuid
import pyodbc
import io
from string import ascii_uppercase
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import sys
import re
from os import listdir
import pandas as pd
import json 

LIEN_SORTIE_FICHIER = ".\\"

def caseType(var):
    if var == 'int64':
        return 'int'
    elif var == 150:
        return 'numeric(9,2)'
    else:
        return 'nvarchar(4000)'

def recuperationJsonModele(Fichier, Feuille, IndicateurHeader):
    

    if(IndicateurHeader == '1'):
        df = pd.read_excel(Fichier, Feuille)
    else:
        df = pd.read_excel(Fichier, Feuille, header=None)

    df_types = df.iloc[1:].infer_objects()
    
    Modele = []

    for column in df:
        data = {}
        print(df[column])
        if(IndicateurHeader == '1'):
            data['Nom'] = column
        else:
            data['Nom'] = "F" + str(column + 1)

        data['Contrainte'] = 'NULL'
        data['Type'] = caseType(df_types[column].dtype.name)
        Modele.append(data)

    return json.dumps(Modele)

def ajoutScriptLigneDefinition(Fichier, Feuille, Modele, IDDomaine, IndicateurHeader):
    try:
        requete = ''
        DefinitionObjetInformation = Fichier
        Description = 'Valeurs généré via script Python pour le fichier Excel finissant par ' + Fichier + ' et dans la feuille ' + Feuille
        IDTypeObjetInformation = 5
        Masque = '{0}[.]{{1}}(XLSX|xlsx|XLS|xls)'.format(os.path.splitext(Fichier)[0])
        ParametreType = "'[{\"NomFeuille\":\"" + Feuille + "\", \"IndicateurHeader\":" + str(IndicateurHeader) + "}]'"
        ValeurEntete = Modele
        requete += "INSERT INTO [dataLake].[DefinitionObjetInformation_Config] ([DefinitionObjetInformation]\t,[Description]\t,[IDTypeObjetInformation]\t,[IDDomaine]\t,[Masque]\t,[InformationSpecifiqueType]\t,[EnteteExterne], [Defaut]\t,[InsertionDifferentielSeulement]\t)\tVALUES\t('{0}'\t,'{1}'\t,{2}\t,{3}\t,'{4}'\t,{5}\t,'{6}'\t, 0, 0)".format(
            DefinitionObjetInformation,
            Description,
            IDTypeObjetInformation,
            IDDomaine,
            Masque,
            ParametreType,
            ValeurEntete) + '\n'
        text_file = open(LIEN_SORTIE_FICHIER + "script.sql", "a", encoding="utf-8")
        text_file.write(requete)
        text_file.close()
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def find_csv_filenames( path_to_dir, suffix=".xlsx" ):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) or filename.endswith( ".xls" ) ]

def main():
    root = tk.Tk()
    root.withdraw()
    text_file = open(LIEN_SORTIE_FICHIER + "script.sql", "w+", encoding="utf-8")
    text_file.write('\n')
    text_file.close()
    while True:
        IDDomaine = input("Entrer le numéro de l'ID Domaine:"+ '\n')
        try: 
            if len(IDDomaine) > 0 and int(IDDomaine) >= 0 and int(IDDomaine) < 100:
                    break
            else:
                print("Le nom n'est pas valide")    
        except ValueError:
            print('Entrer un entier')

    DossierExcel = filedialog.askdirectory()

    ListeFichiers = find_csv_filenames(DossierExcel)
    
    for Fichier in ListeFichiers:
        xls = pd.ExcelFile(DossierExcel + '/' + Fichier)
        ListeFeuilles = xls.sheet_names
        for Feuille in ListeFeuilles:
            while True:
                IndicateurHeader = input("Présence d''une entete pour le fichier: " + Fichier + ' (0-Non, 1-Oui)\n')
                try: 
                    if len(IndicateurHeader) > 0 and int(IndicateurHeader) >= 0 and int(IndicateurHeader) <= 1:
                            break
                    else:
                        print("Le nom n'est pas valide")    
                except ValueError:
                    print('Entrer un entier')
                    
            Modele = recuperationJsonModele(DossierExcel + '/' + Fichier, Feuille, IndicateurHeader)
            ajoutScriptLigneDefinition(Fichier, Feuille, Modele, IDDomaine, IndicateurHeader)


if __name__ == "__main__":
    main()
    