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
import xlsxwriter 

LIEN_SORTIE_FICHIER = ".\\"

def recuperationNomTables(P_FichierScriptTable):
    try:
        ListeAlterTables = []
        with open(P_FichierScriptTable, 'r') as content_file:
            requete = content_file.read()
        
        regex = 'ALTER ([\s\S]*?) GO'
        
        ListeAlterTables = re.findall(re.compile(regex), requete)

        return ListeAlterTables
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def generationScriptCleNaturelle_Config(P_ListeAlterTables, NomDomaine, NomRepertoire):
    try:
        workbook = xlsxwriter.Workbook('BI2020_Modele_{0}_{1}_CleNaturelle_Config.xlsx'
                                       .format(NomDomaine, NomRepertoire))
        worksheet = workbook.add_worksheet('CleNaturelle') 
        row = 0

        i = 0

        my_list = []

        my_list.append('Table')
        my_list.append('Champ')
        
        for col_num, data in enumerate(my_list):
                worksheet.write(row, col_num, data)

        row += 1

        for AlterTables in P_ListeAlterTables:
            ListeChamps = []
            
            regex = 'TABLE([\s\S]*?) WITH [\w\W]+PRIMARY KEY CLUSTERED\(([\w\W]+)\)ON'

            ContenuAlterTables = re.findall(re.compile(regex), AlterTables)

            if len(ContenuAlterTables) > 0:
                NomTable = ContenuAlterTables[0][0]
                ListeChamps = ContenuAlterTables[0][1].split("\n")

                NomTable = re.sub('[ ,\[\].]', '', NomTable)
                NomTable = NomTable.replace('DBO', '')
                NomTable = NomTable.replace('.', '')

                for NomChamp in ListeChamps:
                    my_list = []

                    NomChamp = re.sub('[ ,\[\].]', '', NomChamp)
                    
                    if NomChamp != '':
                        my_list.append(NomTable)
                        my_list.append(NomChamp)

                        for col_num, data in enumerate(my_list):
                            worksheet.write(row, col_num, data)

                        row += 1
        workbook.close()

    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def main():
    root = tk.Tk()
    root.withdraw()


    while True:
        NomDomaine = input("Entrer le nom du Domaine:"+ '\n')
        try: 
            if len(NomDomaine) > 0:
                break
            else:
                print("Le nom n'est pas valide")    
        except ValueError:
            print('Entrer un nom')
    
    while True:
        NomRepertoire = input("Entrer le nom du Repertoire:" + '\n')
        try:
            if len(NomRepertoire) > 0:
                break
            else:
                print("Le nom n'est pas valide")
        except ValueError:
            print('Entrer un nom')

    # --1- Entrez le fichier SCRIPT_INTEGRITE_TABLE
    FichierScriptTable = filedialog.askopenfilename()
    
    # --3- On fait le regex pour avoir tout les noms de tables
    ListeAlterTables = []
    ListeAlterTables = recuperationNomTables(FichierScriptTable)
    
    # --5- On fait l'insert into dynamique
    generationScriptCleNaturelle_Config(ListeAlterTables, NomDomaine, NomRepertoire)


if __name__ == "__main__":
    main()
    