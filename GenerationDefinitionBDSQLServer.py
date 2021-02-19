import os
import uuid
import pyodbc
import io
from string import ascii_uppercase
from datetime import datetime
import sys
import re
import json
import tkinter as tk
import xlsxwriter 

LIEN_SORTIE_FICHIER = ".\\"


def recuperationNomTables(BDD, SERVEUR, cursor):
    try:
        ListeTables = []
        
        requete = "select name from sys.tables"
        cursor.execute(requete)
        rows = cursor.fetchall()
        for row in rows:
            ListeTables.append(row[0])

        return ListeTables
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def recuperationJsonModele(P_ListeTables, BDD, SERVEUR, cursor):
    try:
        ListeModeles = []
        ListeLigne = []

        for NomTable in P_ListeTables:
            # Ancien systeme: On prenait les declarations Centrale direct
            # requete = "SELECT COLUMN_NAME AS Nom , CASE WHEN DATA_TYPE = 'varchar' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX)) ELSE 'max' END + ')' WHEN DATA_TYPE = 'char' THEN 'char(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE = 'nvarchar' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE = 'numeric' THEN 'numeric(' + CAST(NUMERIC_PRECISION AS nvarchar(MAX)) + ',' + CAST(NUMERIC_SCALE AS nvarchar(MAX)) + ')' ELSE DATA_TYPE END AS Type , CASE WHEN IS_NULLABLE = 'NO' THEN 'NOT NULL' ELSE 'NULL' END AS Contrainte FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME IN ('{0}')".format(NomTable)
            # Nouveau systeme: On augmente la taille des champ
            requete = "SELECT COLUMN_NAME AS Nom , CASE WHEN DATA_TYPE like '%%char%%' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 AND CHARACTER_MAXIMUM_LENGTH < 50 THEN CAST(500 AS nvarchar(MAX)) WHEN CHARACTER_MAXIMUM_LENGTH < 100 THEN CAST(1000 AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE like '%%int%%' THEN 'bigint' WHEN DATA_TYPE = 'numeric' THEN 'NUMERIC(38,6)' ELSE DATA_TYPE END AS Type , CASE WHEN IS_NULLABLE = 'NO' THEN 'NOT NULL' ELSE 'NULL' END AS Contrainte FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME IN ('{0}')".format(NomTable)
            cursor.execute(requete)
            rows = cursor.fetchall()
            ListeLigne = []
            for row in rows:
                ligne = {}
                ligne["Nom"] = row[0]
                ligne["Type"] = row[1]
                ligne["Contrainte"] = row[2]
                
                ListeLigne.append(ligne)
                
            ListeModeles.append(json.dumps(ListeLigne))
            
        return ListeModeles
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def generationScriptDefinitionObjetInformation_Config(P_ListeTables, P_ListeModeles, BDD, NomDomaine):
    try:
        workbook = xlsxwriter.Workbook('BI2020_Modele_{0}_{1}_DefinitionObjetInformation_Config.xlsx'
                                       .format(NomDomaine, BDD))
        worksheet = workbook.add_worksheet('DefinitionObjetInformation') 
        row = 0

        i = 0

        my_list = []

        my_list.append('DefinitionObjetInformation')
        my_list.append('Description')
        my_list.append('TypeObjetInformation')
        my_list.append('Domaine')
        my_list.append('Masque')
        my_list.append('InformationSpecifiqueType')
        my_list.append('EnteteExterne')
        my_list.append('Defaut')
        my_list.append('InsertionDifferentielSeulement')
        my_list.append('CodePage')
        my_list.append('SortieDonneesBrut')
        my_list.append('SortieDonneesJson')
        
        for col_num, data in enumerate(my_list):
                worksheet.write(row, col_num, data)

        row += 1

        # requete = ''
        for NomTable in P_ListeTables:
            my_list = []

            DefinitionObjetInformation = NomTable
            Description = 'Valeurs généré via script Python pour le fichier Centrale finissant par ' + NomTable
            IDTypeObjetInformation = 'MS SQL 2012 -'
            Masque = "(?i)\[{0}\].\[[\w]+\].\[{1}\]".format(BDD, NomTable)
            ParametreType = "'[{\"ClauseWhere\":\"\"}]'"
            ValeurEntete = P_ListeModeles[i]

            my_list.append(DefinitionObjetInformation)
            my_list.append(Description)
            my_list.append(IDTypeObjetInformation)
            my_list.append(NomDomaine)
            my_list.append(Masque)
            my_list.append(ParametreType)
            my_list.append(ValeurEntete)
            my_list.append(0)
            my_list.append(1)
            my_list.append(0)
            my_list.append(0)
            my_list.append(0)

            for col_num, data in enumerate(my_list):
                worksheet.write(row, col_num, data)

            row += 1
            i += 1
        workbook.close()
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def main():

    try:
        while True:
            SERVEUR = input("Entrer le nom du serveur:"+ '\n')
            try: 
                if len(SERVEUR) > 0:
                        break
                else:
                    print("Le nom n'est pas valide")    
            except ValueError:
                print('Entrer une chaine de caractere')

        # --1- Entrez le nom de la BD
        while True:
            BDD = input("Entrer le nom de la BD:"+ '\n')
            try: 
                if len(BDD) > 0:
                        break
                else:
                    print("Le nom n'est pas valide")    
            except ValueError:
                print('Entrer une chaine de caractere')

        # --1- Entrez le nom de la BD
        while True:
            NomDomaine = input("Entrer le nom du Domaine:"+ '\n')
            try: 
                if len(NomDomaine) > 0:
                    break
                else:
                    print("Le nom n'est pas valide")    
            except ValueError:
                print('Entrer un nom')

          
        # --2- On récupere les noms de tables
        conn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
        "Server=" + SERVEUR + ";"
        "Database=" + BDD + ";"
        "Trusted_Connection=yes;")
        cursor = conn.cursor()

        ListeTables = []
        ListeTables = recuperationNomTables(BDD, SERVEUR, cursor)
        
        # # --3- Pour chaque nom de tables on prend son schéma JSON
        ListeModeles = []
        ListeModeles = recuperationJsonModele(ListeTables, BDD, SERVEUR, cursor)

        generationScriptDefinitionObjetInformation_Config(ListeTables, ListeModeles, BDD, NomDomaine)


    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


if __name__ == "__main__":
    main()
    