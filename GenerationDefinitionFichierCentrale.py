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

BDD = "DataLake"
SERVEUR =  "dev04bi1-sql.ccq.org"
LIEN_SORTIE_FICHIER = ".\\"
conn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
    "Server=" + SERVEUR + ";"
    "Database=" + BDD + ";"
    "Trusted_Connection=yes;")
cursor = conn.cursor()

def checkIDSystemeSQL(P_Requete):
    retour = 0
    conn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
        "Server=" + SERVEUR + ";"
        "Database=" + BDD + ";"
        "Trusted_Connection=yes;")
    cursor = conn.cursor()
    cursor.execute(P_Requete)
    for row in cursor:
        retour = row[0]
    if retour > 0:
        return True
    return False


def executionScript(P_FichierScriptTable):
    try:
        # conn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
        #     "Server=" + SERVEUR + ";"
        #     "Database=" + BDD + ";"
        #     "Trusted_Connection=yes;")
        # cursor = conn.cursor()

        with open(P_FichierScriptTable, 'r') as content_file:
            # cursor.execute('DROP TABLE IF EXISTS [dbo].[XLADAVI]')
            requete = content_file.read()
            requete = requete.replace(' GO', ' ;\n')
            requete = requete.replace('\n', ' ')
            cursor.execute(requete)
            conn.commit()

    except pyodbc.Error as err:
        print(err)
        sys.exit(1)
    except:
        sys.exit(1)


def recuperationNomTables(P_FichierScriptTable):
    try:
        ListeTables = []
        with open(P_FichierScriptTable, 'r') as content_file:
            requete = content_file.read()
        
        regex = 'CREATE TABLE \[[dboDBO]+\].\[([\w]+)\]'
        
        ListeTables = re.findall(re.compile(regex), requete)

        return ListeTables
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def recuperationJsonModele(P_ListeTables):
    try:
        ListeModeles = []

        for NomTable in P_ListeTables:
            # Ancien systeme: On prenait les declarations Centrale direct
            # requete = "SELECT COLUMN_NAME AS Nom , CASE WHEN DATA_TYPE = 'varchar' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX)) ELSE 'max' END + ')' WHEN DATA_TYPE = 'char' THEN 'char(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE = 'nvarchar' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 THEN CAST(CHARACTER_MAXIMUM_LENGTH AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE = 'numeric' THEN 'numeric(' + CAST(NUMERIC_PRECISION AS nvarchar(MAX)) + ',' + CAST(NUMERIC_SCALE AS nvarchar(MAX)) + ')' ELSE DATA_TYPE END AS Type , CASE WHEN IS_NULLABLE = 'NO' THEN 'NOT NULL' ELSE 'NULL' END AS Contrainte FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME IN ('{0}') FOR JSON PATH".format(NomTable)
            
            # Nouveau systeme: On augmente la taille des champ
            requete = "SELECT COLUMN_NAME AS Nom , CASE WHEN DATA_TYPE like '%%char%%' THEN 'nvarchar(' + CASE WHEN CHARACTER_MAXIMUM_LENGTH > 0 AND CHARACTER_MAXIMUM_LENGTH < 50 THEN CAST(500 AS nvarchar(MAX)) WHEN CHARACTER_MAXIMUM_LENGTH < 100 THEN CAST(1000 AS nvarchar(MAX))  ELSE 'max' END + ')' WHEN DATA_TYPE like '%%int%%' THEN 'bigint' WHEN DATA_TYPE = 'numeric' THEN 'NUMERIC(38,6)' ELSE DATA_TYPE END AS Type , CASE WHEN IS_NULLABLE = 'NO' THEN 'NOT NULL' ELSE 'NULL' END AS Contrainte FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME IN ('{0}') FOR JSON PATH".format(NomTable)
            cursor.execute(requete)
            rows = cursor.fetchall()
            ligne = ''
            for row in rows:
                ligne += row[0]
            
            ListeModeles.append(ligne)
            print(NomTable)
            print(ligne)
            print(requete)
            print('-----')
        return ListeModeles
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def generationScriptDefinitionObjetInformation_Config(P_ListeTables, P_ListeModeles, NomDomaine, NomRepertoire, Delimiteur):
    try:
        workbook = xlsxwriter.Workbook('BI2020_Modele_{0}_{1}_DefinitionObjetInformation_Config.xlsx'
                                       .format(NomDomaine, NomRepertoire))
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
            IDTypeObjetInformation = 'CSV'
            Masque = '[\w.]*{0}[\w.]*[.]{{1}}(TXT|txt)'.format(NomTable)
            # ParametreType = "'[{\"NbLignesEntete\":0,\"NbLignesPiedPage\":0}]'"
            ValeurEntete = P_ListeModeles[i]


            my_list.append(DefinitionObjetInformation)
            my_list.append(Description)
            my_list.append(IDTypeObjetInformation)
            my_list.append(NomDomaine)
            my_list.append(Masque)
            my_list.append('[{"NbLignesEntete":0,"NbLignesPiedPage":0,"Delimiteur":"{0}"}]'.format(Delimiteur))
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
        #     requete += "INSERT INTO [dataLake].[DefinitionObjetInformation_Config] ([DefinitionObjetInformation]\t,[Description]\t,[IDTypeObjetInformation]\t,[IDDomaine]\t,[Masque]\t,[InformationSpecifiqueType]\t,[EnteteExterne], [Defaut]\t, [InsertionDifferentielSeulement]\t)\tVALUES\t('{0}'\t,'{1}'\t,{2}\t,{3}\t,'{4}'\t,{5}\t,'{6}'\t, 0, 0)".format(
        #         DefinitionObjetInformation,
        #         Description,
        #         IDTypeObjetInformation,
        #         IDDomaine,
        #         Masque,
        #         ParametreType,
        #         ValeurEntete) + '\n'
        #     i += 1
        # text_file = open(LIEN_SORTIE_FICHIER + "script.sql", "w+", encoding="utf-8")
        # text_file.write(requete)
        # text_file.close()
        workbook.close()
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)


def dropTable(P_ListeTables):
    try:
        # conn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
        #     "Server=" + SERVEUR + ";"
        #     "Database=" + BDD + ";"
        #     "Trusted_Connection=yes;")
        # cursor = conn.cursor()

        for NomTable in P_ListeTables:
            requete = "DROP TABLE IF EXISTS {0};".format(NomTable)
            print(requete)
            cursor.execute(requete)
            conn.commit()
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

    while True:
        Delimiteur = input("Entrer le délimiteur de ces fichiers CSV:"+ '\n')
        try: 
            if len(Delimiteur) > 0:
                break
            else:
                print("Le nom n'est pas valide")    
        except ValueError:
            print('Entrer un nom')

    # --1- Entrez le fichier SCRIPT TABLE
    FichierScriptTable = filedialog.askopenfilename()
    
    # --2- On lit le fichier et on l'execute en localhost
    executionScript(FichierScriptTable)

    # --3- On fait le regex pour avoir tout les noms de tables
    ListeTables = []
    ListeTables = recuperationNomTables(FichierScriptTable)
    
    # --4- Pour chaque nom de tables on prend son schéma JSON
    ListeModeles = []
    ListeModeles = recuperationJsonModele(ListeTables)

    # --5- On fait l'insert into dynamique
    generationScriptDefinitionObjetInformation_Config(ListeTables, ListeModeles, NomDomaine, NomRepertoire, Delimiteur)

    # --6- Truncate Table des tables
    dropTable(ListeTables)

if __name__ == "__main__":
    main()
    