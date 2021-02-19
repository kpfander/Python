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
import json 
import xlrd
from slugify import slugify
import collections
import xlsxwriter 


LIEN_SORTIE_FICHIER = ".\\"


        

def correspondanceNomExcelNomTable():
    return 'FAIRE CORRESPONDANCE'

def nomVersNomenclatureColonne(value):
    value = slugify(value)

    value = value.replace('-','_')
    value = value.upper()

    return value

def caseType(var):
    if var == 'int64':
        return 'int'
    elif var == 150:
        return 'numeric(9,2)'
    else:
        return 'nvarchar(4000)'

def recuperationNomTable(P_excel):
    cell = P_excel.cell(9,5)
    return cell.value

def recuperationDescription(P_excel):
    cell1 = P_excel.cell(0,3)
    cell2 = P_excel.cell(5,9)

    description = cell1.value + ' | ' + cell2.value

    return description


def recuperationType(Longueur, Decimal, FormatCentral):
    if FormatCentral == 'Numérique':
        if str(Decimal).strip() == '':
            return 'nvarchar({0})'.format(str(Longueur * 2))
            # Du a la regle '-' à droite pour négatif, '+' à droite pour positif on le mets en varchar
            # return 'bigint'
        else:
            return 'nvarchar({0})'.format(str(Longueur * 2))
            # Du a la regle '-' à droite pour négatif, '+' à droite pour positif on le mets en varchar
            # return 'numeric({0}, {1})'.format(str(Longueur), str(int(Decimal)))
    elif FormatCentral == 'Charactère':
        return 'nvarchar({0})'.format(str(Longueur))
    elif FormatCentral == 'Alphanumérique':
        return 'nvarchar({0})'.format(str(Longueur))
    else:
        return 'nvarchar(4000)'

def gestionDuplicatNomColonnes(Colonnes):
    ColonnesRetour = Colonnes

    result = []
    # Generation des increments pour les doublons de noms
    for Colonne in Colonnes:
        fname = Colonne['Nom']
        orig = fname
        i=1
        while fname in result:
            fname = orig + str(i)
            i += 1
        result.append(fname)



    i = 0
    for ColonneRetour in ColonnesRetour:
        ColonneRetour['Nom'] = result[i]
        i = i + 1
    return Colonnes

def recuperationParametres(P_excel):
    ListeLigneAGerer = []
    colonnePositionDe = 5
    lignesDepart = 6
    nbLignes = P_excel.nrows

    FixedWidth = []
    Colonnes = []

    for lignePositionDe in range(lignesDepart,nbLignes):
        try:
            ValeurDe = P_excel.cell(lignePositionDe,colonnePositionDe).value
            # Si la valeur est présente dans Position 'De' prendre la ligne entière
            if ValeurDe is not None and ValeurDe != '' and int(ValeurDe) > 0:
                # print('Valeur est présente ! On récupère la ligne entière pour un traitement après !!')
                ListeLigneAGerer.append(P_excel.row_slice(rowx=lignePositionDe, start_colx=0, end_colx=9))
                # print(P_excel.row_slice(rowx=lignePositionDe, start_colx=0, end_colx=9))
        except ValueError:
            continue
    
    # On va gérer créer nos Déclaration de Colonne, pour chaque elements
    for LigneAGerer in ListeLigneAGerer:
        # 0- Récupération variable
        NomBrut = LigneAGerer[1].value
        Longueur = int(LigneAGerer[2].value)
        Decimal = LigneAGerer[3].value
        FormatCentral = LigneAGerer[4].value
        PositionDe = int(LigneAGerer[5].value)
        PositionA = int(LigneAGerer[6].value)
        
        # 1- Gestion du nom de colonne
        Nom = nomVersNomenclatureColonne(NomBrut)
        # 2- Gestion du Type
        Type = recuperationType(Longueur, Decimal, FormatCentral)
        # 3- Gestion de la contrainte
        Contrainte = 'NULL'
        

        dataColonne = {}
        dataColonne['Nom'] = Nom
        dataColonne['Type'] = Type
        dataColonne['Contrainte'] = Contrainte
        Colonnes.append(dataColonne)

        # 4- Gestion du fixed width
        dataDebutFin = {}
        dataDebutFin['Debut'] = PositionDe
        dataDebutFin['Fin'] = PositionA
        FixedWidth.append(dataDebutFin)


    dataFixedWidth = {}
    dataFixedWidth['NbLignesEntete'] = 0
    dataFixedWidth['NbLignesPiedPage'] = 0
    dataFixedWidth['FixedWidth'] = FixedWidth
    
    dataParametreType = []
    dataParametreType.append(dataFixedWidth)

    Colonnes = gestionDuplicatNomColonnes(Colonnes)

    JsonFinalColonnes = json.dumps(Colonnes)
    JsonFinalParametreType = json.dumps(dataParametreType)
    return JsonFinalColonnes, JsonFinalParametreType

def ajoutScriptLigneDefinition(NomTable, IDDomaine, JsonFinalColonnes, JsonFinalParametreType, workbook, worksheet, row, i):
    try:
        my_list = []

        NomTable = os.path.splitext(NomTable)[0]
        # requete = ''
        DefinitionObjetInformation = NomTable
        Description = 'Valeurs généré via script Python pour le fichier COMIC ' + NomTable
        IDTypeObjetInformation = 'CSV'
        Masque = '{0}[^.]*[.]{{1}}(txt|TXT)'.format(os.path.splitext(NomTable)[0])
        ParametreType = "'{0}'".format(JsonFinalParametreType)
        ValeurEntete = JsonFinalColonnes

        my_list.append(DefinitionObjetInformation)
        my_list.append(Description)
        my_list.append(IDTypeObjetInformation)
        my_list.append(IDDomaine)
        my_list.append(Masque)
        my_list.append(ParametreType)
        my_list.append(ValeurEntete)
        my_list.append(0)
        my_list.append(0)
        my_list.append(65001)
        my_list.append(0)
        my_list.append(0)

        for col_num, data in enumerate(my_list):
            worksheet.write(row, col_num, data)

        row += 1
        i += 1
        
        return row, i  
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def find_csv_filenames( path_to_dir, suffix=".xlsx" ):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) or filename.endswith( ".xls" ) ]

def main():
    root = tk.Tk()
    root.withdraw()
    while True:
        NomDomaine = input("Entrer le Domaine:"+ '\n')
        try: 
            if len(NomDomaine) > 0:
                    break
            else:
                print("Le nom n'est pas valide")    
        except ValueError:
            print('Entrer un texte')

    DossierExcel = filedialog.askdirectory()

    ListeFichiers = find_csv_filenames(DossierExcel)

    workbook = xlsxwriter.Workbook('BI2020_Modele_{0}_DefinitionObjetInformation_Config.xlsx'.format(NomDomaine))
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

    for Fichier in ListeFichiers:
        L_excel = xlrd.open_workbook(DossierExcel + '/' + Fichier)
        first_sheet = L_excel.sheet_by_index(0)

        NomTable = Fichier
        Description = recuperationDescription(first_sheet)
        JsonFinalColonnes, JsonFinalParametreType = recuperationParametres(first_sheet)

        row, i = ajoutScriptLigneDefinition(NomTable, NomDomaine, JsonFinalColonnes, JsonFinalParametreType, workbook, worksheet, row, i)

    workbook.close()

if __name__ == "__main__":
    main()
    