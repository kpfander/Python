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
#import xlsxwriter 
import xlrd
import xlwt
import glob


## Composer la liste des domaines definies dans le fichier BI2020_Modele_systemeLog_Domaine
def GetNomDomainesFromConfig():
    try:
        worksheet_names = []
        MyListeDomaines =[]
        Liste_cols = []
        Liste_cols.append('IDDomaine')
        Liste_cols.append('NomDomaine')
        workbook = xlrd.open_workbook(r''+ PATH_ENTREE_FICHIER + r'FrameworkBI\BI2020_Modele_systemeLog_Domaine.xlsx', 1)       
        
        validate_worksheetName('DomaineTable', workbook, 'BI2020_Modele_systemeLog_Domaine')    
        worksheet = workbook.sheet_by_index(0)
        
        row = 1
        i = 0
        
        MyListeDomaines.append(Liste_cols)

        for row in range(1 , worksheet.nrows):                 
            MyListeDomaines.append(worksheet.row_values(row))
        
        workbook.release_resources()
        return MyListeDomaines
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def validate_worksheetName(nom, workbook, nomFichierConfig):
    liste_noms = workbook.sheet_names()
    for index in range(0, len(liste_noms)):
            if liste_noms[index] == nom:
                break;
            else:
                print("ERROR: Le nom de worksheet  " + liste_noms[index] + "  n'est pas valide dans le fichier de config " + nomFichierConfig +"")
                return 0
            
def validate_nomDomaine(nom):
    for index in range(0, len(ListeDomaines)):        
        if(nom == ListeDomaines[index][1]):
            return 1
        
    return 0    

def compare_twoLists(list1, list2, err_index, worksheet):
   
    if(len(list1)!= len(list2)):
        print(f"Les nombres des definitions dans les deux configs sont different. {len(list1)} <> {len(list2)}")
        worksheet.write(err_index[0], 0, f"Les nombres des definitions dans les deux configs sont different. {len(list1)} <> {len(list2)}")
        err_index[0] = err_index[0] + 1
    res = [x for x in list1 if x in list1 and x in list2]

    print(f" {len(res)} ")
    #print (res)
        
        
            
def CheckDefinitionObjetInformation():
    try:
        nomWorksheet = 'DefinitionObjetInformation'
        MyListe_files = glob.glob( r''+ PATH_ENTREE_FICHIER + Repertoir_A_VERIFIER +"\*" + nomWorksheet + "*.xlsx")
        #MyListeWrongConfigs = []
        ListeTables = []
        ListeDomainesDef = []
        err_workbook = xlwt.Workbook()    
        ws = err_workbook.add_sheet("Erreurs")
        err_index = 0        
        print('Verification Etap 1 -  Le nom de worksheet est bon dans les fichiers \n')
        print('Verification Etap 2 - Domaine corresponde au nom definie dans BI2020_Modele_systemeLog_Domaine \n')
        for index in range(0, len(MyListe_files)):
            workbook = xlrd.open_workbook(MyListe_files[index], 1)  
            if(validate_worksheetName(nomWorksheet, workbook, MyListe_files[index]) ==0):
                #MyListeWrongConfigs.append([MyListe_files[index], "Nom de worksheet n'est pas valid "])
                ws.write(err_index, 0, MyListe_files[index])
                ws.write(err_index, 1, "Nom de worksheet n'est pas valid \n")
                err_index = err_index + 1                 
            
            else :   
                worksheet = workbook.sheet_by_name(nomWorksheet)
                row = 1
                for row_index in range(row, worksheet.nrows):                                      
                    domain_name = worksheet.cell_value(row_index, 3)
                    if validate_nomDomaine(domain_name)==0:
                        #print(f"ERROR: Le nom de Domaine  {domain_name} n'est pas valide dans le fichier de config {MyListe_files[index]} row {row_index}")
                        #MyListeWrongConfigs.append([MyListe_files[index], f"Nom de Domaine {domain_name} n'est pas valid a la ligne {row_index} "])
                        ws.write(err_index, 0, MyListe_files[index])
                        ws.write(err_index, 1, f"Nom de Domaine {domain_name} n'est pas valid a la ligne {row_index} \n")
                        err_index = err_index + 1  
                    else:
                         ListeTables.append((worksheet.cell_value(row_index, 0), worksheet.cell_value(row_index, 6)))  
            workbook.release_resources()
        if(err_index>0):
            err_workbook.save(r''+ PATH_ENTREE_FICHIER + Repertoir_A_VERIFIER +'\ErreursTrouvees_DefinitionObjetInformation.xls')
            print("Il y avait des erreurs dans les fichiers DefinitionObjetInformation. Les erreurs ont ete sauvegardées dans le fichier: ErrorsTrouvees_DefinitionObjetInformation.xls")
        else:
            print("Il n'y avait pas d'erreurs dans les fichiers DefinitionObjetInformation ")
        
        #ws.write(MyListeWrongConfigs)
       
        #print(MyListeWrongConfigs)        
        return ListeTables
        
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def CheckDModelObjetInformation():
    try:
        nomWorksheet = 'ModeleObjetInformation'
        MyListe_files = glob.glob( r''+ PATH_ENTREE_FICHIER + Repertoir_A_VERIFIER +"\*" + nomWorksheet + "*.xlsx")
        ListeTables = []
        err_workbook = xlwt.Workbook()    
        ws = err_workbook.add_sheet("Erreurs")
        err_index = 0
        print('Verification Etap 3 - ModeleObjetInformation contiens tous les tables definies dans DefinitionObjetInformation et colonne EnteteExterne contiens la meme information \n')
        print(f"On a trouvé {len(MyListe_files)} fichiers de {nomWorksheet}")
        for index in range(0, len(MyListe_files)):
            workbook = xlrd.open_workbook(MyListe_files[index], 1)  
            
            if(validate_worksheetName(nomWorksheet, workbook, MyListe_files[index]) ==0):
                #MyListeWrongConfigs.append([MyListe_files[index], "Nom de worksheet n'est pas valid "])
                ws.write(err_index, 0, MyListe_files[index])
                ws.write(err_index, 1, "Nom de worksheet n'est pas valid \n")
                err_index = err_index + 1                 
            
            else :   
                worksheet = workbook.sheet_by_name(nomWorksheet)
                row = 1
                #ListeTables.append(worksheet)
                for row_index in range(row, worksheet.nrows):                                      
                    ListeTables.append((worksheet.cell_value(row_index, 0), worksheet.cell_value(row_index, 1))  )
                    
            workbook.release_resources()
            
        listeErrIndex = [err_index]
       
        compare_twoLists(ListeTables, ListeTablesDefinies, listeErrIndex, ws)
        err_index = listeErrIndex[0]
            
        print(err_index)
        print(listeErrIndex)
        if(err_index>0):
            err_workbook.save(r''+ PATH_ENTREE_FICHIER + Repertoir_A_VERIFIER +'\ErreursTrouvees_ModelObjetInformation.xls')
            print("Il y avait des erreurs dans les fichiers DefinitionObjetInformation. Les erreurs ont ete sauvegardées dans le fichier: ErrorsTrouvees_DefinitionObjetInformation.xls")
        else:
            print(f"Il n'y avait pas d'erreurs dans les fichiers {nomWorksheet} ")
        
        #print(ListeTables)
        #ws.write(MyListeWrongConfigs)
       
        #print(MyListeWrongConfigs)        
        return ListeTables
        
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)        
        
def CheckConfigsInFolder():
    try:
        MyListe_files = []
        MyListeWrongConfigs = []
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)

def GetFileNames(FileType) :
    MyFiles= []
    
    return MyFiles
def main():
    
    LIEN_SORTIE_FICHIER = ".\\"
    global PATH_ENTREE_FICHIER 
    global Repertoir_A_VERIFIER
    global ListeDomaines 
    global ListeTablesDefinies
    global ListeTablesODS
    
    try:        
        while True:
            PATH_ENTREE_FICHIER = input("Entrer le pathe de repertoir fichiers de ''InputFichiersConfiguration'' :"+ '\n')
            try: 
                if len(PATH_ENTREE_FICHIER) > 0:
                    ##print (PATH_ENTREE_FICHIER)
                    break
                else:
                    print("Le nom n'est pas valide")    
            except ValueError:
                print('Entrer une chaine de caractere')
        if(PATH_ENTREE_FICHIER[-1] != "\\"):
            PATH_ENTREE_FICHIER = PATH_ENTREE_FICHIER + "\\"
        while True:
            Repertoir_A_VERIFIER = input("Entrer le nom de repertoir contenant les fichiers de configuaration a verifier :"+ '\n')
            try: 
                if len(Repertoir_A_VERIFIER) > 0:
                    ##print (Repertoir_A_VERIFIER)
                    break
                else:
                    print("Le nom n'est pas valide")    
            except ValueError:
                print('Entrer une chaine de caractere')
        
        ListeDomaines = GetNomDomainesFromConfig()  
        print (f"{str(len(ListeDomaines)-1)} Domaines definies dans le fichier  r'{PATH_ENTREE_FICHIER} r'FrameworkBI\BI2020_Modele_systemeLog_Domaine.xlsx")
        
        ListeTablesDefinies = CheckDefinitionObjetInformation()
        ListeTablesODS = CheckDModelObjetInformation()
        
        
        print('Verification Etap 3 - DeclarationModele contient NN tables, qui sont definies dans DefinitionObjetInformation ')
        print('Verification Etap 4 - Cles naturelles sont definies sur tous les tables qui sont dans DeclarationModele')    
        
    except pyodbc.Error as err:
        print(err)
        sys.exit(1)
        
if __name__ == "__main__":
    main()