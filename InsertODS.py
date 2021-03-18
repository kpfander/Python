# -*- coding: utf-8 -*-
"""
Created on Tue Feb 16 13:52:31 2021

@author: Kristina Pfander
"""
#from azure.storage.blob import BlobServiceClient

# Define parameters
storageAccountURL = "https://ccqbs001.canadacentral.batch.azure.com"
storageKey         = "tL7z5bo+vimM5BCt1VyAq+l0CBTQQ6Nuc/Ee19H1YQwDsqXRC1kFuseqFxk3YyOIF9uZtvQOFlhm28hYeQlxPw=="
containerName      = "output"

import os
import uuid
import pyodbc
import io
import pandas as pd
from string import ascii_uppercase
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import sys
import re
import xlsxwriter 


BDD = "Donnees_POC_Y450_DST"
SERVEUR =  "testathwma.database.windows.net"
username = 'admin_testathwma'
password = 'P@ssw0rd'   
driver= '{ODBC Driver 17 for SQL Server}'

IDDomain = "2"
LIEN_SORTIE_FICHIER = ".\\"
conn = pyodbc.connect('DRIVER='+driver+';SERVER='+SERVEUR+';PORT=1433;DATABASE='+BDD+';UID='+username+';PWD='+ password) 
cursor = conn.cursor()



def GetTablesToLoad():
    
    # connecting to the database  
   
    
    # cursor  
    crsr = conn.cursor() 
    query = """SELECT
			a.[IDModeleObjetInformation],
			b.[TypeHachage],
			c.ObjetStockage,
			f.IDDomaine,
			f.NomDomaine,
			e.DefinitionObjetInformation,
			cn.CleNaturelle,
			e.Masque,
			TOI.TypeObjetInformation
		FROM [ods].[DeclarationModele_Config] a
		JOIN [cryptage].[TypeHachage_Config] b ON b.IDTypeHachage = a.IDTypeHachage
		JOIN [dataLake].[ObjetStockage] c ON c.IDObjetStockage = a.IDObjetStockage
		JOIN [dataLake].[ModeleObjetInformation] d ON d.IDModeleObjetInformation = a.IDModeleObjetInformation
		JOIN [dataLake].[DefinitionObjetInformation_Config] e ON e.IDDefinitionObjetInformation = d.IDDefinitionObjetInformation
		JOIN [systemeLog].[Domaine] f ON f.IDDomaine = e.IDDomaine
		LEFT JOIN [dataLake].[CleNaturelle_Config] CN ON d.IDModeleObjetInformation = cn.IDModeleObjetInformation
		LEFT JOIN datalake.TypeObjetInformation TOI ON e.IDTypeObjetInformation = TOI.IDTypeObjetInformation
		WHERE cn.CleNaturelle IS NOT NULL AND  f.IDDomaine = """ + IDDomain + """		
		ORDER BY c.IDObjetStockage 
		"""
    #print(query)
        
    crsr.execute(query)
        
    # store all the fetched data in the ans variable 
    ans = crsr.fetchall()  
    data = pd.read_sql(query, conn)

    data.describe()    
    # Since we have already selected all the data entries  
    # using the "SELECT *" SQL command and stored them in  
    # the ans variable, all we need to do now is to print  
    # out the ans variable 
   #print(ans) 
   #print(data['TypeObjetInformation'])
   
    return data

def main():
    root = tk.Tk()
    root.withdraw()
    
    GetTablesToLoad()

if __name__ == "__main__":
    main()
    