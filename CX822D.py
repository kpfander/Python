import tkinter as tk
from tkinter import filedialog
import re
from os import listdir

NomFichier = 'CX822D'

def find_filenames( path_to_dir, suffix=".xlsx" ):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.startswith( NomFichier )]

def main():
    root = tk.Tk()
    root.withdraw()
    DossierExcel = filedialog.askdirectory()
    ListeFichiers = find_filenames(DossierExcel)
    
    regex = r".{51}(.)"
    
    for Fichier in ListeFichiers:
        LienFichier = DossierExcel + '/' + Fichier
        LienFichierType1 = DossierExcel + '/' + Fichier.replace(NomFichier, NomFichier + '_TYPE1')
        LienFichierType2 = DossierExcel + '/' + Fichier.replace(NomFichier, NomFichier + '_TYPE2')

        print(LienFichier)
        with open(LienFichier) as f:
            for line in f:
                Type = int(re.findall(regex, line)[0])
                print('Type: ' + str(Type))
                if Type == 1:
                    fType = open(LienFichierType1,"a")
                else:
                    fType = open(LienFichierType2,"a")
                fType.write(line)
                fType.close()
        f.close()


if __name__ == "__main__":
    main()
    