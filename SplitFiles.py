import tkinter as tk
from tkinter import filedialog
import os
from fsplit.filesplit import FileSplit



def splitFile2(file, output_dir):
    lines_per_file = 300
    smallfile = None
    with open('really_big_file.txt') as bigfile:
        for lineno, line in enumerate(bigfile):
            if lineno % lines_per_file == 0:
                if smallfile:
                    smallfile.close()
                small_filename = 'small_file_{}.txt'.format(lineno + lines_per_file)
                smallfile = open(small_filename, "w")
            smallfile.write(line)
        if smallfile:
            smallfile.close()

# 100000000 = 100Mo
def splitFile(file, output_dir):
    print(file)
    print(output_dir)
    fs = FileSplit(file=file, splitsize=100000000, output_dir=output_dir)
    fs.split()
    #fs.split(include_header=True)


def main():
    root = tk.Tk()
    root.withdraw()

    dossierInput = filedialog.askdirectory()

    for subdir, dirs, files in os.walk(dossierInput):
        for file in files:
            if file.endswith(".TXT"):
                output_dir = os.path.join(subdir, 'chunked')
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                print(os.path.join(subdir, file))
                splitFile(os.path.join(subdir, file), output_dir)

if __name__ == "__main__":
    main()
    