"""
This script will input .txt files from the XPS system and output both a large table of all the raw data,
and individual sheets named after each sample
"""

import pandas as pd
from pandas import ExcelWriter
import glob
import os
from openpyxl import Workbook
import numpy as np
import xlsxwriter
import time

abspath = os.path.abspath("ScanSplitter 1.3.py")
path = os.path.dirname(abspath)
filenames1 = glob.glob(path + "/*.txt")

filenames2 = [os.path.split(y)[1] for y in filenames1]
filenames = [os.path.splitext(os.path.basename(z))[0] for z in filenames2]

#parses all .txt files and inputs their names

for file in filenames:
    raw_file = pd.read_csv(file+".txt", sep ="\t") 

    #imports the raw notes as a list
    notes_raw =(list(raw_file["Notes"]))
    notes = []

    #imports the energy and counts data
    energy = (list(raw_file["Region"]))
    counts = (list(raw_file["Enabled"]))

    positions_start = []
    positions_end = []

    #notes the position of notes and collects them as scan names
    for note in notes_raw:
        if type(note) == str:
            if note != "Notes":
                notes.append(note)
                
    #takes note of the indices of each "Layer" entry, which marks scan start point
    for ind, n in enumerate(energy):
        if n == "Layer":
            positions_start.append(ind)

    #now we convert the start points into an end points              
    #adds the final datapoint as the end point, and deletes zero (start marker)
    positions_end = list(positions_start)
    positions_end.append(len(notes_raw))
    del positions_end[0]

    #we open an Excel file    
    writer = pd.ExcelWriter(file+".xls",engine='xlsxwriter')  
    col = 0

    #we create notes_r, which should contain information about the scan region
    notes_unzip = [i.split() for i in notes]    
    notes_r = [u[7] for u in notes_unzip]
    number_of_scans = (len(notes_r))

    #we create a list of scans 
    energy_scanlist = []
    counts_scanlist = []

    #this subdivides both counts and scans, but discriminates between null layers     
    for start, end in zip(positions_start, positions_end):
        
        if energy[start+1] == "1":
                    
            energy_scanlist.append(energy[(start+3):(end-2)])
            counts_scanlist.append(counts[(start+3):(end-2)])

    for note_r, note_name, k, l in zip(notes_r, notes, energy_scanlist, counts_scanlist):

        #converts the strings into floats for ease of processing later on
        energy_s = [float(n) for n in k]
        counts_s = [float(n) for n in l]    

        #we create individual dataframes for each scan and export it into excel in the "Raw" sheet
    
        data = {str(note_name): energy_s, " ": counts_s}       
        file_frame = pd.DataFrame(data, columns = [str(note_name), " "])
        file_frame.to_excel(writer, sheet_name="Raw", startcol = col, index=False)

        #here we also print each dataframe to a new sheet
        file_frame.to_excel(writer, sheet_name = str(note_r), startcol = col, index=False)
        print("Added "+note_name+" to Workbook")
        col+=3
                
    writer.save()
    print(file+".xls completed with "+str(number_of_scans)+" scans.")

print(str(len(filenames))+" files successfully output.")
time.sleep(3)
  
            
    

