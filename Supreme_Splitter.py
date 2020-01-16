"""
This script will input .txt files from the XPS system and output both a large table of all the raw data,
and individual sheets named after each sample

DO NOT INCLUDE NONSTANDARD CHARACTERS, PANDAS HATE DEGREES SIGNS
"""

import pandas as pd
from pandas import ExcelWriter
import glob
import os
from openpyxl import Workbook
import numpy as np
import xlsxwriter
import time

abspath = os.path.abspath("Supreme_Splitter.exe")
path = os.path.dirname(abspath)
filenames1 = glob.glob(path + "/*.txt")

filenames2 = [os.path.split(y)[1] for y in filenames1]
filenames = [os.path.splitext(os.path.basename(z))[0] for z in filenames2]

#parses all .txt files and inputs their names

for file in filenames:
    print("Processing "+file)
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
    writer = pd.ExcelWriter(file+" RAW.xls",engine='xlsxwriter')  
    col = 0

    #we create notes_r, which should contain information about the scan region

    
    #this function searches through the 'chaotic notes' section and returns the most likely scan name
    levels = ["1s", "2s", "2p", "3s", "3p", "3d", "4s", "1s,", "2s,", "2p,", "3s,", "3p,",
              "3d,", "4s,", "SWEEP", "Survey", "Sweep,", "Sweep", "SURVEY", "Survey,", "-SWEEP", "-Sweep", "-Survey", "-SURVEY"]
       
    def orbcheck(list_notes):
        orbs = [level for level in levels if level in (list_notes)]
        elements_pos = [(list_notes.index(level)) for level in levels if level in (list_notes)]
        element = [int(n)-1 if n > 0 else 1 for n in elements_pos]

        #we have to check that there's at least one valid entry!
        if len(element)== 1:
            element_true = list_notes[element[0]]
        else:
            element_true = "UNKNOWN"

        if len(orbs) == 1:
            orbs_true = orbs[0]
        else:
            orbs_true = "UNKNOWN"
        
        label = element_true+" "+orbs_true
        return label

    #notes_r lists all the scans with a reduced name
    notes_r = [orbcheck(i.split()) for i in notes]
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
        file_frame.to_excel(writer, sheet_name = str(note_r), startcol = col, index=False)
        print("Added "+note_name+" to Workbook")
        
        col+=3

    #we first make a RAW xls file to refer back to - this is comparable to what was made with previous scripts
    writer.save()
    print(file+" RAW.xls completed with "+str(number_of_scans)+" scans.")

writer.save()

#now it's party time
for file in filenames:

    print("Tidying up "+file)
    #we open a new xls file
    writer = pd.ExcelWriter(file+" SORTED.xls",engine='xlsxwriter')

    #and we import the entirety of the RAW file
    sn = pd.ExcelFile(file+" RAW.xls")

    
    sns_r = sn.sheet_names
    #now, sheet by sheet...
    sns = [str(n) for n in sns_r]
    for sheet in sns:
        if sheet!="RAW":
            if sheet!="Raw":
                print("Reducing "+sheet)
                #we read the data from the file, ignoring all blank lines
                data = pd.read_excel(file+" RAW.xls", sheet_name=sheet, Index=None, skip_blank_lines=True, false_values="NaN")
                valid_data = data.dropna(axis ='columns', how="all")
                
                valid_scan_list = []
                columns_list = []

                inline_data = pd.DataFrame()

                #we then remove the redundant 'binding energy' columns
                for ind, heading in enumerate(list(valid_data.columns)):
                    if ind == 0 or ind%2 == 1:                        
                            valid_scan_list.append(list(data[heading]))
                            columns_list.append(heading)
                            inline_data[heading] = (list(data[heading]))
                
                inline_data.to_excel(writer, sheet_name = str(sheet), index=False)

    
    #and finally, save the new file!
    writer.save()
    print("Saved "+file+" after a tidy-up")

print("Thanks for choosing laziness, hope your day is PRETTY WIZARD")
time.sleep(4)

  
            
    

