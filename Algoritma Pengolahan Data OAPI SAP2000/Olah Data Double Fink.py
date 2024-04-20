# Olah data

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import lognorm
import os
import math

Kecepatan = []
Sudut = []
Panjang = []
CFSR = []
JD = []

csvPath = 'C:/Users/febyf/Documents/SAP 25 TES/DOUBLE FINK/CSV/'

#data keamanan tiap kec,sudut
for segmen in range (1100,1400,100):
    for kec_angin_awal in range(1,51,1) :
        for Sudut_atap in range(15,46,5):
            PanjangBatang = segmen * 5
            Path = csvPath+'Cold Formed Summary Result'+' angin '+str(kec_angin_awal)+' '+str(Sudut_atap)+' '+str(PanjangBatang)
            CFSR_test = 0
            JD_test = 0
            
            if(os.path.exists(Path)):

                data_CFSR = pd.read_csv(Path)
                
                for i in range(len(data_CFSR.index)):
                    if data_CFSR.Message[i] == 'Overstress':
                        CFSR_test = 'Fail'
                        break
                
            Kecepatan.append(kec_angin_awal)
            Sudut.append(Sudut_atap)
            Panjang.append(PanjangBatang)
            
            if CFSR_test == 'Fail':
                CFSR.append('Tidak Aman')
            else:
                CFSR.append('Aman')          

data ={
    'Panjang' :Panjang,
    'Kecepatan Angin': Kecepatan,
    'Sudut' : Sudut,
    'Aman Terhadap Penampang' : CFSR
} 


data = pd.DataFrame(data)
data.to_csv((csvPath+'/Rekap Data/'+'DATA KEAMANAN TUNGGAL'), index=False)

#######################################################################3

#data keamanan kec, berdasarkan sudut
Kecepatan = []
Penampang = []

for i in range (1,51,1):
    count = 0
    Kecepatan.append(i)
    selected_data = data.loc[data['Kecepatan Angin'] == i, 'Aman Terhadap Penampang'].values
    for j in range(len(selected_data)):
        if selected_data[j] == 'Tidak Aman':
            count += 1
    Penampang.append((count/len(selected_data))*100)

data ={
    'Kecepatan Angin': Kecepatan,
    'Aman Terhadap Penampang' : Penampang
} 
data = pd.DataFrame(data)

data.to_csv((csvPath+'/Rekap Data/'+'DATA KEAMANAN KELOMPOK'), index=False)

############################################################################
