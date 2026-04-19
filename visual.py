import pandas as pd
import os
import glob
import streamlit as st

outl = r"C:\Users\dluminsa\Desktop\QPR\EXTRA\DSD"
outv = r"C:\Users\dluminsa\Desktop\QPR\EXTRA\AUDIT"
outp = r"C:\Users\dluminsa\Desktop\QPR\EXTRA\PALD"

dsds = glob.glob(os.path.join(outl, '*.csv'))


dsdr = []
for file in dsds:
    dfa = pd.read_csv(file)
    dsdr.append(dfa)

dsd = pd.concat(dsdr)

dsd2 = dsd.groupby('STATUS').sum().reset_index()

st.write(dsd2)

dsdv = glob.glob(os.path.join(outv, '*.csv'))


dsdrv = []
for fil in dsdv:
    dfb = pd.read_csv(fil)
    dsdrv.append(dfb)

dsdv = pd.concat(dsdrv)

dsdv = dsdv.groupby('AGEBND').sum().reset_index()

st.write(dsdv)

dsdp = glob.glob(os.path.join(outp, '*.csv'))


dfp = []
for fin in dsdp:
    dfc = pd.read_csv(fin)
    dfp.append(dfc)

dsdp = pd.concat(dfp)

dsdp = dsdp.groupby('WTBD').sum().reset_index()
mapper = {
    '6-9.9': 1, '10-13.9':2, '14-19.9':3, '20-24.9':4
}
dsdp['R'] = dsdp['WTBD'].map(mapper)
dsdp = dsdp.sort_values(by = 'R')
dsdp = dsdp.drop(columns ='R')
st.write(dsdp)