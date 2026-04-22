import os
import io
import time
import traceback
import datetime as dt
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st
import gspread
from openpyxl import load_workbook
from google.oauth2.service_account import Credentials
from streamlit_gsheets import GSheetsConnection

a = 2
# def extract():
#VARIABLES
lyear = 2026
lmonth = 3
if a == 2:
    cola,colb,colc = st.columns([1,3,1])
    colb.subheader('PIVOT TABLES FOR ECHO DATA')
    facility = st.text_input('Enter Facility Name')
    if not facility:
        st.warning('**Please enter the facility name to proceed**')
        st.stop()
    else:
         pass   
    

    file = st.file_uploader("Upload your EMR extract here", type=['csv']) 
    if file is not None:        
        st.session_state.df = None
        time.sleep(1)
    if 'df' not in st.session_state:
        st.session_state.df = None

    if 'reader' not in st.session_state:
        st.session_state.reader = None
   

    if file is not None:
                    st.session_state.df = pd.read_csv(file)
                    df = st.session_state.df
    
                    df = df.rename(columns= {'ART  ':'ART',  'RD  ':'RD',  'VD  ':'VD', 'LD  ': 'LD','ARVS  ':'ARVS', 'ARVD ':'ARVD',
         'AG  ':'AG','WT  ':'WT'})#, 'TPT ': 'TPT'})
                    df = df.rename(columns= {'ART ':'ART', 'RD ':'RD', 'VD ':'VD', 'LD ': 'LD','WT ': 'WT','ARVS ':'ARVS',
                            'AG ':'AG', 'ARVD ':'ARVD'})#, 'TPT  ': 'TPT'})
                    columns = ['ART','AG','AS', 'VD', 'RD','LD','WT', 'ARVS', 'ARVD']
                    cols = df.columns.to_list()
                    needed = set(columns)
                    there = set(cols)
                    missing = needed - there
                    missing = list(missing)
                    if not all(column in cols for column in columns):
                        missing_columns = [column for column in columns if column not in cols]
                        for column in missing_columns:
                            st.markdown(f' **ERROR !!! MISSING COLUMN(S): {missing}**')
                            st.markdown('**First rename all the columns as guided above**')

                            st.stop()
                    st.session_state.reader= True
    if st.session_state.reader:
                        st.session_state.df = st.session_state.df.rename(columns= {'ART  ':'ART', 'AS  ':'AS', 'RD  ':'RD',  'VD  ':'VD', 'LD  ': 'LD','ARVS  ':'ARVS','ARVD  ':'ARVD',
                'AG  ':'AG'})
                        st.session_state.df = st.session_state.df.rename(columns= {'ART ':'ART', 'RD ':'RD', 'VD ':'VD',  'LD ': 'LD', 'AG ':'AG', 'ARVS ':'ARVS', 'ARVD ':'ARVD'})#, 'TPT  ': 'TPT'})
                        df = st.session_state.df.copy()
                    
                        df = df[['ART','AG','AS', 'VD', 'RD','LD','WT', 'ARVS', 'ARVD']].copy()
                        
                        df['ART'] = df['ART'].astype(str)
                        df['A'] = df['ART'].str.replace('[^0-9]', '', regex=True)
                        df['A'] = pd.to_numeric(df['A'], errors= 'coerce')
                        df = df[df['A']>0].copy()
                                          

                        testrt = df.copy()
                            
                        testrc = testrt[testrt['VD'].isnull()].copy()    

                        if testrc.shape[0]>1000:
                            st.warning('VD is empty, use the correct HIV Viral Load Date, it can not be blank')
                            testrc =testrc[['ART', 'AG','RD', 'VD']].copy()
                            st.write(testrc.head(5))
                            st.stop()



                        testrd = df[testrt['LD'].isnull()].copy()    
                        if testrd.shape[0]>1000:
                            st.warning('LD is empty, use the correct Last Encouter Date, it can not be blank')
                            testrd =testrd[['ART', 'RD', 'LD']].copy()
                            st.write(testrd.head(5))
                            st.stop()
                        testre = df[~testrt['AS'].isnull()].copy()  

                        testrde = df[~testrt['RD'].isnull()].copy()    
                        if testrde.shape[0]<10:
                            st.warning('RD is empty, use the correct Return Visit Date, it can not be blank')
                            testrde =df[['ART', 'RD', 'LD']].copy()
                            st.write(testrde.head(5))
                            st.stop()

                      
                   
                    
                        df['RD'] = df['RD'].astype(str)
                   
                        df['VD'] = df['VD'].astype(str)
                        df['LD'] = df['LD'].astype(str)
                                    
                        y = pd.DataFrame({'ART' :['2','3','4','5'], 'RD':['1-1-1',1,'1/1/1','3 8 2001'], 
                                        'VD':['1-1-1',1,'1/1/1','3 8 2001'], 'LD':['1-1-1',1,'1/1/1','3 8 2001'],'ARVS':['1-1-1',1,'1/1/1','3 8 2001']})                        
                        
                        df['RD'] = df['RD'].astype(str)
                
                        df['VD'] = df['VD'].astype(str)
                        
                        df['LD'] = df['LD'].astype(str)
                    
                        df['RD'] = df['RD'].str.replace('00:00:00', '', regex=True)
                    
                        df['VD'] = df['VD'].str.replace('00:00:00', '', regex=True)
                        
                        df['LD'] = df['LD'].str.replace('00:00:00', '', regex=True)
                     
                        df["RD"] = df["RD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                       
                        df["VD"] = df["VD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                       
                        df["RD"] = df["RD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                    
                        df["VD"] = df["VD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        
                        df = pd.concat([df,y])
                        df = df.copy()
                        
                        df['RD'] = df['RD'].astype(str) ###
                        
                        df['VD'] = df['VD'].astype(str) ###
                        
                        df['LD'] = df['LD'].astype(str)
                        df['ARVS'] = df['ARVS'].astype(str)
                
         

                        # SORTING THE RETURN VISIT DATE
                        A = df[df['RD'].str.contains('-')].copy()
                        a = df[~df['RD'].str.contains('-')].copy()
                        B = a[a['RD'].str.contains('/')].copy()
                        C = a[~a['RD'].str.contains('/')].copy()
                        E = C[C['RD'].str.contains(' ')].copy()
                        D = C[~C['RD'].str.contains(' ')].copy()                     
                        #D = C[C['RD'].apply(lambda x: isinstance(x, (int, float)) or x.isdigit())].copy()
                        #E = C[~C['RD'].apply(lambda x: isinstance(x, (int, float)) or x.isdigit())].copy()              
                        A[['Ryear', 'Rmonth', 'Rday']] = A['RD'].str.split('-', expand = True)
                        B[['Ryear', 'Rmonth', 'Rday']] = B['RD'].str.split('/', expand = True)
                        try:
                            D['RD'] = pd.to_numeric(D['RD'], errors='coerce')
                            D['RD'] = pd.to_datetime(D['RD'], origin='1899-12-30', unit='D', errors='coerce')
                            D['RD'] =  D['RD'].astype(str)
                            D[['Ryear', 'Rmonth', 'Rday']] = D['RD'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['RD'] = pd.to_datetime(E['RD'],format='%d %m %Y', errors='coerce')
                            E['RD'] =  E['RD'].astype(str)
                            E[['Ryear', 'Rmonth', 'Rday']] = E['RD'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E])   

                        #SORTING THE VD DATE
                        A = df[df['VD'].str.contains('-')].copy()
                        a = df[~df['VD'].str.contains('-')].copy()
                        B = a[a['VD'].str.contains('/')].copy()
                        C = a[~a['VD'].str.contains('/')].copy()
                        E = C[C['VD'].str.contains(' ')].copy()
                        D = C[~C['VD'].str.contains(' ')].copy()      
                        A[['Vyear', 'Vmonth', 'Vday']] = A['VD'].str.split('-', expand = True)
                        B[['Vyear', 'Vmonth', 'Vday']] = B['VD'].str.split('/', expand = True)
                        try:
                            D['VD'] = pd.to_numeric(D['VD'], errors='coerce')
                            D['VD'] = pd.to_datetime(D['VD'], origin='1899-12-30', unit='D', errors='coerce')
                            D['VD'] =  D['VD'].astype(str)
                            D[['Vyear', 'Vmonth', 'Vday']] = D['VD'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['VD'] = pd.to_datetime(E['VD'],format='%d %m %Y', errors='coerce')
                            E['VD'] =  E['VD'].astype(str)
                            E[['Vyear', 'Vmonth', 'Vday']] = E['VD'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E])  
                        df = df.copy()
                        
          



                        # SORTING THE LAST ENCOUNTER DATES
                        A = df[df['LD'].str.contains('-')].copy()
                        a = df[~df['LD'].str.contains('-')].copy()
                        B = a[a['LD'].str.contains('/')].copy()
                        C = a[~a['LD'].str.contains('/')].copy()
                        E = C[C['LD'].str.contains(' ')].copy()
                        D = C[~C['LD'].str.contains(' ')].copy()
                        A[['Lyear', 'Lmonth', 'Lday']] = A['LD'].str.split('-', expand = True)
                        B[['Lyear', 'Lmonth', 'Lday']] = B['LD'].str.split('/', expand = True)
                        try:
                            D['LD'] = pd.to_numeric(D['LD'], errors='coerce')
                            D['LD'] = pd.to_datetime(D['LD'], origin='1899-12-30', unit='D', errors='coerce')
                            D['LD'] =  D['LD'].astype(str)
                            D[['Lyear', 'Lmonth', 'Lday']] = D['LD'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['LD'] = pd.to_datetime(E['LD'],format='%d %m %Y', errors='coerce')
                            E['LD'] =  E['LD'].astype(str)
                            E[['Lyear', 'Lmonth', 'Lday']] = E['LD'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E])

                        A = df[df['ARVS'].str.contains('-')].copy()
                        a = df[~df['ARVS'].str.contains('-')].copy()
                        B = a[a['ARVS'].str.contains('/')].copy()
                        C = a[~a['ARVS'].str.contains('/')].copy()
                        E = C[C['ARVS'].str.contains(' ')].copy()
                        D = C[~C['ARVS'].str.contains(' ')].copy()                                
                        A[['TDF', '3TC', 'DTG']] = A['ARVS'].str.split('-', expand = True)
                        B[['TDF', '3TC', 'DTG']] = B['ARVS'].str.split('/', expand = True)
                        try:
                            D['ARVS'] =  D['ARVS'].astype(str)
                            D[['TDF', '3TC', 'DTG']] = D['ARVS'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['ARVS'] =  E['ARVS'].astype(str)
                            E[['TDF', '3TC', 'DTG']] = E['ARVS'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E]) 
                        #st.write(df.head(10))
                       
  

                        df['RD'] = df['RD'].astype(str)
   
                        df['VD'] = df['VD'].astype(str)
                      
                        df['LD'] = df['LD'].astype(str)
                        
            #             #Clearing NaT from te dates
                     
                        df['RD'] = df['RD'].str.replace('NaT', '',regex=True)
                       
                        df['VD'] = df['VD'].str.replace('NaT', '',regex=True)
                       
                        df['LD'] = df['LD'].str.replace('NaT', '',regex=True)
                      
                        df[['Vyear', 'Vmonth', 'Vday']] =df[['Vyear', 'Vmonth', 'Vday']].apply(pd.to_numeric, errors = 'coerce') 
                        df['Vyear'] = df['Vyear'].fillna(994)
                        a = df[df['Vyear']>31].copy()
                        b = df[df['Vyear']<32].copy()
                        #c = df[]
                        b = b.rename(columns={'Vyear': 'Vday2', 'Vday': 'Vyear'})
                        b = b.rename(columns={'Vday2': 'Vday'})
                        df = pd.concat([a,b])
                        
                        # #SORTING THE RETURN VISIT DATE YEARS
                        df[['Rday', 'Ryear']] = df[['Rday', 'Ryear']].apply(pd.to_numeric, errors='coerce')
                        df['Ryear'] = df['Ryear'].fillna(994)
                        a = df[df['Ryear']>31].copy()
                        b = df[df['Ryear']<32].copy()
                        b = b.rename(columns={'Ryear': 'Rday2', 'Rday': 'Ryear'})
                        b = b.rename(columns={'Rday2': 'Rday'})
                        df = pd.concat([a,b])
                           
                 
                        # #SORTING THE LAST ENCOUNTER
                        df[['Lday', 'Lyear']] = df[['Lday', 'Lyear']].apply(pd.to_numeric, errors='coerce')
                        df['Lyear'] = df['Lyear'].fillna(994)
                        a = df[df['Lyear']>31].copy()
                        b = df[df['Lyear']<32].copy()
                        b = b.rename(columns={'Lyear': 'Lday2', 'Lday': 'Lyear'})
                        b = b.rename(columns={'Lday2': 'Lday'})
                        df = pd.concat([a,b])
                        df = df.copy()

                        def ager(a):
                            if a < 5:
                                return '< 5 Yrs'
                            elif a < 10:
                                return '5-9 Yrs'
                            elif a < 15:
                                return '10-14 Yrs'
                            else:
                                return '15-19 Yrs'

                        def weight(a):
                             if a < 6:
                                  return '3-5.9 KG'
                             if a < 10:
                                  return '6-9.9 KG'
                             elif a < 15:
                                  return '10-13.9 KG'
                             elif a < 20:
                                  return '14-19.9 KG'
                             elif a < 25:
                                  return '20-24.9 KG'
                             elif a < 30:
                                  return '25-29.9 KG'
                             else:
                                  return '>= 30 KG'
                        datyw =pd.DataFrame({'WEIGHT BANDS': ['3-5.9 KG', '6-9.9 KG', '10-13.9 KG', '14-19.9 KG', '20-24.9 KG', '25-29.9 KG', '>= 30 KG']})
                        datya = pd.DataFrame({'AGE BANDS': ['< 5 Yrs', '5-9 Yrs', '10-14 Yrs', '15-19 Yrs']})
                        wmapper = {'3-5.9 KG':1, '6-9.9 KG':2, '10-13.9 KG':3, '14-19.9 KG':4, '20-24.9 KG':5, '25-29.9 KG':6, '>= 30 KG':7}
                        wmapper2 = {'3-5.9 KG':1, '6-9.9 KG':2, '10-13.9 KG':3, '14-19.9 KG':4, '20-24.9 KG':5}
                        df['WT'] = pd.to_numeric(df['WT'], errors='coerce')
                        df['WEIGHT BANDS'] = df['WT'].apply(weight)

                        amapper ={'< 5 Yrs':1, '5-9 Yrs':2, '10-14 Yrs':3, '15-19 Yrs':4}
                        df['AG'] = pd.to_numeric(df['AG'], errors='coerce')
                        df['AGE BANDS'] = df['AG'].apply(ager)

                        df['AG'] = pd.to_numeric(df['AG'], errors='coerce')
                        df = df[df['AG']<20].copy()
                        testrw = df[df['WT'].isnull()].copy()  
                        at = testrw.shape[0] 

                        if at > 0:
                            st.info('**ERROR!!!!**')
                            st.warning(f'**{at} CALHIV listed below have no weight in the uploaded extract**')
                            st.write('**first input their weight to avoid errors**')
                            testrw =testrw[['ART','AG', 'RD', 'WT']].copy()
                            st.write(testrw)
                            st.stop()
                        else:
                            #  st.write('**All CALHIV have weight, proceeding with the analysis**')
                            pass
                        b11 = df.shape[0]

                        if b11<1:
                             st.info("**THIS EXTRACT DOESN'T AHVE ANY CALHIV BTN 0 to 19 YEAR, CHECK MANUALLY**")
                             st.stop()
                        else:
                            dfb11 = pd.DataFrame({ 'Total number of CALHIV currently recieving care at this HF?': [b11]})
                            #  st.write(dfb11)
                            df['WT'] = pd.to_numeric(df['WT'], errors='coerce')
                            b2 = df[df['WT']>2.9].copy()
                            dfb2 = df.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                            dfb2 = pd.merge(dfb2, datyw, on='WEIGHT BANDS', how='right')
                            dfb2['TOTAL'] = dfb2['TOTAL'].fillna(0)
                            dfb2['R'] = dfb2['WEIGHT BANDS'].map(wmapper)
                            dfb2 = dfb2.sort_values('R').drop(columns='R')
                            #  st.write(dfb2)
                             
                            dfb3 = df.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                            dfb3 = pd.merge(dfb3, datya, on='AGE BANDS', how='right')
                            dfb3['TOTAL'] = dfb3['TOTAL'].fillna(0)
                            dfb3['R'] = dfb3['AGE BANDS'].map(amapper)
                            dfb3 = dfb3.sort_values('R').drop(columns='R')
                            #  st.write(dfb3)


                            df[['TDF', 'DTG']] = df[['TDF', 'DTG']].astype(str)
                            dfc = df[((df['TDF']=='ABC') & (df['DTG']=='DTG'))].copy()
                            dfn = df[~((df['TDF']=='ABC') & (df['DTG']=='DTG'))].copy()
                            a1 = dfc.shape[0]
                            a2 = b11 - a1
                            if a2 > 0:
                                cola, colb = st.columns(2)
                                cola.markdown(f'**You have {a2} CALHIV are not on ABC-DTG regimen**')
                                with colb.expander('**Click here to see them**'):
                                    dfn = dfn[['ART', 'AG', 'WT', 'ARVS']].copy()
                                    dfn = dfn.reset_index(drop=True)
                                    st.write(dfn)
                            else:
                                 pass
                            #children on ABC-DTG]
                            #st.write(dfc[['ART','AG', 'WT', 'ARVS']])
                            dfc['WT'] = pd.to_numeric(dfc['WT'], errors='coerce')
                            dfc = dfc[((dfc['WT']<6) | (dfc['WT']>24.9))].copy()
                            a3 = dfc.shape[0]
                            if a3 > 0:
                                 #st.write(dfc[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                dfc1 = dfc.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfc1 = pd.merge(dfc1, datyw, on='WEIGHT BANDS', how='right')
                                dfc1['TOTAL'] = dfc1['TOTAL'].fillna(0)
                                dfc1['R'] = dfc1['WEIGHT BANDS'].map(wmapper)
                                dfc1 = dfc1.sort_values('R').drop(columns='R')
                                #st.write(dfc1)
                                dfc2 = dfc.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfc2 = pd.merge(dfc2, datya, on='AGE BANDS', how='right')
                                dfc2['TOTAL'] = dfc2['TOTAL'].fillna(0)
                                dfc2['R'] = dfc2['AGE BANDS'].map(amapper)
                                dfc2 = dfc2.sort_values('R').drop(columns='R')
                                # st.write(dfc2)  
                            else:
                                 dfc1 = datyw.copy()
                                 dfc1['TOTAL'] = 0
                                 #st.write(dfc1)  
                                 dfc2 = datya.copy()
                                 dfc2['TOTAL'] = 0
                                 #st.write(dfc2) 
                            
                            #pALD ELIGIBILITY BY AGE BANDS FROM THE WHOLE CURR C4
                            df['AG'] = pd.to_numeric(df['AG'], errors='coerce')
                            dfcpa =  df[(df['AG']<10)].copy()

                            if dfcpa.shape[0]>0:
                                #st.write(dfcp.shape[0])
                                #st.write(dfcp[['ART', 'AG',ARVS']].reset_index(drop=True)) 
                                dfc4 = dfcpa.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                c4map = {'< 5 Yrs':1, '5-9 Yrs':2}
                                dfc4['R'] = dfc4['AGE BANDS'].map(c4map)
                                dfc4 = dfc4.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(dfc4)

                                #OF THESE, WHO IS ON PALD? (WEIGHT MUST BE BTN 6 AND 24.9 KG)
                                dfcpa[['TDF', 'DTG']] = dfcpa[['TDF', 'DTG']].astype(str)
                                dfc5 = dfcpa[((dfcpa['TDF']=='ABC') & (dfcpa['DTG']=='DTG'))].copy()
                                dfc5['WT'] = pd.to_numeric(dfc5['WT'], errors='coerce')
                                dfc5 = dfc5[(dfc5['WT']>=6) & (dfc5['WT']<=24.9)].copy()
                                if dfc5.shape[0]>0:
                                    #st.write(dfc5[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                    dfc5 = dfc5.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfc5['R'] = dfc5['AGE BANDS'].map(c4map)
                                    dfc5 = dfc5.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    #st.write(dfc5)
                                else:
                                     dfc5 = datyw.copy()
                                     dfc5['TOTAL'] = 0
                            else:
                                 dfc4 = datya.copy()
                                 dfc4['TOTAL'] = 0
                                 dfc4 = dfc4.head(2)
                                 #st.write(dfc4)  
                                 dfc5 = datya.copy()
                                 dfc5['TOTAL'] = 0
                                 dfc5 = dfc5.head(2)
                                 #st.write(dfc5)

                            #pALD ELIGIBILITY BY WEIGHT BANDS FROM THE WHOLE CURR C6
                            df['WT'] = pd.to_numeric(df['WT'], errors='coerce')
                            dfcpc = df[(df['WT']>5.9) & (df['WT']<25)].copy()

                            if dfcpc.shape[0] >0:
                                dfc6 = dfcpc.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                c6map = {'3-5.9 KG':1, '6-9.9 KG':2, '10-13.9 KG':3, '14-19.9 KG':4, '20-24.9 KG':5}
                                dfc6['R'] = dfc6['WEIGHT BANDS'].map(c6map)
                                dfc6 = dfc6.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(dfc6)

                                #OF THESE, WHO IS ON PALD? 
                                dfcpc[['TDF', 'DTG']] = dfcpc[['TDF', 'DTG']].astype(str)
                                dfc7 = dfcpc[((dfcpc['TDF']=='ABC') & (dfcpc['DTG']=='DTG'))].copy()
                                
                                #NOT ON PALD BUT ELIGIBLE BY WEIGHT
                                dfc8 = dfcpc[~((dfcpc['TDF']=='ABC') & (dfcpc['DTG']=='DTG'))].copy()
                                if dfc8.shape[0]>0:
                                        st.info('**You have CALHIV who are eligible for PALD by weight but are not on ABC-DTG regimen**')
                                        st.write(dfc8[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                        confirm = st.radio('Is this correct?', ('Yes', 'No'), horizontal=True, index=None)
                                        if not confirm or confirm == 'No':
                                            st.warning('**Correct their regimen in the extract before uploading again**')
                                            st.stop()
                                        else:
                                             pass
                                #ON ABC-DTG
                                if dfc7.shape[0]>0:
                                    #st.write(dfc7[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                    dfc7 = dfc7.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfc7['R'] = dfc7['WEIGHT BANDS'].map(c6map)
                                    dfc7 = dfc7.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    #st.write(dfc7)
                                else:
                                     dfc7 = datyw.copy()
                                     dfc7['TOTAL'] = 0
                                     dfc7 = dfc7.head(5)
                                     #st.write(dfc7)
                            else:
                                 dfc6 = datyw.copy()
                                 dfc6['TOTAL'] = 0
                                 dfc6 = dfc6.head(5)
                                 #st.write(dfc6)  
                                 dfc7 = datyw.copy()
                                 dfc7['TOTAL'] = 0
                                 dfc7 = dfc7.head(5)
                                 #st.write(dfc7)
                            #TOTAL CONSUMPTION OF DRUGS
                            #CAME LAST MONTH
                            df[['Lmonth', 'Lyear']] = df[['Lmonth', 'Lyear']].apply(pd.to_numeric, errors='coerce')
                            dff = df[((df['Lyear']==lyear)  & (df['Lmonth']==lmonth))].copy()
                            #st.write(dff[['ART', 'LD', 'WT', 'ARVS', 'TDF', '3TC', 'DTG']].reset_index(drop=True))
                            
                            if dff.shape[0]>0: #if there are clients that attended
                                 #OF THOSE WHO ATTENDED, HOW MANY ARE ON ABC DTG
                                 dff[['TDF', 'DTG']] = dff[['TDF', 'DTG']].astype(str)
                                 dffp = dff[((dff['TDF']=='ABC') & (dff['DTG']=='DTG'))].copy() #ON ABC/DTG
                                 dfaz = dff[((dff['TDF']=='AZT') & (dff['3TC']=='3TC'))].copy()   #ON AZT/3TC 

                                 #OF THOSE ON ABC DTG, HOW MANY ARE FOR PALD
                                 dffp['WT'] = pd.to_numeric(dffp['WT'], errors='coerce')
                                 dfsp = dffp[(dffp['WT']>2.9) & (dffp['WT']<25)].copy()
                                 dfnp = dffp[((dffp['WT']<3) | (dffp['WT']>24.9))].copy()

                                 if dfsp.shape[0]>0: #if there are pld clients that attended
                                    st.write(dfsp[['ART', 'AG','LD', 'ARVS','ARVD']])
                                       
                                    dfsp['ARVD'] = pd.to_numeric(dfsp['ARVD'], errors='coerce')

                                    dfp1 = dfsp[dfsp['ARVD']<180].copy() #USED THE PACK OF 90
                                    if dfp1.shape[0]>0:
                                        p1 = dfp1.shape[0]
                                    else:
                                        p1 = 0

                                    dfp2 = dfsp[dfsp['ARVD']==180].copy() #USED THE PACK OF 180
                                    if dfp2.shape[0]>0:
                                        p2 = dfp2.shape[0]
                                    else:
                                        p2 = 0


                                 else:
                                    p1 = 0
                                    p2 = 0
                            else:
                                 p1 = 0
                                 p2 = 0

                        


                            


                              
                             


                             
    