import os
import io
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import numpy as np
import time
import streamlit as st

def extract():
    cola,colb,colc = st.columns([1,3,1])
    st.subheader('PIVOT TABLES FOR ECHO DATA USE TX CURR, NOT EVER ENROLLED')
    st.image(r'rename.png')
    #VARIABLES
    lyear = 2026
    lmonth = 6 #LAST MONTH OF THE SAID QTR
    qmonths = [4,5,6] #MONTHS IN THE QUARTER

    vyear = 2025 #VL DATE 6 MONTHS AGO
    vmonth =12  #VL MONTH 6 MONTHS AGO

    vmonths = [1,2,3]

    facility = st.text_input('Enter Facility Name')
    if not facility:
        st.warning('**Please first enter the name of the facility you want to analyse**')
        st.stop()
    else:
         pass   
    

    file = st.file_uploader(f"Upload a CSV containing the {facility}'s TX CURR", type=['csv']) 
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
    
                    df = df.rename(columns= {'ART  ':'ART',  'RD  ':'RD', 'DSD  ':'DSD', 'VD  ':'VD', 'LD  ': 'LD','ARVS  ':'ARVS', 'ARVD ':'ARVD','BCD4  ':'BCD4','AG  ':'AG', 'AS  ':'AS', 'WT  ':'WT', 'VR  ':'VR'})#, 'TPT ': 'TPT'})  

                    df = df.rename(columns= {'ART ':'ART', 'RD ':'RD', 'VD ':'VD', 'LD ': 'LD','WT ':'WT','ARVS ':'ARVS',
                            'AG ':'AG', 'ARVD ':'ARVD', 'AS ':'AS', 'VR ':'VR', 'BCD4 ':'BCD4', 'DSD ':'DSD'})
                    
                    columns = ['ART','AG','AS', 'VD', 'RD','LD','WT', 'ARVS', 'ARVD', 'BCD4', 'VR','DSD']
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
                        st.session_state.df = st.session_state.df.rename(columns= {'ART  ':'ART', 'AS  ':'AS', 'RD  ':'RD', 'DSD  ':'DSD', 'VD  ':'VD', 'LD  ': 'LD','ARVS  ':'ARVS','ARVD  ':'ARVD','AG  ':'AG','BCD4  ':'BCD4','VR  ':'VR'})
                        st.session_state.df = st.session_state.df.rename(columns= {'ART ':'ART', 'RD ':'RD', 'DSD ':'DSD','VD ':'VD',  'LD ': 'LD', 'AG ':'AG', 'ARVS ':'ARVS', 'ARVD ':'ARVD', 'BCD4 ':'BCD4', 'VR ':'VR'})
                        df = st.session_state.df.copy()
                    
                        df = df[['ART','AG','AS', 'VD', 'RD','LD','WT', 'ARVS', 'ARVD', 'BCD4', 'VR', 'DSD']].copy()
                        
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
                        df['AS'] = df['AS'].astype(str)
                                    
                        y = pd.DataFrame({'ART' :['2','3','4','5'], 'RD':['1-1-1',1,'1/1/1','3 8 2001'], 'AS':['1-1-1',1,'1/1/1','3 8 2001'],
                                        'VD':['1-1-1',1,'1/1/1','3 8 2001'], 'LD':['1-1-1',1,'1/1/1','3 8 2001'],'ARVS':['1-1-1-1-1',1,'1/1/1/1/1','3 8 2001 4 6']})                        
                        
                        df['RD'] = df['RD'].astype(str)
                        df['AS'] = df['AS'].astype(str)
                        df['VD'] = df['VD'].astype(str)      
                        df['LD'] = df['LD'].astype(str)
                    
                        df['RD'] = df['RD'].str.replace('00:00:00', '', regex=True)
                        df['AS'] = df['AS'].str.replace('00:00:00', '', regex=True)
                        df['VD'] = df['VD'].str.replace('00:00:00', '', regex=True)
                        df['LD'] = df['LD'].str.replace('00:00:00', '', regex=True)
                     
                        df["RD"] = df["RD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["AS"] = df["AS"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["VD"] = df["VD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["LD"] = df["LD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip() 
                       
                        df["RD"] = df["RD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["VD"] = df["VD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["LD"] = df["LD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["AS"] = df["AS"].str.replace(r"\s*\..*$", "", regex=True).str.strip()

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

                        #SORTING THE REGIMENS
                        A = df[df['ARVS'].str.contains('-')].copy()
                        a = df[~df['ARVS'].str.contains('-')].copy()
                        B = a[a['ARVS'].str.contains('/')].copy()
                        C = a[~a['ARVS'].str.contains('/')].copy()
                        E = C[C['ARVS'].str.contains(' ')].copy()
                        D = C[~C['ARVS'].str.contains(' ')].copy()                                
                        A[['TDF', '3TC', 'DTG','DR','DR2']] = A['ARVS'].str.split('-', expand = True)
                        B[['TDF', '3TC', 'DTG','DR','DR2']] = B['ARVS'].str.split('/', expand = True)
                        try:
                            D['ARVS'] =  D['ARVS'].astype(str)
                            D[['TDF', '3TC', 'DTG','DR','DR2']] = D['ARVS'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['ARVS'] =  E['ARVS'].astype(str)
                            E[['TDF', '3TC', 'DTG']] = E['ARVS'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E]) 
                    

                        # SPLITTING ART START DATE
                        df['AS'] = df['AS'].astype(str)
                        A = df[df['AS'].str.contains('-')].copy()
                        a = df[~df['AS'].str.contains('-')].copy()
                        B = a[a['AS'].str.contains('/')].copy()
                        C = a[~a['AS'].str.contains('/')].copy()
                        E = C[C['AS'].str.contains(' ')].copy()
                        D = C[~C['AS'].str.contains(' ')].copy()
                        A[['Ayear', 'Amonth', 'Aday']] = A['AS'].str.split('-', expand = True)
                        B[['Ayear', 'Amonth', 'Aday']] = B['AS'].str.split('/', expand = True)
                        try:
                            D['AS'] = pd.to_numeric(D['AS'], errors='coerce')
                            D['AS'] = pd.to_datetime(D['AS'], origin='1899-12-30', unit='D', errors='coerce')
                            D['AS'] =  D['AS'].astype(str)
                            D[['Ayear', 'Amonth', 'Aday']] = D['AS'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['AS'] = pd.to_datetime(E['AS'],format='%d %m %Y', errors='coerce')
                            E['AS'] =  E['AS'].astype(str)
                            E[['Ayear', 'Amonth', 'Aday']] = E['AS'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E]) 
                       
  

                        df['RD'] = df['RD'].astype(str)
                        df['AS'] = df['AS'].astype(str)
                        df['VD'] = df['VD'].astype(str)
                        df['LD'] = df['LD'].astype(str)
                        
            #             #Clearing NaT from te dates
                     
                        df['RD'] = df['RD'].str.replace('NaT', '',regex=True)
                        df['AS'] = df['AS'].str.replace('NaT', '',regex=True)
                        df['VD'] = df['VD'].str.replace('NaT', '',regex=True)
                        df['LD'] = df['LD'].str.replace('NaT', '',regex=True)       
                           
                 
                        # #SORTING THE LAST ENCOUNTER
                        df[['Lday', 'Lyear']] = df[['Lday', 'Lyear']].apply(pd.to_numeric, errors='coerce')
                        df['Lyear'] = df['Lyear'].fillna(994)
                        a = df[df['Lyear']>31].copy()
                        b = df[df['Lyear']<32].copy()
                        b = b.rename(columns={'Lyear': 'Lday2', 'Lday': 'Lyear'})
                        b = b.rename(columns={'Lday2': 'Lday'})
                        df = pd.concat([a,b])
                        df = df.copy()

                        # #SORTING THE RETURN VISIT DATE YEARS
                        df[['Rday', 'Ryear']] = df[['Rday', 'Ryear']].apply(pd.to_numeric, errors='coerce')
                        df['Ryear'] = df['Ryear'].fillna(994)
                        a = df[df['Ryear']>31].copy()
                        b = df[df['Ryear']<32].copy()
                        b = b.rename(columns={'Ryear': 'Rday2', 'Rday': 'Ryear'})
                        b = b.rename(columns={'Rday2': 'Rday'})
                        df = pd.concat([a,b])

                        #SORTING THE VIRAL LOAD MONTHS AND YEAR
                        df[['Vyear', 'Vmonth', 'Vday']] =df[['Vyear', 'Vmonth', 'Vday']].apply(pd.to_numeric, errors = 'coerce') 
                        df['Vyear'] = df['Vyear'].fillna(994)
                        a = df[df['Vyear']>31].copy()
                        b = df[df['Vyear']<32].copy()
                        #c = df[]
                        b = b.rename(columns={'Vyear': 'Vday2', 'Vday': 'Vyear'})
                        b = b.rename(columns={'Vday2': 'Vday'})
                        df = pd.concat([a,b])

                        #SORTING THE ART START YEARS
                        df[['Ayear', 'Amonth', 'Aday']] =df[['Ayear', 'Amonth', 'Aday']].apply(pd.to_numeric, errors = 'coerce')
                        df['Ayear'] = df['Ayear'].fillna(994)
                        a = df[df['Ayear']>31].copy()
                        b = df[df['Ayear']<32].copy()
                        b = b.rename(columns={'Ayear': 'Aday2', 'Aday': 'Ayear'})
                        b = b.rename(columns={'Aday2': 'Aday'})
                        df = pd.concat([a,b])
                        dfe = df.shape[0]

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
                            st.info(f'**{at} CALHIV listed below have no weight in the uploaded extract**')
                            st.write('**first input their weight to avoid errors**')
                            testrw =testrw[['ART','AG', 'RD', 'WT']].copy()
                            colh, colj = st.columns([2,1])
                            csvd = testrw.to_csv(index=False)
                            colh.write(testrw)
                            with colj:
                                st.download_button(
                                        label="Download them from here",
                                        data= csvd,
                                        file_name=f"{facility}_missing_wight.csv",
                                        mime="text/csv",
                                        key="download_missing_weight"
                                        )
                            
                            st.stop()
                        else:
                            #  st.write('**All CALHIV have weight, proceeding with the analysis**')
                            pass
                        b11 = df.shape[0]

                        if b11<1:
                             st.info("**THIS EXTRACT DOESN'T AHVE ANY CALHIV BTN 0 to 19 YEAR, CHECK MANUALLY**")
                             st.stop()
                        else:
                            dfb11 = pd.DataFrame({ " ":['TOTAL'],
                                                        'CALHIV': [b11],
                                                        '': 'Total number of CALHIV currently recieving care at this HF?'})
                            #st.write(dfb11)
                            df['WT'] = pd.to_numeric(df['WT'], errors='coerce')
                            b2 = df[df['WT']>2.9].copy()
                            dfb2 = df.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                            dfb2 = pd.merge(dfb2, datyw, on='WEIGHT BANDS', how='right')
                            dfb2['TOTAL'] = dfb2['TOTAL'].fillna(0)
                            dfb2['R'] = dfb2['WEIGHT BANDS'].map(wmapper)
                            dfb2 = dfb2.sort_values('R').drop(columns='R')
                            #st.write(dfb2)
                             
                            dfb3 = df.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                            dfb3 = pd.merge(dfb3, datya, on='AGE BANDS', how='right')
                            dfb3['TOTAL'] = dfb3['TOTAL'].fillna(0)
                            dfb3['R'] = dfb3['AGE BANDS'].map(amapper)
                            dfb3 = dfb3.sort_values('R').drop(columns='R')
                            #st.write(dfb3)


                            df[['TDF', 'DTG']] = df[['TDF', 'DTG']].astype(str)
                            dfc = df[((df['TDF']=='ABC') & (df['DTG']=='DTG'))].copy()
                            dfc2 = df[((df['TDF']=='TDF') & (df['DTG']=='DTG'))].copy()
                            
                            a3cn = dfc2[(dfc['WT']>29.9)].copy()
                            dfn = df[~((df['TDF']=='ABC') & (df['DTG']=='DTG'))].copy()
                            a1 = dfc.shape[0]
                            a2 = b11 - a1
                            if a2 > 0:
                                cola, colb = st.columns(2)
                                cola.markdown(f'**You have {a2} CALHIV that are not on ABC-DTG regimen**')
                                with colb.expander('**Click here to see them**'):
                                    dfn = dfn[['ART', 'AG', 'WT', 'ARVS']].copy()
                                    dfn = dfn.reset_index(drop=True)
                                    st.write(dfn)
                            else:
                                 pass
                            #children on ABC-DTG]
                            #st.write(dfc[['ART','AG', 'WT', 'ARVS']])
                            dfc['WT'] = pd.to_numeric(dfc['WT'], errors='coerce')
                            dfc = dfc[((dfc['WT']>2.9) & (dfc['WT']<6))].copy()
                            a3b = dfc[((dfc['WT']>24.9) & (dfc['WT']<30))].copy()
                            a3 = pd.concat([dfc, a3b,a3cn])
                            st.write(a3)
                            st.stop()
                            if a3 > 0:
                                 #st.write(dfc[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                dfc1 = dfc.groupby('WEIGHT BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfc1 = pd.merge(dfc1, datyw, on='WEIGHT BANDS', how='right')
                                dfc1['TOTAL'] = dfc1['TOTAL'].fillna(0)
                                dfc1['R'] = dfc1['WEIGHT BANDS'].map(wmapper)
                                dfc1 = dfc1.sort_values('R').drop(columns='R')
                                # st.write(dfc1)
                                dfc2 = dfc.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfc2 = pd.merge(dfc2, datya, on='AGE BANDS', how='right')
                                dfc2['TOTAL'] = dfc2['TOTAL'].fillna(0)
                                dfc2['R'] = dfc2['AGE BANDS'].map(amapper)
                                dfc2 = dfc2.sort_values('R').drop(columns='R')
                                # st.write(dfc2)  
                            else:
                                 dfc1 = datyw.copy()
                                 dfc1['TOTAL'] = 0
                                #  st.write(dfc1)  
                                 dfc2 = datya.copy()
                                 dfc2['TOTAL'] = 0
                                #  st.write(dfc2) 
                            
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
                                    # st.write(dfc5)
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
                                # st.write(dfc6)

                                #OF THESE, WHO IS ON PALD? 
                                dfcpc[['TDF', 'DTG']] = dfcpc[['TDF', 'DTG']].astype(str)
                                dfc7 = dfcpc[((dfcpc['TDF']=='ABC') & (dfcpc['DTG']=='DTG'))].copy()
                                
                                #NOT ON PALD BUT ELIGIBLE BY WEIGHT
                                dfc8 = dfcpc[~((dfcpc['TDF']=='ABC') & (dfcpc['DTG']=='DTG'))].copy()
                                if dfc8.shape[0]>0:
                                        st.info('**You have CALHIV who are eligible for PALD by weight but are not on ABC-DTG regimen**')
                                        st.write(dfc8[['ART', 'AG', 'WT', 'ARVS']].reset_index(drop=True))
                                        confirm = st.radio('Is this correct?', ('Yes', 'No'), horizontal=True, index=None)
                                        if not confirm:
                                             st.stop()
                                        elif confirm == 'No':
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
                            aptc = dff.shape[0]
                            #st.write(aptc)
                            #st.write(dff[['ART', 'LD', 'WT', 'ARVS', 'TDF', '3TC', 'DTG']].reset_index(drop=True))
                            
                            if dff.shape[0]>0: #if there are clients that attended
                                p4 = 0
                                #st.write(dff)
                                #OF THOSE WHO ATTENDED, HOW MANY ARE ON ABC DTG
                                dff[['TDF', 'DTG']] = dff[['TDF', 'DTG']].astype(str)
                                dffp = dff[((dff['TDF']=='ABC') & (dff['DTG']=='DTG'))].copy() #ON ABC/DTG
                                dfaz = dff[((dff['TDF']=='AZT') & (dff['3TC']=='3TC'))].copy()   #ON AZT/3TC 
                                dfdar = dff[(dff['DTG']=='DRV/r')].copy()   #ON DRV/r

                                #OF THOSE ON ABC DTG, HOW MANY ARE FOR PALD
                                dffp['WT'] = pd.to_numeric(dffp['WT'], errors='coerce')
                                dfsp = dffp[(dffp['WT']>5.9) & (dffp['WT']<25)].copy() #PALD
                                dfnp = dffp[(dffp['WT']<6)].copy() #ABC/3TC 120/60

                                if dfsp.shape[0]>0: #if there are pld clients that attended
                                    # st.write(dfsp[['ART', 'AG','LD', 'ARVS','ARVD']])
                                       
                                    dfsp['ARVD'] = pd.to_numeric(dfsp['ARVD'], errors='coerce')

                                    dfp1 = dfsp[dfsp['ARVD']<180].copy() #USED THE PACK OF 90
                                    if dfp1.shape[0]>0:
                                        p1 = dfp1['ARVD'].sum()
                                        p1 = p1/90
                                        p1 = int(Decimal(p1).quantize(0, rounding=ROUND_HALF_UP))
                                        #st.write(p1)
                                    else:
                                        p1 = 0

                                    dfp2 = dfsp[dfsp['ARVD']==180].copy() #USED THE PACK OF 180
                                    if dfp2.shape[0]>0:
                                        p2 = dfp2['ARVD'].sum()
                                        p2 = p2/180 
                                        p2 = int(Decimal(p2).quantize(0, rounding=ROUND_HALF_UP))
                                    
                                    else:
                                        p2 = 0
                                else:
                                    p1 = 0
                                    p2 = 0
                                if dfnp.shape[0]>0: #if there are non pld clients that attended
                                    st.write(dfnp[['ART', 'AG','LD', 'ARVS','ARVD']])
                                    dfnp['ARVD'] = pd.to_numeric(dfnp['ARVD'], errors='coerce')
                                    p3 = dfnp['ARVD'].sum()
                                    p3 = p3/30
                                    p3 = int(Decimal(p3).quantize(0, rounding=ROUND_HALF_UP))
                                else:
                                    p3 = 0
                                #THOSE ON AZT/3TC
                                if dfaz.shape[0]>0:
                                    #st.write(dfaz[['ART', 'AG','LD', 'ARVS','ARVD']])
                                    dfaz['ARVD'] = pd.to_numeric(dfaz['ARVD'], errors='coerce')
                                    p5 = dfaz['ARVD'].sum()
                                    p5 = p5/60
                                    p5 = int(Decimal(p5).quantize(0, rounding=ROUND_HALF_UP))
                                else:
                                    p5 = 0 

                                if dfdar.shape[0]>0: #THOSE ON DRV/r
                                    #st.write(dfdar[['ART', 'AG','LD', 'ARVS','ARVD']])
                                    dfdar['ARVD'] = pd.to_numeric(dfdar['ARVD'], errors='coerce')
                                    p6 = dfdar['ARVD'].sum()
                                    p6 = p6/60
                                    p6 = int(Decimal(p6).quantize(0, rounding=ROUND_HALF_UP))  
                                else:
                                    p6 = 0           

                            else:
                                 p1 = 0
                                 p2 = 0
                                 p3 = 0
                                 p4 = 0
                                 p5 = 0
                                 p6 = 0
                            df3 = pd.DataFrame({
                                'DRUG': ['ABC/3TC/DTG 60/30/5 90 PACK','ABC/3TC/DTG 60/30/5 180 PACK','ABC/3TC 120/60','DTG 10 MG','AZT/3TC 60/30','DARUNAVIR'],
                                'STOCK USED': [p1, p2, p3, p4, p5, p6],
                                'NOTE': ['','','','CONFIRM WITH FACILITY FOR THIS','',"CHECK IF YOU DON'T HAVE 2nd AND 3rd LINERS"]
                                })
                                
                            #st.write(df3)

                            ###TX NEWS
                            df[['Ayear', 'Amonth']] = df[['Ayear', 'Amonth']].apply(pd.to_numeric, errors='coerce')
                            dftn = df[((df['Ayear']==lyear)  & (df['Amonth'].isin(qmonths)))].copy()
                        

                            if dftn.shape[0]>0: #IF THEY ARE THERE
                                news = dftn[['ART', 'AS', 'RD', 'BCD4']].copy()
                                #st.write(dftn[['ART', 'AG','RD', 'ARVS','BCD4']])
                                dfg1 = dftn.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfg1 = pd.merge(dfg1, datya, on='AGE BANDS', how='right')
                                dfg1['TOTAL'] = dfg1['TOTAL'].fillna(0)
                                dfg1['R'] = dfg1['AGE BANDS'].map(amapper)
                                dfg1 = dfg1.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(dfg1)
                                dftn['BCD4'] = pd.to_numeric(dftn['BCD4'], errors='coerce') 

                                #THOSE WITH LOW CD4 COUNT
                                dfcd = dftn[dftn['BCD4']<200].copy()
                                if dfcd.shape[0]>0:
                                    lows = dfcd[['ART', 'AS', 'RD', 'BCD4']].copy()
                                    lows['NEEDED'] = 'Check if they had TB LAM and all the AHD Cascade as required in section G.1'
                                    #st.write(dfcd[['ART', 'AG','AS', 'ARVS','BCD4']])
                                    dfg12 = dfcd.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfg12 = pd.merge(dfg12, datya, on='AGE BANDS', how='right')
                                    dfg12['TOTAL'] = dfg12['TOTAL'].fillna(0)
                                    dfg12['R'] = dfg12['AGE BANDS'].map(amapper)
                                    dfg12 = dfg12.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    #st.write(dfg12)
                                else:
                                     dfg12 = datya.copy()
                                     dfg12['TOTAL'] = 0
                                    #  st.write(dfg2)
                                     lows = pd.DataFrame()
                                    #  st.write(lows)

                                
                            else:
                                 dfg1 = datya.copy()
                                 dfg1['TOTAL'] = 0
                                #  st.write(dfg1)
                                 dfg12 = dfg1.copy()
                                
                            #NONE SUPPRESSORS
                            df[['Vyear', 'Vmonth']] = df[['Vyear', 'Vmonth']].apply(pd.to_numeric, errors='coerce')
                            has = df[((df['Vyear']>vyear)| ((df['Vyear']==vyear) & (df['Vmonth']>vmonth)))].copy()
                            has['VR'] = pd.to_numeric(has['VR'], errors='coerce')
                            dfns = has[has['VR']>999].copy()
                            if dfns.shape[0]>0:
                                nones = dfns[['ART', 'RD', 'VD', 'VR']].copy()
                                #  st.write(nones)
                                nones['NEEDED'] = 'Check if they had TB LAM and all the AHD Cascade as required in section G.2'
                                dfg2 = dfns.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfg2 = pd.merge(dfg2, datya, on='AGE BANDS', how='right')
                                dfg2['TOTAL'] = dfg2['TOTAL'].fillna(0)
                                dfg2['R'] = dfg2['AGE BANDS'].map(amapper)
                                dfg2 = dfg2.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(dfg2)
                            else:
                                 dfg2 = datya.copy()
                                 dfg2['TOTAL'] = 0
                                 nones = pd.DataFrame()

                            
                            #ON APPOINTMENT
                            df[['Rmonth', 'Ryear']] = df[['Rmonth', 'Ryear']].apply(pd.to_numeric, errors='coerce')
                            dfm = df[((df['Ryear']==lyear)& (df['Rmonth']==lmonth))].copy()
                            aptm = dfm.shape[0]

                            apt = aptm + aptc
                            if apt>0:
                                dfh1 = pd.DataFrame( {'RETENTION':['CALHIV this HF expecting (H.1)','how many actually returned (H.1.1)'],
                                                                   'TOTAL': [apt, aptc]
                                                                   })
                                # st.write(dfh1)
                            else:
                                dfh1 = pd.DataFrame( {'RETENTION':['CALHIV this HF expecting (H.1)','how many actually returned (H.1.1)'],
                                                                   'TOTAL': [0,0]
                                                                   })

                            #VL DATA
                            df[['Vyear', 'Vmonth']] = df[['Vyear', 'Vmonth']].apply(pd.to_numeric, errors='coerce')
                            dfvl = df[((df['Vyear']>vyear)|((df['Vyear']==vyear) & (df['Vmonth']>vmonth)))].copy()
                            dfnvl = df[((df['Vyear']<vyear)|((df['Vyear']==vyear) & (df['Vmonth']<=vmonth)))].copy()
                            # st.write(dfvl.shape[0])
                            # st.write(dfnvl.shape[0])

                            if dfvl.shape[0]>0: #THOSE WITH VLS
                                #st.write(dfvl[['ART', 'AG','VD', 'ARVS','VR']])
                                dfi = dfvl.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                dfi = pd.merge(dfi, datya, on='AGE BANDS', how='right')
                                dfi['TOTAL'] = dfi['TOTAL'].fillna(0)
                                dfi['R'] = dfi['AGE BANDS'].map(amapper)
                                dfi = dfi.sort_values('R').drop(columns='R').reset_index(drop=True)
                                # st.write(dfi)

                                dfvl['VR'] = pd.to_numeric(dfvl['VR'], errors='coerce').fillna(0)
                                dfvls = dfvl[dfvl['VR']<1000].copy() #SUPPRESSED
                                dfvlns = dfvl[dfvl['VR']>=1000].copy() #NONE SUPPESSED

                                if dfvls.shape[0]>0:
                                    #st.write(dfvls[['ART', 'AG','VD', 'ARVS','VR']])
                                    dfi3 = dfvls.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfi3 = pd.merge(dfi3, datya, on='AGE BANDS', how='right')
                                    dfi3['TOTAL'] = dfi3['TOTAL'].fillna(0)
                                    dfi3['R'] = dfi3['AGE BANDS'].map(amapper)
                                    dfi3 = dfi3.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfi3)
                                else:
                                    dfi3 = datya.copy()
                                    dfi3['TOTAL'] = 0
                                    #st.write(dfj)

                                if dfvlns.shape[0]>0:
                                    #st.write(dfvls[['ART', 'AG','VD', 'ARVS','VR']])
                                    dfi4 = dfvlns.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfi4 = pd.merge(dfi4, datya, on='AGE BANDS', how='right')
                                    dfi4['TOTAL'] = dfi4['TOTAL'].fillna(0)
                                    dfi4['R'] = dfi4['AGE BANDS'].map(amapper)
                                    dfi4 = dfi4.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfi4)
                                else:
                                    dfi4 = datya.copy()
                                    dfi4['TOTAL'] = 0
                                    #st.write(dfj)
                            else:
                                 dfi = datya.copy()
                                 dfi['TOTAL'] = 0
                                 dfi3 = dfi.copy()
                                 dfi4 = dfi.copy()
                            if dfnvl.shape[0] > 0:
                                    #st.write(dfnvl[['ART', 'AG','VD', 'ARVS','VR']])
                                    dfi2 = dfnvl.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfi2 = pd.merge(dfi2, datya, on='AGE BANDS', how='right')
                                    dfi2['TOTAL'] = dfi2['TOTAL'].fillna(0)
                                    dfi2['R'] = dfi2['AGE BANDS'].map(amapper)
                                    dfi2 = dfi2.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfi2)
                            else:
                                dfi2 = datya.copy()
                                dfi2['TOTAL'] = 0 
                            hlvs = pd.DataFrame()
                            dfj1 = datya.copy()
                            dfj1["TOTAL"] = 0
                            llvs = pd.DataFrame()

                            if dfvl.shape[0]>0:
                                dfvl[['Vyear', 'Vmonth']] = dfvl[['Vyear', 'Vmonth']].apply(pd.to_numeric, errors='coerce')
                                #VLS DOE LAST MONTH, DUE FOR IAC
                                dfvj = dfvl[((dfvl['Vyear']== vyear) & (dfvl['Vmonth'].isin(vmonths)))].copy()     
                                dfvj['VR'] = pd.to_numeric(dfvj['VR'], errors = 'coerce')
                                #LLVs
                                dfvk = dfvj[((dfvj['VR']>200) & (dfvj['VR']<1000))].copy() 
                                dfvk['VR'] = pd.to_numeric(dfvk['VR'], errors = 'coerce')
                                dfvk = dfvk[dfvk['VR']!=400].copy()
                                #HLVS
                                dfvj = dfvj[dfvj['VR']>999].copy()
                                st.write(dfvj)
                                
                                if dfvj.shape[0]>0: #THE HLVS
                                     #st.write(dfvj[['ART', 'AG','VD', 'ARVS','VR']])
                                    hlvs = dfvj[['ART', 'RD', 'VD', 'VR']].copy()
                                    hlvs['NEEDED'] = 'Check if they completed IACs as needed in section J'
                                    dfj1 = dfvj.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfj2 = pd.merge(dfj1, datya, on='AGE BANDS', how='right')
                                    dfj1['TOTAL'] = dfj1['TOTAL'].fillna(0)
                                    dfj1['R'] = dfj1['AGE BANDS'].map(amapper)
                                    dfj1 = dfj1.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    #st.write(dfj1)
                                else:
                                    dfj1 = datya.copy()
                                    dfj1['TOTAL'] = 0
                                    hlvs = pd.DataFrame()

                                if dfvk.shape[0]>0: #THE LLVS
                                     #st.write(dfvk[['ART', 'AG','VD', 'ARVS','VR']])
                                    llvs = dfvk[['ART', 'RD', 'VD', 'VR']].copy()
                                    llvs['NEEDED'] = 'Check if they completed IACs as needed in section K'
                                    dfk1 = dfvk.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfk2 = pd.merge(dfk1, datya, on='AGE BANDS', how='right')
                                    dfk1['TOTAL'] = dfk1['TOTAL'].fillna(0)
                                    dfk1['R'] = dfk1['AGE BANDS'].map(amapper)
                                    dfk1 = dfk1.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfk1)

                                else:
                                    dfk1 = datya.copy()
                                    dfk1['TOTAL'] = 0
                                    # st.write(dfk1)
                                    llvs = pd.DataFrame()
                            else:
                                 dfj1 = datya.copy()
                                 dfj1['TOTAL'] = 0
                                 dfk1 = dfj1.copy()
                                #  st.writ(dfk1)
                            dfL = pd.DataFrame({ " ":['TOTAL'],
                                                        'CALHIV': [b11],
                                                        '': 'Total CALHIV in care'})
                            dfL11 = datya.copy()
                            dfL11['TOTAL'] = 0
                            dfL11['NOTE'] = 'Only put zero if this facility has no teen clubs'
                            #st.write(dfL12)
                           #MMD SECTION
                            df['ARVD'] = pd.to_numeric(df['ARVD'], errors = 'coerce')

                            dfL1 = df[df['ARVD']>89].copy()
                            dfL1['ARVD'] = pd.to_numeric(dfL1['ARVD'], errors = 'coerce')

                            #3 months
                            dfL12 = dfL1[dfL1['ARVD'] < 179].copy()
                            if dfL12.shape[0] > 0:
                                    #st.write(dfL12[['ART', 'AG','VD', 'ARVS','VR']])
                                    dfL12 = dfL12.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfL12 = pd.merge(dfL12, datya, on='AGE BANDS', how='right')
                                    dfL12['TOTAL'] = dfL12['TOTAL'].fillna(0)
                                    dfL12['R'] = dfL12['AGE BANDS'].map(amapper)
                                    dfL12 = dfL12.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfL12)
                            else:
                                dfL12 = datya.copy()
                                dfL12['TOTAL'] = 0
                                 
                            #6 MONTHS
                            dfL13 = dfL1[dfL1['ARVD'] >= 180].copy()
                            if dfL13.shape[0] > 0:
                                    #st.write(dfL13[['ART', 'AG','VD', 'ARVS','VR']])
                                    dfL13 = dfL13.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                    dfL13 = pd.merge(dfL13, datya, on='AGE BANDS', how='right')
                                    dfL13['TOTAL'] = dfL13['TOTAL'].fillna(0)
                                    dfL13['R'] = dfL13['AGE BANDS'].map(amapper)
                                    dfL13 = dfL13.sort_values('R').drop(columns='R').reset_index(drop=True)
                                    # st.write(dfL13)
                            else:
                                dfL13 = datya.copy()
                                dfL13['TOTAL'] = 0
                        
                        df['DSD'] =  df['DSD'].astype(str)
                        dmap = {
                             'Community Drug Distribution Point': 'CDDP',
                             'Community Client Led ART Delivery': 'GCLAD',
                             'Facility Based Groups':'FBG',
                             'Fast Track Drug Refill':'FTDR',
                             'Facility Based Individual Management': 'FBIM'

                        }
                        df['DSD'] = df['DSD'].astype(str)
                        df['DSD'] = df['DSD'].fillna('Fast Track Drug Refill')
                        df['DSD'] = df['DSD'].str.replace('NaT','Fast Track Drug Refill')
                        df['DSD'] = df['DSD'].str.replace('nan','Fast Track Drug Refill')
                        df['DSDM'] = df['DSD'].map(dmap)
                        df['DSDM'] = df['DSDM'].astype(str)
                        

                        cddp = df[df['DSDM']== 'CDDP'].copy()
                        if cddp.shape[0]>0:
                                #st.write(cddp[['ART', 'AG','VD', 'ARVS','DSD']])
                                cddp = cddp.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                cddp = pd.merge(cddp, datya, on='AGE BANDS', how='right')
                                cddp['TOTAL'] = cddp['TOTAL'].fillna(0)
                                cddp['R'] = cddp['AGE BANDS'].map(amapper)
                                cddp = cddp.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(cddp)
                        else:
                                cddp = datya.copy()
                                cddp['TOTAL'] = 0

                        cclad = df[df['DSDM']== 'CCLAD'].copy()
                        if cclad.shape[0]>0:
                                #st.write(cclad['ART', 'AG','VD', 'ARVS','DSD']])
                                cclad = cclad.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                cclad = pd.merge(cclad, datya, on='AGE BANDS', how='right')
                                cclad['TOTAL'] = cclad['TOTAL'].fillna(0)
                                cclad['R'] = cclad['AGE BANDS'].map(amapper)
                                cclad = cclad.sort_values('R').drop(columns='R').reset_index(drop=True)
                                # st.write(cclad)
                        else:
                                cclad = datya.copy()
                                cclad['TOTAL'] = 0
                                #st.write(cclad)
                        crddp = df[df['DSDM']== 'CRDDP'].copy()
                        if crddp.shape[0]>0:
                                #st.write(crddp[['ART', 'AG','VD', 'ARVS','DSD']])
                                crddp = crddp.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                crddp = pd.merge(crddp, datya, on='AGE BANDS', how='right')
                                crddp['TOTAL'] = crddp['TOTAL'].fillna(0)
                                crddp['R'] = crddp['AGE BANDS'].map(amapper)
                                crddp = crddp.sort_values('R').drop(columns='R').reset_index(drop=True)
                                # st.write(crddp)
                        else:
                                crddp = datya.copy()
                                crddp['TOTAL'] = 0
                                crddp['NOTE'] = 'IF THIS FACILITY HAS A PHARMAVY, CHECK ART ACCESS, AND SUBTRACT THEM FROM CDDP DATA'
                                #st.write(crddp)

                        fbim = df[df['DSDM']== 'FBIM'].copy()
                        if fbim.shape[0]>0:
                                #st.write(fbim[['ART', 'AG','VD', 'ARVS','DSD']])
                                fbim = fbim.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                fbim = pd.merge(fbim, datya, on='AGE BANDS', how='right')
                                fbim['TOTAL'] = fbim['TOTAL'].fillna(0)
                                fbim['R'] = fbim['AGE BANDS'].map(amapper)
                                fbim = fbim.sort_values('R').drop(columns='R').reset_index(drop=True)
                                # st.write(fbim)
                        else:
                                fbim = datya.copy()
                                fbim['TOTAL'] = 0 
                                #st.write(fbim)

                        fbg = df[df['DSDM']== 'FBG'].copy()
                        if fbg.shape[0]>0:
                                #st.write(fbg[['ART', 'AG','VD', 'ARVS','DSD']])
                                fbg = fbg.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                fbg = pd.merge(fbg, datya, on='AGE BANDS', how='right')
                                fbg['TOTAL'] = fbg['TOTAL'].fillna(0)
                                fbg['R'] = fbg['AGE BANDS'].map(amapper)
                                fbg = fbg.sort_values('R').drop(columns='R').reset_index(drop=True)
                                #st.write(fbg)
                        else:
                                fbg = datya.copy()
                                fbg['TOTAL'] = 0   
                        
                        ftdr = df[df['DSDM']== 'FTDR'].copy()
                        if ftdr.shape[0]>0:
                                #st.write(ftdr[['ART', 'AG','VD', 'ARVS','DSD']])
                                ftdr = ftdr.groupby('AGE BANDS')['ART'].size().reset_index().rename(columns={'ART': 'TOTAL'})
                                ftdr = pd.merge(ftdr, datya, on='AGE BANDS', how='right')
                                ftdr['TOTAL'] = ftdr['TOTAL'].fillna(0)
                                ftdr['R'] = ftdr['AGE BANDS'].map(amapper)
                                ftdr = ftdr.sort_values('R').drop(columns='R').reset_index(drop=True)
                                # st.write(ftdr)
                        else:
                                ftdr = datya.copy()
                                ftdr['TOTAL'] = 0
                             
    
    # if file is not None:# and st.session_state.dfw:
        
    #         # Create an in-memory BytesIO buffer
                        output = io.BytesIO()
 
    
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                dfb11.to_excel(writer, sheet_name="B.1.1", index=False) 
                                dfb2.to_excel(writer, sheet_name="B.2", index=False)
                                dfb3.to_excel(writer, sheet_name="B.3", index=False)
                                dfc1.to_excel(writer, sheet_name="C.1", index=False)
                                dfc2.to_excel(writer, sheet_name="C.2", index=False)
                                dfc2.to_excel(writer, sheet_name="C.4", index=False)
                                dfc5.to_excel(writer, sheet_name="C.5", index=False)
                                dfc1.to_excel(writer, sheet_name="C.6", index=False)
                                dfc7.to_excel(writer, sheet_name="C.7", index=False)
                                df3.to_excel(writer, sheet_name="F.3.2", index=False)
                                dfg1.to_excel(writer, sheet_name="G.1", index=False)
                                dfg12.to_excel(writer, sheet_name="G.1.2", index=False)
                                dfg2.to_excel(writer, sheet_name="G.2.1", index=False)
                                dfh1.to_excel(writer, sheet_name="H", index=False)
                                dfi.to_excel(writer, sheet_name="I.1", index=False)
                                dfi2.to_excel(writer, sheet_name="I.2", index=False)
                                dfi3.to_excel(writer, sheet_name="I.3", index=False)
                                dfi4.to_excel(writer, sheet_name="I.4", index=False)
                                dfj1.to_excel(writer, sheet_name="J.1.1", index=False)
                                dfk1.to_excel(writer, sheet_name="K.1.1", index=False)
                                dfL.to_excel(writer, sheet_name="L", index=False)
                                dfL11.to_excel(writer, sheet_name="L.1.1", index=False)
                                dfL12.to_excel(writer, sheet_name="L.1.2", index=False)
                                dfL13.to_excel(writer, sheet_name="L.1.3", index=False)
                                cddp.to_excel(writer, sheet_name="L.1.4(CDDP)", index=False)
                                cclad.to_excel(writer, sheet_name="L.1.5(CCLAD)", index=False)
                                crddp.to_excel(writer, sheet_name="L.1.6(CRPDDP)", index=False)
                                fbim.to_excel(writer, sheet_name="L.1.7(FBIM)", index=False)
                                fbg.to_excel(writer, sheet_name="L.1.8(FBG)", index=False)
                                ftdr.to_excel(writer, sheet_name="L.1.9(FTDR)", index=False)                            

                                output.seek(0) 

    
                        col1, col2 = st.columns (2)
                                        
                        col1.download_button(
                                            label="📥 DOWNLOAD ECHO DATA",
                                            data=output,
                                            file_name=f"{facility}_ECHO DATA.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                            )
            
                
                        # Create an in-memory BytesIO buffer
                        
                        if dftn.shape[0]>0:
                            dftn['BCD4'] = pd.to_numeric(dftn['BCD4'], errors='coerce') 

                            #THOSE WITH LOW CD4 COUNT
                            dfcd = dftn[dftn['BCD4']<200].copy()
                            if dfcd.shape[0]>0:
                                    lows = dfcd[['ART', 'AS', 'RD', 'BCD4']].copy()
                                    lows['NEEDED'] = 'Check if they had TB LAM and all the AHD Cascade as required in section G.1'
                            else:
                                lows = pd.DataFrame()
                        else:
                            lows = pd.DataFrame()

                        output2 = io.BytesIO()  
                        with pd.ExcelWriter(output2, engine="xlsxwriter") as writer:
                                if lows.shape[0]>0:
                                    lows.to_excel(writer, sheet_name=" <200 (SECTION G.1)", index=False) 
                                else:
                                    pass

                                if nones.shape[0]>0:
                                    nones.to_excel(writer, sheet_name="NS (SECTION G.2)", index=False) 
                                else:
                                    pass
                            
                                if hlvs.shape[0]>0:
                                    hlvs.to_excel(writer, sheet_name="HLVS (SECTION J)", index=False) 
                                else:
                                    pass

                                if llvs.shape[0]>0:
                                    llvs.to_excel(writer, sheet_name="LLVS (SECTION K)", index=False) 
                                else:
                                    pass

                        output2.seek(0) 
            
                        if nones.shape[0]>1 or hlvs.shape[0]>0 or llvs.shape[0]> 0 or lows.shape[0]>0:
                            col2.download_button(
                                label="📥 DOWNLOAD LINELISTS FOR FOLLOW UP",
                                data=output2,
                                file_name=f"{facility}_LINELISTS.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key = 'LINELISTS'
                                )    
                        else:                                
                                 pass  

                        st.success('**CREATED BY Dr. LUMINSA DESIRE**')
                        
pages = {
    "READER:": [
        st.Page(extract, title="ECHO DATA CALL"),
    ],
   
}

pg = st.navigation(pages)
pg.run()
