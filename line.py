import os
import io
import time
import traceback
import datetime as dt
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

#Clear cache at the very start of the app
st.cache_data.clear()
st.cache_resource.clear()
columnss = ['ART','AS', 'VD', 'RD','DD','LD', 'TO','NTO']
# st.write(columnss)

def extract():
    cola,colb,colc = st.columns([1,3,1])
    colb.subheader('VL LINELISTS')   
    today = datetime.now()
    todayd = today.strftime("%Y-%m-%d")# %H:%M")
    week = today.strftime("%V")
    wk = int(week) + 13
    # wk = int(week) - 39 # USE ONLY IN Q4
    thisweek = int(week)-1
    woke = wk-2
    cola,colb = st.columns(2)
    cola.write(f"**DATE TODAY:    {todayd}**")
    colb.write(f"**CURRENT SURGE WEEK:    {wk}**")

    file = st.file_uploader("Upload your EMR extract here", type=['xlsx']) 
    if file is not None:   
        if 'fd' not in st.session_state:
            fileN = file.name
            name = os.path.basename(fileN).split('.')[0]
            st.session_state.fd = name
        else:
            pass
    else:
        pass
    if file is not None: 
       fileN = file.name
       namey = os.path.basename(fileN).split('.')[0]
       if str(namey) != str(st.session_state.fd):
                #st.info(f'DATA FOR {facy} NOT SUBMITTED')
                st.session_state.submited = False
                st.cache_data.clear()
                st.session_state.fd = namey
                st.cache_resource.clear()
                st.session_state.submited =False
                st.session_state.df = None
                st.session_state.reader =False#
                time.sleep(1)
                        
    if 'submited' not in st.session_state:
        st.session_state.submited =False
    if 'df' not in st.session_state:
        st.session_state.df = None

    if 'reader' not in st.session_state:
        st.session_state.reader =False#
    #ext = None
    if file is not None and not st.session_state.reader:
        # Get the file name
        fileN = file.name
        ext = os.path.basename(fileN).split('.')[1]

    #df = None
    if file is not None and not st.session_state.reader:
        wb = load_workbook(file)
        sheets = wb.sheetnames
        if len(sheets)>1:
            st. warning('THIS EXTRACT HAS MULTIPLE SHEETS, I CAN NOT TELL WHICH ONE TO READ')
            st.write(sheets)
            time.sleep(3)
            st.info('DELETE ALL THE OTHER SHEETS AND REMAIN WITH ONE THAT HAS THE EVER ENROLLED')
            st.stop()
        else:
            pass

    if file is not None and not st.session_state.reader:
                    st.session_state.df = pd.read_excel(file)
                    df = st.session_state.df
                    df.columns = df.columns.str.strip()
        
                    columns = ['ART','AS', 'VD', 'RD','DD','LD', 'TO','NTO']
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
                          # Convert 'ART' column to string and create 'ART' column with numeric part to remove blanks
                        st.session_state.df.colums = st.session_state.df.columns.str.strip()
                        df = st.session_state.df.copy()
                        
                        df['ART'] = df['ART'].astype(str)
                        df['A'] = df['ART'].str.replace('[^0-9]', '', regex=True)
                        df['A'] = pd.to_numeric(df['A'], errors= 'coerce')
                        df = df[df['A']>0].copy()
                    
                        df['AS'] = df['AS'].astype(str)
                        df['RD'] = df['RD'].astype(str)
                        df['TO'] = df['TO'].astype(str)
                        df['VD'] = df['VD'].astype(str)
                        df['DD'] = df['DD'].astype(str)                 
                        y = pd.DataFrame({'ART' :['2','3','4','5'],  'RD':['1-1-1',1,'1/1/1','3 8 2001'],'DD':['1-1-1',1,'1/1/1','3 8 2001'],'LD':['1-1-1',1,'1/1/1','3 8 2001'],  
                                        'TO':['1-1-1',1,'1/1/1','3 8 2001'], 'AS':['1-1-1',1,'1/1/1','3 8 2001'], 'VD':['1-1-1',1,'1/1/1','3 8 2001'],
                                        })                        
                        df['AS'] = df['AS'].astype(str)
                        df['RD'] = df['RD'].astype(str)
                        df['LD'] = df['LD'].astype(str)
                        df['TO'] = df['TO'].astype(str)
                        df['VD'] = df['VD'].astype(str)
                        df['DD'] = df['DD'].astype(str)
          
                        df['AS'] = df['AS'].str.replace('00:00:00', '', regex=True)
                        df['RD'] = df['RD'].str.replace('00:00:00', '', regex=True)
                        df['TO'] = df['TO'].str.replace('00:00:00', '', regex=True)
                        df['VD'] = df['VD'].str.replace('00:00:00', '', regex=True)
                        df['DD'] = df['DD'].str.replace('00:00:00', '', regex=True)
                        df['LD'] = df['LD'].str.replace('00:00:00', '', regex=True)
            
                        df["LD"] = df["LD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["DD"] = df["DD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["RD"] = df["RD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        
                        df["VD"] = df["VD"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["TO"] = df["TO"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        df["AS"] = df["AS"].str.replace(r"\s*\d{1,2}:\d{2}.*$", "", regex=True).str.strip()
                        
                        df["DD"] = df["DD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["RD"] = df["RD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["LD"] = df["LD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["VD"] = df["VD"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["TO"] = df["TO"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df["AS"] = df["AS"].str.replace(r"\s*\..*$", "", regex=True).str.strip()
                        df = pd.concat([df,y])
                        df = df.copy()
                        df['AS'] = df['AS'].astype(str) ###
                        df['RD'] = df['RD'].astype(str) ###
                                                                        
                        df['TO'] = df['TO'].astype(str) ##
                        df['VD'] = df['VD'].astype(str) ###
                        df['DD'] = df['DD'].astype(str) ####
                        df['LD'] = df['LD'].astype(str)

                        # SPLITITNG THE LAST ENCOUNTER DATES
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
                    

                        # SPLITTING ART START DATE
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

                        # SPLITTING DEATH DATE
                        A = df[df['DD'].str.contains('-')].copy()
                        a = df[~df['DD'].str.contains('-')].copy()
                        B = a[a['DD'].str.contains('/')].copy()
                        C = a[~a['DD'].str.contains('/')].copy()
                        E = C[C['DD'].str.contains(' ')].copy()
                        D = C[~C['DD'].str.contains(' ')].copy()
                        A[['Dyear', 'Dmonth', 'Dday']] = A['DD'].str.split('-', expand = True)
                        B[['Dyear', 'Dmonth', 'Dday']] = B['DD'].str.split('/', expand = True)
                        try:
                            D['DD'] = pd.to_numeric(D['DD'], errors='coerce')
                            D['DD'] = pd.to_datetime(D['DD'], origin='1899-12-30', unit='D', errors='coerce')
                            D['DD'] =  D['DD'].astype(str)
                            D[['Dyear', 'Dmonth', 'Dday']] = D['DD'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['DD'] = pd.to_datetime(E['DD'],format='%d %m %Y', errors='coerce')
                            E['DD'] =  E['DD'].astype(str)
                            E[['Dyear', 'Dmonth', 'Dday']] = E['DD'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E])       

                        # SORTING THE RETURN VISIT DATE
                        A = df[df['RD'].str.contains('-')].copy()
                        a = df[~df['RD'].str.contains('-')].copy()
                        B = a[a['RD'].str.contains('/')].copy()
                        C = a[~a['RD'].str.contains('/')].copy()
                        E = C[C['RD'].str.contains(' ')].copy()
                        D = C[~C['RD'].str.contains(' ')].copy()                                   
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

                        #SORTING THE TO DATE
                        A = df[df['TO'].str.contains('-')].copy()
                        a = df[~df['TO'].str.contains('-')].copy()
                        B = a[a['TO'].str.contains('/')].copy()
                        C = a[~a['TO'].str.contains('/')].copy()
                        E = C[C['TO'].str.contains(' ')].copy()
                        D = C[~C['TO'].str.contains(' ')].copy()         
                        A[['Tyear', 'Tmonth', 'Tday']] = A['TO'].str.split('-', expand = True)
                        B[['Tyear', 'Tmonth', 'Tday']] = B['TO'].str.split('/', expand = True)
                        try:
                            D['TO'] = pd.to_numeric(D['TO'], errors='coerce')
                            D['TO'] = pd.to_datetime(D['TO'], origin='1899-12-30', unit='D', errors='coerce')
                            D['TO'] =  D['TO'].astype(str)
                            D[['Tyear', 'Tmonth', 'Tday']] = D['TO'].str.split('-', expand = True)
                        except:
                            pass
                        try:  
                            E['TO'] = pd.to_datetime(E['TO'],format='%d %m %Y', errors='coerce')
                            E['TO'] =  E['TO'].astype(str)
                            E[['Tyear', 'Tmonth', 'Tday']] = E['TO'].str.split('-', expand = True)
                        except:
                            pass
                        df = pd.concat([A,B,D,E])     

                       
                        #BRINGING BACK THE / IN DATES
                        df['AS'] = df['AS'].astype(str)
                        df['RD'] = df['RD'].astype(str)
                        df['TO'] = df['TO'].astype(str)
                        df['VD'] = df['VD'].astype(str)
                        df['DD'] = df['DD'].astype(str)
                        df['LD'] = df['LD'].astype(str)

            #             #Clearing NaT from te dates
                        df['AS'] = df['AS'].str.replace('NaT', '',regex=True)
                        df['RD'] = df['RD'].str.replace('NaT', '',regex=True)
                        df['LD'] = df['LD'].str.replace('NaT', '',regex=True)
                        df['TO'] = df['TO'].str.replace('NaT', '',regex=True)
                        df['VD'] = df['VD'].str.replace('NaT', '',regex=True)
                        df['DD'] = df['DD'].str.replace('NaT', '',regex=True)
          
                                    #SORTING THE VIRAL LOAD YEARS
                        df[['Vyear', 'Vmonth', 'Vday']] =df[['Vyear', 'Vmonth', 'Vday']].apply(pd.to_numeric, errors = 'coerce') 
                        df['Vyear'] = df['Vyear'].fillna(994)
                        a = df[df['Vyear']>31].copy()
                        b = df[df['Vyear']<32].copy()
                        #c = df[]
                        b = b.rename(columns={'Vyear': 'Vday2', 'Vday': 'Vyear'})
                        b = b.rename(columns={'Vday2': 'Vday'})
                        df = pd.concat([a,b])
                        dfa = df.shape[0]
                
                        # #SORTING THE RETURN VISIT DATE YEARS
                        df[['Rday', 'Ryear']] = df[['Rday', 'Ryear']].apply(pd.to_numeric, errors='coerce')
                        df['Ryear'] = df['Ryear'].fillna(994)
                        a = df[df['Ryear']>31].copy()
                        b = df[df['Ryear']<32].copy()
                        b = b.rename(columns={'Ryear': 'Rday2', 'Rday': 'Ryear'})
                        b = b.rename(columns={'Rday2': 'Rday'})
                        df = pd.concat([a,b])

                            #SORTING THE TRANSFER OUT DATE YEAR
                        df[['Tday', 'Tyear']] = df[['Tday', 'Tyear']].apply(pd.to_numeric, errors='coerce')
                        df['Tyear'] = df['Tyear'].fillna(994)
                        a = df[df['Tyear']>31].copy()
                        b = df[df['Tyear']<32].copy()
                        b = b.rename(columns={'Tyear': 'Tday2', 'Tday': 'Tyear'})
                        b = b.rename(columns={'Tday2': 'Tday'})
                        df = pd.concat([a,b]) 

                        #SORTING THE ART START YEARS
                        df[['Ayear', 'Amonth', 'Aday']] =df[['Ayear', 'Amonth', 'Aday']].apply(pd.to_numeric, errors = 'coerce')
                        df['Ayear'] = df['Ayear'].fillna(994)
                        a = df[df['Ayear']>31].copy()
                        b = df[df['Ayear']<32].copy()
                        b = b.rename(columns={'Ayear': 'Aday2', 'Aday': 'Ayear'})
                        b = b.rename(columns={'Aday2': 'Aday'})
                        df = pd.concat([a,b])


                        #SORTING THE DEAD YEARS
                        df[['Dyear', 'Dmonth', 'Dday']] =df[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors = 'coerce')
                        df['Dyear'] = df['Dyear'].fillna(994)
                        a = df[df['Dyear']>31].copy()
                        b = df[df['Dyear']<32].copy()
                        b = b.rename(columns={'Dyear': 'Dday2', 'Dday': 'Dyear'})
                        b = b.rename(columns={'Dday2': 'Dday'})
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
     
                       
                        #CREATE WEEKS 
                        df['Rdaya'] = df['Rday'].astype(str).str.split('.').str[0]
                        df['Rmontha'] = df['Rmonth'].astype(str).str.split('.').str[0]
                        df['Ryeara'] = df['Ryear'].astype(str).str.split('.').str[0]
                        df['RETURN DATE'] = df['Rdaya'] + '/' + df['Rmontha'] + '/' + df['Ryeara']
                        df['RETURN DATE'] = pd.to_datetime(df['RETURN DATE'], format='%d/%m/%Y', errors='coerce')
                        #CREATING WEEEK FOR RETURN VISIT DATE
                        df['RWEEK'] = df['RETURN DATE'].dt.strftime('%V')
                        df['RWEEK'] = pd.to_numeric(df['RWEEK'], errors='coerce')
                        # df['RWEEK1'] = df['RWEEK'] + 13
                        df['RWEEK1'] = df['RWEEK'] - 39

                        #       #PARAMETERS TO SORT OUT FALSE TOs, USING LD AND T
                        df['Taya'] = df['Tday'].astype(str).str.split('.').str[0]
                        df['Tmontha'] = df['Tmonth'].astype(str).str.split('.').str[0]
                        df['Tyeara'] = df['Tyear'].astype(str).str.split('.').str[0]
                        df['TO DATE'] = df['Taya'] + '/' + df['Tmontha'] + '/' + df['Tyeara']
                        df['TO DATE'] = pd.to_datetime(df['TO DATE'], format='%d/%m/%Y', errors='coerce')

                       #LAST ENCOUTER TO DATES
                        df['Ldaya'] = df['Lday'].astype(str).str.split('.').str[0]
                        df['Lmontha'] = df['Lmonth'].astype(str).str.split('.').str[0]
                        df['Lyeara'] = df['Lyear'].astype(str).str.split('.').str[0]
                        df['LAST DATE'] = df['Ldaya'] + '/' + df['Lmontha'] + '/' + df['Lyeara']
                        df['LAST DATE'] = pd.to_datetime(df['LAST DATE'], format='%d/%m/%Y', errors='coerce')
                        df['DURA'] = round((df['LAST DATE']-df['TO DATE']).dt.days)

  
            
###############################################################################################
                        #Q1 parameters
            

                        cyear = 2026  #curr year
                        cyp = 2027 # a year after
                        cyp1 = cyp +1
                        cmonth = 6 #last month of this qtr
                        cmp = 7 # a month after
                        cday  = 3 #starting day
                        cdm = 2 #a day before
                        qmonths = [4, 5, 6] # months of the qtr

                        lmonth = 3 # last qtr month, used for txcur
                        lyear = 2026 # last qtr  yea
                        ldm = 3 # a day before


                        vyeara = 2025 # current vl year
                        vmm = 7   # cut off month for VL, first month of next qtr, (watch out for Q1)
                        vayear = 2026 #for art start date in vl, cutt of six months
                        vamonth = 4 # for art start date in vl cutt off sixmonth <


                        fmonth =  4 # first month of this current qtr


     
        
##################################################################################################################################
                        #POTENTIAL TXCUR ALTER... 
                        df[['Rmonth', 'Rday', 'Ryear']] = df[['Rmonth', 'Rday', 'Ryear']].apply(pd.to_numeric, errors='coerce')
                        df25 = df[df['Ryear']>lyear].copy()
                        df24 = df[df['Ryear'] == lyear].copy()
                        df24[['Rmonth', 'Rday']] = df24[['Rmonth', 'Rday']].apply(pd.to_numeric, errors='coerce')
                        df24 = df24[((df24['Rmonth']>lmonth) | ((df24['Rmonth']==lmonth) & (df24['Rday']>ldm)))].copy()
                        df = pd.concat([df25, df24]).copy()
                        
                        df = df.copy()
                        # DSD = df.copy()
        
                        #REMOVE TO of the last reporting month
                        df[ 'Tyear'] = pd.to_numeric(df['Tyear'], errors='coerce')
                        dfto = df[df['Tyear']!=994].copy() #HAVE TOs
                        dfnot = df[df['Tyear'] == 994].copy() #NO TO
        
                       #REMOVE THE TO
                        dfto[['Ryear', 'Rmonth']] = dfto[['Ryear', 'Rmonth']].apply(pd.to_numeric, errors='coerce')
                        dftoy = dfto[((dfto['Ryear']!=lyear) |((dfto['Ryear']==lyear) & (dfto['Rmonth']>lmonth)))].copy() #OTHERS WOULD BE FALSE TOs, even those made last Q since they were brought as false if their RRDs were this year
                        
                        dftox = dfto[((dfto['Ryear']==lyear) & (dfto['Rmonth']==lmonth))].copy() #CLIENTS WITH RD OF REPORTING MONTH, DIDN'T RETURN BUT WERE TO LATER
                        dftox[['Tyear', 'Tmonth']] = dftox[['Tyear', 'Tmonth']].apply(pd.to_numeric, errors='coerce')
                        dftox = dftox[((dftox['Tyear']==lyear) & (dftox['Tmonth']>lmonth))].copy()

        
                        df = pd.concat([dftoy,dfnot])
                        
                        df = df.copy()
                        if dftox.shape[0]>0:
                            df = pd.concat([df,dftox])
                        else:
                            df =df.copy()
                        #REMOVE the dead of the reporting month
                        df[ 'Dyear'] = pd.to_numeric(df['Dyear'], errors='coerce')
                        dfdd = df[df['Dyear']!=994].copy()
                        dfnot = df[df['Dyear'] == 994].copy()
                        #THOSE WHO DIED BEFORE FIRST MONTH OF THE Q
                        dfdd[['Dyear', 'Dmonth']] = dfdd[['Dyear', 'Dmonth']].apply(pd.to_numeric, errors='coerce')
                        dfdd = dfdd[((dfdd['Dyear']>lyear) |((dfdd['Dyear']==lyear) & (dfdd['Dmonth']>lmonth)))].copy() #DOESN'T MAKE SENSE
                        
                        df = pd.concat([ dfdd,dfnot])
                           
                        df = df.copy()
                       
        ##A             
                        #QUARTERLY TX ML\
                        dfcurr = df.copy()
                        # #DEAD
                        dfcurr['Dyear'] = pd.to_numeric(dfcurr['Dyear'], errors='coerce')
                        deadq = dfcurr[dfcurr['Dyear']!=994].copy()  #THE DEAD
                  
                        dfcurr = dfcurr[dfcurr['Dyear']==994].copy() #REMOVED THE DEAD
    
                        # ####TO
                        dfcurr['Tyear'] = pd.to_numeric(dfcurr['Tyear'], errors='coerce')
                        dfcurra = dfcurr[dfcurr['Tyear']==994].copy()  #NO TO 

                        dfcto = dfcurr[dfcurr['Tyear']!=994].copy() #HAS TOs, false TOs that are txml and active
                        
                        dfcto[['Ryear', 'Rmonth', 'RWEEK']] =  dfcto[['Ryear', 'Rmonth', 'RWEEK']].apply(pd.to_numeric) #TO USE WEEKS FOR NOW    
                        dfctoF = dfcto[ ((dfcto['Ryear']> cyear) | ((dfcto['Ryear'] ==cyear) & (dfcto['Rmonth']>cmonth))) ].copy()
                        #dfctoF = dfcto[ ((dfcto['Ryear']> cyear) | ((dfcto['Ryear'] ==cyear) & (dfcto['RWEEK']>=thisweek))) ].copy()
        
                        dfctoT = dfcto[ ((dfcto['Ryear']< cyear) | ((dfcto['Ryear'] ==cyear) & (dfcto['Rmonth']<cmp))) ].copy() #HAVE TRUE TOS, OLD TOS
                        #dfctoT = dfcto[ ((dfcto['Ryear']< cyear) | ((dfcto['Ryear'] ==cyear) & (dfcto['RWEEK']<thisweek))) ].copy() #HAVE TRUE TOS, OLD TOS
                        dfctoT[['Tyear', 'Tmonth']] = dfctoT[['Tyear', 'Tmonth']].apply(pd.to_numeric, errors='coerce')
                        
                        #OLD TO VS TO OF THE Q        
                        dfctold = dfctoT[((dfctoT['Tyear']<cyear)| ((dfctoT['Tyear'] ==cyear) & (dfctoT['Tmonth']<fmonth)))].copy() #OLD TOs, MADE BEFORE FIRST MONTH OF THE QTR
                        
                        dfctoT = dfctoT[((dfctoT['Tyear'] ==cyear) & (dfctoT['Tmonth'].isin(qmonths)))].copy() #TOs made this Q, ARE THE TOs
                        
                        dfctold[['Rmonth', 'Rday']] = dfctold[['Rmonth', 'Rday']].apply(pd.to_numeric)

                        #OLD TOs returned, that are still active
                        dfctox = dfctold[((dfctold['Rmonth'] ==cmonth) & (dfctold['Rday']>cdm))].copy() #ADD THEM BACK TO CURR
                        #dfctox = dfctold[dfctold['RWEEK']>=thisweek].copy() #ADD THEM BACK TO CURR
        
                        #OLD TOs THAT ARE LOST
                        dfctoyx = dfctold[((dfctold['Rmonth']< cmonth) | ((dfctold['Rmonth']==cmonth) & (dfctold['Rday']<cday)))].copy() #ADD THEM TO LOST
                       
                        dfctoyx['DURA'] = pd.to_numeric(dfctoyx['DURA'], errors = 'coerce')
                        dfctoy = dfctoyx[dfctoyx['DURA']>0].copy() #RETURNED GOT LOST, ADD TO TXML
                        dfctoyz= dfctoyx[dfctoyx['DURA']<1].copy() #NEVER RETURNED, IS A TO
        
                        #stos vs not
                        dfctoT['NTO'] = dfctoT['NTO'].astype(str)
                        words = ['sto', 's\\.t\\.o', 'st\\.o', 'self']
                        pattern = '|'.join(words)
                        dfsto = dfctoT[dfctoT['NTO'].str.contains(pattern, case=False, na=False)]
                        dfsto = dfsto.copy()
                        dfsto['A'] = pd.to_numeric(dfsto['A'], errors= 'coerce')
                        dfctoT['A'] = pd.to_numeric(dfctoT['A'], errors= 'coerce')
                        dfTO =  dfctoT[~dfctoT['A']. isin(dfsto['A'])].copy()
                        dfTO = pd.concat([dfTO, dfctoyz])
                        dfTO = dfTO.copy()
                     
                      
                        dfcur = pd.concat([dfcurra, dfctoF])
                        dfcur = dfcur.copy()

                        #ON APPT
                        dfcur[['Rday','Rmonth', 'Ryear']] = dfcur[['Rday','Rmonth', 'Ryear']].apply(pd.to_numeric, errors = 'coerce')
                         
                        lacks = dfcur[((dfcur['Vyear']< vyeara) | ((dfcur['Vyear'] ==vyeara) & (dfcur['Vmonth']<vmm)))].copy()
                        lacks[['Ayear', 'Amonth']] = lacks[['Ayear', 'Amonth']].apply(pd.to_numeric, errors ='coerce')
                        lacks = lacks[((lacks['Ayear']<vayear) |((lacks['Ayear']==vayear)& (lacks['Amonth'] <vamonth)))].copy()
                        lacks = lacks[lacks['Ayear']!=994].copy()
                        # lacks = lacks[lacks['Tyear']==994].copy()
         
                        
                        dfcur[['Ryear', 'Rmonth']] = dfcur[['Ryear', 'Rmonth']].apply(pd.to_numeric, errors ='coerce')
                        #LOST LAST QTR
                        currlosta = dfcur[((dfcur['Ryear'] == lyear) & (dfcur['Rmonth']==lmonth))].copy()
            
                        #LOST THIS QTR
                        curlostb = dfcur[((dfcur['Ryear'] == cyear) & (dfcur['Rmonth'].isin(qmonths)))].copy() #LOST THIS QTR
                        curlostb[['Ryear', 'Rmonth', 'Rday', 'RWEEK']] = curlostb[['Ryear', 'Rmonth', 'Rday','RWEEK']].apply(pd.to_numeric, errors ='coerce')
                        curlostc = curlostb[ ((curlostb['Rmonth']<cmonth) |(( curlostb['Rmonth']== cmonth) & (curlostb['Rday']<cday)))].copy()
                        #curlostc = curlostb[curlostb['RWEEK']<thisweek].copy()
        
                        #currlost = pd.concat([curlostc, dfctoy]) #x will be the TOs that returned and got lost, in second q include curlosta
                        framers = [curlostc, dfctoy, currlosta]
                        framers = [f for f in framers if not f.empty]
                        if framers:  # at least one is non-empty
                            currlost = pd.concat(framers, ignore_index=True)
                        else:       # all three are empty
                            currlost = pd.DataFrame()
                            
                        currlost = currlost.copy()
                  
                        
                        cur26 = dfcur[dfcur['Ryear'] >cyear].copy() #ACTIVE NEXT OTHER YEARS
                        cur25 = dfcur[dfcur['Ryear'] == cyear].copy() # ACTIVE THIS YEAR
                        cur25[['Ryear', 'Rmonth', 'Rday']] = cur25[['Ryear', 'Rmonth', 'Rday']].apply(pd.to_numeric, errors ='coerce')
                        cur25 = cur25[ ((cur25['Rmonth']>cmonth) |(( cur25['Rmonth']==cmonth) & (cur25['Rday']>cdm)))].copy()
       
                        frames = [cur25, cur26, dfctox]
                        frames = [f for f in frames if not f.empty]
                        if frames:  # at least one is non-empty
                            dfcur = pd.concat(frames, ignore_index=True)
                        else:       # all three are empty
                            dfcur = pd.DataFrame()
                        if dfcur.shape[0] ==0:
                            st.warning('NO ACTIVE CLIENTS IN THIS EXTRACT, TRY MANUAL FILTERING WITH YOUR RD COLUMN TO VERIFY FOR YOUR SELF')
                            st.stop()
                        else:
                            pass
                        dfcur[['Ayear', 'Amonth']] = dfcur[['Ayear', 'Amonth']].apply(pd.to_numeric, errors='coerce')
                        dfcur = dfcur[((dfcur['Ayear'] <cyear) | ((dfcur['Ayear'] == cyear) & dfcur['Amonth']<cmp))].copy() #REMOVES TX NEW DATA COLL MONTH
                        dfcur['Ryear'] =  pd.to_numeric(dfcur['Ryear'], errors='coerce')
                        dfcur = dfcur[dfcur['Ryear'] <cyp1].copy() #REMOVES EXTREME YEARS
        
                        dfcur = dfcur.copy()
                        a1 = dfcur.shape[0]                   
                                     

                        st.write(f'**TX CUR IS {a1}**')
     
                        name = st.text_input('FACILITY NAME')
                        if not name:
                            st.warning('PLEASE ENTER FACILITY NAME')
                            st.stop()

                        out = r"C:\Users\Desire Luminsa\Desktop\ECHO\CURR"
                        path = os.path.join(out, name + '.csv')
                        dfcur = dfcur[['ART','RD','AG','WT','ARVD','VR','AS','VD','LD','ARVS','BCD4','DSD']]
                        dfcur.to_csv(path, index=False)
                        st.write('**FILE SAVED**')   
                        
pages = {
    "READER:": [
        st.Page(extract, title="EMR EXTRACT READER"),
    ],
   
}

pg = st.navigation(pages)
pg.run()
                                
    

