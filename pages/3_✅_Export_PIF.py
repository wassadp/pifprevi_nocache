import pandas as pd
import streamlit as st
import os
import time
import pandas as pd
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter
import datetime
import calendar
import locale
from openpyxl.styles import Font
import itertools
from datetime import datetime


st.title("Macro")
st.write("Test Macro du fichier Export_pif avant l'ajout à l'outil OUTILSPIF")

def findDay(date):
    born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
    return (calendar.day_name[born])   
 
uploaded_file = st.file_uploader("Choose a file")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df1 = pd.DataFrame(columns=df.columns)
    wb= Workbook()
    writer = pd.ExcelWriter('multiple3.xlsx', engine='xlsxwriter')
    st.write(locale.locale_alias)
    def clean(df,i):
        df['Numéro de Jour'] = df['jour'].dt.day
        df['Date complète'] = df['jour'].dt.strftime('%d/%m/%Y')
        df['Jour de la semaine'] = df['jour'].dt.day_name(locale="fr_FR.ISO8859-1")
        #df['SOMME PAX LOCAUX DE LA JOURNEE'] = df.iloc[:,4:].sum()
        df['SOMME PAX LOCAUX DE LA JOURNEE'] = df.iloc[:, 4:].sum(axis=1)    
        g = str(i).replace(" ", "_")
        df[str(i).replace(" ", "_")] = df['jour'].dt.month_name(locale="fr_FR.ISO8859-1")
        first_column = df.pop('Numéro de Jour')
        df.insert(1, 'Numéro de Jour', first_column)
        first_column = df.pop('Date complète')
        df.insert(2, 'Date complète', first_column)
        first_column = df.pop('Jour de la semaine')
        df.insert(1, 'Jour de la semaine', first_column)
        first_column = df.pop(str(i).replace(" ", "_"))
        df.insert(0, str(i).replace(" ", "_"), first_column)
        df.pop('jour')
        df[str(i).replace(" ", "_")] = list(itertools.chain.from_iterable([key] + [float('nan')]*(len(list(val))-1) 
                            for key, val in itertools.groupby(df[str(i).replace(" ", "_")].tolist())))

    
    def findDay(date):
        born = datetime.datetime.strptime(date, '%d %m %Y').weekday()
        return (calendar.day_name[born])    

    st.write(df)
    site = []
    for i in df.site.unique():
        name = str(i).replace(" ", "_")
        site += [name]
        name = df.copy()
        name = name[name['site'] == i]
        name = df.pivot_table(values='charge', index='jour', columns=['heure'], aggfunc='first')
        name.reset_index(inplace=True)
        clean(name,i)
        name.to_excel(writer, sheet_name=str(i).replace(" ", "_"), index=False)

    writer.save()  

    df = pd.read_excel("multiple3.xlsx")
    st.write(df)

    import io
    from pyxlsb import open_workbook as open_xlsb

    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        df.to_excel(writer)
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.save()

        st.download_button(
        label="Télécharger fichier Export pif",
        data=buffer,
        file_name="test-export.xlsx",
        mime="application/vnd.ms-excel"
        )