from datetime import datetime
import pandas as pd 
import numpy as np
from datetime import date, timedelta
from xlwt.Workbook import *
import streamlit as st


with open('style.css')as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html = True)

c1, c2 = st.columns(2)

uploaded_file = c1.file_uploader("Ancien code :", key=1)
if uploaded_file is not None:
    @st.cache(suppress_st_warning=True,allow_output_mutation=True)
    def oldd():
        return pd.read_csv(uploaded_file, sep=";")
    old = oldd()
    uploaded_file1 = c2.file_uploader("Nouveau code :", key=2)
    if uploaded_file1 is not None:
        @st.cache(suppress_st_warning=True,allow_output_mutation=True)
        def neww():
            return pd.read_excel(uploaded_file1)
        new = neww()
        c3, c4 = st.columns(2)
        uploaded_file2 = c3.file_uploader("Nouveau code affiné :", key=3)    
        if uploaded_file2 is not None:
            @st.cache(suppress_st_warning=True,allow_output_mutation=True)
            def new_courbee():
                return pd.read_excel(uploaded_file2)
            new_courbe = new_courbee()
            uploaded_file3 = c4.file_uploader("Réalisé :", key=4)
            if uploaded_file3 is not None:
                @st.cache(suppress_st_warning=True,allow_output_mutation=True)
                def reall():
                    return pd.read_excel(uploaded_file3, skiprows=3)
                real = reall()
# @st.cache(suppress_st_warning=True,allow_output_mutation=True)
# def openn():
#     columns = ['jour','heure','charge','site']
#     new  = pd.read_excel("export_pif_du_2023-03-16_au_2023-03-26 (V2).xlsx")
#     new_courbe = pd.read_excel("export_pif_du_2023-03-16_au_2023-03-26 (1).xlsx")
#     old  = pd.read_csv("export_compilation_pif (mars 2023).csv", sep=";") 
#     real = pd.read_excel("réalisé_PIF_[Wassim] (9).xlsx", skiprows=3)
#     x=788800000
#     return new, new_courbe, old, columns, real

                columns = ['jour','heure','charge','site']

                col1, col2, col3 = st.columns(3)
                with col1:
                    debut = st.date_input("Date de début :", key=5)
                with col2:    
                    fin = st.date_input("Date de fin :", key=6)
                with col3:
                    if st.button("Valider dates"):
                        next = True

                

                start_date = pd.to_datetime(debut)
                end_date = pd.to_datetime(fin) 
                #start_date = pd.to_datetime('2023-04-06')
                #end_date = pd.to_datetime('2023-04-16')

                wb= Workbook()
                writer = pd.ExcelWriter('Comp.xlsx', engine='xlsxwriter')

                def CLEAN_TIME(m):
                    m = m[0:8]
                    time_r = str(m[0:2]) + ":" + str(m[3:5]) + ":" + str(m[6:8])
                    return time_r


                @st.cache(suppress_st_warning=True,allow_output_mutation=True)
                def old1(old):
                    start_date = pd.to_datetime(debut)
                    end_date = pd.to_datetime(fin) 
                    old.replace(regex=[r'[\r,;]+', 'foo'], value='.',inplace=True)
                    old["charge"] = old.charge.astype(float)
                    #old["jour"] = old.jour.astype(str)
                    old['jour'] = pd.to_datetime(old['jour'])
                    old.replace({"Jour":'jour'}, inplace =True)
                    mask2 = (old['jour'] >= start_date) & (old['jour'] <= end_date)
                    old = old.loc[mask2]
                    old.replace({'C2F centraux':'C2F',
                            'C2G centraux':'C2G',
                            'salle M centraux':'M CTR',
                            'salle K centraux':'K CTR',
                            'salle L centraux':'L CTR',
                            'salle L corres':'L CNT',
                            'galerie E vers F':'Galerie EF',
                            'PIF_liaison_AC':'Liaison AC',
                            'PIF_liaison_BD':'Liaison BD',
                            'PIF_T3':'T3',
                            'PIF_T1_Inter':'Terminal 1',
                            'PIF_T1_Schengen':'Terminal 1_5',
                            'PIF_T1_sat5':'Terminal 1_5',
                            'PIF_T1_sat6':'Terminal 1_6'


                            },inplace=True)
                    old = old[columns]
                    #old.to_excel(writer, sheet_name='Ancien_code', index=False)
                    return old

                @st.cache(suppress_st_warning=True,allow_output_mutation=True)
                def new1(new):
                    start_date = pd.to_datetime(debut)
                    end_date = pd.to_datetime(fin)
                    new['jour'] = pd.to_datetime(new['jour'])
                    mask = (new['jour'] >= start_date) & (new['jour'] <= end_date)
                    new = new.loc[mask]
                    new = new[columns]
                    #new.to_excel(writer, sheet_name='Nouveau_code', index=False)
                    return new


                @st.cache(suppress_st_warning=True,allow_output_mutation=True)
                def new_courbe1(new_courbe):
                    start_date = pd.to_datetime(debut)
                    end_date = pd.to_datetime(fin)
                    new_courbe['jour'] = pd.to_datetime(new_courbe['jour'])
                    maskn = (new_courbe['jour'] >= start_date) & (new_courbe['jour'] <= end_date)
                    new_courbe = new_courbe.loc[maskn]
                    new_courbe = new_courbe[columns]
                    #new_courbe.to_excel(writer, sheet_name='new_courbe', index=False)
                    return new_courbe

                @st.cache(suppress_st_warning=True,allow_output_mutation=True)
                def real1(real): 
                    start_date = pd.to_datetime(debut)
                    end_date = pd.to_datetime(fin)
                    real['Jour'] = pd.to_datetime(real['Jour'])
                    mask3 = (real['Jour'] >= start_date) & (real['Jour'] <= end_date)
                    real = real.loc[mask3]
                    #real = real[real['Groupe Position']== 'LBD']
                    real.rename(columns={'Tranche 10 minutes passage': 'heure',
                                "Groupe Position": "site",
                                "Jour":"jour",
                                "Nb de passages":"charge"},
                                inplace = True)
                    real["charge"] = real.charge.astype(float)            
                    real = real[columns]
                    #real.to_excel(writer, sheet_name='real', index=False)
                    real['heure'] = real['heure'].apply(lambda x: CLEAN_TIME(x))
                    real.replace({'C2F centraux':'C2F',
                            'C2G centraux':'C2G',
                            'M centraux':'M CTR',
                            'K centraux':'K CTR',
                            'L centraux':'L CTR',
                            'G centraux':'C2G',
                            'salle K centraux':'K CTR',
                            'salle L centraux':'L CTR',
                            'L corresp':'L CNT',
                            'galerie EF':'Galerie EF',
                            'C2AC - PIFS - AUTO':'Liaison AC',
                            'LBD':'Liaison BD',
                            'PIFs T3':'T3',
                            'C2E3TPIF-09':'L CNT',
                            'C2EJTPIF-003':'L CNT',
                            'C2E3CPIF-L02':'L CTR',
                            'C2E3CPIF-L13':'L CTR',
                            'C2E3CPIF-L16':'L CTR',
                            'CT-PIF-A':'T3',
                            'F centraux':'C2F',
                            'Puits K ?':'K CNT',
                            'C2E4CPIF-M04':'M CTR',
                            'C2E4CPIF-M12':'M CTR',
                            'C2AC - PIFS - AUTO':'Liaison AC',
                            'C2AC-PIFS-03':'Liaison AC',
                            'C2AC-PIFS-04':'Liaison AC',
                            'LAC':'Liaison AC'
                            },inplace=True)

                    real = real.groupby(by=["jour", "heure", "site"]).sum().reset_index()
                    return real


                #new, new_courbe, old, columns, real = openn()

                old = old1(old)
                new = new1(new)
                new_courbe = new_courbe1(new_courbe)
                real = real1(real)

                import datetime
                import time 

                
                df_final = new.merge(old,on=('site','heure','jour'),  how='left')
                df_final = df_final.rename(columns={"charge_x":"Nouveau_code",
                                                    "charge_y":"Ancien_code"})
                df_final = df_final.merge(real,on=('site','heure','jour'),  how='left')
                df_final = df_final.rename(columns={"charge":"real"})
                df_final = df_final.merge(new_courbe,on=('site','heure','jour'),  how='left')
                df_final = df_final.rename(columns={"charge":"Nouveau_code_affiné"})
                df_final.drop_duplicates(inplace=True)


                writer.save()
                df_final.to_excel(writer, sheet_name='Comparaison', index=False)
                df_final.to_excel("comparaison.xlsx", sheet_name='Comparaison', index=False)
                
                
                def df_final1():
                    return df_final.copy()
                
                df_final = df_final1()
                df_comp = df_final1()
                st.title('Dashboard Comparaison')

                st.header("Cumul PIFs confondues")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric(label = "Ancien code",value = str(round(sum(old.charge))))
                col2.metric(label = "Nouveau code",value = str(round(sum(new.charge))))
                col3.metric(label = "Nouveau code affiné",value = str(round(sum(new_courbe.charge))))
                col4.metric(label = "Réalisé",value = str(round(sum(real.charge))))


                df_final = df_final.astype({'Nouveau_code_affiné':'float64'})
                df_final['Nouveau_code_affiné'].astype('float')

                st.markdown('--------------------')

                
                date = st.selectbox('Choisir jour(s)',
                                    options=df_final['jour'].unique())

                mask_jour = df_final['jour'] == date

                df = df_final[mask_jour]


                Charge_tot_new = df['Nouveau_code'].sum()
                Charge_tot_old = df['Ancien_code'].sum()
                Charge_tot_real = df['real'].sum()

                site = st.multiselect('Choisir site(s)',
                                    options=df['site'].unique(), default= 'C2F')

                mask_site = df['site'].isin(site)
                mask_site2 = df_final['site'].isin(site)

                df = df[mask_site]
                df_semaine = df_final[mask_site2]
                #df_semaine['Nouveau_code_affiné'] = df_semaine['Nouveau_code_affiné'].rolling(window=3).mean()                
                df["heure_pleine"] = df["heure"].str[:2]
                df_semaine["heure_pleine"] = df_semaine["heure"].str[:2]
                df_heure = df.copy().groupby(['heure_pleine']).sum()
                df_semaine_heure = df_semaine.copy().groupby(['heure_pleine']).sum()
                df = df.groupby(['heure']).sum()



                tab1, tab2 = st.tabs(["Jour", "10 jours"])

                with tab1:
                    st.subheader("Tranche de 10 min :")
                    #st.subheader("Vue du " + str(date.astype('datetime64[D]')) + ":")
                    st.line_chart(df[['Nouveau_code', 'Ancien_code', 'real', 'Nouveau_code_affiné']])

                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        Charge_tot_old = round(df['Ancien_code'].sum())
                        col1.metric(label = "Ancien code", value = str(Charge_tot_old))
                    with col2:    
                        Charge_tot_new = round(df['Nouveau_code'].sum())
                        col2.metric(label = "Nouveau code", value = str(Charge_tot_new), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_new))))
                    with col3:    
                        Charge_tot_real = round(df['real'].sum())
                        Charge_tot_new_courbe = round(df['Nouveau_code_affiné'].sum())
                        col3.metric(label = "Nouveau code affiné", value = str(Charge_tot_new_courbe), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_new_courbe))))
                    with col4:    
                        col4.metric(label = "Réalisé", value = str(Charge_tot_real), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_real))))



                with tab2:

                    st.subheader("Tranche de 10 min :")
                    df_semaine = df_semaine.groupby(['heure']).sum()
                    st.line_chart(df_semaine[['Nouveau_code', 'Ancien_code', 'real', 'Nouveau_code_affiné']])

                    st.subheader("KPI")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        Charge_tot_old = round(df_semaine['Ancien_code'].sum())
                        col1.metric(label = "Ancien code", value = str(Charge_tot_old))
                    with col2:    
                        Charge_tot_new = round(df_semaine['Nouveau_code'].sum())
                        col2.metric(label = "Nouveau code", value = str(Charge_tot_new), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_new))))
                    with col3:    
                        Charge_tot_real = round(df_semaine['real'].sum())
                        Charge_tot_new_courbe = round(df_semaine['Nouveau_code_affiné'].sum())
                        col3.metric(label = "Nouveau code affiné", value = str(Charge_tot_new_courbe), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_new_courbe))))
                    with col4:    
                        col4.metric(label = "Réalisé", value = str(Charge_tot_real), delta=str(np.round(0 - (Charge_tot_old - Charge_tot_real))))
                

                st.markdown('--------------------')

                with tab1:

                    st.subheader("Tranche de 1 heure :")
                    st.line_chart(df_heure[['Nouveau_code', 'Ancien_code', 'real', 'Nouveau_code_affiné']])


                with tab2:
                    st.subheader("Tranche de 1 heure: ")
                    st.line_chart(df_semaine_heure[['Nouveau_code', 'Ancien_code', 'real', 'Nouveau_code_affiné']])


                directory_exp = "export_comparaison_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"
                import io
                from pyxlsb import open_workbook as open_xlsb

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_comp.to_excel(writer, sheet_name='Comparaison', index=False)
                    writer.save()

                    st.download_button(
                    label="Télécharger fichier comparaison",
                    data=buffer,
                    file_name=directory_exp,
                    mime="application/vnd.ms-excel"
                    )    
