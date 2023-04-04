import pandas as pd  
import streamlit as st
import numpy as np
import datetime
from functools import reduce
import time as tm
import openpyxl
   
st.set_page_config(page_title="Pif Previ V2", page_icon="🚨", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

st.title('New_hyp')
st.subheader("Programme complet :")
uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache(suppress_st_warning=True,allow_output_mutation=True)
    def df():
        with st.spinner('Chargemement Programme complet ...'):
            df = pd.read_excel(uploaded_file, "pgrm_complet")
            sat5 = ['FI', 'LO', 'A3', 'SK', 'DY', 'D8', 'S4']
            sat6 = ['LH', 'LX', 'OS', 'EW', 'GQ', 'SN']
            df.loc[df['Cie Ope'].isin(sat6), 'Libellé terminal'] = 'Terminal 1_6'
            df.loc[df['Cie Ope'].isin(sat5), 'Libellé terminal'] = 'Terminal 1_5'
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_Inter","Terminal 1")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_5","Terminal 1_5")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_6","Terminal 1_6")
        st.success("Programme complet chargée !")
        return df

    df_pgrm = df()         
    start_all = tm.time()
    l_date = pd.to_datetime(df_pgrm['Local Date'].unique().tolist()).date
    l_date = sorted(l_date)

    uploaded_file_config = st.file_uploader("Choisir un fichier de congig :", key=15)
    if uploaded_file_config is not None:
        @st.cache(suppress_st_warning=True,allow_output_mutation=True)
        def get_pif_in_fichier_config(pif):
            return pd.read_excel(uploaded_file_config, sheet_name=pif)
    
    # a définir en fonction du fichier de congig

    L_pif = ['K CNT', 'K CTR', 
                    'L CNT', 'L CTR', 
                    'M CTR', 
                    'Galerie EF', 'C2F', 
                    'C2G', 
                    'Liaison AC', 
                    'Liaison BD', 
                    'T3',
                    'Terminal 1',
                    'Terminal 1_5',
                    'Terminal 1_6']
    table_faisceau_iata = pd.read_excel("table_faisceau_IATA.xlsx")
    table_faisceau_iata = table_faisceau_iata[['Prov Dest','Faisceau géographique']]
    df_pgrm = df_pgrm.merge(table_faisceau_iata,how='left', left_on='Prov Dest', right_on='Prov Dest')

    uploaded_file1 = st.file_uploader("Choisir le fichier hypotheses_repartition_correspondances.xlsx :", key=4)

    @st.cache(suppress_st_warning=True,allow_output_mutation=True)
    def COURBE_PRES(t):
        df = pd.read_excel('courbes_presentation.xlsx', t)   
    #     L=[]
    #     if t == 'T2E':
    #         for i in ['Amérique Centre + Sud', 'Amérique Nord', 'Autre Afrique',
    #    'Autre Europe', 'DOM TOM', 'Extrême Orient', 'Moyen Orient',
    #    'Métropole', 'Schengen', 'U.E. hors M & S']:
    #             i = df[df['faisceau_geographique']==i].copy()
    #             st.write(i)
    #             i['pourc'] = i['pourc'].shift(2)
    #             i['pourc'] = i['pourc'].fillna(0)
    #             L += [i]
    #         dff = pd.concat(L)               
    #         return dff
    #     else:
    #         return df 
        return df       
    col1, col2 = st.columns(2)
    with col1:
        debut = st.date_input("Date de début :", key=10)
    with col2:    
        fin = st.date_input("Date de fin :", key=2)
    
    start_date = pd.to_datetime(debut)
    end_date = pd.to_datetime(fin) 

    if st.button('Créer Export PIF'):
    


        #Fonction qui regroupe les sous fonctions de traitement


        
        st.warning('La requête a bien été prise en compte, début du traitement.\nNe tentez pas de fermer la fenêtre même si celle-ci semble figée')
        placeholder = st.empty()
        my_bar2 = placeholder.progress(5)
    ### path files ###
        path_hyp = r"" + "hypotheses_repartition_correspondances.xlsx"
        name_hyp = "Feuil1"
        
        path_faisceaux = r"" + "faisceaux_escales.xlsx"
        name_faisceaux = "escales"
        
    #        ancienne courbes de prés globale, sans distinction de terminal
    #        path_courbes = r"" + source_outils_previ.chemin_fichier_source(4)
    #        name_courbes = "nouvellesCourbesPresentation"
        
        path_courbes_term = r"" + "nouvelles_courbes_presentation_PIF.xlsx"
        list_terminaux = ['Terminal 2A', 'Terminal 2B', 'Terminal 2C', 'Terminal 2D',
                          'EK', 'EL', 'EM', 'F', 'G', 'Terminal 3','Terminal 1',
                          'Terminal 1_5','Terminal 1_6']
        
        path_output = r"" + "output_export_pif"
        name_output = "export_pif"
        


        
        def FAISCEAUX_IATA():
            df = pd.read_excel(path_faisceaux, name_faisceaux)
            del df['faisceau_facturation']
            del df['faisceau_commercial']
            del df['cl_long']
            del df['pays']
            del df['ville']
            del df['aeroport']
            del df['escale_OACI']
            del df['jour_ref']
            del df['statut']
            return df 
        
        df_faisceaux = FAISCEAUX_IATA()
        
        
    #        Pour la courbe de pres unique, inutile
    #        def COURBE_PRESENTATION():
    #            return pd.read_excel(path_courbes, name_courbes)
        
    #        df_courbe_presentation = COURBE_PRESENTATION()
        
        
    #        Entre pgrm ADP et pgrm AF les heures ne sont pas au même format. On les transforme ici. A terme migrer cette fonction dans Concat
        def STR_TO_DT(df):
            df_temp = df
            l_dt = []
            for t in range(df.shape[0]):
                TSTR =  str(df['Horaire théorique'][t])
                if len(TSTR)<10:
                    l = [int(i) for i in TSTR.split(':')]
                    l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
                else:
                    TSTR = TSTR[10:]
                    l = [int(i) for i in TSTR.split(':')]
                    l_dt.append(datetime.time(hour=l[0], minute=l[1], second=0))
            
            df['Horaire théorique'] = l_dt
                
            return df_temp
        
        df_pgrm_dt = STR_TO_DT(df_pgrm)
        df_pgrm_dt = df_pgrm_dt.loc[(df_pgrm_dt['Local Date'] >= start_date) & (df_pgrm_dt['Local Date'] <= end_date)]
        df_pgrm_dt.reset_index(inplace=True, drop=True)
        df_pgrm_dt['Unnamed: 0'] = df_pgrm_dt.index
        
    
        
        faisceaux = ['Métropole', 'Schengen', 'U.E. hors M & S', 'Afrique du Nord',
       'Amérique du Nord', 'Autre Afrique', 'Autre Europe', 'DOM TOM',
       'Extrême Orient', 'Moyen Orient', 'Amérique Centre + Sud']
        
            ### DISPATCH ###       


        import numpy
        from datetime import datetime, timedelta
        def HYP_REP(sheet):
            df = pd.read_excel(uploaded_file1, sheet)
            df['heure'] = pd.to_datetime(df['heure'].str[:8],format='%H.%M.%S')
            # df['heure_heure'] = df['heure'].str[:9]
            # df['heure_heure'] = pd.to_datetime(df['heure_heure'], format='%H', exact=False)
            # df['heure_heure'] = df.heure_heure.apply(lambda x: x.hour)
            return df

        df_pgrm_dt['Horaire théorique'] = pd.to_datetime(df_pgrm_dt['Horaire théorique'],format='%H:%M:%S')
        # df_pgrm_dt['Horaire théorique'] = df_pgrm_dt["Horaire théorique"].apply(lambda x: x.hour)
        df_pgrm_dt = df_pgrm_dt.drop_duplicates(subset=df_pgrm_dt.columns.difference(['Unnamed: 0']))

        def DISPATCH_NEW(df):
            """Permet la création d'un DF dispatch qui facilite le tri par batterie de PIF"""
            col = ['Local Date', 'Horaire théorique', 'Prov Dest', 'A/D', 'Libellé terminal', 'Faisceau géographique',
                    'K CNT', 'K CTR', 
                    'L CNT', 'L CTR', 
                    'M CTR', 
                    'Galerie EF', 'C2F', 
                    'C2G', 
                    'Liaison AC', 
                    'Liaison BD', 
                    'T3',
                    'Terminal 1',
                    'Terminal 1_5',
                    'Terminal 1_6']

            dispatch_df = pd.DataFrame(columns = col, index = df['Unnamed: 0'])

            dispatch_df['Local Date'] = df['Local Date']
            dispatch_df['Horaire théorique'] = df['Horaire théorique']
            dispatch_df['Prov Dest'] = df['Prov Dest']
            dispatch_df['A/D'] = df['A/D']
            dispatch_df['Libellé terminal'] = df['Libellé terminal']
            dispatch_df['Faisceau géographique'] = df['Faisceau géographique']
            
            def dispatch_term(terminal, salle_apport, salle_emport, AD):
                hyp_rep = HYP_REP(salle_apport + "_" + salle_emport)
                
                L_df = []
                for i,n in zip(hyp_rep['heure'].values, numpy.roll(hyp_rep['heure'].values,-1)):
                    for j in faisceaux:
                        x = hyp_rep.loc[(hyp_rep['heure'] == i)][j].tolist()[0] 
                        if AD == 'D':
                            x = 1 
                        if x != 0:
                            temp = df.loc[(df['A/D'] == AD) & (df['Libellé terminal'] == terminal)].copy()
                            temp = temp.loc[(df['Faisceau géographique'] == j)]
                            temp = temp.loc[(temp['Horaire théorique'] > i) & (temp['Horaire théorique'] < n) ]['Pax CNT TOT']*x
                    
                        
                            L_df += [temp]
                return reduce(lambda a, b: a.add(b, fill_value = 0),L_df)


            def dispatch_term_D(terminal, type_pax = 'PAX TOT'):                
                temp = df.loc[(df['A/D'] == 'D') & (df['Libellé terminal'] == terminal)].copy()
                return temp[type_pax]

            
            
            for pif in L_pif:
                dispatch = []
                df_config = get_pif_in_fichier_config(pif)
                
                for index, row in df_config.iterrows():
                    if row['Arr/Dep'] == 'D':
                        dispatch += [dispatch_term_D(row['terminal'], row['type_pax'])]
                    else:
                        dispatch += [dispatch_term(row['terminal'], row['salle_apport'], row['salle_emport'], row['Arr/Dep'])]

                       
                dispatch_df[pif] = reduce(lambda a, b: a.add(b, fill_value = 0),
                                          dispatch)  

            dispatch_df.fillna(0, inplace=True)

            return dispatch_df






        dispatch = DISPATCH_NEW(df_pgrm_dt)

        dispatch.to_excel("dispatch.xlsx", sheet_name="dispatch")
        

        liste_df_courbe_presentation_terminal = {}
        
        for t in list_terminaux:
            liste_df_courbe_presentation_terminal[t] = COURBE_PRES(t)
        
        def courbe(df_c):
            l_f = df_c['faisceau_geographique'].unique().tolist()
            
            courbe = {}
            for i in l_f:    
                courbe[i] = ( df_c['pourc'].loc[(df_c['faisceau_geographique'] == i)
                                                & (df_c['heure_debut'] == df_c['heure_debut'][0])].tolist())
            return courbe

        l_courbe_geo_t = {}

        for t in list_terminaux:    
            l_courbe_geo_t[t] = courbe(liste_df_courbe_presentation_terminal[t])
        
        with st.info:
            st.write(l_courbe_geo_t)


        global pb_index
        pb_index = 0


        dispatch_paf = dispatch.copy()

        dispatch_paf_D = dispatch_paf.copy()
        dispatch_paf_D = dispatch_paf_D[dispatch_paf_D["A/D"] == "D"]
        dispatch_paf_A = dispatch_paf.copy()
        dispatch_paf_A = dispatch_paf_A[dispatch_paf_A["A/D"] == "A"]


        n_D = 24
        n_A = 5 #len(L_A)

        # Create a list to store the duplicated rows
        rows = []
        L_A = [0, 0, 0, 0.5, 0.5]
        L_pif = ['K CNT', 'K CTR', 
                    'L CNT', 'L CTR', 
                    'M CTR', 
                    'Galerie EF', 'C2F', 
                    'C2G', 
                    'Liaison AC', 
                    'Liaison BD', 
                    'T3',
                    'Terminal 1',
                    'Terminal 1_5',
                    'Terminal 1_6']


        # DEPART
        # Loop through each row in the dataframe
        for index, row in dispatch_paf_D.iterrows():
            # Loop n times to duplicate the row and subtract 10 minutes from the datetime column each time
            for i in range(n_D):
                # Create a copy of the original row
                new_row = row.copy()
                if new_row['Faisceau géographique'] == 0:
                    x = "Extrême Orient"
                else:
                    x = new_row['Faisceau géographique']       
                L = l_courbe_geo_t[new_row['Libellé terminal']][x]
                # Subtract 10 minutes from the datetime column
                new_row['Horaire théorique'] -= timedelta(minutes=10*i)
                for pif in L_pif:
                    new_row[pif] = L[i]*new_row[pif]
                
                # Append the modified row to the list
                rows.append(new_row)
                
                
        # Create a new dataframe from the list of duplicated rows
        new_df = pd.DataFrame(rows)


        # ARRIVER

        for index, row1 in dispatch_paf_A.iterrows():
            # Loop n times to duplicate the row and subtract 10 minutes from the datetime column each time
            for i in range(n_A):
                # Create a copy of the original row
                new_row = row1.copy()
                # Subtract 10 minutes from the datetime column
                new_row['Horaire théorique'] += timedelta(minutes=10*i)
                for pif in L_pif:
                    new_row[pif] = L_A[i]*new_row[pif]
                
                # Append the modified row to the list
                rows.append(new_row)
                
                
        # Create a new dataframe from the list of duplicated rows
        new_df_A = pd.DataFrame(rows)

        
        df_final = pd.melt(new_df_A, id_vars=['Local Date', 'Horaire théorique'], value_vars=L_pif)
        
        def ceil_dt(x):
            return x + (datetime.min - x) % timedelta(minutes=10)

        df_final['Horaire théorique'] = df_final['Horaire théorique'].apply(lambda x: ceil_dt(x))

        
        df_final['Horaire théorique'] = df_final['Horaire théorique'].dt.time
        df_final = df_final.groupby(['Local Date', 'Horaire théorique', 'variable']).sum().reset_index()


        
        df_final.rename(columns={"Local Date":"jour",
                         'Horaire théorique':'heure',
                         'variable':'site',
                         'value':'charge'}, inplace=True)
        

        import time
        def CLEAN_TIME(m):
            t = '0:00'.join(str(m).rsplit('5:00', 1))
            #l = [int(k) for k in t.split(':')]
            time_r = time(hour = int(t[11:13]), minute = int(t[14:16]), second = int(t[17:19]))

            return time_r
        


        directory_exp = "export_pif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"
        import io
        from pyxlsb import open_workbook as open_xlsb

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name=name_output)
            writer.save()

            st.download_button(
            label="Télécharger fichier Export pif",
            data=buffer,
            file_name=directory_exp,
            mime="application/vnd.ms-excel"
            )
                        

        st.info("Export PIF créé avec succès !" + "\n\nPour lancer une nouvelle étude, lancer uniquement 'CHOISIR LES DATES'")
        


