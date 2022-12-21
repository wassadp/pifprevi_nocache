import pandas as pd  
import streamlit as st
import numpy as np
import datetime
from functools import reduce
import time as tm
import openpyxl
   
st.set_page_config(page_title="Pif Previ", page_icon="🚨", layout="centered", initial_sidebar_state="auto", menu_items=None)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 


st.subheader("Programme complet :")
uploaded_file = st.file_uploader("Choisir un fichier :", key=1)
if uploaded_file is not None:
    @st.cache(suppress_st_warning=True,allow_output_mutation=True)
    def df():
        with st.spinner('Chargemement Programme complet ...'):
            df = pd.read_excel(uploaded_file, "pgrm_complet")
            #sat5 = ['FI', 'LO', 'A3', 'SK', 'DY', 'D8']
            #sat6 = ['LH', 'LX', 'OS', 'EW', 'SN']
            #df.loc[df['Cie Ope'].isin(sat6), 'Libellé terminal'] = 'Terminal 1_6'
            #df.loc[df['Cie Ope'].isin(sat5), 'Libellé terminal'] = 'Terminal 1_5'
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_Inter","Terminal 1")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_5","Terminal 1_5")
            df['Libellé terminal'] = df['Libellé terminal'].str.replace("T1_6","Terminal 1_6")
        st.success("Programme complet chargée !")
        return df

    df_pgrm = df()         
    start_all = tm.time()
    l_date = pd.to_datetime(df_pgrm['Local Date'].unique().tolist()).date
    l_date = sorted(l_date)
    uploaded_file1 = st.file_uploader("Choisir le fichier hypotheses_repartition_correspondances.xlsx :", key=4)
    if uploaded_file1 is not None:
        @st.cache(suppress_st_warning=True,allow_output_mutation=True)
        def HYPOTHESE_REP():
            df = pd.read_excel(uploaded_file1, name_hyp)
            df['plage'] = 'am'
            df.loc[df['heure_debut']>=(datetime.time(17)) , 'plage'] = 'pm'             
            return df
    @st.cache(suppress_st_warning=True,allow_output_mutation=True)
    def COURBE_PRES(t):
        df = pd.read_excel('courbe_presentation.xlsx', t)             
        return df
    col1, col2 = st.columns(2)
    with col1:
        debut = st.date_input("Date de début :",datetime.today().strftime('%Y-%m-%d'), key=10)
    with col2:    
        fin = st.date_input("Date de fin :",datetime.today().strftime('%Y-%m-%d'), key=2)
    
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
        list_terminaux = ['T2AC', 'T2BD', 'T2E', 'T2F', 'T2G', 'T3','T1_Inter','T1_5','T1_6']
        
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
        df_hyp_rep = HYPOTHESE_REP()
        
        
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
        

        

        
    ### DISPATCH ###       
        def DISPATCH_NEW(df, hyp_rep):
            """Permet la création d'un DF dispatch qui facilite le tri par batterie de PIF"""
            col = ['Local Date', 'Horaire théorique', 'Prov Dest', 'A/D', 'Libellé terminal', 'K CNT', 'K CTR', 
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

            #                IMPLEMENTATION T1

            dispatch_df = pd.DataFrame(columns = col, index = df['Unnamed: 0'])

            dispatch_df['Local Date'] = df['Local Date']
            dispatch_df['Horaire théorique'] = df['Horaire théorique']
            dispatch_df['Prov Dest'] = df['Prov Dest']
            dispatch_df['A/D'] = df['A/D']
            dispatch_df['Libellé terminal'] = df['Libellé terminal']

        #           variable 1ere ligne a lire : "hypothèse de répartition K vers terminal2ABCD le matin (am = matin, pm = soir cad après 17h)

        #            Si une erreur de flottant survient, cela provient certainement d'ici : les valeurs ne sont pas considérées comme des flottants mais en série d'un element 
        #            donc on les transforme en liste puis on récupère le 1er (et normalement unique élément). Contrairement aux 6 d'avant qui eux sont directement des flottants 
        #           grace au "1 - valeur"
        #            En cas de bug Retirez le .tolist()[0] 
            def hyp_rep_salle(salle_apport, salle_emport, periode):
                return hyp_rep.loc[(hyp_rep['salle_apport'] == salle_apport) & (hyp_rep['salle_emport'] == salle_emport) & (hyp_rep['heure_debut'] == hyp_rep[periode][0])]['taux'].tolist()[0]
                
            def hyp_rep_salle_abcd(salle_apport, periode):
                return (1 - hyp_rep.loc[(hyp_rep['salle_apport'] == salle_apport) & (hyp_rep['heure_debut'] == hyp_rep[periode][0])].sum()['taux'])
            
            def dispatch_term(AD, terminal, periode, type_pax = 'Pax CNT TOT', _terminal2 = None):
                if _terminal2 == None:
                    temp = df.loc[(df['A/D'] == AD) & (df['Libellé terminal'] == terminal)]
                else:
                    temp = df.loc[(df['Libellé terminal'] == terminal) | (df['Libellé terminal'] == _terminal2)]
                    temp = temp.loc[temp['A/D'] == periode] #erreur mettre "AD" à la place de periode (à vérifier)
                    
                if periode == 'am':
                    return temp.loc[(temp['Horaire théorique'] >= hyp_rep['heure_debut'][0]) & (temp['Horaire théorique'] < hyp_rep['heure_fin'][0])][type_pax]
                elif periode == 'pm' :
                    return temp.loc[(temp['Horaire théorique'] >= hyp_rep['heure_fin'][0])][type_pax]
                else:
                    return "erreur periode"


        #            Dans chaque colonne de dispatch on a les batteries de PIF, 
        #           et comme on a filtré les vols de la logique des PIF dans l_a_k_am par exemple 
        #           on le multiplie par la proportion de gens allant de K vers T2ABDC (ces PAX utilisent le K CNT), ainsi de suite pour chaque PIF
        #           
        #           Reduce permet ici d'additionner les sous dataframe ensembles et de combler les nan par 0. L'index est tjr celui de df_pgrm_dt, ligne de vol à vols
        #    K CNT
            dispatch_df['K CNT'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EK", "am") * hyp_rep_salle_abcd('salle K', 'heure_debut'), 
                        dispatch_term('A', "EK", "pm") * hyp_rep_salle_abcd('salle K', 'heure_fin'), 
                        dispatch_term('A', "EL", "am") * hyp_rep_salle_abcd('salle L', 'heure_debut'), 
                        dispatch_term('A', "EL", "pm") * hyp_rep_salle_abcd('salle L', 'heure_fin'), 
                        dispatch_term('A', "EM", "am") * hyp_rep_salle_abcd('salle M', 'heure_debut'), 
                        dispatch_term('A', "EM", "pm") * hyp_rep_salle_abcd('salle M', 'heure_fin')])


        #    K CTR
            dispatch_df['K CTR'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EK", "am") * hyp_rep_salle('salle K', 'salle K', 'heure_debut'),
                        dispatch_term('A', "EK", "pm") * hyp_rep_salle('salle K', 'salle K', 'heure_fin'), 
                        dispatch_term('D', "EK", "am", "Pax LOC TOT"), 
                        dispatch_term('D', "EK", "pm", "Pax LOC TOT"),
                        dispatch_term('A', "EL", "am") * hyp_rep_salle('salle L', 'salle K', 'heure_debut'),
                        dispatch_term('A', "EL", "pm") * hyp_rep_salle('salle L', 'salle K', 'heure_fin'), 
                        dispatch_term('A', "EM", "am") * hyp_rep_salle('salle M', 'salle K', 'heure_debut'),
                        dispatch_term('A', "EM", "pm") * hyp_rep_salle('salle M', 'salle K', 'heure_fin')])

        #    L CNT
            dispatch_df['L CNT'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EL", "am") * (hyp_rep_salle('salle L', 'salle L', 'heure_debut') + 
                                                    hyp_rep_salle('salle L', 'C2F', 'heure_debut') + 
                                                    hyp_rep_salle('salle L', 'C2G', 'heure_debut')),
                        dispatch_term('A', "EL", "pm") *  (hyp_rep_salle('salle L', 'salle L', 'heure_fin') + 
                                                    hyp_rep_salle('salle L', 'C2F', 'heure_fin') + 
                                                    hyp_rep_salle('salle L', 'C2G', 'heure_fin'))])


        #    L CTR
            dispatch_df['L CTR'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EK", "am") * hyp_rep_salle('salle K', 'salle L', 'heure_debut'),
                        dispatch_term('A', "EK", "pm") * hyp_rep_salle('salle K', 'salle L', 'heure_fin'),
                        dispatch_term('D', "EL", "am", 'Pax LOC TOT'), 
                        dispatch_term('D', "EL", "pm", 'Pax LOC TOT'),
                        dispatch_term('A', "EM", "am") * hyp_rep_salle('salle M', 'salle L', 'heure_debut'), 
                        dispatch_term('A', "EM", "pm") * hyp_rep_salle('salle M', 'salle L', 'heure_fin')])

        #    M CTR
            dispatch_df['M CTR'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EK", "am") * hyp_rep_salle('salle K', 'salle M', 'heure_debut'),
                        dispatch_term('A', "EK", "pm") * hyp_rep_salle('salle K', 'salle M', 'heure_fin'),
                        dispatch_term('A', "EL", "am") * hyp_rep_salle('salle L', 'salle M', 'heure_debut'),
                        dispatch_term('A', "EL", "pm") * hyp_rep_salle('salle L', 'salle M', 'heure_fin'),
                        dispatch_term('A', "EM", "am") * hyp_rep_salle('salle M', 'salle M', 'heure_debut'),
                        dispatch_term('A', "EM", "pm") * hyp_rep_salle('salle M', 'salle M', 'heure_fin'),
                        dispatch_term('D', "EM", "am", 'Pax LOC TOT'), 
                        dispatch_term('D', "EM", "pm", 'Pax LOC TOT')])

        #    Galerie EF
            dispatch_df['Galerie EF'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('A', "EK", "am") * (hyp_rep_salle('salle K', 'C2F', 'heure_debut') + hyp_rep_salle('salle K', 'C2G', 'heure_debut')),
                        dispatch_term('A', "EK", "pm") * (hyp_rep_salle('salle K', 'C2F', 'heure_fin') + hyp_rep_salle('salle K', 'C2G', 'heure_fin')),
                        dispatch_term('A', "EM", "am") * hyp_rep_salle('salle M', 'C2F', 'heure_debut'), 
                        dispatch_term('A', "EM", "pm") * hyp_rep_salle('salle M', 'C2F', 'heure_fin')])

        #    C2F
            dispatch_df['C2F'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "F", "am", 'Pax LOC TOT'), 
                        dispatch_term('D', "F", "pm", 'Pax LOC TOT')])

        #    C2G
            dispatch_df['C2G'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "G", "am", 'Pax LOC TOT'), 
                        dispatch_term('D', "G", "pm", 'Pax LOC TOT')])


        #    LAC
            dispatch_df['Liaison AC'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 2A", "am", "PAX TOT"), 
                        dispatch_term('D', "Terminal 2A", "pm", 'PAX TOT'),
                        dispatch_term('D', "Terminal 2C", "am", 'PAX TOT'),
                        dispatch_term('D', "Terminal 2C", "pm", 'PAX TOT')])


        #    LBD
            dispatch_df['Liaison BD'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 2B", "am", 'PAX TOT'), 
                        dispatch_term('D', "Terminal 2B", "pm", 'PAX TOT'),
                        dispatch_term('D', "Terminal 2D", "am", 'PAX TOT'),
                        dispatch_term('D', "Terminal 2D", "pm", 'PAX TOT')])

        #    T3
            dispatch_df['T3'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 3", "am", 'PAX TOT'), 
                        dispatch_term('D', "Terminal 3", "pm", 'PAX TOT')])
            

            #    Terminal 1 Jonction
            dispatch_df['Terminal 1'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 1", "am", 'PAX TOT'), 
                        dispatch_term('D', "Terminal 1", "pm", 'PAX TOT')])

            
            #    Terminal 1 Schengen (5 et 6)
            dispatch_df['Terminal 1_5'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 1_5", "am", 'PAX TOT'), 
                        dispatch_term('D', "Terminal 1_5", "pm", 'PAX TOT')])
            
            dispatch_df['Terminal 1_6'] = reduce(lambda a, b: a.add(b, fill_value = 0), 
                        [dispatch_term('D', "Terminal 1_6", "am", 'PAX TOT'), 
                        dispatch_term('D', "Terminal 1_6", "pm", 'PAX TOT')])         

            dispatch_df.fillna(0, inplace=True)

            return dispatch_df

 
        dispatch = DISPATCH_NEW(df_pgrm_dt, df_hyp_rep)
 
        #dispatch.to_excel("dispatch.xlsx", sheet_name="dispatch")
        

        liste_df_courbe_presentation_terminal = []
        
        for t in list_terminaux:
            liste_df_courbe_presentation_terminal.append((t, COURBE_PRES(t)))
        
        def courbe(decalage, df_c):
            l_f = df_c['faisceau_geographique'].unique().tolist()
            
            courbe = []
            for i in range(len(l_f)):    
                courbe.append((l_f[i], df_c['pourc'].loc[(df_c['faisceau_geographique'] == l_f[i])
                                                & (df_c['heure_debut'] == df_c['heure_debut'][0])].tolist()))
            for c in range(len(l_f)):
                for i in range(decalage):
                    courbe[c][1].append(0)
            return courbe
        
        #Décalage de la courbe de présentation pour ramener l quart des pax après l'heure théorique avant. On rajoute 25% de zéro, mathématiquement exacte pour la convolution
        dec = 2 + int(0.25 * len(liste_df_courbe_presentation_terminal[3][1]['pourc'].loc[(liste_df_courbe_presentation_terminal[3][1]['faisceau_geographique'] == 'Métropole')
                                                & (liste_df_courbe_presentation_terminal[3][1]['heure_debut'] == liste_df_courbe_presentation_terminal[3][1]['heure_debut'][0])].tolist()))
            
        
        l_courbe_geo_t = []
        
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[0][0], courbe(dec + 6, liste_df_courbe_presentation_terminal[0][1]))) #AC
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[1][0], courbe(dec + 6, liste_df_courbe_presentation_terminal[1][1]))) #BD
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[2][0], courbe(dec + 6, liste_df_courbe_presentation_terminal[2][1]))) #2E
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[3][0], courbe(dec, liste_df_courbe_presentation_terminal[3][1]))) #2F
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[4][0], courbe(dec + 4, liste_df_courbe_presentation_terminal[4][1]))) #2G
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[5][0], courbe(dec, liste_df_courbe_presentation_terminal[5][1]))) #T3
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[6][0], courbe(dec + 3, liste_df_courbe_presentation_terminal[6][1]))) #T1 Inter
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[7][0], courbe(dec + 3, liste_df_courbe_presentation_terminal[7][1]))) #T1_5
        l_courbe_geo_t.append((liste_df_courbe_presentation_terminal[8][0], courbe(dec + 3, liste_df_courbe_presentation_terminal[8][1]))) #T1_6
        
        #st.write(l_courbe_geo_t)
        
        l_faisceaux = [l_courbe_geo_t[0][1][i][0] for i in range(len(l_courbe_geo_t[0][1]))]
        
        from scipy import signal 
            
        def CREATE_DF_SITE(dispatch_df, site):
            """Permet de créer le format de l'export pif final"""
            c = ['jour', 'heure', 'site', 'charge', 'type']
            l_pas10min = pd.date_range(pd.datetime(2022,1,1), periods=144, freq="10T").time.tolist()
            df = pd.DataFrame(columns=c)
            l_jour = dispatch_df['Local Date'].sort_values(ascending = True).unique().tolist()
            nb_jour = len(l_jour)
            df['heure'] = l_pas10min * nb_jour
            df['site'] = site
            df['charge'] = 0
            df['type'] = "pifbi_python"
            for i in range(len(l_jour)):
                df.iloc[144*i:144*(i+1), 0] = pd.to_datetime(l_jour[i])
            return df
        
        def ITERATE_SITE(dispatch_df):
            l_df_site = {}
            l_site = dispatch_df.columns.tolist()
            for site_i in range(5, dispatch_df.shape[1]):
                l_df_site[l_site[site_i]] = CREATE_DF_SITE(dispatch_df, l_site[site_i])
                #l_df_site.append((l_site[site_i]: CREATE_DF_SITE(dispatch_df, l_site[site_i])))

            return l_df_site
        
      
        global pb_index
        pb_index = 0
                        
        
        def EXPORT_PIF(dispatch_df, df_faisceaux_geo, l_f, l_courbe_t):
            o = 10
            my_bar = placeholder.progress(o)
            #Tout le calcul pour avoir les faisceaux et les terminaux lissées par les courbes de présentation et de débarquement"""
            from datetime import time
            
            l_f_iata = []
            for f in l_f:
                l_f_iata.append(df_faisceaux_geo['escale_IATA'].loc[
                        (df_faisceaux_geo['faisceau_geographique'] == f)].tolist())
            
        #    33% des gens débarquent toutes les 5 min pour un total de 15 min ce qui fait 1.
    #           Le temps total est de 30 min, chaque valeur d'une liste dure 10 min, grain du pgrm_complet)
    #            courbe_deb_generique = [0, 0.33, 0.66]
        
    #            Meme principe pour k, l, m sauf que l'on regarde le temps de trajet pour les corres k -> k, k -> l, k -> m etc. et on fait la moyenne
    #            pour une salle. Ex pour k : k->k + k->l + k->m = 44 min donc pour courbe_deb_k on prend en moyenne 44/3 ~ 12 min de trajet. donc 18 min de deb tot
            courbe_deb_k = [0, 0.4, 0.6]
            courbe_deb_l = [0.1, 0.45, 0.45]
            courbe_deb_m = [0, 0.5, 0.5]
            
            df_site = ITERATE_SITE(dispatch_df)
            
            
            pax_mixte_k_ctr = [[] for i in range(len(l_courbe_t[2][1]))]
            pax_mixte_k_ctr_i = [[] for i in range(len(l_courbe_t[2][1]))]

            pax_mixte_l_ctr = [[] for i in range(len(l_courbe_t[2][1]))]
            pax_mixte_l_ctr_i = [[] for i in range(len(l_courbe_t[2][1]))]

            pax_mixte_m_ctr = [[] for i in range(len(l_courbe_t[2][1]))]
            pax_mixte_m_ctr_i = [[] for i in range(len(l_courbe_t[2][1]))]

            pax_od_c2f = [[] for i in range(len(l_courbe_t[3][1]))]
            pax_od_c2f_i = [[] for i in range(len(l_courbe_t[3][1]))]
            
            pax_od_c2g = [[] for i in range(len(l_courbe_t[4][1]))]
            pax_od_c2g_i = [[] for i in range(len(l_courbe_t[4][1]))]
            
            pax_od_ac = [[] for i in range(len(l_courbe_t[0][1]))]
            pax_od_ac_i = [[] for i in range(len(l_courbe_t[0][1]))]
            
            pax_od_bd = [[] for i in range(len(l_courbe_t[1][1]))]
            pax_od_bd_i = [[] for i in range(len(l_courbe_t[1][1]))]
            
            pax_od_t3 = [[] for i in range(len(l_courbe_t[5][1]))]
            pax_od_t3_i = [[] for i in range(len(l_courbe_t[5][1]))]
            
            my_bar.progress(o + 10)
            o += 10
            my_bar.progress(o +10)
            o += 10

    #                IMPLEMENTATION T1
            
            pax_od_t1_j = [[] for i in range(len(l_courbe_t[6][1]))]
            pax_od_t1_j_i = [[] for i in range(len(l_courbe_t[6][1]))]
                
            pax_od_t1_5 = [[] for i in range(len(l_courbe_t[7][1]))]
            pax_od_t1_5_i = [[] for i in range(len(l_courbe_t[7][1]))]

            pax_od_t1_6 = [[] for i in range(len(l_courbe_t[7][1]))]
            pax_od_t1_6_i = [[] for i in range(len(l_courbe_t[7][1]))]

            import time as tm
            start5 = tm.time()

            def CLEAN_TIME(m):
                t = '0:00'.join(str(m).rsplit('5:00', 1))
                l = [int(k) for k in t.split(':')]
                time_r = time(hour = l[0], minute = l[1], second = l[2])
                return time_r

            dispatch_df['Horaire théorique'] = dispatch_df['Horaire théorique'].apply(lambda x: CLEAN_TIME(x))

            for r in range(dispatch_df.shape[0]):
                 
        #        K CNT
                if dispatch_df['K CNT'][r] != 0:
                    index_k_cnt = df_site['K CNT'].loc[(dispatch_df['Horaire théorique'][r] == df_site['K CNT']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['K CNT']['jour'])].index
                    df_site['K CNT']['charge'][index_k_cnt] += dispatch_df['K CNT'][r]
                    
                
        #        K CTR
                if dispatch_df['K CTR'][r] != 0:             
                    index_k_ctr = df_site['K CTR'].loc[(dispatch_df['Horaire théorique'][r] == df_site['K CTR']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['K CTR']['jour'])].index
                    
                    if dispatch_df['A/D'][r] == 'A':
                        df_site['K CTR']['charge'][index_k_ctr] += dispatch_df['K CTR'][r]
                    else:
                        for i in range(len(l_courbe_t[2][1])):
                            if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                                pax_mixte_k_ctr[i].append(index_k_ctr)
                                pax_mixte_k_ctr_i[i].append(r)
                    
        #        L CNT
                if dispatch_df['L CNT'][r] != 0:
                    index_l_cnt = df_site['L CNT'].loc[(dispatch_df['Horaire théorique'][r] == df_site['L CNT']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['L CNT']['jour'])].index
                    
                    df_site['L CNT']['charge'][index_l_cnt] += dispatch_df['L CNT'][r]
                            
        #        L CTR
                if dispatch_df['L CTR'][r] != 0:
                    index_l_ctr = df_site['L CTR'].loc[(dispatch_df['Horaire théorique'][r] == df_site['L CTR']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['L CTR']['jour'])].index
                    
                    if dispatch_df['A/D'][r] == 'A':
                        df_site['L CTR']['charge'][index_l_ctr] += dispatch_df['L CTR'][r]
                    else:
                        for i in range(len(l_courbe_t[2][1])):
                            if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                                pax_mixte_l_ctr[i].append(index_l_ctr)
                                pax_mixte_l_ctr_i[i].append(r)
                    
                    
    #                    M CTR
                if dispatch_df['M CTR'][r] != 0:
                    index_m_ctr = df_site['M CTR'].loc[(dispatch_df['Horaire théorique'][r] == df_site['M CTR']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['M CTR']['jour'])].index
                    
                    if dispatch_df['A/D'][r] == 'A':
                        df_site['M CTR']['charge'][index_m_ctr] += dispatch_df['M CTR'][r]
                    else:
                        for i in range(len(l_courbe_t[2][1])):
                            if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                                pax_mixte_m_ctr[i].append(index_m_ctr)
                                pax_mixte_m_ctr_i[i].append(r)
                    
        #        Galerie EF
                if dispatch_df['Galerie EF'][r] != 0:
                    m = dispatch_df['Horaire théorique'][r]
                    index_g_ef = df_site['Galerie EF'].loc[(df_site['Galerie EF']['heure'] == m)
                                & (dispatch_df['Local Date'][r] == df_site['Galerie EF']['jour'])].index
                    
                    df_site['Galerie EF']['charge'][index_g_ef] += dispatch_df['Galerie EF'][r]
                    
        #        C2F
                if dispatch_df['C2F'][r] != 0:
                    index_c2f = df_site['C2F'].loc[(dispatch_df['Horaire théorique'][r] == df_site['C2F']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['C2F']['jour'])].index
                    
                    
                    for i in range(len(l_courbe_t[3][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_c2f[i].append(index_c2f)
                            pax_od_c2f_i[i].append(r)
        
        #        C2G
                if dispatch_df['C2G'][r] != 0:
                    index_c2g = df_site['C2G'].loc[(dispatch_df['Horaire théorique'][r] == df_site['C2G']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['C2G']['jour'])].index
                    
                    for i in range(len(l_courbe_t[4][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_c2g[i].append(index_c2g)
                            pax_od_c2g_i[i].append(r)
                            
        #        Liaison AC
                if dispatch_df['Liaison AC'][r] != 0:
                    index_ac = df_site['Liaison AC'].loc[(dispatch_df['Horaire théorique'][r] == df_site['Liaison AC']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['Liaison AC']['jour'])].index
                    
                    for i in range(len(l_courbe_t[0][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_ac[i].append(index_ac)
                            pax_od_ac_i[i].append(r)
                            
        #        Liaison BD
                if dispatch_df['Liaison BD'][r] != 0:

                    index_bd = df_site['Liaison BD'].loc[(dispatch_df['Horaire théorique'][r] == df_site['Liaison BD']['heure'])
                                & (dispatch_df['Local Date'][r] == df_site['Liaison BD']['jour'])].index
        
                    for i in range(len(l_courbe_t[1][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_bd[i].append(index_bd)
                            pax_od_bd_i[i].append(r)
                            
        #        T3
                if dispatch_df['T3'][r] != 0:
                    index_t3 = df_site['T3'].loc[(dispatch_df['Horaire théorique'][r] == df_site['T3']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['T3']['jour'])].index
                    
                    for i in range(len(l_courbe_t[5][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_t3[i].append(index_t3)
                            pax_od_t3_i[i].append(r)
                            
        #        Terminal 1 jonction
                if dispatch_df['Terminal 1'][r] != 0:
                    index_t1_j = df_site['Terminal 1'].loc[(dispatch_df['Horaire théorique'][r] == df_site['Terminal 1']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['Terminal 1']['jour'])].index
                    
                    for i in range(len(l_courbe_t[6][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_t1_j[i].append(index_t1_j)
                            pax_od_t1_j_i[i].append(r)
                
        #        Terminal 1 Schengen 5 et 6
                if dispatch_df['Terminal 1_5'][r] != 0:
                    index_t1_5 = df_site['Terminal 1_5'].loc[(dispatch_df['Horaire théorique'][r] == df_site['Terminal 1_5']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['Terminal 1_5']['jour'])].index
                    
                    for i in range(len(l_courbe_t[7][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_t1_5[i].append(index_t1_5)
                            pax_od_t1_5_i[i].append(r)

                if dispatch_df['Terminal 1_6'][r] != 0:
                    index_t1_6 = df_site['Terminal 1_6'].loc[(dispatch_df['Horaire théorique'][r] == df_site['Terminal 1_6']['heure'])
                                    & (dispatch_df['Local Date'][r] == df_site['Terminal 1_6']['jour'])].index
                    
                    for i in range(len(l_courbe_t[8][1])):
                        if dispatch_df['Prov Dest'][r] in l_f_iata[i]:
                            pax_od_t1_6[i].append(index_t1_6)
                            pax_od_t1_6_i[i].append(r)                    
                
                if r%500==0:
                    print(str(r)+"/"+str(dispatch_df.shape[0]))

            end = tm.time()
            #st.write('dispatch')
            #st.write(end - start5)         
            my_bar.progress(o +10)     
            o += 10   

            start2 = tm.time()
            l_ch_k_ctr = df_site["K CTR"]['charge'].tolist()
            convo_k_ctr = list(signal.convolve(l_ch_k_ctr, courbe_deb_k, mode='same'))
            df_site["K CTR"]['charge'] = convo_k_ctr 
            
            l_ch_l_ctr = df_site["L CTR"]['charge'].tolist()
            convo_l_ctr = list(signal.convolve(l_ch_l_ctr, courbe_deb_l, mode='same'))
            df_site["L CTR"]['charge'] = convo_l_ctr 
            
            l_ch_m_ctr = df_site["M CTR"]['charge'].tolist()
            convo_m_ctr = list(signal.convolve(l_ch_m_ctr, courbe_deb_m, mode='same'))
            df_site["M CTR"]['charge'] = convo_m_ctr 
                
            nb_s = 8
            print("\n1/"+str(nb_s))
            for i in range(len(pax_mixte_k_ctr)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_mixte_k_ctr[i])):
                    df_charge['charge'][pax_mixte_k_ctr[i][value_index]] += dispatch_df['K CTR'][pax_mixte_k_ctr_i[i][value_index]]
                
                print(str(l_courbe_t[2][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_mixte_k_ctr), 0)) +"%")
                convo_k_ctr = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[2][1][i][1], mode='same'))
                df_site["K CTR"]['charge'] += convo_k_ctr
            

                
            print("\n2/"+str(nb_s))
            for i in range(len(pax_mixte_l_ctr)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_mixte_l_ctr[i])):
                    df_charge['charge'][pax_mixte_l_ctr[i][value_index]] += dispatch_df['L CTR'][pax_mixte_l_ctr_i[i][value_index]]
                    
                print(str(l_courbe_t[2][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_mixte_l_ctr), 0)) +"%")    
                convo_l_ctr = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[2][1][i][1], mode='same'))
                df_site['L CTR']['charge'] += convo_l_ctr
            

            
            print("\n3/"+str(nb_s))
            for i in range(len(pax_mixte_m_ctr)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_mixte_m_ctr[i])):
                    df_charge['charge'][pax_mixte_m_ctr[i][value_index]] += dispatch_df['M CTR'][pax_mixte_m_ctr_i[i][value_index]]
                
                print(str(l_courbe_t[2][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_mixte_m_ctr), 0)) +"%")
                convo_m_ctr = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[2][1][i][1], mode='same'))
                df_site['M CTR']['charge'] += convo_m_ctr
            

            
            print("\n4/"+str(nb_s))
            for i in range(len(pax_od_c2f)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_c2f[i])):
                    df_charge['charge'][pax_od_c2f[i][value_index]] += dispatch_df['C2F'][pax_od_c2f_i[i][value_index]]
                
                print(str(l_courbe_t[3][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_c2f), 0)) +"%")
                convo_c2f = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[3][1][i][1], mode='same'))
                df_site['C2F']['charge'] += convo_c2f
            
            my_bar.progress(o +10)
            o += 10
            
            print("\n5/"+str(nb_s))
            for i in range(len(pax_od_c2g)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_c2g[i])):
                    df_charge['charge'][pax_od_c2g[i][value_index]] += dispatch_df['C2G'][pax_od_c2g_i[i][value_index]]
                
                print(str(l_courbe_t[4][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_c2g), 0)) +"%")
                convo_c2g = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[4][1][i][1], mode='same'))
                df_site['C2G']['charge'] += convo_c2g

            my_bar.progress(o +10)
            o += 10
            print("\n6/"+str(nb_s))
            for i in range(len(pax_od_ac)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_ac[i])):
                    df_charge['charge'][pax_od_ac[i][value_index]] += dispatch_df['Liaison AC'][pax_od_ac_i[i][value_index]]
                    
                print(str(l_courbe_t[0][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_ac), 0)) +"%")
                convo_ac = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[0][1][i][1], mode='same'))
                df_site['Liaison AC']['charge'] += convo_ac
            

            my_bar.progress(o +10)
            o += 10
            print("\n7/"+str(nb_s))
            for i in range(len(pax_od_bd)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_bd[i])):
                    df_charge['charge'][pax_od_bd[i][value_index]] += dispatch_df['Liaison BD'][pax_od_bd_i[i][value_index]]
                

                convo_bd = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[1][1][i][1], mode='same'))
                df_site['Liaison BD']['charge'] += convo_bd
            
            my_bar.progress(o +10)
            o += 10
            
            for i in range(len(pax_od_t3)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")             
                for value_index in range(len(pax_od_t3[i])):
                    df_charge['charge'][pax_od_t3[i][value_index]] += dispatch_df['T3'][pax_od_t3_i[i][value_index]]
            
                convo_t3 = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[5][1][i][1], mode='same'))
                df_site['T3']['charge'] += convo_t3
            my_bar.progress(o +10)
            o += 10

                       
            for i in range(len(pax_od_t1_j)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_t1_j[i])):
                    df_charge['charge'][pax_od_t1_j[i][value_index]] += dispatch_df['Terminal 1'][pax_od_t1_j_i[i][value_index]]
                
                print(str(l_courbe_t[6][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_t1_j), 0)) +"%")
                convo_t1 = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[6][1][i][1], mode='same'))
                df_site['Terminal 1']['charge'] += convo_t1

            for i in range(len(pax_od_t1_5)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_t1_5[i])):
                    df_charge['charge'][pax_od_t1_5[i][value_index]] += dispatch_df['Terminal 1_5'][pax_od_t1_5_i[i][value_index]]
                
                print(str(l_courbe_t[7][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_t1_5), 0)) +"%")
                convo_t1_5 = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[7][1][i][1], mode='same'))
                df_site['Terminal 1_5']['charge'] += convo_t1_5

            for i in range(len(pax_od_t1_6)):
                df_charge = CREATE_DF_SITE(dispatch_df, "temp")
                for value_index in range(len(pax_od_t1_6[i])):
                    df_charge['charge'][pax_od_t1_6[i][value_index]] += dispatch_df['Terminal 1_6'][pax_od_t1_6_i[i][value_index]]
                
                print(str(l_courbe_t[8][1][i][0]) + " " + str(round(100 * (i + 1) / len(pax_od_t1_6), 0)) +"%")
                convo_t1_6 = list(signal.convolve(df_charge['charge'].tolist(), l_courbe_t[8][1][i][1], mode='same'))
                df_site['Terminal 1_6']['charge'] += convo_t1_6


            
        #   PAX CNT : on utilise courbe_deb
            l_ch_k_cnt = df_site['K CNT']['charge'].tolist()
            convo_k_cnt = list(signal.convolve(l_ch_k_cnt, courbe_deb_k, mode='same'))
            df_site['K CNT']['charge'] = convo_k_cnt 
            
            l_ch_l_cnt = df_site['L CNT']['charge'].tolist()
            convo_l_cnt = list(signal.convolve(l_ch_l_cnt, courbe_deb_l, mode='same'))
            df_site['L CNT']['charge'] = convo_l_cnt 
            
            l_ch_g_ef = df_site['Galerie EF']['charge'].tolist()
            convo_g_ef = list(signal.convolve(l_ch_g_ef, courbe_deb_m, mode='same'))
            df_site['Galerie EF']['charge'] = convo_g_ef
                    
            my_bar.progress(o +10)
            o += 10
            st.success('Traitement terminé !')

            return df_site

        
        def TO_EXCEL(df_site, path):
            col = df_site["K CNT"].columns.tolist()
            
            df = pd.DataFrame(columns=col)
            df['jour'] = df_site["K CNT"]['jour'].tolist() * len(df_site)
            
            for i,j in zip(df_site,range(len(df_site))):

                df['charge'][len(df_site[i])*j:len(df_site[i])*(j+1)] = df_site[i]['charge']
                df['heure'][len(df_site[i])*j:len(df_site[i])*(j+1)] = df_site[i]['heure']
                df['site'][len(df_site[i])*j:len(df_site[i])*(j+1)] = df_site[i]['site']
                df['type'][len(df_site[i])*j:len(df_site[i])*(j+1)] = df_site[i]['type']
                 
            return df


        directory_exp = "export_pif_du_" + str(start_date.date()) + "_au_" + str(end_date.date()) + ".xlsx"
        x = TO_EXCEL(EXPORT_PIF(dispatch, df_faisceaux, l_faisceaux, l_courbe_geo_t), path = directory_exp)
        end3 = tm.time()

        st.info("Export PIF créé avec succès !" + "\n\nPour lancer une nouvelle étude, lancer uniquement 'CHOISIR LES DATES'")
        
        import io
        from pyxlsb import open_workbook as open_xlsb

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            x.to_excel(writer, sheet_name=name_output)
            writer.save()

            st.download_button(
            label="Télécharger fichier Export pif",
            data=buffer,
            file_name=directory_exp,
            mime="application/vnd.ms-excel"
            )

