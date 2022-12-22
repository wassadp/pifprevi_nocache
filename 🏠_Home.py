import pandas as pd
import streamlit as st

st.set_page_config(page_title="OutilsPIF V2", page_icon="🏠", layout="centered", initial_sidebar_state="auto", 
                    menu_items={
                        'About': "# This is a Test. This is an *extremely* cool app!"
                    })


st.title('OutilsPIF V2') 

st.write("Cet outil sert à regrouper plusieurs actions effectués au sein de l'IngeX de CDGD.")
st.write("Vous retrouverez ainsi :")
st.markdown("Onglet **Concat** : Un outil de concaténation des programmes AF Skyteam et des programmes ADP.")
st.markdown("Onglet **Pif Previ** : Un outil de prévisions des flux aux différents sites de PIF dans l'aéroport CDG")
st.markdown("Onglet **Export PIF** : Un outil de mise en forme des réalisés PIF.")

st.sidebar.info("Version : 1.0")


hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)