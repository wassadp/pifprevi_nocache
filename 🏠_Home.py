import pandas as pd
import streamlit as st

st.set_page_config(page_title="OutilsPIF V2", page_icon="🏠", layout="centered", initial_sidebar_state="auto", 
                    menu_items={
                        'About': "# This is a Test. This is an *extremely* cool app!"
                    })


st.title('OutilsPIF V2') 

st.write("Cet outil sert à regrouper plusieurs actions effectués au sein de l'IngeX de CDGD. Vous retrouverez ainsi :\n\nUn outil de concaténation des programmes "
                        + "AF Skyteam et des programmes ADP.\n\nUn outil de prévisions des flux aux différents sites de PIF dans l'aéroport CDG.\n"
                        + "\nUn outil de mise en forme des réalisés PIF.")

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)                         