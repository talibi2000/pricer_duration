import streamlit as st
import numpy as np
import pandas as pd
import openpyxl
from datetime import datetime
from PIL import Image
from dateutil.relativedelta import relativedelta
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


def calcul_emprunt(profil, capital, taux_interet, date_maturite, VALUE_DATE, periodicite):
    def extraire_valeurs(d):
        if d < 1:
            d1, d2 = 1, 1
        elif d < 2:
            d1, d2 = 1, 2
        elif d < 3:
            d1, d2 = 2, 3
        elif d < 6:
            d1, d2 = 3, 6
        elif d < 9:
            d1, d2 = 6, 9
        elif d < 12:
            d1, d2 = 9, 12
        else:
            multiplicateur = int(d / 12)
            d1 = multiplicateur * 12
            d2 = d1 + 12

        return d1, d2

    dff = pd.read_excel('C:\\Users\\HP\\Desktop\\Output\\TCI.xlsx', sheet_name='DATA')
    if periodicite == 'annuel':
        increment = relativedelta(months=12)
        frequence_paiement = 1
    elif periodicite == 'semestriel':
        increment = relativedelta(months=6)
        frequence_paiement = 2
    elif periodicite == 'trimestriel':
        increment = relativedelta(months=3)
        frequence_paiement = 4
    else:
        increment = relativedelta(months=1)
        frequence_paiement = 12
    periodes = []
    flux = []
    discount_factors = []
    date_actuelle = VALUE_DATE
    nombre_paiements = int(((date_maturite - VALUE_DATE).days / 365) * frequence_paiement)
    table_amortissement = np.zeros((nombre_paiements, 6), dtype=object)
    if profil == 'ECHCONST' or profil == "A":
        taux_interet_periodique = (taux_interet / frequence_paiement) / 100
        annuite = (capital * taux_interet_periodique) / (1 - (1 + taux_interet_periodique) ** (-nombre_paiements))
        capital_rest = capital
        for j in range(nombre_paiements):
            date_actuelle += increment
            interets = capital_rest * taux_interet_periodique
            flux.append(annuite)
            amortissement = annuite - interets
            capital_rest -= amortissement
            n = int((date_actuelle - VALUE_DATE).days / 360)
            periodes.append(n)
            discount_factor = 1 / (1 + ((taux_interet_periodique) ** (periodes[j])))
            discount_factors.append(discount_factor)
            table_amortissement[j] = j + 1, pd.to_datetime(date_actuelle,
                                                           format='%Y-%m-%d'), annuite, interets, amortissement, capital_rest
        duree_vie_ponderee = 12 * sum([periodes[i] * flux[i] * discount_factors[i] for i in range(len(periodes))]) \
                             / sum([flux[i] * discount_factors[i] for i in range(len(periodes))])
        df_amortissement = pd.DataFrame(table_amortissement,
                                        columns=['id', "Date d'échéance", 'annuité', 'Intérêts', 'Amortissement',
                                                 'capital rest'])
    elif profil == 'LINEAIRE' or profil == "L":
        taux_interet_periodique = (taux_interet / 100) / frequence_paiement
        amortissement = capital / nombre_paiements
        capital_rest = capital
        for j in range(nombre_paiements):
            date_actuelle += increment
            interets = capital_rest * taux_interet_periodique
            annuite = amortissement + interets
            flux.append(interets)
            capital_rest -= amortissement
            n = int((date_actuelle - VALUE_DATE).days / 360)
            periodes.append(n)
            discount_factor = 1 / (1 + ((taux_interet_periodique) ** (periodes[j])))
            discount_factors.append(discount_factor)
            flux.append(annuite)

            table_amortissement[j] = 1, pd.to_datetime(date_actuelle,
                                                       format='%Y-%m-%d'), annuite, interets, amortissement, capital_rest
        duree_vie_ponderee = 12 * sum([periodes[i] * flux[i] * discount_factors[i] for i in range(len(periodes))]) \
                             / sum([flux[i] * discount_factors[i] for i in range(len(periodes))])
        df_amortissement = pd.DataFrame(table_amortissement,
                                        columns=['id', "Date d'échéance", 'annuité', 'Intérêts', 'Amortissement',
                                                 'capital rest'])

    else:
        taux_interet_periodique = (taux_interet / 100) / frequence_paiement
        annuite = capital * taux_interet_periodique
        capital_rest = capital
        for j in range(nombre_paiements):

            interets = capital_rest * taux_interet_periodique

            if j == nombre_paiements - 1:
                amortissement = capital
            else:
                amortissement = 0
            annuite = interets + amortissement
            date_actuelle += increment
            capital_rest -= amortissement
            n = int((date_actuelle - VALUE_DATE).days / 360)
            periodes.append(n)
            discount_factor = 1 / (1 + ((taux_interet_periodique) ** (periodes[j])))
            discount_factors.append(discount_factor)
            flux.append(annuite)
            table_amortissement[j] = 1, pd.to_datetime(date_actuelle,
                                                       format='%Y-%m-%d'), annuite, interets, amortissement, capital_rest

        duree_vie_ponderee = 12 * sum([periodes[i] * flux[i] * discount_factors[i] for i in range(len(periodes))]) \
                             / sum([flux[i] * discount_factors[i] for i in range(len(periodes))])
        df_amortissement = pd.DataFrame(table_amortissement,
                                        columns=['id', "Date d'échéance", 'annuité', 'Intérêts', 'Amortissement',
                                                 'capital rest'])

    aujourdhui = datetime.now()
    if VALUE_DATE.month == aujourdhui.month and VALUE_DATE.year == aujourdhui.year:
        VALUE_DATE = VALUE_DATE.replace(month=VALUE_DATE.month - 1)
    d1, d2 = extraire_valeurs(duree_vie_ponderee)

    if duree_vie_ponderee < 1:
        valeur = dff.loc[dff['Date'].dt.strftime('%Y-%m') == VALUE_DATE.strftime('%Y-%m'), 1]
        TSR_interpole = valeur.iloc[0]
        TCI_interpole = valeur.iloc[1]
    else:
        valeur1 = dff.loc[dff['Date'].dt.strftime('%Y-%m') == VALUE_DATE.strftime('%Y-%m'), d1]
        valeur2 = dff.loc[dff['Date'].dt.strftime('%Y-%m') == VALUE_DATE.strftime('%Y-%m'), d2]
        TSR1 = valeur1.iloc[0]
        TCI1 = valeur1.iloc[1]
        TSR2 = valeur2.iloc[0]
        TCI2 = valeur2.iloc[1]
        resultats = pd.DataFrame({'TSR1': [TSR1], 'TCI1': [TCI1], 'TSR2': [TSR2], 'TCI2': [TCI2]})

        def linear_interpolation(x1, y1, x2, y2, x):
            return y1 + (y2 - y1) * (x - x1) / (x2 - x1)

        TSR_interpole = linear_interpolation(d1, TSR1, d2, TSR2, duree_vie_ponderee)
        TCI_interpole = linear_interpolation(d1, TCI1, d2, TCI2, duree_vie_ponderee)
    TSR = f"TSR :  {TSR_interpole * 100:.6f} %"
    TCI = f"TCI :  {TCI_interpole * 100:.6f} %"
    dureee = f"Durée :  {duree_vie_ponderee:.4f} mois {duree_vie_ponderee / 12:.4f} ans"
    return dureee, TCI, TSR, df_amortissement

st.title("Calculateur de remboursement de prêt : Durée, TCI et Écoulement")
image = Image.open("attijari_logo.png")

st.image(image, caption='Attijariwafa Bank', width=200)
profil = st.selectbox("Profil", ["ECHCONST", "LINEAIRE", "INFINE"])
capital = st.number_input("Capital", min_value=0.0, value=0.0, step=100.0)
taux_interet = st.number_input("Taux d'intérêt (%)", min_value=0.0, value=5.0, step=0.5)
if taux_interet == 0:
    st.warning("Veuillez changer la valeur du taux d'intérêt.")
date_maturite = st.date_input("Date de maturité")
value_date = st.date_input("Value date")
periodicite = st.selectbox("Périodicité", ["annuel", "semestriel", "trimestriel", "mensuel"])

if st.button("Calculer"):
    duree, tci, tsr, df_amortissement = calcul_emprunt(
        profil, capital, taux_interet, date_maturite, value_date, periodicite
    )
    st.markdown("<div style='font-weight: bold; color: blue; text-decoration: underline;'>Durée :</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-weight: bold;'>{duree}</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-weight: bold; color: blue; text-decoration: underline;'>TCI :</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-weight: bold;'>{tci}</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-weight: bold; color: blue; text-decoration: underline;'>TSR :</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-weight: bold;'>{tsr}</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-weight: bold; color: blue; text-decoration: underline;'>Tableau d'amortissement :</div>", unsafe_allow_html=True)
    st.write(df_amortissement)





