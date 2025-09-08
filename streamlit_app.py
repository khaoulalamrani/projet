import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ----------------------------
# 1. Charger les données CSV
# ----------------------------
df = pd.read_csv("entreprises.csv")  # Remplacer par ton CSV

# ----------------------------
# 2. Nettoyage des colonnes numériques
# ----------------------------
numeric_cols = ["Note_Google", "Nb_Avis_Google", "Score_Presence_Digitale",
                "Distance-TARMIZ(KM)", "Anciennete_Estimee"]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")  # force conversion

# ----------------------------
# 3. Titre de l'app Streamlit
# ----------------------------
st.title("Dashboard - Secteur Optique")

# ----------------------------
# 4. Top villes
# ----------------------------
st.subheader("Top 10 Villes")
top_cities = df["Ville"].value_counts().head(10)
fig1 = px.bar(x=top_cities.index, y=top_cities.values,
              labels={"x":"Ville","y":"Nombre d'entreprises"},
              title="Top 10 des villes")
st.plotly_chart(fig1)

# ----------------------------
# 5. Carte géographique
# ----------------------------
if {"Latitude","Longitude"}.issubset(df.columns):
    st.subheader("Répartition Géographique")
    geo_df = df.dropna(subset=["Latitude","Longitude"])
    fig2 = px.scatter_mapbox(
        geo_df,
        lat="Latitude",
        lon="Longitude",
        hover_name="Nom",
        hover_data=["Ville","Note_Google","Nb_Avis_Google"],
        color="Note_Google",
        size="Nb_Avis_Google",
        color_continuous_scale="Viridis",
        mapbox_style="open-street-map",
        height=500,
        title="Répartition Géographique des Optiques"
    )
    st.plotly_chart(fig2)

# ----------------------------
# 6. Analyse Notes Google
# ----------------------------
st.subheader("Analyse des Notes Google")
fig3 = make_subplots(
    rows=2, cols=2,
    subplot_titles=("Distribution des Notes", "Notes vs Nombre d'Avis",
                    "Notes par Ville (Top 5)", "Histogram Nombre d'Avis")
)

# Distribution des notes
fig3.add_trace(go.Histogram(x=df["Note_Google"].dropna(), nbinsx=20, name="Notes"),
               row=1, col=1)

# Notes vs Nb d'avis
fig3.add_trace(go.Scatter(x=df["Note_Google"], y=df["Nb_Avis_Google"],
                          mode="markers", name="Notes vs Avis"), row=1, col=2)

# Notes par ville (Top 5)
top5 = df["Ville"].value_counts().head(5).index
city_notes = df[df["Ville"].isin(top5)].groupby("Ville")["Note_Google"].mean()
fig3.add_trace(go.Bar(x=city_notes.index, y=city_notes.values, name="Moyenne par ville"),
               row=2, col=1)

# Histogram Nb d'avis
fig3.add_trace(go.Histogram(x=df["Nb_Avis_Google"].dropna(), nbinsx=30, name="Nb Avis"),
               row=2, col=2)

fig3.update_layout(height=700, title_text="Analyse des Métriques Google")
st.plotly_chart(fig3)

# ----------------------------
# 7. Présence digitale
# ----------------------------
st.subheader("Présence Digitale")
digital_cols = ["Site web","Réseaux sociaux","Email"]
digital_presence = []
labels = []
for col in digital_cols:
    if col in df.columns:
        count = df[col].notna().sum()
        digital_presence.append(count)
        labels.append(col)

fig4 = go.Figure()
fig4.add_trace(go.Bar(x=labels, y=digital_presence,
                      text=[f"{v} ({v/len(df)*100:.1f}%)" for v in digital_presence],
                      textposition="auto"))
fig4.update_layout(yaxis_title="Nombre d'entreprises", title="Présence Digitale")
st.plotly_chart(fig4)

# ----------------------------
# 8. Distance TARMIZ
# ----------------------------
if "Distance-TARMIZ(KM)" in df.columns:
    st.subheader("Analyse de la Distance par rapport à TARMIZ")
    fig5 = make_subplots(rows=1, cols=2, subplot_titles=("Distribution des Distances","Distance vs Note Google"))
    fig5.add_trace(go.Histogram(x=df["Distance-TARMIZ(KM)"].dropna(), name="Distance"), row=1, col=1)
    fig5.add_trace(go.Scatter(x=df["Distance-TARMIZ(KM)"], y=df["Note_Google"], mode="markers", name="Distance vs Note"), row=1, col=2)
    fig5.update_layout(height=400)
    st.plotly_chart(fig5)

# ----------------------------
# 9. Score Présence Digitale
# ----------------------------
if "Score_Presence_Digitale" in df.columns:
    st.subheader("Score de Présence Digitale")
    fig6 = px.histogram(df, x="Score_Presence_Digitale", nbins=20, title="Distribution du Score Digital")
    st.plotly_chart(fig6)

    fig7 = px.scatter(df, x="Score_Presence_Digitale", y="Note_Google",
                      hover_data=["Nom","Ville"], title="Score Digital vs Note Google")
    st.plotly_chart(fig7)

# ----------------------------
# 10. Taille des entreprises
# ----------------------------
if "Taille_Entreprise" in df.columns:
    st.subheader("Répartition par Taille d'Entreprise")
    fig8 = px.pie(df, names="Taille_Entreprise", title="Taille des Entreprises")
    st.plotly_chart(fig8)

# ----------------------------
# 11. Résumé statistique
# ----------------------------
st.subheader("Résumé Statistique")
st.write(f"Total entreprises : {len(df)}")
if "Note_Google" in df.columns:
    st.write(f"Note moyenne Google : {df['Note_Google'].mean():.2f}")
if "Distance-TARMIZ(KM)" in df.columns:
    st.write(f"Distance moyenne de TARMIZ : {df['Distance-TARMIZ(KM)'].mean():.1f} km")
if "Score_Presence_Digitale" in df.columns:
    st.write(f"Score digital moyen : {df['Score_Presence_Digitale'].mean():.1f}")
top_city = df["Ville"].value_counts().idxmax()
st.write(f"Ville leader : {top_city} ({df['Ville'].value_counts().max()} optiques)")

st.success("✅ Toutes les visualisations ont été générées !")
