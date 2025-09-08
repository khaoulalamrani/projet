import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl  # Pour lire les fichiers Excel

# Charger les données
df = pd.read_excel('OPTIQUESS.xlsx', engine='openpyxl')

# Configuration du style
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")

# 1. Vue d'ensemble des données
print(f"Données chargées: {df.shape[0]} entreprises optiques")

# 2. Distribution par ville
fig1 = px.bar(
    x=df['Ville'].value_counts().head(10).index,
    y=df['Ville'].value_counts().head(10).values,
    title="Top 10 des villes - Entreprises Optiques",
    labels={'x': 'Ville', 'y': 'Nombre d\'entreprises'}
)
fig1.show()

# 3. Carte géographique
if 'Latitude' in df.columns and 'Longitude' in df.columns:
    geo_df = df.dropna(subset=['Latitude', 'Longitude'])
    fig2 = px.scatter_mapbox(
        geo_df,
        lat="Latitude",
        lon="Longitude",
        hover_name="Nom",
        hover_data=['Ville', 'Note_Google', 'Nb_Avis_Google'],
        color='Note_Google',
        size='Nb_Avis_Google',
        color_continuous_scale="Viridis",
        title="Répartition Géographique des Optiques",
        mapbox_style="open-street-map",
        height=600
    )
    fig2.show()

# 4. Analyse des notes Google
fig3 = make_subplots(
    rows=2, cols=2,
    subplot_titles=('Distribution des Notes', 'Notes vs Nombre d\'Avis', 
                   'Notes par Ville (Top 5)', 'Histogram Nombre d\'Avis')
)

# Distribution des notes
fig3.add_trace(
    go.Histogram(x=df['Note_Google'].dropna(), name='Notes', nbinsx=20),
    row=1, col=1
)

# Scatter Notes vs Avis
fig3.add_trace(
    go.Scatter(x=df['Note_Google'], y=df['Nb_Avis_Google'], 
              mode='markers', name='Notes vs Avis'),
    row=1, col=2
)

# Notes par ville (top 5)
top_cities = df['Ville'].value_counts().head(5).index
city_notes = df[df['Ville'].isin(top_cities)].groupby('Ville')['Note_Google'].mean()
fig3.add_trace(
    go.Bar(x=city_notes.index, y=city_notes.values, name='Moyenne par ville'),
    row=2, col=1
)

# Histogram nombre d'avis
fig3.add_trace(
    go.Histogram(x=df['Nb_Avis_Google'].dropna(), name='Nb Avis', nbinsx=30),
    row=2, col=2
)

fig3.update_layout(height=800, title_text="Analyse des Métriques Google")
fig3.show()

# 5. Présence digitale
digital_cols = ['Site web', 'Réseaux sociaux', 'Email']
digital_presence = []
labels = []

for col in digital_cols:
    if col in df.columns:
        count = df[col].notna().sum()
        digital_presence.append(count)
        labels.append(col)

fig4 = go.Figure()
fig4.add_trace(go.Bar(x=labels, y=digital_presence, 
                     text=[f"{val} ({val/len(df)*100:.1f}%)" for val in digital_presence],
                     textposition='auto'))
fig4.update_layout(title="Présence Digitale des Optiques", 
                  yaxis_title="Nombre d'entreprises")
fig4.show()

# 6. Analyse de la distance par rapport à TARMIZ
if 'Distance-TARMIZ(KM)' in df.columns:
    fig5 = make_subplots(
        rows=1, cols=2,
        subplot_titles=('Distribution des Distances', 'Distance vs Note Google')
    )
    
    fig5.add_trace(
        go.Histogram(x=df['Distance-TARMIZ(KM)'].dropna(), name='Distance'),
        row=1, col=1
    )
    
    fig5.add_trace(
        go.Scatter(x=df['Distance-TARMIZ(KM)'], y=df['Note_Google'], 
                  mode='markers', name='Distance vs Note'),
        row=1, col=2
    )
    
    fig5.update_layout(height=400, title_text="Analyse des Distances par rapport à TARMIZ")
    fig5.show()

# 7. Score de présence digitale
if 'Score_Presence_Digitale' in df.columns:
    fig6 = px.histogram(df, x='Score_Presence_Digitale', 
                       title="Distribution du Score de Présence Digitale",
                       nbins=20)
    fig6.show()
    
    # Corrélation Score Digital vs Note Google
    fig7 = px.scatter(df, x='Score_Presence_Digitale', y='Note_Google',
                     hover_data=['Nom', 'Ville'],
                     title="Score Présence Digitale vs Note Google")
    fig7.show()

# 8. Taille des entreprises
if 'Taille_Entreprise' in df.columns:
    fig8 = px.pie(df, names='Taille_Entreprise', 
                 title="Répartition par Taille d'Entreprise")
    fig8.show()

# 9. Dashboard final avec métriques clés
fig_final = make_subplots(
    rows=3, cols=2,
    subplot_titles=('Top 10 Villes', 'Notes Google Distribution',
                   'Présence Digitale', 'Distance de TARMIZ',
                   'Score Digital vs Note', 'Ancienneté Estimée'),
    specs=[[{"type": "bar"}, {"type": "histogram"}],
           [{"type": "bar"}, {"type": "histogram"}],
           [{"type": "scatter"}, {"type": "histogram"}]]
)

# Top villes
top_cities = df['Ville'].value_counts().head(10)
fig_final.add_trace(
    go.Bar(x=top_cities.index, y=top_cities.values, name='Villes'),
    row=1, col=1
)

# Notes Google
fig_final.add_trace(
    go.Histogram(x=df['Note_Google'].dropna(), name='Notes'),
    row=1, col=2
)

# Présence digitale
fig_final.add_trace(
    go.Bar(x=labels, y=digital_presence, name='Digital'),
    row=2, col=1
)

# Distance TARMIZ
if 'Distance-TARMIZ(KM)' in df.columns:
    fig_final.add_trace(
        go.Histogram(x=df['Distance-TARMIZ(KM)'].dropna(), name='Distance'),
        row=2, col=2
    )

# Score vs Note
if 'Score_Presence_Digitale' in df.columns:
    fig_final.add_trace(
        go.Scatter(x=df['Score_Presence_Digitale'], y=df['Note_Google'], 
                  mode='markers', name='Score vs Note'),
        row=3, col=1
    )

# Ancienneté
if 'Anciennete_Estimee' in df.columns:
    fig_final.add_trace(
        go.Histogram(x=df['Anciennete_Estimee'].dropna(), name='Ancienneté'),
        row=3, col=2
    )

fig_final.update_layout(height=1000, title_text="Dashboard Complet - Secteur Optique", 
                       showlegend=False)
fig_final.show()

# 10. Résumé statistique
print("\n=== RÉSUMÉ STATISTIQUE ===")
print(f"Total entreprises optiques: {len(df)}")
if 'Note_Google' in df.columns:
    print(f"Note Google moyenne: {df['Note_Google'].mean():.2f}")
if 'Distance-TARMIZ(KM)' in df.columns:
    print(f"Distance moyenne de TARMIZ: {df['Distance-TARMIZ(KM)'].mean():.1f} km")
if 'Score_Presence_Digitale' in df.columns:
    print(f"Score présence digitale moyen: {df['Score_Presence_Digitale'].mean():.1f}")

# Ville avec le plus d'optiques
top_city = df['Ville'].value_counts().index[0]
top_city_count = df['Ville'].value_counts().iloc[0]
print(f"Ville leader: {top_city} ({top_city_count} optiques)")

print("\n✅ Toutes les visualisations ont été générées!")
