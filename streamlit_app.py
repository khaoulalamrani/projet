import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import numpy as np
from datetime import datetime

# ----------------------------
# CONFIGURATION DE LA PAGE
# ----------------------------
st.set_page_config(
    page_title="Dashboard Optiques",
    page_icon="👓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------
# CSS PERSONNALISÉ AMÉLIORÉ
# ----------------------------
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    /* Global Styles */
    .main {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3c72, #2a5298);
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        border: 1px solid rgba(255,255,255,0.2);
    }
    
    .main-header h1 {
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 1rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem 1.5rem;
        border-radius: 15px;
        text-align: center;
        color: white;
        box-shadow: 0 8px 32px rgba(31, 38, 135, 0.37);
        backdrop-filter: blur(4px);
        border: 1px solid rgba(255, 255, 255, 0.18);
        margin-bottom: 1.5rem;
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 40px rgba(31, 38, 135, 0.5);
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
    }
    
    .metric-label {
        font-size: 1rem;
        opacity: 0.9;
        font-weight: 500;
    }
    
    .section-header {
        background: linear-gradient(90deg, #667eea, #764ba2);
        padding: 1rem 2rem;
        border-radius: 10px;
        color: white;
        margin: 2rem 0 1rem 0;
        font-weight: 600;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    /* Graphique containers */
    .chart-container {
        background: white;
        border-radius: 15px;
        padding: 1.5rem;
        box-shadow: 0 8px 25px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
        border: 1px solid rgba(255,255,255,0.2);
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 60px;
        padding: 0px 24px;
        background-color: rgba(255,255,255,0.1);
        border-radius: 10px;
        color: white;
        font-weight: 600;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: white;
        color: #667eea;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# TITRE PRINCIPAL
# ----------------------------
st.markdown("""
<div class="main-header">
    <h1>👓 Dashboard Secteur Optique</h1>
    <p style="font-size: 1.2rem; margin: 0;">Analyse Complète et Interactive du Marché Optique</p>
</div>
""", unsafe_allow_html=True)

# ----------------------------
# CHARGEMENT DES DONNÉES
# ----------------------------
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('OPTIQUESS.xlsx', engine='openpyxl')
        
        # Nettoyage des colonnes numériques
        numeric_cols = ["Note_Google", "Nb_Avis_Google", "Score_Presence_Digitale",
                       "Distance-TARMIZ(KM)", "Anciennete_Estimee"]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        
        return df
    except Exception as e:
        st.error(f"Erreur de chargement: {e}")
        return None

df = load_data()

if df is not None:
    # ----------------------------
    # SIDEBAR AVEC FILTRES AMÉLIORÉ
    # ----------------------------
    with st.sidebar:
        st.markdown("## 🔍 Filtres et Options")
        
        # Filtre par ville avec style
        st.markdown("### 🏙️ Localisation")
        cities = ['Toutes'] + sorted(df['Ville'].dropna().unique().tolist())
        selected_city = st.selectbox("Filtrer par ville", cities, key="city_filter")
        
        # Filtre par note
        if 'Note_Google' in df.columns:
            st.markdown("### ⭐ Performance")
            note_range = st.slider(
                "Plage de notes Google",
                float(df['Note_Google'].min()),
                float(df['Note_Google'].max()),
                (float(df['Note_Google'].min()), float(df['Note_Google'].max())),
                step=0.1,
                key="note_range"
            )
        
        # Filtre par distance
        if 'Distance-TARMIZ(KM)' in df.columns:
            st.markdown("### 📍 Proximité")
            max_distance = st.slider(
                "Distance max de TARMIZ (km)",
                0.0,
                float(df['Distance-TARMIZ(KM)'].max()),
                float(df['Distance-TARMIZ(KM)'].max()),
                key="distance_filter"
            )
        
        st.markdown("### 🎨 Apparence")
        color_theme = st.selectbox(
            "Thème de couleurs",
            ["Viridis", "Plasma", "Inferno", "Magma", "Cividis", "Turbo"],
            index=0
        )
        
        # Nouvelles options
        chart_height = st.slider("Hauteur des graphiques", 400, 800, 600)
        show_animations = st.checkbox("Animations", True)
    
    # Application des filtres
    df_filtered = df.copy()
    
    if selected_city != 'Toutes':
        df_filtered = df_filtered[df_filtered['Ville'] == selected_city]
    
    if 'Note_Google' in df.columns:
        df_filtered = df_filtered[
            (df_filtered['Note_Google'] >= note_range[0]) & 
            (df_filtered['Note_Google'] <= note_range[1])
        ]
    
    if 'Distance-TARMIZ(KM)' in df.columns:
        df_filtered = df_filtered[df_filtered['Distance-TARMIZ(KM)'] <= max_distance]
    
    # ----------------------------
    # MÉTRIQUES PRINCIPALES AMÉLIORÉES
    # ----------------------------
    st.markdown("## 📊 Métriques Clés")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        total_optiques = len(df_filtered)
        delta_total = len(df_filtered) - len(df)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_optiques:,}</div>
            <div class="metric-label">📊 Total Optiques</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Note_Google' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{avg_note:.2f}/5</div>
                <div class="metric-label">⭐ Note Moyenne</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        if 'Site web' in df.columns:
            web_presence = (df_filtered['Site web'].notna().sum() / len(df_filtered)) * 100
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{web_presence:.1f}%</div>
                <div class="metric-label">🌐 Présence Web</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col4:
        if 'Distance-TARMIZ(KM)' in df.columns:
            avg_distance = df_filtered['Distance-TARMIZ(KM)'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{avg_distance:.1f} km</div>
                <div class="metric-label">📍 Distance Moy.</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col5:
        if 'Score_Presence_Digitale' in df.columns:
            avg_digital = df_filtered['Score_Presence_Digitale'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{avg_digital:.0f}/100</div>
                <div class="metric-label">📱 Score Digital</div>
            </div>
            """, unsafe_allow_html=True)
    
    # ----------------------------
    # ONGLETS PRINCIPAUX AMÉLIORÉS
    # ----------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🌍 Géographie", 
        "⭐ Performance", 
        "📱 Digital", 
        "📊 Analytics", 
        "📋 Données"
    ])
    
    with tab1:
        st.markdown('<div class="section-header"><h3>🌍 Analyse Géographique Complète</h3></div>', 
                   unsafe_allow_html=True)
        
        # Carte géographique grande et améliorée
        if {"Latitude","Longitude"}.issubset(df.columns):
            geo_df = df_filtered.dropna(subset=["Latitude","Longitude"])
            if len(geo_df) > 0:
                fig_map = px.scatter_mapbox(
                    geo_df,
                    lat="Latitude",
                    lon="Longitude",
                    hover_name="Nom",
                    hover_data=["Ville","Note_Google","Nb_Avis_Google"],
                    color="Note_Google",
                    size="Nb_Avis_Google",
                    color_continuous_scale=color_theme,
                    mapbox_style="open-street-map",
                    height=chart_height + 100,
                    title="🗺️ Répartition Géographique Interactive des Optiques",
                    size_max=20
                )
                fig_map.update_layout(
                    title_font_size=20,
                    title_x=0.5,
                    margin=dict(t=80, b=20, l=20, r=20),
                    paper_bgcolor='white',
                    font=dict(size=12)
                )
                fig_map.update_coloraxes(colorbar_title="Note Google")
                st.plotly_chart(fig_map, use_container_width=True)
        
        # Analyses géographiques en colonnes
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # Top villes avec graphique horizontal amélioré
            st.markdown("### 🏆 Top 15 Villes par Nombre d'Optiques")
            top_cities = df_filtered["Ville"].value_counts().head(15)
            
            fig_cities = go.Figure(data=[
                go.Bar(
                    y=top_cities.index[::-1],  # Inverser pour avoir le plus grand en haut
                    x=top_cities.values[::-1],
                    orientation='h',
                    marker=dict(
                        color=top_cities.values[::-1],
                        colorscale=color_theme,
                        showscale=True,
                        colorbar=dict(title="Nombre d'optiques")
                    ),
                    text=top_cities.values[::-1],
                    textposition='outside',
                    hovertemplate="<b>%{y}</b><br>Optiques: %{x}<extra></extra>"
                )
            ])
            
            fig_cities.update_layout(
                title="Distribution par Ville",
                height=chart_height,
                xaxis_title="Nombre d'optiques",
                yaxis_title="Ville",
                showlegend=False,
                paper_bgcolor='white',
                plot_bgcolor='white',
                title_font_size=16,
                font=dict(size=11)
            )
            fig_cities.update_xaxes(showgrid=True, gridcolor='lightgray')
            fig_cities.update_yaxes(showgrid=False)
            st.plotly_chart(fig_cities, use_container_width=True)
        
        with col2:
            # Analyse de performance par ville
            st.markdown("### 📊 Performance Moyenne par Ville (Top 10)")
            if 'Note_Google' in df.columns:
                city_performance = df_filtered.groupby('Ville').agg({
                    'Note_Google': 'mean',
                    'Nb_Avis_Google': 'mean',
                    'Nom': 'count'
                }).rename(columns={'Nom': 'Nombre'})
                
                # Filtrer les villes avec au moins 2 optiques
                city_performance = city_performance[city_performance['Nombre'] >= 2]
                top_perf_cities = city_performance.nlargest(10, 'Note_Google')
                
                fig_perf = go.Figure()
                
                fig_perf.add_trace(go.Scatter(
                    x=top_perf_cities['Note_Google'],
                    y=top_perf_cities['Nb_Avis_Google'],
                    mode='markers+text',
                    text=top_perf_cities.index,
                    textposition="top center",
                    marker=dict(
                        size=top_perf_cities['Nombre']*5,
                        color=top_perf_cities['Note_Google'],
                        colorscale=color_theme,
                        showscale=True,
                        colorbar=dict(title="Note Google"),
                        line=dict(width=2, color='white')
                    ),
                    hovertemplate="<b>%{text}</b><br>" +
                                 "Note: %{x:.2f}<br>" +
                                 "Avis moyens: %{y:.0f}<br>" +
                                 "Nombre d'optiques: %{marker.size}<extra></extra>"
                ))
                
                fig_perf.update_layout(
                    title="Performance vs Popularité par Ville",
                    xaxis_title="Note Google Moyenne",
                    yaxis_title="Nombre d'Avis Moyen",
                    height=chart_height,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font_size=16
                )
                fig_perf.update_xaxes(showgrid=True, gridcolor='lightgray')
                fig_perf.update_yaxes(showgrid=True, gridcolor='lightgray')
                st.plotly_chart(fig_perf, use_container_width=True)
            
            # Stats géographiques améliorées
            st.markdown("### 📈 Statistiques Géographiques")
            geo_stats = {
                "🏙️ Villes couvertes": df_filtered['Ville'].nunique(),
                "🎯 Ville dominante": df_filtered['Ville'].value_counts().index[0],
                "📍 Concentration": f"{(df_filtered['Ville'].value_counts().iloc[0] / len(df_filtered)) * 100:.1f}%"
            }
            
            for label, value in geo_stats.items():
                st.metric(label, value)
    
    with tab2:
        st.markdown('<div class="section-header"><h3>⭐ Analyse de Performance Détaillée</h3></div>', 
                   unsafe_allow_html=True)
        
        if 'Note_Google' in df.columns:
            # Grand graphique de distribution des notes
            col1, col2 = st.columns([2, 1])
            
            with col1:
                fig_notes = go.Figure()
                
                fig_notes.add_trace(go.Histogram(
                    x=df_filtered["Note_Google"].dropna(),
                    nbinsx=30,
                    marker=dict(
                        color=df_filtered["Note_Google"].dropna(),
                        colorscale=color_theme,
                        line=dict(width=1, color='white')
                    ),
                    hovertemplate="Note: %{x}<br>Nombre: %{y}<extra></extra>"
                ))
                
                # Ajouter des lignes pour les seuils importants
                fig_notes.add_vline(x=4.0, line_dash="dash", line_color="red", 
                                   annotation_text="Seuil Excellence (4.0)")
                fig_notes.add_vline(x=4.5, line_dash="dash", line_color="green", 
                                   annotation_text="Seuil Premium (4.5)")
                
                fig_notes.update_layout(
                    title="📊 Distribution des Notes Google (Analyse Détaillée)",
                    xaxis_title="Note Google",
                    yaxis_title="Nombre d'optiques",
                    height=chart_height,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font_size=18,
                    showlegend=False
                )
                fig_notes.update_xaxes(showgrid=True, gridcolor='lightgray')
                fig_notes.update_yaxes(showgrid=True, gridcolor='lightgray')
                st.plotly_chart(fig_notes, use_container_width=True)
            
            with col2:
                # Métriques de performance
                st.markdown("### 🎯 Métriques Performance")
                
                excellent = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
                good = len(df_filtered[df_filtered['Note_Google'] >= 4.0])
                average = len(df_filtered[df_filtered['Note_Google'] >= 3.5])
                
                perf_data = pd.DataFrame({
                    'Catégorie': ['Excellent\n(≥4.5)', 'Bon\n(≥4.0)', 'Moyen\n(≥3.5)', 'Faible\n(<3.5)'],
                    'Nombre': [excellent, good-excellent, average-good, len(df_filtered)-average],
                    'Pourcentage': [
                        excellent/len(df_filtered)*100,
                        (good-excellent)/len(df_filtered)*100,
                        (average-good)/len(df_filtered)*100,
                        (len(df_filtered)-average)/len(df_filtered)*100
                    ]
                })
                
                fig_perf_pie = px.pie(
                    perf_data, 
                    values='Nombre', 
                    names='Catégorie',
                    title="Répartition Performance",
                    color_discrete_sequence=px.colors.qualitative.Set3,
                    hole=0.4
                )
                fig_perf_pie.update_layout(height=400)
                fig_perf_pie.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_perf_pie, use_container_width=True)
            
            # Analyse corrélation Notes vs Avis (grand graphique)
            st.markdown("### 🔍 Relation Performance vs Popularité")
            
            fig_scatter = go.Figure()
            
            # Ajouter les points avec couleur selon la ville
            cities_for_color = df_filtered['Ville'].value_counts().head(10).index
            df_scatter = df_filtered[df_filtered['Ville'].isin(cities_for_color)]
            
            for city in cities_for_color:
                city_data = df_scatter[df_scatter['Ville'] == city]
                fig_scatter.add_trace(go.Scatter(
                    x=city_data["Note_Google"],
                    y=city_data["Nb_Avis_Google"],
                    mode='markers',
                    name=city,
                    marker=dict(
                        size=8,
                        line=dict(width=1, color='white'),
                        opacity=0.8
                    ),
                    hovertemplate="<b>%{hovertext}</b><br>" +
                                 "Note: %{x}<br>" +
                                 "Avis: %{y}<br>" +
                                 "Ville: " + city + "<extra></extra>",
                    hovertext=city_data["Nom"]
                ))
            
            # Ajouter ligne de tendance
            if len(df_filtered) > 1:
                z = np.polyfit(df_filtered["Note_Google"].dropna(), 
                              df_filtered["Nb_Avis_Google"].dropna(), 1)
                p = np.poly1d(z)
                x_trend = np.linspace(df_filtered["Note_Google"].min(), 
                                     df_filtered["Note_Google"].max(), 100)
                fig_scatter.add_trace(go.Scatter(
                    x=x_trend, 
                    y=p(x_trend),
                    mode='lines',
                    name='Tendance',
                    line=dict(color='red', dash='dash', width=2)
                ))
            
            fig_scatter.update_layout(
                title="📈 Corrélation Notes Google vs Nombre d'Avis par Ville",
                xaxis_title="Note Google",
                yaxis_title="Nombre d'Avis",
                height=chart_height + 50,
                paper_bgcolor='white',
                plot_bgcolor='white',
                title_font_size=18,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            fig_scatter.update_xaxes(showgrid=True, gridcolor='lightgray')
            fig_scatter.update_yaxes(showgrid=True, gridcolor='lightgray')
            st.plotly_chart(fig_scatter, use_container_width=True)
    
    with tab3:
        st.markdown('<div class="section-header"><h3>📱 Présence Digitale Avancée</h3></div>', 
                   unsafe_allow_html=True)
        
        # Analyse présence digitale avec graphiques plus grands
        digital_cols = ["Site web", "Réseaux sociaux", "Email"]
        digital_data = []
        
        for col in digital_cols:
            if col in df.columns:
                count = df_filtered[col].notna().sum()
                percentage = (count / len(df_filtered)) * 100
                digital_data.append({
                    'Canal': col.replace('_', ' ').title(),
                    'Nombre': count,
                    'Pourcentage': percentage,
                    'Manquant': len(df_filtered) - count
                })
        
        if digital_data:
            col1, col2 = st.columns([1, 1])
            
            with col1:
                # Graphique en barres amélioré
                digital_df = pd.DataFrame(digital_data)
                
                fig_digital = go.Figure()
                
                fig_digital.add_trace(go.Bar(
                    name='Présent',
                    x=digital_df['Canal'],
                    y=digital_df['Pourcentage'],
                    marker_color=px.colors.qualitative.Set2[0],
                    text=[f"{p:.1f}%" for p in digital_df['Pourcentage']],
                    textposition='outside'
                ))
                
                fig_digital.add_trace(go.Bar(
                    name='Absent',
                    x=digital_df['Canal'],
                    y=[100-p for p in digital_df['Pourcentage']],
                    marker_color=px.colors.qualitative.Set2[1],
                    text=[f"{100-p:.1f}%" for p in digital_df['Pourcentage']],
                    textposition='inside'
                ))
                
                fig_digital.update_layout(
                    title="📊 Taux de Présence Digitale par Canal",
                    xaxis_title="Canal Digital",
                    yaxis_title="Pourcentage (%)",
                    height=chart_height,
                    barmode='stack',
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    title_font_size=16,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_digital, use_container_width=True)
            
            with col2:
                # Graphique radar pour la maturité digitale
                if "Score_Presence_Digitale" in df.columns:
                    st.markdown("### 🎯 Maturité Digitale")
                    
                    # Créer des segments de maturité
                    digital_segments = {
                        'Débutant (0-30)': len(df_filtered[df_filtered['Score_Presence_Digitale'] <= 30]),
                        'Intermédiaire (31-60)': len(df_filtered[(df_filtered['Score_Presence_Digitale'] > 30) & 
                                                                 (df_filtered['Score_Presence_Digitale'] <= 60)]),
                        'Avancé (61-80)': len(df_filtered[(df_filtered['Score_Presence_Digitale'] > 60) & 
                                                          (df_filtered['Score_Presence_Digitale'] <= 80)]),
                        'Expert (81-100)': len(df_filtered[df_filtered['Score_Presence_Digitale'] > 80])
                    }
                    
                    fig_maturity = go.Figure(data=[
                        go.Bar(
                            x=list(digital_segments.keys()),
                            y=list(digital_segments.values()),
                            marker=dict(
                                color=list(digital_segments.values()),
                                colorscale=color_theme,
                                showscale=True,
                                colorbar=dict(title="Nombre")
                            ),
                            text=list(digital_segments.values()),
                            textposition='outside'
                        )
                    ])
                    
                    fig_maturity.update_layout(
                        title="🚀 Segments de Maturité Digitale",
                        xaxis_title="Niveau de Maturité",
                        yaxis_title="Nombre d'optiques",
                        height=chart_height,
                        paper_bgcolor='white',
                        plot_bgcolor='white',
                        title_font_size=16
                    )
                    fig_maturity.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_maturity, use_container_width=True)
        
        # Analyse corrélation digital vs performance
        if all(col in df.columns for col in ["Score_Presence_Digitale", "Note_Google"]):
            st.markdown("### 🔗 Impact Digital sur la Performance")
            
            fig_impact = go.Figure()
            
            # Créer des bulles par ville
            top_cities = df_filtered['Ville'].value_counts().head(8).index
            colors = px.colors.qualitative.Set1
            
            for i, city in enumerate(top_cities):
                city_data = df_filtered[df_filtered['Ville'] == city]
                fig_impact.add_trace(go.Scatter(
                    x=city_data["Score_Presence_Digitale"],
                    y=city_data["Note_Google"],
                    mode='markers',
                    name=city,
                    marker=dict(
                        size=city_data["Nb_Avis_Google"]/3,  # Taille proportionnelle aux avis
                        color=colors[i % len(colors)],
                        line=dict(width=2, color='white'),
                        opacity=0.8
                    ),
                    hovertemplate="<b>%{hovertext}</b><br>" +
                                 "Score Digital: %{x}<br>" +
                                 "Note Google: %{y}<br>" +
                                 "Ville: " + city + "<extra></extra>",
                    hovertext=city_data["Nom"]
                ))
            
            fig_impact.update_layout(
                title="🎯 Corrélation Score Digital vs Performance Google",
                xaxis_title="Score Présence Digitale",
                yaxis_title="Note Google",
                height=chart_height,
                paper_bgcolor='white',
                plot_bgcolor='white',
                title_font_size=18,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            fig_impact.update_xaxes(showgrid=True, gridcolor='lightgray')
            fig_impact.update_yaxes(showgrid=True, gridcolor='lightgray')
            st.plotly_chart(fig_impact, use_container_width=True)
    
    with tab4:
        st.markdown('<div class="section-header"><h3>📊 Analytics Avancés et Insights</h3></div>', 
                   unsafe_allow_html=True)
        
        # Matrice de corrélation grande et interactive
        numeric_columns = df_filtered.select_dtypes(include=[np.number]).columns
        if len(numeric_columns) > 1:
            st.markdown("### 🔗 Matrice de Corrélation Interactive")
            
            correlation_matrix = df_filtered[numeric_columns].corr()
            
            # Créer une heatmap personnalisée
            fig_corr = go.Figure(data=go.Heatmap(
                z=correlation_matrix.values,
                x=correlation_matrix.columns,
                y=correlation_matrix.columns,
                colorscale=color_theme,
                text=np.round(correlation_matrix.values, 2),
                texttemplate="%{text}",
                textfont={"size": 12},
                hoverongaps=False,
                hovertemplate="<b>%{y} vs %{x}</b><br>Corrélation: %{z:.3f}<extra></extra>"
            ))
            
            fig_corr.update_layout(
                title="🎯 Matrice de Corrélation des Variables Numériques",
                height=chart_height + 100,
                paper_bgcolor='white',
                title_font_size=18,
                xaxis=dict(tickangle=45),
                yaxis=dict(tickangle=0)
            )
            st.plotly_chart(fig_corr, use_container_width=True)
        
        # Analyses multi-dimensionnelles
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # Analyse par distance de TARMIZ
            if "Distance-TARMIZ(KM)" in df.columns:
                st.markdown("### 📍 Analyse par Proximité TARMIZ")
                
                # Créer des segments de distance
                df_filtered['Distance_Segment'] = pd.cut(
                    df_filtered['Distance-TARMIZ(KM)'],
                    bins=[0, 5, 10, 20, 50, float('inf')],
                    labels=['Très proche\n(0-5km)', 'Proche\n(5-10km)', 
                           'Moyen\n(10-20km)', 'Loin\n(20-50km)', 'Très loin\n(>50km)']
                )
                
                distance_analysis = df_filtered.groupby('Distance_Segment').agg({
                    'Note_Google': 'mean',
                    'Nb_Avis_Google': 'mean',
                    'Score_Presence_Digitale': 'mean',
                    'Nom': 'count'
                }).rename(columns={'Nom': 'Nombre'})
                
                fig_distance = go.Figure()
                
                fig_distance.add_trace(go.Bar(
                    name='Note Google',
                    x=distance_analysis.index,
                    y=distance_analysis['Note_Google'],
                    yaxis='y',
                    offsetgroup=1,
                    marker_color=px.colors.qualitative.Set1[0]
                ))
                
                fig_distance.add_trace(go.Scatter(
                    name='Score Digital',
                    x=distance_analysis.index,
                    y=distance_analysis['Score_Presence_Digitale'],
                    yaxis='y2',
                    mode='lines+markers',
                    marker=dict(size=8),
                    line=dict(width=3, color=px.colors.qualitative.Set1[1])
                ))
                
                fig_distance.update_layout(
                    title="Performance vs Distance de TARMIZ",
                    xaxis_title="Segment de Distance",
                    height=chart_height,
                    paper_bgcolor='white',
                    plot_bgcolor='white',
                    yaxis=dict(
                        title="Note Google",
                        side="left"
                    ),
                    yaxis2=dict(
                        title="Score Digital",
                        side="right",
                        overlaying="y"
                    ),
                    legend=dict(x=0.7, y=1)
                )
                st.plotly_chart(fig_distance, use_container_width=True)
        
        with col2:
            # Analyse temporelle si ancienneté disponible
            if "Anciennete_Estimee" in df.columns:
                st.markdown("### 📅 Analyse Temporelle")
                
                # Créer des segments d'ancienneté
                df_filtered['Anciennete_Segment'] = pd.cut(
                    df_filtered['Anciennete_Estimee'],
                    bins=[0, 5, 10, 20, float('inf')],
                    labels=['Récent\n(0-5 ans)', 'Établi\n(5-10 ans)', 
                           'Mature\n(10-20 ans)', 'Vétéran\n(>20 ans)']
                )
                
                age_analysis = df_filtered.groupby('Anciennete_Segment').agg({
                    'Note_Google': 'mean',
                    'Score_Presence_Digitale': 'mean',
                    'Nom': 'count'
                }).rename(columns={'Nom': 'Nombre'})
                
                fig_age = go.Figure(data=[
                    go.Scatterpolar(
                        r=age_analysis['Note_Google'],
                        theta=age_analysis.index,
                        fill='toself',
                        name='Note Google',
                        line_color=px.colors.qualitative.Set1[0]
                    ),
                    go.Scatterpolar(
                        r=age_analysis['Score_Presence_Digitale']/20,  # Normaliser pour le radar
                        theta=age_analysis.index,
                        fill='toself',
                        name='Score Digital (/20)',
                        line_color=px.colors.qualitative.Set1[1],
                        opacity=0.7
                    )
                ])
                
                fig_age.update_layout(
                    polar=dict(
                        radialaxis=dict(
                            visible=True,
                            range=[0, 5]
                        )),
                    title="🕐 Performance par Ancienneté",
                    height=chart_height,
                    paper_bgcolor='white'
                )
                st.plotly_chart(fig_age, use_container_width=True)
        
        # Analyse de benchmarking avancée
        st.markdown("### 🎯 Benchmarking Complet")
        
        benchmark_col1, benchmark_col2, benchmark_col3, benchmark_col4 = st.columns(4)
        
        with benchmark_col1:
            if 'Note_Google' in df.columns:
                percentiles = df_filtered['Note_Google'].quantile([0.25, 0.5, 0.75, 0.9]).round(2)
                st.markdown("**📊 Notes Google**")
                st.metric("Top 10%", f"{percentiles[0.9]}+")
                st.metric("Top 25%", f"{percentiles[0.75]}+")
                st.metric("Médiane", f"{percentiles[0.5]}")
        
        with benchmark_col2:
            if 'Nb_Avis_Google' in df.columns:
                avis_percentiles = df_filtered['Nb_Avis_Google'].quantile([0.25, 0.5, 0.75, 0.9]).round(0)
                st.markdown("**💬 Nombre d'Avis**")
                st.metric("Top 10%", f"{int(avis_percentiles[0.9])}+")
                st.metric("Top 25%", f"{int(avis_percentiles[0.75])}+")
                st.metric("Médiane", f"{int(avis_percentiles[0.5])}")
        
        with benchmark_col3:
            if 'Score_Presence_Digitale' in df.columns:
                digital_percentiles = df_filtered['Score_Presence_Digitale'].quantile([0.25, 0.5, 0.75, 0.9]).round(0)
                st.markdown("**📱 Score Digital**")
                st.metric("Top 10%", f"{int(digital_percentiles[0.9])}+")
                st.metric("Top 25%", f"{int(digital_percentiles[0.75])}+")
                st.metric("Médiane", f"{int(digital_percentiles[0.5])}")
        
        with benchmark_col4:
            st.markdown("**🏆 Champions**")
            if 'Note_Google' in df.columns:
                best_rated = df_filtered.nlargest(1, 'Note_Google')['Nom'].iloc[0] if len(df_filtered) > 0 else "N/A"
                st.metric("Meilleure Note", best_rated[:15] + "..." if len(best_rated) > 15 else best_rated)
            
            if 'Nb_Avis_Google' in df.columns:
                most_reviewed = df_filtered.nlargest(1, 'Nb_Avis_Google')['Nom'].iloc[0] if len(df_filtered) > 0 else "N/A"
                st.metric("Plus d'Avis", most_reviewed[:15] + "..." if len(most_reviewed) > 15 else most_reviewed)
    
    with tab5:
        st.markdown('<div class="section-header"><h3>📋 Données Détaillées et Export</h3></div>', 
                   unsafe_allow_html=True)
        
        # Options d'affichage améliorées
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            show_all = st.checkbox("Toutes les colonnes", False)
        
        with col2:
            if not show_all:
                default_cols = ['Nom', 'Ville', 'Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale']
                available_cols = [col for col in default_cols if col in df_filtered.columns]
                display_cols = st.multiselect(
                    "Colonnes à afficher:",
                    df_filtered.columns.tolist(),
                    default=available_cols
                )
            else:
                display_cols = df_filtered.columns.tolist()
        
        with col3:
            sort_options = ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale', 'Distance-TARMIZ(KM)']
            available_sort = [col for col in sort_options if col in df_filtered.columns]
            sort_by = st.selectbox("Trier par:", available_sort + ['Nom', 'Ville'])
        
        with col4:
            ascending = st.checkbox("Tri croissant", False)
        
        # Affichage du tableau avec style
        if display_cols:
            df_display = df_filtered[display_cols].copy()
            
            if sort_by in df_display.columns:
                df_display = df_display.sort_values(sort_by, ascending=ascending)
            
            # Mise en forme des données numériques
            numeric_cols_display = df_display.select_dtypes(include=[np.number]).columns
            for col in numeric_cols_display:
                if 'Note' in col:
                    df_display[col] = df_display[col].round(2)
                elif 'Score' in col or 'Distance' in col:
                    df_display[col] = df_display[col].round(1)
            
            st.markdown("### 📊 Tableau de Données")
            st.dataframe(
                df_display,
                use_container_width=True,
                height=500
            )
            
            # Statistiques du tableau améliorées
            st.markdown("### 📈 Statistiques du Tableau")
            
            stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
            
            with stats_col1:
                st.metric("📊 Lignes", f"{len(df_display):,}")
                st.metric("📋 Colonnes", len(display_cols))
            
            with stats_col2:
                if 'Note_Google' in df_display.columns:
                    avg_note = df_display['Note_Google'].mean()
                    st.metric("⭐ Note Moy.", f"{avg_note:.2f}")
                if 'Nb_Avis_Google' in df_display.columns:
                    total_avis = df_display['Nb_Avis_Google'].sum()
                    st.metric("💬 Total Avis", f"{int(total_avis):,}")
            
            with stats_col3:
                if 'Score_Presence_Digitale' in df_display.columns:
                    avg_digital = df_display['Score_Presence_Digitale'].mean()
                    st.metric("📱 Score Digital Moy.", f"{avg_digital:.0f}")
                villes_uniques = df_display['Ville'].nunique() if 'Ville' in df_display.columns else 0
                st.metric("🏙️ Villes", villes_uniques)
            
            with stats_col4:
                if 'Distance-TARMIZ(KM)' in df_display.columns:
                    avg_distance = df_display['Distance-TARMIZ(KM)'].mean()
                    st.metric("📍 Distance Moy.", f"{avg_distance:.1f} km")
                # Pourcentage du dataset total
                percentage_shown = (len(df_display) / len(df)) * 100
                st.metric("📊 % Dataset", f"{percentage_shown:.1f}%")
        
        # Export amélioré
        st.markdown("### 📥 Options d'Export")
        
        export_col1, export_col2, export_col3 = st.columns(3)
        
        with export_col1:
            if st.button("📊 Exporter Données Filtrées", use_container_width=True):
                csv = df_filtered.to_csv(index=False, encoding='utf-8')
                st.download_button(
                    label="💾 Télécharger CSV Filtré",
                    data=csv,
                    file_name=f"optiques_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        with export_col2:
            if display_cols and st.button("📋 Exporter Sélection", use_container_width=True):
                csv_display = df_display.to_csv(index=False, encoding='utf-8')
                st.download_button(
                    label="💾 Télécharger Sélection",
                    data=csv_display,
                    file_name=f"optiques_selection_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        
        with export_col3:
            if st.button("📈 Rapport Statistiques", use_container_width=True):
                # Générer un rapport statistique
                rapport = f"""
# Rapport Statistique - Optiques
## Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}

### Données Générales
- Total optiques analysées: {len(df_filtered):,}
- Villes couvertes: {df_filtered['Ville'].nunique()}
- Ville principale: {df_filtered['Ville'].value_counts().index[0]} ({df_filtered['Ville'].value_counts().iloc[0]} optiques)

### Performance Google
- Note moyenne: {df_filtered['Note_Google'].mean():.2f}/5
- Médiane des notes: {df_filtered['Note_Google'].median():.2f}
- Optiques excellentes (≥4.5): {len(df_filtered[df_filtered['Note_Google'] >= 4.5])} ({len(df_filtered[df_filtered['Note_Google'] >= 4.5])/len(df_filtered)*100:.1f}%)

### Présence Digitale
- Score digital moyen: {df_filtered['Score_Presence_Digitale'].mean():.0f}/100
- Avec site web: {df_filtered['Site web'].notna().sum()} ({df_filtered['Site web'].notna().sum()/len(df_filtered)*100:.1f}%)

### Proximité TARMIZ
- Distance moyenne: {df_filtered['Distance-TARMIZ(KM)'].mean():.1f} km
- Optiques à moins de 10km: {len(df_filtered[df_filtered['Distance-TARMIZ(KM)'] <= 10])}
"""
                
                st.download_button(
                    label="💾 Télécharger Rapport",
                    data=rapport,
                    file_name=f"rapport_optiques_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                    mime="text/markdown",
                    use_container_width=True
                )
    
    # ----------------------------
    # FOOTER AVEC RÉSUMÉ EXÉCUTIF AMÉLIORÉ
    # ----------------------------
    st.markdown("---")
    st.markdown("## 📈 Résumé Exécutif")
    
    summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)
    
    with summary_col1:
        st.markdown("### 🏆 Performance")
        if 'Note_Google' in df.columns:
            excellent = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            good = len(df_filtered[df_filtered['Note_Google'] >= 4.0])
            avg_note = df_filtered['Note_Google'].mean()
            
            st.write(f"• **Note moyenne**: {avg_note:.2f}/5")
            st.write(f"• **Excellentes** (≥4.5): {excellent} ({excellent/len(df_filtered)*100:.1f}%)")
            st.write(f"• **Bonnes** (≥4.0): {good} ({good/len(df_filtered)*100:.1f}%)")
            
            # Indicateur de performance
            if avg_note >= 4.2:
                st.success("✅ Performance globalement excellente")
            elif avg_note >= 3.8:
                st.info("ℹ️ Performance satisfaisante")
            else:
                st.warning("⚠️ Marge d'amélioration importante")
    
    with summary_col2:
        st.markdown("### 📱 Maturité Digitale")
        if 'Score_Presence_Digitale' in df.columns:
            avg_digital = df_filtered['Score_Presence_Digitale'].mean()
            high_digital = len(df_filtered[df_filtered['Score_Presence_Digitale'] >= 70])
            
            st.write(f"• **Score moyen**: {avg_digital:.0f}/100")
            st.write(f"• **Maturité élevée** (≥70): {high_digital} ({high_digital/len(df_filtered)*100:.1f}%)")
            
            if 'Site web' in df.columns:
                web_presence = df_filtered['Site web'].notna().sum()
                st.write(f"• **Avec site web**: {web_presence} ({web_presence/len(df_filtered)*100:.1f}%)")
            
            # Indicateur de maturité digitale
            if avg_digital >= 60:
                st.success("✅ Bonne maturité digitale")
            elif avg_digital >= 40:
                st.info("ℹ️ Maturité digitale modérée")
            else:
                st.warning("⚠️ Retard digital significatif")
    
    with summary_col3:
        st.markdown("### 🌍 Couverture Géographique")
        total_cities = df_filtered['Ville'].nunique()
        top_city = df_filtered['Ville'].value_counts().index[0]
        top_count = df_filtered['Ville'].value_counts().iloc[0]
        concentration = (top_count / len(df_filtered)) * 100
        
        st.write(f"• **Villes couvertes**: {total_cities}")
        st.write(f"• **Ville leader**: {top_city}")
        st.write(f"• **Concentration**: {concentration:.1f}% dans {top_city}")
        
        # Indicateur de diversité géographique
        if concentration < 30:
            st.success("✅ Bonne diversité géographique")
        elif concentration < 50:
            st.info("ℹ️ Concentration modérée")
        else:
            st.warning("⚠️ Forte concentration géographique")
    
    with summary_col4:
        st.markdown("### 📊 Insights Clés")
        if all(col in df.columns for col in ['Note_Google', 'Nb_Avis_Google']):
            # Trouver les top performers
            top_performers = df_filtered[
                (df_filtered['Note_Google'] >= 4.0) & 
                (df_filtered['Nb_Avis_Google'] >= 20)
            ]
            
            st.write(f"• **Top performers**: {len(top_performers)} ({len(top_performers)/len(df_filtered)*100:.1f}%)")
            
            # Corrélation performance/avis
            if len(df_filtered) > 10:
                correlation = df_filtered[['Note_Google', 'Nb_Avis_Google']].corr().iloc[0,1]
                st.write(f"• **Corrélation Note/Avis**: {correlation:.2f}")
            
            if 'Distance-TARMIZ(KM)' in df.columns:
                proche_tarmiz = len(df_filtered[df_filtered['Distance-TARMIZ(KM)'] <= 15])
                st.write(f"• **Proche TARMIZ** (<15km): {proche_tarmiz}")
        
        # Recommandation globale
        if 'Note_Google' in df.columns and 'Score_Presence_Digitale' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            avg_digital = df_filtered['Score_Presence_Digitale'].mean()
            
            if avg_note >= 4.0 and avg_digital >= 60:
                st.success("🚀 Secteur performant et mature")
            elif avg_note >= 3.8 and avg_digital >= 40:
                st.info("📈 Secteur en développement")
            else:
                st.warning("🎯 Potentiel d'amélioration")
    
    # Timestamp avec informations sur les filtres
    filters_applied = []
    if selected_city != 'Toutes':
        filters_applied.append(f"Ville: {selected_city}")
    if 'Note_Google' in df.columns and (note_range[0] != df['Note_Google'].min() or note_range[1] != df['Note_Google'].max()):
        filters_applied.append(f"Notes: {note_range[0]:.1f}-{note_range[1]:.1f}")
    if 'Distance-TARMIZ(KM)' in df.columns and max_distance != df['Distance-TARMIZ(KM)'].max():
        filters_applied.append(f"Distance: ≤{max_distance:.0f}km")
    
    filter_text = " | ".join(filters_applied) if filters_applied else "Aucun filtre"
    
    st.markdown(f"""
    ---
    **📊 Dashboard Optiques** • *Dernière mise à jour: {datetime.now().strftime('%d/%m/%Y à %H:%M')}*  
    **🔍 Filtres actifs**: {filter_text} • **📈 Données**: {len(df_filtered):,} optiques sur {len(df):,} total
    """)

else:
    # Page d'erreur améliorée
    st.markdown("""
    <div style="text-align: center; padding: 3rem; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 15px; color: white;">
        <h1>❌ Erreur de Chargement</h1>
        <p style="font-size: 1.2rem;">Impossible de charger le fichier OPTIQUESS.xlsx</p>
        <p>Vérifiez que le fichier existe dans le même répertoire que ce script.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 Instructions:")
    st.markdown("""
    1. Assurez-vous que le fichier `OPTIQUESS.xlsx` est présent
    2. Vérifiez les permissions de lecture du fichier
    3. Confirmez que le fichier n'est pas corrompu
    4. Redémarrez l'application si nécessaire
    """)
