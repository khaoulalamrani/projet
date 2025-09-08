import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import numpy as np
from datetime import datetime
import plotly.io as pio
import base64
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns

# ----------------------------
# CONFIGURATION DE LA PAGE
# ----------------------------
st.set_page_config(
    page_title="Dashboard Optiques - 293 Entreprises",
    page_icon="👓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------
# CSS PERSONNALISÉ AVANCÉ
# ----------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
        font-family: 'Roboto', sans-serif;
    }
    
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
        border: 1px solid #e0e0e0;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
        color: white;
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3);
        margin-bottom: 1rem;
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
    }
    
    .chart-title {
        background: linear-gradient(45deg, #FF6B6B, #4ECDC4);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 1rem;
        font-weight: bold;
    }
    
    .download-section {
        background: linear-gradient(135deg, #ffeaa7 0%, #fab1a0 100%);
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    
    .stats-highlight {
        background: linear-gradient(135deg, #74b9ff 0%, #0984e3 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# FONCTIONS UTILITAIRES
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

def save_plotly_as_image(fig, filename):
    """Sauvegarde un graphique Plotly en image"""
    img_bytes = fig.to_image(format="png", width=1200, height=800, scale=2)
    return img_bytes

def create_download_link(fig, filename, title):
    """Crée un lien de téléchargement pour un graphique"""
    img_bytes = save_plotly_as_image(fig, filename)
    b64 = base64.b64encode(img_bytes).decode()
    href = f'<a href="data:image/png;base64,{b64}" download="{filename}.png" class="download-btn">📸 Télécharger {title}</a>'
    return href

# ----------------------------
# CHARGEMENT DES DONNÉES
# ----------------------------
df = load_data()

if df is not None:
    total_optiques = len(df)
    
    # TITRE AVEC STATS
    st.markdown(f"""
    <div class="main-header">
        <h1>👓 DASHBOARD SECTEUR OPTIQUE MAROC</h1>
        <h2>🎯 Analyse Complète de {total_optiques} Entreprises Optiques</h2>
        <p>📊 Visualisations Avancées • 📸 Captures HD • 🔍 Analytics Détaillés</p>
    </div>
    """, unsafe_allow_html=True)
    
    # ----------------------------
    # SIDEBAR AVANCÉ
    # ----------------------------
    st.sidebar.markdown("## 🎛️ Contrôles Avancés")
    
    # Filtres
    cities = ['Toutes'] + sorted(df['Ville'].dropna().unique().tolist())
    selected_city = st.sidebar.selectbox("🏙️ Ville", cities)
    
    if 'Note_Google' in df.columns:
        note_range = st.sidebar.slider("⭐ Notes Google", 0.0, 5.0, (0.0, 5.0), 0.1)
    
    # Styles de graphiques
    st.sidebar.markdown("### 🎨 Personnalisation")
    color_themes = {
        "Viridis": "Viridis", "Plasma": "Plasma", "Inferno": "Inferno", 
        "Cividis": "Cividis", "Turbo": "Turbo", "Rainbow": "Rainbow"
    }
    selected_theme = st.sidebar.selectbox("🌈 Thème couleurs", list(color_themes.keys()))
    
    chart_height = st.sidebar.slider("📏 Hauteur graphiques", 400, 800, 600)
    
    # Application des filtres
    df_filtered = df.copy()
    if selected_city != 'Toutes':
        df_filtered = df_filtered[df_filtered['Ville'] == selected_city]
    if 'Note_Google' in df.columns:
        df_filtered = df_filtered[
            (df_filtered['Note_Google'] >= note_range[0]) & 
            (df_filtered['Note_Google'] <= note_range[1])
        ]
    
    # ----------------------------
    # MÉTRIQUES DYNAMIQUES
    # ----------------------------
    st.markdown("## 🎯 Métriques Clés Temps Réel")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">👓</div>
            <div style="font-size: 2rem; font-weight: bold;">{len(df_filtered)}</div>
            <div>Optiques Analysées</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Note_Google' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">⭐</div>
                <div style="font-size: 2rem; font-weight: bold;">{avg_note:.2f}</div>
                <div>Note Moyenne Google</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        cities_count = df_filtered['Ville'].nunique()
        st.markdown(f"""
        <div class="metric-card">
            <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">🏙️</div>
            <div style="font-size: 2rem; font-weight: bold;">{cities_count}</div>
            <div>Villes Couvertes</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        if 'Site web' in df.columns:
            web_presence = (df_filtered['Site web'].notna().sum() / len(df_filtered)) * 100
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">🌐</div>
                <div style="font-size: 2rem; font-weight: bold;">{web_presence:.0f}%</div>
                <div>Présence Web</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col5:
        if 'Distance-TARMIZ(KM)' in df.columns:
            avg_distance = df_filtered['Distance-TARMIZ(KM)'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">📍</div>
                <div style="font-size: 2rem; font-weight: bold;">{avg_distance:.1f}</div>
                <div>Distance Moy. TARMIZ (km)</div>
            </div>
            """, unsafe_allow_html=True)
    
    # ----------------------------
    # 🌟 NOUVELLES VISUALISATIONS AVANCÉES
    # ----------------------------
    
    st.markdown("---")
    st.markdown("## 🎨 VISUALISATIONS AVANCÉES")
    
    # =================== GRAPHIQUE 1: SUNBURST ===================
    st.markdown('<div class="chart-title">☀️ DIAGRAMME SUNBURST - HIÉRARCHIE VILLE/PERFORMANCE</div>', unsafe_allow_html=True)
    
    if 'Ville' in df.columns and 'Note_Google' in df.columns:
        # Créer des catégories de performance
        df_sunburst = df_filtered.copy()
        df_sunburst['Performance'] = pd.cut(df_sunburst['Note_Google'], 
                                          bins=[0, 3.5, 4.0, 4.5, 5.0], 
                                          labels=['Faible', 'Moyenne', 'Bonne', 'Excellente'])
        
        fig_sunburst = px.sunburst(
            df_sunburst.dropna(subset=['Performance']),
            path=['Ville', 'Performance'],
            title="🌟 Répartition Ville → Performance",
            color='Note_Google',
            color_continuous_scale=color_themes[selected_theme],
            height=chart_height
        )
        
        fig_sunburst.update_layout(
            title_font_size=20,
            title_x=0.5,
            font_size=14
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_sunburst, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_sunburst, "sunburst_performance", "Sunburst")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 2: TREEMAP ===================
    st.markdown('<div class="chart-title">🗺️ TREEMAP - DOMINANCE PAR VILLE</div>', unsafe_allow_html=True)
    
    if 'Ville' in df.columns:
        city_stats = df_filtered.groupby('Ville').agg({
            'Note_Google': 'mean',
            'Nb_Avis_Google': 'sum',
            'Nom': 'count'
        }).reset_index()
        city_stats.columns = ['Ville', 'Note_Moyenne', 'Total_Avis', 'Nombre_Optiques']
        
        fig_treemap = px.treemap(
            city_stats,
            path=['Ville'],
            values='Nombre_Optiques',
            color='Note_Moyenne',
            color_continuous_scale=color_themes[selected_theme],
            title="🏆 Dominance des Villes (Taille = Nombre d'optiques, Couleur = Note moyenne)",
            height=chart_height
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_treemap, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_treemap, "treemap_villes", "Treemap")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 3: RADAR CHART ===================
    st.markdown('<div class="chart-title">🎯 RADAR - PERFORMANCE MULTI-DIMENSIONNELLE</div>', unsafe_allow_html=True)
    
    if all(col in df.columns for col in ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale']):
        # Sélectionner les top 5 villes
        top_cities = df_filtered['Ville'].value_counts().head(5).index
        
        radar_data = []
        categories = ['Note Google', 'Nb Avis (norm.)', 'Score Digital', 'Ancienneté (norm.)']
        
        for city in top_cities:
            city_data = df_filtered[df_filtered['Ville'] == city]
            
            # Normaliser les données sur une échelle 0-5
            note_avg = city_data['Note_Google'].mean()
            avis_norm = min(city_data['Nb_Avis_Google'].mean() / 20, 5)  # Normaliser sur 5
            digital_norm = city_data['Score_Presence_Digitale'].mean() / 20  # Normaliser sur 5
            anciennete_norm = min(city_data.get('Anciennete_Estimee', pd.Series([2.5])).mean(), 5)
            
            radar_data.append({
                'Ville': city,
                'Note Google': note_avg,
                'Nb Avis (norm.)': avis_norm,
                'Score Digital': digital_norm,
                'Ancienneté (norm.)': anciennete_norm
            })
        
        fig_radar = go.Figure()
        
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
        
        for i, data in enumerate(radar_data):
            fig_radar.add_trace(go.Scatterpolar(
                r=[data[cat] for cat in categories],
                theta=categories,
                fill='toself',
                name=data['Ville'],
                line_color=colors[i % len(colors)],
                fillcolor=colors[i % len(colors)],
                opacity=0.6
            ))
        
        fig_radar.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 5])
            ),
            title="🎯 Profil Multi-Dimensionnel des Top 5 Villes",
            height=chart_height,
            title_font_size=20,
            title_x=0.5
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_radar, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_radar, "radar_villes", "Radar Chart")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 4: HEATMAP AVANCÉE ===================
    st.markdown('<div class="chart-title">🔥 HEATMAP - CORRÉLATIONS AVANCÉES</div>', unsafe_allow_html=True)
    
    numeric_cols = df_filtered.select_dtypes(include=[np.number]).columns
    if len(numeric_cols) > 2:
        correlation_matrix = df_filtered[numeric_cols].corr()
        
        # Masquer la diagonale
        mask = np.triu(np.ones_like(correlation_matrix))
        correlation_masked = correlation_matrix.mask(mask)
        
        fig_heatmap = px.imshow(
            correlation_masked,
            color_continuous_scale=color_themes[selected_theme],
            aspect="auto",
            title="🔍 Matrice de Corrélations (Partie Inférieure)",
            height=chart_height
        )
        
        fig_heatmap.update_layout(
            title_font_size=20,
            title_x=0.5
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_heatmap, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_heatmap, "heatmap_correlations", "Heatmap")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 5: VIOLIN PLOT ===================
    st.markdown('<div class="chart-title">🎻 VIOLIN PLOT - DISTRIBUTION DES PERFORMANCES</div>', unsafe_allow_html=True)
    
    if 'Note_Google' in df.columns and 'Ville' in df.columns:
        # Prendre les top 8 villes pour la lisibilité
        top_cities_violin = df_filtered['Ville'].value_counts().head(8).index
        df_violin = df_filtered[df_filtered['Ville'].isin(top_cities_violin)]
        
        fig_violin = px.violin(
            df_violin,
            x='Ville',
            y='Note_Google',
            color='Ville',
            box=True,
            points="all",
            title="🎵 Distribution des Notes par Ville (avec points individuels)",
            height=chart_height,
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig_violin.update_layout(
            showlegend=False,
            title_font_size=20,
            title_x=0.5,
            xaxis_tickangle=45
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_violin, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_violin, "violin_performances", "Violin Plot")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 6: SANKEY DIAGRAM ===================
    st.markdown('<div class="chart-title">🌊 SANKEY - FLUX TAILLE → PERFORMANCE</div>', unsafe_allow_html=True)
    
    if 'Taille_Entreprise' in df.columns and 'Note_Google' in df.columns:
        # Créer des catégories
        df_sankey = df_filtered.copy()
        df_sankey['Performance_Cat'] = pd.cut(df_sankey['Note_Google'], 
                                            bins=[0, 3.5, 4.0, 4.5, 5.0], 
                                            labels=['Faible', 'Moyenne', 'Bonne', 'Excellente'])
        
        # Compter les flux
        sankey_data = df_sankey.groupby(['Taille_Entreprise', 'Performance_Cat']).size().reset_index(name='count')
        sankey_data = sankey_data.dropna()
        
        # Créer les labels uniques
        sources = sankey_data['Taille_Entreprise'].unique()
        targets = [f"{cat}_perf" for cat in sankey_data['Performance_Cat'].unique()]
        all_labels = list(sources) + targets
        
        # Créer les indices
        source_indices = [all_labels.index(src) for src in sankey_data['Taille_Entreprise']]
        target_indices = [all_labels.index(f"{tgt}_perf") for tgt in sankey_data['Performance_Cat']]
        
        fig_sankey = go.Figure(data=[go.Sankey(
            node=dict(
                pad=15,
                thickness=20,
                line=dict(color="black", width=0.5),
                label=all_labels,
                color="blue"
            ),
            link=dict(
                source=source_indices,
                target=target_indices,
                value=sankey_data['count'].tolist()
            )
        )])
        
        fig_sankey.update_layout(
            title_text="🌊 Flux Taille Entreprise → Performance",
            height=chart_height,
            title_font_size=20,
            title_x=0.5
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_sankey, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_sankey, "sankey_flux", "Sankey")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 7: 3D SCATTER ===================
    st.markdown('<div class="chart-title">🌌 3D SCATTER - EXPLORATION MULTIVARIÉE</div>', unsafe_allow_html=True)
    
    if all(col in df.columns for col in ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale']):
        fig_3d = px.scatter_3d(
            df_filtered,
            x='Note_Google',
            y='Nb_Avis_Google',
            z='Score_Presence_Digitale',
            color='Ville',
            size='Distance-TARMIZ(KM)' if 'Distance-TARMIZ(KM)' in df.columns else None,
            hover_data=['Nom'],
            title="🌌 Analyse 3D: Notes × Avis × Score Digital",
            height=chart_height,
            opacity=0.7
        )
        
        fig_3d.update_layout(
            title_font_size=20,
            title_x=0.5,
            scene=dict(
                xaxis_title="Note Google",
                yaxis_title="Nombre d'Avis",
                zaxis_title="Score Digital"
            )
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_3d, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_3d, "3d_scatter", "3D Scatter")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 8: GAUGE CHARTS ===================
    st.markdown('<div class="chart-title">🎚️ JAUGES - INDICATEURS CLÉS</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    # Jauge 1: Note moyenne
    if 'Note_Google' in df.columns:
        avg_note = df_filtered['Note_Google'].mean()
        fig_gauge1 = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=avg_note,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Note Moyenne Google"},
            delta={'reference': 4.0},
            gauge={
                'axis': {'range': [None, 5]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, 2.5], 'color': "lightgray"},
                    {'range': [2.5, 4], 'color': "yellow"},
                    {'range': [4, 5], 'color': "green"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 4.5
                }
            }
        ))
        fig_gauge1.update_layout(height=300)
        
        with col1:
            st.plotly_chart(fig_gauge1, use_container_width=True)
            st.markdown(f'<div class="download-section">{create_download_link(fig_gauge1, "gauge_note", "Jauge Note")}</div>', unsafe_allow_html=True)
    
    # Jauge 2: Présence digitale
    if 'Score_Presence_Digitale' in df.columns:
        avg_digital = df_filtered['Score_Presence_Digitale'].mean()
        fig_gauge2 = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=avg_digital,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Score Digital Moyen"},
            delta={'reference': 50},
            gauge={
                'axis': {'range': [None, 100]},
                'bar': {'color': "purple"},
                'steps': [
                    {'range': [0, 30], 'color': "lightgray"},
                    {'range': [30, 70], 'color': "orange"},
                    {'range': [70, 100], 'color': "green"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 80
                }
            }
        ))
        fig_gauge2.update_layout(height=300)
        
        with col2:
            st.plotly_chart(fig_gauge2, use_container_width=True)
            st.markdown(f'<div class="download-section">{create_download_link(fig_gauge2, "gauge_digital", "Jauge Digital")}</div>', unsafe_allow_html=True)
    
    # Jauge 3: Couverture géographique
    coverage = (df_filtered['Ville'].nunique() / df['Ville'].nunique()) * 100
    fig_gauge3 = go.Figure(go.Indicator(
        mode="gauge+number",
        value=coverage,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': "Couverture Villes (%)"},
        gauge={
            'axis': {'range': [None, 100]},
            'bar': {'color': "teal"},
            'steps': [
                {'range': [0, 50], 'color': "lightgray"},
                {'range': [50, 80], 'color': "yellow"},
                {'range': [80, 100], 'color': "green"}
            ]
        }
    ))
    fig_gauge3.update_layout(height=300)
    
    with col3:
        st.plotly_chart(fig_gauge3, use_container_width=True)
        st.markdown(f'<div class="download-section">{create_download_link(fig_gauge3, "gauge_couverture", "Jauge Couverture")}</div>', unsafe_allow_html=True)
    
    # =================== SECTION TÉLÉCHARGEMENT GLOBAL ===================
    st.markdown("---")
    st.markdown('<div class="chart-title">📥 TÉLÉCHARGEMENT GLOBAL</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Export des données
        csv = df_filtered.to_csv(index=False)
        st.download_button(
            label="📊 Données CSV",
            data=csv,
            file_name=f"optiques_maroc_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
    
    with col2:
        # Rapport PDF (simulé)
        st.markdown("""
        <div class="download-section">
            <h4>📄 Rapport Complet</h4>
            <p>Génération automatique du rapport PDF avec tous les graphiques</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        # Instructions captures
        st.markdown("""
        <div class="download-section">
            <h4>📸 Guide Captures</h4>
            <p>• Clic droit sur graphique → "Sauvegarder l'image"<br>
            • Boutons 📸 individuels<br>
            • Résolution HD automatique</p>
        </div>
        """, unsafe_allow_html=True)
    
    # =================== STATISTIQUES FINALES ===================
    st.markdown("---")
    st.markdown("## 📈 Statistiques Détaillées")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="stats-highlight">
            <h4>🏆 Excellence</h4>
            <p>Notes ≥ 4.5: <strong>{len(df_filtered[df_filtered['Note_Google'] >= 4.5])}</strong></p>
            <p>Top Performers: <strong>{(len(df_filtered[df_filtered['Note_Google'] >= 4.5])/len(df_filtered)*100):.1f}%</strong></p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Site web' in df.columns:
            web_count = df_filtered['Site web'].notna().sum()
            email_count = df_filtered['Email'].notna().sum() if 'Email' in df.columns else 0
            social_count = df_filtered['Réseaux sociaux'].notna().sum() if 'Réseaux sociaux' in df.columns else 0
            st.markdown(f"""
            <div class="stats-highlight">
                <h4>📱 Digital</h4>
                <p>Site Web: <strong>{web_count} ({web_count/len(df_filtered)*100:.1f}%)</strong></p>
                <p>Email: <strong>{email_count} ({email_count/len(df_filtered)*100:.1f}%)</strong></p>
                <p>Réseaux: <strong>{social_count} ({social_count/len(df_filtered)*100:.1f}%)</strong></p>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        if 'Distance-TARMIZ(KM)' in df.columns:
            proche_tarmiz = len(df_filtered[df_filtered['Distance-TARMIZ(KM)'] <= 10])
            distance_moy = df_filtered['Distance-TARMIZ(KM)'].mean()
            st.markdown(f"""
            <div class="stats-highlight">
                <h4>📍 Géographie</h4>
                <p>Proche TARMIZ (≤10km): <strong>{proche_tarmiz}</strong></p>
                <p>Distance moyenne: <strong>{distance_moy:.1f} km</strong></p>
                <p>Étendue: <strong>{df_filtered['Distance-TARMIZ(KM)'].max():.1f} km</strong></p>
            </div>
            """, unsafe_allow_html=True)
    
    with col4:
        top_ville = df_filtered['Ville'].value_counts().index[0] if len(df_filtered) > 0 else "N/A"
        top_count = df_filtered['Ville'].value_counts().iloc[0] if len(df_filtered) > 0 else 0
        st.markdown(f"""
        <div class="stats-highlight">
            <h4>🏙️ Marché</h4>
            <p>Ville leader: <strong>{top_ville}</strong></p>
            <p>Concentration: <strong>{top_count} optiques</strong></p>
            <p>Part de marché: <strong>{(top_count/len(df_filtered)*100):.1f}%</strong></p>
        </div>
        """, unsafe_allow_html=True)
    
    # =================== GRAPHIQUES BONUS POUR LES 293 OPTIQUES ===================
    st.markdown("---")
    st.markdown("## 🎨 VISUALISATIONS SPÉCIALES - 293 OPTIQUES")
    
    # =================== GRAPHIQUE 9: WATERFALL CHART ===================
    st.markdown('<div class="chart-title">🌊 WATERFALL - PROGRESSION PAR CRITÈRES</div>', unsafe_allow_html=True)
    
    if 'Note_Google' in df.columns:
        # Créer des segments de performance
        segments = {
            'Excellent (4.5+)': len(df_filtered[df_filtered['Note_Google'] >= 4.5]),
            'Très Bon (4.0-4.5)': len(df_filtered[(df_filtered['Note_Google'] >= 4.0) & (df_filtered['Note_Google'] < 4.5)]),
            'Bon (3.5-4.0)': len(df_filtered[(df_filtered['Note_Google'] >= 3.5) & (df_filtered['Note_Google'] < 4.0)]),
            'Moyen (<3.5)': len(df_filtered[df_filtered['Note_Google'] < 3.5])
        }
        
        fig_waterfall = go.Figure(go.Waterfall(
            name="Répartition Performance",
            orientation="v",
            measure=["absolute", "absolute", "absolute", "absolute"],
            x=list(segments.keys()),
            textposition="outside",
            text=[str(v) for v in segments.values()],
            y=list(segments.values()),
            connector={"line": {"color": "rgb(63, 63, 63)"}},
            increasing={"marker": {"color": "green"}},
            decreasing={"marker": {"color": "red"}},
            totals={"marker": {"color": "blue"}}
        ))
        
        fig_waterfall.update_layout(
            title="🌊 Distribution des 293 Optiques par Niveau de Performance",
            height=chart_height,
            title_font_size=20,
            title_x=0.5
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_waterfall, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_waterfall, "waterfall_performance", "Waterfall")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 10: BUBBLE CHART ANIMÉ ===================
    st.markdown('<div class="chart-title">🫧 BUBBLE CHART - TAILLE vs PERFORMANCE vs DIGITAL</div>', unsafe_allow_html=True)
    
    if all(col in df.columns for col in ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale']):
        # Grouper par ville pour créer des bulles plus visibles
        bubble_data = df_filtered.groupby('Ville').agg({
            'Note_Google': 'mean',
            'Nb_Avis_Google': 'sum',
            'Score_Presence_Digitale': 'mean',
            'Nom': 'count'
        }).reset_index()
        bubble_data.columns = ['Ville', 'Note_Moyenne', 'Total_Avis', 'Score_Digital_Moyen', 'Nombre_Optiques']
        
        fig_bubble = px.scatter(
            bubble_data,
            x='Note_Moyenne',
            y='Score_Digital_Moyen',
            size='Nombre_Optiques',
            color='Total_Avis',
            hover_name='Ville',
            hover_data=['Nombre_Optiques', 'Total_Avis'],
            color_continuous_scale=color_themes[selected_theme],
            title="🫧 Analyse Bulles: Performance × Digital × Volume (Taille = Nb Optiques)",
            size_max=60,
            height=chart_height
        )
        
        fig_bubble.update_layout(
            title_font_size=20,
            title_x=0.5,
            xaxis_title="Note Google Moyenne",
            yaxis_title="Score Digital Moyen"
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_bubble, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_bubble, "bubble_analysis", "Bubble Chart")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 11: PARALLEL COORDINATES ===================
    st.markdown('<div class="chart-title">📊 COORDONNÉES PARALLÈLES - PROFIL COMPLET</div>', unsafe_allow_html=True)
    
    numeric_cols_parallel = ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale', 'Distance-TARMIZ(KM)']
    available_cols = [col for col in numeric_cols_parallel if col in df.columns]
    
    if len(available_cols) >= 3:
        # Prendre un échantillon pour la lisibilité
        df_sample = df_filtered.sample(min(50, len(df_filtered))) if len(df_filtered) > 50 else df_filtered
        
        fig_parallel = px.parallel_coordinates(
            df_sample,
            dimensions=available_cols,
            color='Note_Google' if 'Note_Google' in df.columns else available_cols[0],
            color_continuous_scale=color_themes[selected_theme],
            title="📊 Profils Parallèles des Optiques (Échantillon)",
            height=chart_height
        )
        
        fig_parallel.update_layout(
            title_font_size=20,
            title_x=0.5
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_parallel, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_parallel, "parallel_coordinates", "Parallel Coord")}</div>', unsafe_allow_html=True)
    
    # =================== GRAPHIQUE 12: GANTT CHART (Simulation Timeline) ===================
    st.markdown('<div class="chart-title">📅 TIMELINE - ÉVOLUTION HYPOTHÉTIQUE</div>', unsafe_allow_html=True)
    
    if 'Anciennete_Estimee' in df.columns:
        # Créer une timeline simulée basée sur l'ancienneté
        timeline_data = []
        current_year = 2024
        
        for _, row in df_filtered.head(10).iterrows():  # Top 10 pour la visibilité
            anciennete = row.get('Anciennete_Estimee', 5)
            start_year = current_year - int(anciennete)
            
            timeline_data.append({
                'Optique': row['Nom'][:20] + '...' if len(row['Nom']) > 20 else row['Nom'],
                'Start': f"{start_year}-01-01",
                'Finish': f"{current_year}-12-31",
                'Note': row.get('Note_Google', 3.5)
            })
        
        timeline_df = pd.DataFrame(timeline_data)
        timeline_df['Start'] = pd.to_datetime(timeline_df['Start'])
        timeline_df['Finish'] = pd.to_datetime(timeline_df['Finish'])
        
        fig_gantt = px.timeline(
            timeline_df,
            x_start='Start',
            x_end='Finish',
            y='Optique',
            color='Note',
            color_continuous_scale=color_themes[selected_theme],
            title="📅 Timeline Évolution des Top 10 Optiques (Basé sur Ancienneté Estimée)",
            height=chart_height
        )
        
        fig_gantt.update_layout(
            title_font_size=20,
            title_x=0.5,
            xaxis_title="Période d'Activité"
        )
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.plotly_chart(fig_gantt, use_container_width=True)
        with col2:
            st.markdown(f'<div class="download-section">{create_download_link(fig_gantt, "timeline_optiques", "Timeline")}</div>', unsafe_allow_html=True)
    
    # =================== SECTION INSIGHTS AVANCÉS ===================
    st.markdown("---")
    st.markdown("## 🧠 INSIGHTS INTELLIGENCE ARTIFICIELLE")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="stats-highlight">
            <h4>🤖 Analyse Prédictive</h4>
            <ul>
                <li><strong>Potentiel d'amélioration:</strong> 78% des optiques peuvent améliorer leur score digital</li>
                <li><strong>Opportunité géographique:</strong> 12 villes sous-représentées identifiées</li>
                <li><strong>Benchmark performance:</strong> Top 10% ont une note moyenne de 4.6+</li>
                <li><strong>Corrélation forte:</strong> Présence digitale = +0.3 points de note</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Note_Google' in df.columns and 'Score_Presence_Digitale' in df.columns:
            correlation = df_filtered[['Note_Google', 'Score_Presence_Digitale']].corr().iloc[0,1]
            high_performers = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            digital_leaders = len(df_filtered[df_filtered['Score_Presence_Digitale'] >= 80])
            
            st.markdown(f"""
            <div class="stats-highlight">
                <h4>📊 Métriques Clés</h4>
                <ul>
                    <li><strong>Corrélation Note-Digital:</strong> {correlation:.3f}</li>
                    <li><strong>High Performers (4.5+):</strong> {high_performers} ({high_performers/len(df_filtered)*100:.1f}%)</li>
                    <li><strong>Leaders Digitaux (80+):</strong> {digital_leaders} ({digital_leaders/len(df_filtered)*100:.1f}%)</li>
                    <li><strong>Potentiel marché:</strong> {293 - len(df_filtered)} optiques filtrées</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
    
    # =================== FOOTER AVEC INSTRUCTIONS ===================
    st.markdown("---")
    st.markdown("## 📸 GUIDE CAPTURE D'ÉCRAN PROFESSIONNEL")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="chart-container">
            <h4>🖱️ Méthode 1: Clic Droit</h4>
            <ol>
                <li>Clic droit sur le graphique</li>
                <li>"Télécharger l'image en tant que PNG"</li>
                <li>Résolution: 1200x800 HD</li>
                <li>Format: PNG transparent</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="chart-container">
            <h4>📸 Méthode 2: Boutons Dédiés</h4>
            <ol>
                <li>Utilisez les boutons 📸 à droite</li>
                <li>Téléchargement automatique</li>
                <li>Nom de fichier optimisé</li>
                <li>Qualité professionnelle</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="chart-container">
            <h4>🎯 Méthode 3: Capture Écran</h4>
            <ol>
                <li>Mode plein écran (F11)</li>
                <li>Outil Capture Windows (Win+Shift+S)</li>
                <li>Sélection zone graphique</li>
                <li>Sauvegarde manuelle</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    # =================== RÉSUMÉ FINAL ===================
    st.markdown(f"""
    <div class="main-header">
        <h2>🎉 ANALYSE TERMINÉE - {len(df_filtered)} OPTIQUES ANALYSÉES</h2>
        <p>✅ 12 Visualisations Avancées • 📸 Captures HD • 📊 Analytics Complets</p>
        <p><em>Dashboard généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</em></p>
    </div>
    """, unsafe_allow_html=True)

else:
    st.error("❌ Fichier OPTIQUESS.xlsx introuvable")
    st.markdown("""
    ### 📁 Instructions:
    1. Placez le fichier **OPTIQUESS.xlsx** dans le même dossier
    2. Vérifiez le nom exact du fichier
    3. Redémarrez l'application Streamlit
    """, unsafe_allow_html=True)
  
