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
# CSS PERSONNALISÉ
# ----------------------------
st.markdown("""
<style>
    /* Thème principal */
    .main-header {
        background: linear-gradient(90deg, #1f77b4, #ff7f0e);
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 15px 0 rgba(31, 38, 135, 0.37);
        margin-bottom: 1rem;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.8;
    }
    
    .section-header {
        background: linear-gradient(45deg, #667eea, #764ba2);
        padding: 0.5rem 1rem;
        border-radius: 5px;
        color: white;
        margin: 1rem 0;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    .filter-container {
        background: rgba(255, 255, 255, 0.1);
        padding: 1rem;
        border-radius: 10px;
        backdrop-filter: blur(10px);
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# TITRE PRINCIPAL
# ----------------------------
st.markdown("""
<div class="main-header">
    <h1>👓 Dashboard Secteur Optique - Analyse Complète</h1>
    <p>Tableau de bord interactif pour l'analyse du marché optique</p>
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
    # SIDEBAR AVEC FILTRES
    # ----------------------------
    st.sidebar.markdown("## 🔍 Filtres et Options")
    
    # Filtre par ville
    cities = ['Toutes'] + sorted(df['Ville'].dropna().unique().tolist())
    selected_city = st.sidebar.selectbox("🏙️ Filtrer par ville", cities)
    
    # Filtre par note
    if 'Note_Google' in df.columns:
        note_range = st.sidebar.slider(
            "⭐ Plage de notes Google",
            float(df['Note_Google'].min()),
            float(df['Note_Google'].max()),
            (float(df['Note_Google'].min()), float(df['Note_Google'].max())),
            step=0.1
        )
    
    # Filtre par distance
    if 'Distance-TARMIZ(KM)' in df.columns:
        max_distance = st.sidebar.slider(
            "📍 Distance max de TARMIZ (km)",
            0.0,
            float(df['Distance-TARMIZ(KM)'].max()),
            float(df['Distance-TARMIZ(KM)'].max())
        )
    
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
    
    # Options d'affichage
    st.sidebar.markdown("## 🎨 Options d'affichage")
    color_theme = st.sidebar.selectbox(
        "Thème de couleurs",
        ["Viridis", "Plasma", "Inferno", "Magma", "Cividis"]
    )
    
    # ----------------------------
    # MÉTRIQUES PRINCIPALES
    # ----------------------------
    st.markdown("## 📊 Métriques Clés")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        total_optiques = len(df_filtered)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_optiques}</div>
            <div class="metric-label">📊 Total Optiques</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Note_Google' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{avg_note:.2f}</div>
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
                <div class="metric-value">{avg_distance:.1f}</div>
                <div class="metric-label">📍 Distance Moy. TARMIZ</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col5:
        if 'Score_Presence_Digitale' in df.columns:
            avg_digital = df_filtered['Score_Presence_Digitale'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{avg_digital:.0f}</div>
                <div class="metric-label">📱 Score Digital</div>
            </div>
            """, unsafe_allow_html=True)
    
    # ----------------------------
    # ONGLETS PRINCIPAUX
    # ----------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🌍 Géographie", 
        "⭐ Performance", 
        "📱 Digital", 
        "📊 Analytics", 
        "📋 Données"
    ])
    
    with tab1:
        st.markdown('<div class="section-header"><h3>🌍 Analyse Géographique</h3></div>', 
                   unsafe_allow_html=True)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # Carte géographique améliorée
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
                        height=600,
                        title="🗺️ Répartition Géographique Interactive"
                    )
                    fig_map.update_layout(
                        title_font_size=16,
                        title_x=0.5,
                        margin=dict(t=50, b=0, l=0, r=0)
                    )
                    st.plotly_chart(fig_map, use_container_width=True)
        
        with col2:
            # Top villes avec style amélioré
            st.markdown("### 🏆 Top 10 Villes")
            top_cities = df_filtered["Ville"].value_counts().head(10)
            fig_cities = px.bar(
                y=top_cities.index,
                x=top_cities.values,
                orientation='h',
                color=top_cities.values,
                color_continuous_scale=color_theme,
                title="Nombre d'optiques par ville"
            )
            fig_cities.update_layout(
                height=400,
                showlegend=False,
                title_font_size=14
            )
            st.plotly_chart(fig_cities, use_container_width=True)
            
            # Statistiques géographiques
            st.markdown("### 📈 Stats Géo")
            total_cities = df_filtered['Ville'].nunique()
            st.metric("🏙️ Villes couvertes", total_cities)
            
            if len(top_cities) > 0:
                concentration = (top_cities.iloc[0] / len(df_filtered)) * 100
                st.metric("🎯 Concentration", f"{concentration:.1f}%")
    
    with tab2:
        st.markdown('<div class="section-header"><h3>⭐ Analyse de Performance</h3></div>', 
                   unsafe_allow_html=True)
        
        if 'Note_Google' in df.columns:
            # Dashboard des notes avec sous-graphiques
            fig_perf = make_subplots(
                rows=2, cols=3,
                subplot_titles=(
                    "Distribution des Notes", 
                    "Notes vs Nombre d'Avis",
                    "Top 5 Villes - Notes Moyennes",
                    "Évolution par Score Digital",
                    "Distribution Nb d'Avis",
                    "Performance par Distance"
                ),
                specs=[[{"type": "histogram"}, {"type": "scatter"}, {"type": "bar"}],
                       [{"type": "scatter"}, {"type": "histogram"}, {"type": "scatter"}]]
            )
            
            # Distribution des notes
            fig_perf.add_trace(
                go.Histogram(x=df_filtered["Note_Google"].dropna(), 
                           nbinsx=20, name="Notes", marker_color='lightblue'),
                row=1, col=1
            )
            
            # Notes vs Nb d'avis
            fig_perf.add_trace(
                go.Scatter(x=df_filtered["Note_Google"], y=df_filtered["Nb_Avis_Google"],
                          mode="markers", name="Performance", 
                          marker=dict(color='orange', size=6)),
                row=1, col=2
            )
            
            # Top 5 villes - notes moyennes
            top5_cities = df_filtered["Ville"].value_counts().head(5).index
            city_notes = df_filtered[df_filtered["Ville"].isin(top5_cities)].groupby("Ville")["Note_Google"].mean()
            fig_perf.add_trace(
                go.Bar(x=city_notes.index, y=city_notes.values, 
                      name="Moyenne", marker_color='green'),
                row=1, col=3
            )
            
            # Score digital vs notes
            if "Score_Presence_Digitale" in df.columns:
                fig_perf.add_trace(
                    go.Scatter(x=df_filtered["Score_Presence_Digitale"], 
                             y=df_filtered["Note_Google"],
                             mode="markers", name="Digital vs Note",
                             marker=dict(color='purple', size=6)),
                    row=2, col=1
                )
            
            # Distribution Nb d'avis
            fig_perf.add_trace(
                go.Histogram(x=df_filtered["Nb_Avis_Google"].dropna(), 
                           nbinsx=30, name="Nb Avis", marker_color='red'),
                row=2, col=2
            )
            
            # Distance vs notes
            if "Distance-TARMIZ(KM)" in df.columns:
                fig_perf.add_trace(
                    go.Scatter(x=df_filtered["Distance-TARMIZ(KM)"], 
                             y=df_filtered["Note_Google"],
                             mode="markers", name="Distance vs Note",
                             marker=dict(color='brown', size=6)),
                    row=2, col=3
                )
            
            fig_perf.update_layout(
                height=800, 
                title_text="📊 Dashboard Performance Complet",
                title_x=0.5,
                showlegend=False
            )
            st.plotly_chart(fig_perf, use_container_width=True)
            
            # Insights de performance
            col1, col2, col3 = st.columns(3)
            
            with col1:
                high_rated = len(df_filtered[df_filtered['Note_Google'] >= 4.0])
                st.metric("🌟 Notes ≥ 4.0", f"{high_rated} ({high_rated/len(df_filtered)*100:.1f}%)")
            
            with col2:
                high_reviews = len(df_filtered[df_filtered['Nb_Avis_Google'] >= 50])
                st.metric("💬 Avis ≥ 50", f"{high_reviews} ({high_reviews/len(df_filtered)*100:.1f}%)")
            
            with col3:
                top_performers = len(df_filtered[
                    (df_filtered['Note_Google'] >= 4.0) & 
                    (df_filtered['Nb_Avis_Google'] >= 20)
                ])
                st.metric("🏆 Top Performers", f"{top_performers} ({top_performers/len(df_filtered)*100:.1f}%)")
    
    with tab3:
        st.markdown('<div class="section-header"><h3>📱 Présence Digitale</h3></div>', 
                   unsafe_allow_html=True)
        
        # Analyse présence digitale
        digital_cols = ["Site web","Réseaux sociaux","Email"]
        digital_data = []
        
        for col in digital_cols:
            if col in df.columns:
                count = df_filtered[col].notna().sum()
                percentage = (count / len(df_filtered)) * 100
                digital_data.append({
                    'Canal': col,
                    'Nombre': count,
                    'Pourcentage': percentage
                })
        
        if digital_data:
            col1, col2 = st.columns(2)
            
            with col1:
                digital_df = pd.DataFrame(digital_data)
                fig_digital = px.bar(
                    digital_df,
                    x='Canal',
                    y='Pourcentage',
                    color='Pourcentage',
                    color_continuous_scale=color_theme,
                    title="📊 Taux de Présence par Canal",
                    text='Pourcentage'
                )
                fig_digital.update_traces(
                    texttemplate='%{text:.1f}%', 
                    textposition='outside'
                )
                fig_digital.update_layout(height=400)
                st.plotly_chart(fig_digital, use_container_width=True)
            
            with col2:
                fig_pie = px.pie(
                    digital_df,
                    values='Nombre',
                    names='Canal',
                    title="🥧 Répartition des Canaux",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
        
        # Score présence digitale
        if "Score_Presence_Digitale" in df.columns:
            col1, col2 = st.columns(2)
            
            with col1:
                fig_score_dist = px.histogram(
                    df_filtered, 
                    x="Score_Presence_Digitale", 
                    nbins=20,
                    color_discrete_sequence=['lightcoral'],
                    title="📈 Distribution Score Digital"
                )
                st.plotly_chart(fig_score_dist, use_container_width=True)
            
            with col2:
                if 'Note_Google' in df.columns:
                    fig_correlation = px.scatter(
                        df_filtered, 
                        x="Score_Presence_Digitale", 
                        y="Note_Google",
                        color="Ville",
                        size="Nb_Avis_Google" if "Nb_Avis_Google" in df.columns else None,
                        hover_data=["Nom"],
                        title="🔗 Score Digital vs Performance"
                    )
                    st.plotly_chart(fig_correlation, use_container_width=True)
    
    with tab4:
        st.markdown('<div class="section-header"><h3>📊 Analytics Avancés</h3></div>', 
                   unsafe_allow_html=True)
        
        # Matrice de corrélation
        numeric_columns = df_filtered.select_dtypes(include=[np.number]).columns
        if len(numeric_columns) > 1:
            st.markdown("### 🔗 Matrice de Corrélation")
            correlation_matrix = df_filtered[numeric_columns].corr()
            
            fig_corr = px.imshow(
                correlation_matrix,
                color_continuous_scale=color_theme,
                title="Corrélations entre variables numériques"
            )
            fig_corr.update_layout(height=500)
            st.plotly_chart(fig_corr, use_container_width=True)
        
        # Analyse par segments
        col1, col2 = st.columns(2)
        
        with col1:
            if "Taille_Entreprise" in df.columns:
                st.markdown("### 🏢 Analyse par Taille")
                size_analysis = df_filtered.groupby('Taille_Entreprise').agg({
                    'Note_Google': 'mean',
                    'Nb_Avis_Google': 'mean',
                    'Score_Presence_Digitale': 'mean'
                }).round(2)
                
                st.dataframe(size_analysis, use_container_width=True)
                
                fig_size = px.pie(
                    df_filtered, 
                    names="Taille_Entreprise",
                    title="Répartition par Taille",
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                st.plotly_chart(fig_size, use_container_width=True)
        
        with col2:
            if "Anciennete_Estimee" in df.columns:
                st.markdown("### 📅 Analyse Temporelle")
                fig_age = px.histogram(
                    df_filtered,
                    x="Anciennete_Estimee",
                    nbins=15,
                    color_discrete_sequence=['lightseagreen'],
                    title="Distribution de l'Ancienneté"
                )
                st.plotly_chart(fig_age, use_container_width=True)
                
                # Ancienneté vs Performance
                if 'Note_Google' in df.columns:
                    fig_age_perf = px.scatter(
                        df_filtered,
                        x="Anciennete_Estimee",
                        y="Note_Google",
                        color="Score_Presence_Digitale" if "Score_Presence_Digitale" in df.columns else None,
                        title="Ancienneté vs Performance"
                    )
                    st.plotly_chart(fig_age_perf, use_container_width=True)
        
        # Benchmarking
        st.markdown("### 🎯 Benchmarking")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if 'Note_Google' in df.columns:
                top_25_pct = df_filtered['Note_Google'].quantile(0.75)
                st.metric("🥇 Top 25% Notes", f"{top_25_pct:.2f}+")
        
        with col2:
            if 'Nb_Avis_Google' in df.columns:
                median_reviews = df_filtered['Nb_Avis_Google'].median()
                st.metric("📊 Médiane Avis", f"{median_reviews:.0f}")
        
        with col3:
            if 'Score_Presence_Digitale' in df.columns:
                top_digital = df_filtered['Score_Presence_Digitale'].quantile(0.90)
                st.metric("🚀 Top 10% Digital", f"{top_digital:.0f}+")
        
        with col4:
            if 'Distance-TARMIZ(KM)' in df.columns:
                close_to_tarmiz = len(df_filtered[df_filtered['Distance-TARMIZ(KM)'] <= 10])
                st.metric("📍 Proche TARMIZ (<10km)", close_to_tarmiz)
    
    with tab5:
        st.markdown('<div class="section-header"><h3>📋 Données Détaillées</h3></div>', 
                   unsafe_allow_html=True)
        
        # Options d'affichage
        col1, col2, col3 = st.columns(3)
        
        with col1:
            show_all = st.checkbox("Afficher toutes les colonnes", False)
        
        with col2:
            if not show_all:
                display_cols = st.multiselect(
                    "Colonnes à afficher:",
                    df_filtered.columns.tolist(),
                    default=['Nom', 'Ville', 'Note_Google', 'Nb_Avis_Google'][:4]
                )
            else:
                display_cols = df_filtered.columns.tolist()
        
        with col3:
            sort_by = st.selectbox(
                "Trier par:",
                ['Note_Google', 'Nb_Avis_Google', 'Score_Presence_Digitale', 'Distance-TARMIZ(KM)']
            )
        
        # Affichage du tableau
        if display_cols:
            df_display = df_filtered[display_cols].copy()
            
            if sort_by in df_display.columns:
                df_display = df_display.sort_values(sort_by, ascending=False)
            
            st.dataframe(
                df_display,
                use_container_width=True,
                height=400
            )
            
            # Statistiques du tableau
            st.markdown("### 📈 Statistiques")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📊 Lignes affichées", len(df_display))
            
            with col2:
                st.metric("📋 Colonnes", len(display_cols))
            
            with col3:
                if 'Note_Google' in df_display.columns:
                    avg_note_filtered = df_display['Note_Google'].mean()
                    st.metric("⭐ Moyenne filtrée", f"{avg_note_filtered:.2f}")
            
            with col4:
                if 'Score_Presence_Digitale' in df_display.columns:
                    avg_digital_filtered = df_display['Score_Presence_Digitale'].mean()
                    st.metric("📱 Score moy. filtré", f"{avg_digital_filtered:.0f}")
        
        # Export des données
        st.markdown("### 📥 Export")
        
        col1, col2 = st.columns(2)
        
        with col1:
            csv = df_filtered.to_csv(index=False)
            st.download_button(
                label="📊 Télécharger CSV (Filtré)",
                data=csv,
                file_name=f"optiques_filtered_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        with col2:
            if display_cols:
                csv_display = df_display.to_csv(index=False)
                st.download_button(
                    label="📋 Télécharger Sélection",
                    data=csv_display,
                    file_name=f"optiques_selection_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
    
    # ----------------------------
    # FOOTER AVEC RÉSUMÉ
    # ----------------------------
    st.markdown("---")
    st.markdown("## 📈 Résumé Exécutif")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🏆 Performance")
        if 'Note_Google' in df.columns:
            excellent = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            st.write(f"• **Excellentes** (≥4.5): {excellent} ({excellent/len(df_filtered)*100:.1f}%)")
            good = len(df_filtered[df_filtered['Note_Google'] >= 4.0])
            st.write(f"• **Bonnes** (≥4.0): {good} ({good/len(df_filtered)*100:.1f}%)")
    
    with col2:
        st.markdown("### 📱 Digital")
        if all(col in df.columns for col in ['Site web', 'Email', 'Réseaux sociaux']):
            complete_digital = len(df_filtered[
                df_filtered['Site web'].notna() & 
                df_filtered['Email'].notna() & 
                df_filtered['Réseaux sociaux'].notna()
            ])
            st.write(f"• **Présence complète**: {complete_digital} ({complete_digital/len(df_filtered)*100:.1f}%)")
    
    with col3:
        st.markdown("### 🌍 Géographie")
        total_cities = df_filtered['Ville'].nunique()
        top_city = df_filtered['Ville'].value_counts().index[0]
        top_count = df_filtered['Ville'].value_counts().iloc[0]
        st.write(f"• **Villes couvertes**: {total_cities}")
        st.write(f"• **Leader**: {top_city} ({top_count})")
    
    # Timestamp
    st.markdown(f"*Dernière mise à jour: {datetime.now().strftime('%d/%m/%Y à %H:%M')}*")

else:
    st.error("❌ Impossible de charger le fichier OPTIQUESS.xlsx")
    st.info("Vérifiez que le fichier existe dans le même répertoire que ce script.")
