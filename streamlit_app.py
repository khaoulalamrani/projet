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
    page_title="Dashboard Optiques - Simple",
    page_icon="👓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------
# CSS ULTRA SIMPLE ET CLAIR
# ----------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;600;700&display=swap');
    
    .main {
        font-family: 'Roboto', sans-serif;
        background: #f8f9fa;
    }
    
    .big-title {
        background: linear-gradient(90deg, #4CAF50, #2196F3);
        color: white;
        padding: 3rem 2rem;
        border-radius: 20px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }
    
    .big-title h1 {
        font-size: 3.5rem;
        margin: 0;
        font-weight: 700;
    }
    
    .big-title p {
        font-size: 1.5rem;
        margin: 1rem 0 0 0;
        opacity: 0.9;
    }
    
    .super-card {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 8px 25px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
        border-left: 5px solid #4CAF50;
    }
    
    .super-number {
        font-size: 4rem;
        font-weight: 700;
        color: #2196F3;
        margin: 0;
        line-height: 1;
    }
    
    .super-text {
        font-size: 1.3rem;
        color: #666;
        margin: 1rem 0 0 0;
        font-weight: 500;
    }
    
    .section-title {
        background: #4CAF50;
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        font-size: 1.8rem;
        font-weight: 600;
        margin: 2rem 0 1.5rem 0;
    }
    
    .explanation-box {
        background: #e8f5e8;
        border: 2px solid #4CAF50;
        border-radius: 10px;
        padding: 1.5rem;
        margin: 1rem 0;
        font-size: 1.1rem;
        line-height: 1.6;
    }
    
    .good-news {
        background: #d4edda;
        border: 2px solid #28a745;
        color: #155724;
    }
    
    .warning-news {
        background: #fff3cd;
        border: 2px solid #ffc107;
        color: #856404;
    }
    
    .info-news {
        background: #d1ecf1;
        border: 2px solid #17a2b8;
        color: #0c5460;
    }
    
    .emoji-big {
        font-size: 2rem;
        margin-right: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# TITRE PRINCIPAL SIMPLE
# ----------------------------
st.markdown("""
<div class="big-title">
    <h1>👓 Mes Optiques</h1>
    <p>Comprendre simplement le marché des opticiens</p>
</div>
""", unsafe_allow_html=True)

# ----------------------------
# CHARGEMENT DES DONNÉES
# ----------------------------
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('OPTIQUESS.xlsx', engine='openpyxl')
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
    # FILTRES SIMPLES
    # ----------------------------
    st.sidebar.markdown("## 🔍 Choisir les optiques à voir")
    
    # Filtre ville simple
    cities = ['Toutes les villes'] + sorted(df['Ville'].dropna().unique().tolist())
    selected_city = st.sidebar.selectbox("Dans quelle ville ?", cities)
    
    # Filtre note simple
    if 'Note_Google' in df.columns:
        note_choice = st.sidebar.radio(
            "Quelles notes afficher ?",
            ["Toutes les notes", "Seulement les bonnes (≥4.0)", "Seulement les excellentes (≥4.5)"]
        )
    
    # Application des filtres
    df_filtered = df.copy()
    
    if selected_city != 'Toutes les villes':
        df_filtered = df_filtered[df_filtered['Ville'] == selected_city]
    
    if 'Note_Google' in df.columns:
        if note_choice == "Seulement les bonnes (≥4.0)":
            df_filtered = df_filtered[df_filtered['Note_Google'] >= 4.0]
        elif note_choice == "Seulement les excellentes (≥4.5)":
            df_filtered = df_filtered[df_filtered['Note_Google'] >= 4.5]
    
    # ----------------------------
    # CHIFFRES CLÉS TRÈS SIMPLES
    # ----------------------------
    st.markdown('<div class="section-title">📊 Les Chiffres Importants</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total = len(df_filtered)
        st.markdown(f"""
        <div class="super-card">
            <div class="super-number">{total:,}</div>
            <div class="super-text">🏪 Optiques trouvées</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'Note_Google' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            note_emoji = "🌟" if avg_note >= 4.5 else "⭐" if avg_note >= 4.0 else "🔸"
            st.markdown(f"""
            <div class="super-card">
                <div class="super-number">{avg_note:.1f}/5</div>
                <div class="super-text">{note_emoji} Note moyenne</div>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        cities_count = df_filtered['Ville'].nunique()
        st.markdown(f"""
        <div class="super-card">
            <div class="super-number">{cities_count}</div>
            <div class="super-text">🏙️ Villes différentes</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        if 'Note_Google' in df.columns:
            excellent = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            pct_excellent = (excellent / len(df_filtered)) * 100
            st.markdown(f"""
            <div class="super-card">
                <div class="super-number">{pct_excellent:.0f}%</div>
                <div class="super-text">🏆 Optiques excellentes</div>
            </div>
            """, unsafe_allow_html=True)
    
    # ----------------------------
    # ONGLETS ULTRA-SIMPLES
    # ----------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "🗺️ Où sont les optiques ?", 
        "⭐ Qui a les meilleures notes ?", 
        "💻 Qui est sur internet ?",
        "📋 Voir toutes les données"
    ])
    
    with tab1:
        st.markdown('<div class="section-title">🗺️ Où se trouvent les optiques ?</div>', unsafe_allow_html=True)
        
        # Graphique simple des villes
        st.markdown("### 📊 Combien d'optiques par ville ?")
        
        city_counts = df_filtered['Ville'].value_counts().head(15)
        
        # Graphique en barres colorées très simple
        fig_cities = go.Figure(data=[
            go.Bar(
                x=city_counts.values,
                y=city_counts.index[::-1],  # Inverser pour le plus grand en haut
                orientation='h',
                marker=dict(
                    color=['#4CAF50' if i == 0 else '#2196F3' if i < 3 else '#FFC107' if i < 5 else '#FF9800' 
                           for i in range(len(city_counts))],
                    line=dict(color='white', width=2)
                ),
                text=city_counts.values[::-1],
                textposition='outside',
                textfont=dict(size=14, color='black', family="Roboto")
            )
        ])
        
        fig_cities.update_layout(
            title=dict(
                text="🏆 Le classement des villes avec le plus d'optiques",
                font=dict(size=20, color='#333', family="Roboto"),
                x=0.5
            ),
            xaxis=dict(
                title="Nombre d'optiques",
                titlefont=dict(size=16),
                tickfont=dict(size=12),
                showgrid=True,
                gridcolor='lightgray'
            ),
            yaxis=dict(
                titlefont=dict(size=16),
                tickfont=dict(size=12)
            ),
            height=600,
            paper_bgcolor='white',
            plot_bgcolor='white',
            margin=dict(l=100, r=50, t=80, b=50)
        )
        
        st.plotly_chart(fig_cities, use_container_width=True)
        
        # Explication simple
        top_city = city_counts.index[0]
        top_count = city_counts.iloc[0]
        st.markdown(f"""
        <div class="explanation-box good-news">
            <span class="emoji-big">🏆</span>
            <strong>La ville avec le plus d'optiques est {top_city}</strong> avec {top_count} optiques.
            <br><br>
            💡 <strong>Ce que ça veut dire :</strong> {top_city} a plus de concurrence entre optiques, 
            mais aussi plus de choix pour les clients.
        </div>
        """, unsafe_allow_html=True)
        
        # Carte si coordonnées disponibles
        if {"Latitude", "Longitude"}.issubset(df.columns):
            st.markdown("### 🗺️ Sur la carte")
            geo_df = df_filtered.dropna(subset=["Latitude", "Longitude"]).head(100)  # Limiter pour la performance
            
            if len(geo_df) > 0:
                fig_map = px.scatter_mapbox(
                    geo_df,
                    lat="Latitude",
                    lon="Longitude",
                    hover_name="Nom",
                    hover_data=["Ville", "Note_Google"],
                    color="Note_Google",
                    size_max=15,
                    zoom=6,
                    mapbox_style="open-street-map",
                    height=500,
                    color_continuous_scale="RdYlGn"
                )
                
                fig_map.update_layout(
                    title=dict(
                        text="📍 Où se trouvent les optiques sur la carte",
                        font=dict(size=18, family="Roboto"),
                        x=0.5
                    ),
                    margin=dict(t=50, b=0, l=0, r=0)
                )
                
                st.plotly_chart(fig_map, use_container_width=True)
                
                st.markdown("""
                <div class="explanation-box info-news">
                    <span class="emoji-big">💡</span>
                    <strong>Comment lire cette carte :</strong>
                    <br>• Les points verts = optiques avec de bonnes notes
                    <br>• Les points rouges = optiques avec des notes plus faibles
                    <br>• Cliquez sur un point pour voir les détails
                </div>
                """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown('<div class="section-title">⭐ Qui a les meilleures notes ?</div>', unsafe_allow_html=True)
        
        if 'Note_Google' in df.columns:
            # Graphique de performance simple
            st.markdown("### 🎯 Les notes des optiques")
            
            # Créer des catégories simples
            excellent = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            tres_bon = len(df_filtered[(df_filtered['Note_Google'] >= 4.0) & (df_filtered['Note_Google'] < 4.5)])
            moyen = len(df_filtered[(df_filtered['Note_Google'] >= 3.5) & (df_filtered['Note_Google'] < 4.0)])
            faible = len(df_filtered[df_filtered['Note_Google'] < 3.5])
            
            categories = ['🌟 Excellent\n(4.5-5.0)', '⭐ Très bon\n(4.0-4.4)', '🔸 Moyen\n(3.5-3.9)', '🔻 À améliorer\n(<3.5)']
            values = [excellent, tres_bon, moyen, faible]
            colors = ['#4CAF50', '#8BC34A', '#FFC107', '#FF5722']
            
            fig_notes = go.Figure(data=[
                go.Bar(
                    x=categories,
                    y=values,
                    marker=dict(color=colors, line=dict(color='white', width=2)),
                    text=values,
                    textposition='outside',
                    textfont=dict(size=16, color='black', family="Roboto")
                )
            ])
            
            fig_notes.update_layout(
                title=dict(
                    text="📊 Combien d'optiques dans chaque catégorie ?",
                    font=dict(size=20, color='#333', family="Roboto"),
                    x=0.5
                ),
                xaxis=dict(
                    tickfont=dict(size=12),
                    title=""
                ),
                yaxis=dict(
                    title="Nombre d'optiques",
                    titlefont=dict(size=16),
                    tickfont=dict(size=12),
                    showgrid=True,
                    gridcolor='lightgray'
                ),
                height=500,
                paper_bgcolor='white',
                plot_bgcolor='white',
                margin=dict(t=80, b=50)
            )
            
            st.plotly_chart(fig_notes, use_container_width=True)
            
            # Explication personnalisée
            total_good = excellent + tres_bon
            pct_good = (total_good / len(df_filtered)) * 100
            
            if pct_good >= 70:
                message_type = "good-news"
                emoji = "🎉"
                message = f"Excellente nouvelle ! {pct_good:.0f}% des optiques ont une bonne note (4.0 ou plus)."
            elif pct_good >= 50:
                message_type = "info-news"
                emoji = "👍"
                message = f"Plutôt bien ! {pct_good:.0f}% des optiques ont une bonne note (4.0 ou plus)."
            else:
                message_type = "warning-news"
                emoji = "⚠️"
                message = f"Attention ! Seulement {pct_good:.0f}% des optiques ont une bonne note (4.0 ou plus)."
            
            st.markdown(f"""
            <div class="explanation-box {message_type}">
                <span class="emoji-big">{emoji}</span>
                <strong>{message}</strong>
                <br><br>
                💡 <strong>Ce que ça veut dire :</strong> Les notes Google montrent la satisfaction des clients. 
                Plus la note est haute, plus les clients sont contents !
            </div>
            """, unsafe_allow_html=True)
            
            # Top des meilleures optiques
            st.markdown("### 🏆 Le top 10 des meilleures optiques")
            
            top_optiques = df_filtered.nlargest(10, 'Note_Google')[['Nom', 'Ville', 'Note_Google', 'Nb_Avis_Google']]
            
            # Créer un graphique pour le top 10
            fig_top = go.Figure(data=[
                go.Bar(
                    x=top_optiques['Note_Google'],
                    y=[f"{row['Nom'][:25]}{'...' if len(row['Nom']) > 25 else ''}\n({row['Ville']})" 
                       for _, row in top_optiques.iterrows()][::-1],
                    orientation='h',
                    marker=dict(
                        color=top_optiques['Note_Google'],
                        colorscale='RdYlGn',
                        showscale=False,
                        line=dict(color='white', width=1)
                    ),
                    text=[f"{note:.1f} ⭐" for note in top_optiques['Note_Google']][::-1],
                    textposition='outside',
                    textfont=dict(size=12, color='black')
                )
            ])
            
            fig_top.update_layout(
                title=dict(
                    text="🥇 Les 10 optiques avec les meilleures notes",
                    font=dict(size=18, family="Roboto"),
                    x=0.5
                ),
                xaxis=dict(
                    title="Note Google",
                    range=[3.5, 5.1],
                    tickfont=dict(size=12)
                ),
                yaxis=dict(tickfont=dict(size=10)),
                height=600,
                paper_bgcolor='white',
                plot_bgcolor='white',
                margin=dict(l=200, r=50, t=60, b=50)
            )
            
            st.plotly_chart(fig_top, use_container_width=True)
            
            # Analyse des avis
            if 'Nb_Avis_Google' in df.columns:
                st.markdown("### 💬 Qui a le plus d'avis clients ?")
                
                avg_avis = df_filtered['Nb_Avis_Google'].mean()
                median_avis = df_filtered['Nb_Avis_Google'].median()
                
                # Distribution des avis en catégories simples
                peu_avis = len(df_filtered[df_filtered['Nb_Avis_Google'] < 10])
                moyen_avis = len(df_filtered[(df_filtered['Nb_Avis_Google'] >= 10) & (df_filtered['Nb_Avis_Google'] < 50)])
                beaucoup_avis = len(df_filtered[df_filtered['Nb_Avis_Google'] >= 50])
                
                fig_avis = go.Figure(data=[go.Pie(
                    labels=['👥 Beaucoup d\'avis\n(50+)', '💬 Quelques avis\n(10-49)', '🔇 Peu d\'avis\n(<10)'],
                    values=[beaucoup_avis, moyen_avis, peu_avis],
                    hole=.3,
                    marker=dict(colors=['#4CAF50', '#FFC107', '#FF5722']),
                    textinfo='label+percent+value',
                    textfont=dict(size=12)
                )])
                
                fig_avis.update_layout(
                    title=dict(
                        text="📈 Répartition du nombre d'avis par optique",
                        font=dict(size=18, family="Roboto"),
                        x=0.5
                    ),
                    height=500,
                    paper_bgcolor='white'
                )
                
                st.plotly_chart(fig_avis, use_container_width=True)
                
                st.markdown(f"""
                <div class="explanation-box info-news">
                    <span class="emoji-big">📊</span>
                    <strong>En moyenne, chaque optique a {avg_avis:.0f} avis clients.</strong>
                    <br><br>
                    💡 <strong>Pourquoi c'est important :</strong> Plus une optique a d'avis, plus on peut faire confiance à sa note. 
                    Une note de 5/5 avec 2 avis est moins fiable qu'une note de 4.5/5 avec 100 avis !
                </div>
                """, unsafe_allow_html=True)
    
    with tab3:
        st.markdown('<div class="section-title">💻 Qui est présent sur internet ?</div>', unsafe_allow_html=True)
        
        # Analyser la présence digitale de manière simple
        digital_analysis = {}
        
        if 'Site web' in df.columns:
            avec_site = df_filtered['Site web'].notna().sum()
            sans_site = len(df_filtered) - avec_site
            digital_analysis['Site web'] = {'Oui': avec_site, 'Non': sans_site}
        
        if 'Email' in df.columns:
            avec_email = df_filtered['Email'].notna().sum()
            sans_email = len(df_filtered) - avec_email
            digital_analysis['Email'] = {'Oui': avec_email, 'Non': sans_email}
        
        if 'Réseaux sociaux' in df.columns:
            avec_social = df_filtered['Réseaux sociaux'].notna().sum()
            sans_social = len(df_filtered) - avec_social
            digital_analysis['Réseaux sociaux'] = {'Oui': avec_social, 'Non': sans_social}
        
        if digital_analysis:
            st.markdown("### 🌐 Présence sur internet")
            
            # Créer un graphique empilé simple
            canaux = list(digital_analysis.keys())
            oui_values = [digital_analysis[canal]['Oui'] for canal in canaux]
            non_values = [digital_analysis[canal]['Non'] for canal in canaux]
            
            fig_digital = go.Figure()
            
            fig_digital.add_trace(go.Bar(
                name='✅ Présent',
                x=canaux,
                y=oui_values,
                marker_color='#4CAF50',
                text=[f"{val} optiques" for val in oui_values],
                textposition='inside'
            ))
            
            fig_digital.add_trace(go.Bar(
                name='❌ Absent',
                x=canaux,
                y=non_values,
                marker_color='#FF5722',
                text=[f"{val} optiques" for val in non_values],
                textposition='inside'
            ))
            
            fig_digital.update_layout(
                title=dict(
                    text="📊 Combien d'optiques sont présentes sur chaque canal ?",
                    font=dict(size=18, family="Roboto"),
                    x=0.5
                ),
                barmode='stack',
                xaxis=dict(
                    title="Type de présence internet",
                    titlefont=dict(size=14),
                    tickfont=dict(size=12)
                ),
                yaxis=dict(
                    title="Nombre d'optiques",
                    titlefont=dict(size=14),
                    tickfont=dict(size=12)
                ),
                height=500,
                paper_bgcolor='white',
                plot_bgcolor='white',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )
            
            st.plotly_chart(fig_digital, use_container_width=True)
            
            # Analyse simple
            if 'Site web' in digital_analysis:
                pct_site = (digital_analysis['Site web']['Oui'] / len(df_filtered)) * 100
                
                if pct_site >= 80:
                    message_type = "good-news"
                    emoji = "🚀"
                    message = f"Excellent ! {pct_site:.0f}% des optiques ont un site web."
                elif pct_site >= 50:
                    message_type = "info-news"
                    emoji = "👍"
                    message = f"Pas mal ! {pct_site:.0f}% des optiques ont un site web."
                else:
                    message_type = "warning-news"
                    emoji = "⚠️"
                    message = f"Attention ! Seulement {pct_site:.0f}% des optiques ont un site web."
                
                st.markdown(f"""
                <div class="explanation-box {message_type}">
                    <span class="emoji-big">{emoji}</span>
                    <strong>{message}</strong>
                    <br><br>
                    💡 <strong>Pourquoi c'est important :</strong> Aujourd'hui, les clients cherchent d'abord sur internet. 
                    Une optique sans site web perd beaucoup de clients potentiels !
                </div>
                """, unsafe_allow_html=True)
        
        # Score digital si disponible
        if 'Score_Presence_Digitale' in df.columns:
            st.markdown("### 📱 Niveau de maturité digitale")
            
            # Créer des catégories simples pour le score digital
            debutant = len(df_filtered[df_filtered['Score_Presence_Digitale'] <= 30])
            intermediaire = len(df_filtered[(df_filtered['Score_Presence_Digitale'] > 30) & (df_filtered['Score_Presence_Digitale'] <= 60)])
            avance = len(df_filtered[(df_filtered['Score_Presence_Digitale'] > 60) & (df_filtered['Score_Presence_Digitale'] <= 80)])
            expert = len(df_filtered[df_filtered['Score_Presence_Digitale'] > 80])
            
            labels = ['🟢 Expert\n(80-100)', '🔵 Avancé\n(60-80)', '🟡 Intermédiaire\n(30-60)', '🔴 Débutant\n(0-30)']
            values = [expert, avance, intermediaire, debutant]
            colors = ['#4CAF50', '#2196F3', '#FFC107', '#FF5722']
            
            fig_maturity = go.Figure(data=[go.Pie(
                labels=labels,
                values=values,
                hole=.4,
                marker=dict(colors=colors),
                textinfo='label+percent+value',
                textfont=dict(size=12)
            )])
            
            fig_maturity.update_layout(
                title=dict(
                    text="🎯 À quel niveau sont les optiques sur le digital ?",
                    font=dict(size=18, family="Roboto"),
                    x=0.5
                ),
                height=500,
                paper_bgcolor='white'
            )
            
            st.plotly_chart(fig_maturity, use_container_width=True)
            
            avg_score = df_filtered['Score_Presence_Digitale'].mean()
            
            if avg_score >= 70:
                message_type = "good-news"
                emoji = "🚀"
                level = "très bon"
            elif avg_score >= 50:
                message_type = "info-news"
                emoji = "📈"
                level = "correct"
            else:
                message_type = "warning-news"
                emoji = "📉"
                level = "faible"
            
            st.markdown(f"""
            <div class="explanation-box {message_type}">
                <span class="emoji-big">{emoji}</span>
                <strong>Le niveau digital moyen est {level} avec un score de {avg_score:.0f}/100.</strong>
                <br><br>
                💡 <strong>Ce que ça veut dire :</strong> Ce score combine site web, réseaux sociaux, présence Google, etc. 
                Plus le score est élevé, plus l'optique est visible sur internet !
            </div>
            """, unsafe_allow_html=True)
    
    with tab4:
        st.markdown('<div class="section-title">📋 Toutes les données</div>', unsafe_allow_html=True)
        
        st.markdown("### 🔍 Explorer les données")
        
        # Options simples
        col1, col2 = st.columns(2)
        
        with col1:
            colonnes_importantes = ['Nom', 'Ville', 'Note_Google', 'Nb_Avis_Google', 'Site web']
            colonnes_disponibles = [col for col in colonnes_importantes if col in df_filtered.columns]
            
            afficher_colonnes = st.multiselect(
                "Quelles informations voulez-vous voir ?",
                df_filtered.columns.tolist(),
                default=colonnes_disponibles
            )
        
        with col2:
            tri_options = ['Note_Google', 'Nb_Avis_Google', 'Nom', 'Ville']
            tri_disponible = [col for col in tri_options if col in df_filtered.columns]
            
            if tri_disponible:
                trier_par = st.selectbox("Comment trier ?", 
                                       ["Meilleures notes d'abord", "Plus d'avis d'abord", "Par nom", "Par ville"])
        
        # Affichage du tableau simplifié
        if afficher_colonnes:
            df_show = df_filtered[afficher_colonnes].copy()
            
            # Appliquer le tri
            if tri_disponible and 'trier_par' in locals():
                if trier_par == "Meilleures notes d'abord" and 'Note_Google' in df_show.columns:
                    df_show = df_show.sort_values('Note_Google', ascending=False)
                elif trier_par == "Plus d'avis d'abord" and 'Nb_Avis_Google' in df_show.columns:
                    df_show = df_show.sort_values('Nb_Avis_Google', ascending=False)
                elif trier_par == "Par nom" and 'Nom' in df_show.columns:
                    df_show = df_show.sort_values('Nom')
                elif trier_par == "Par ville" and 'Ville' in df_show.columns:
                    df_show = df_show.sort_values('Ville')
            
            # Formater les nombres pour qu'ils soient plus lisibles
            if 'Note_Google' in df_show.columns:
                df_show['Note_Google'] = df_show['Note_Google'].round(1).astype(str) + "/5"
            
            if 'Score_Presence_Digitale' in df_show.columns:
                df_show['Score_Presence_Digitale'] = df_show['Score_Presence_Digitale'].round(0).astype(str) + "/100"
            
            if 'Distance-TARMIZ(KM)' in df_show.columns:
                df_show['Distance-TARMIZ(KM)'] = df_show['Distance-TARMIZ(KM)'].round(1).astype(str) + " km"
            
            # Renommer les colonnes pour qu'elles soient plus claires
            rename_dict = {
                'Nom': '🏪 Nom de l\'optique',
                'Ville': '🏙️ Ville',
                'Note_Google': '⭐ Note Google',
                'Nb_Avis_Google': '💬 Nombre d\'avis',
                'Site web': '🌐 Site internet',
                'Email': '📧 Email',
                'Réseaux sociaux': '📱 Réseaux sociaux',
                'Score_Presence_Digitale': '📊 Score digital',
                'Distance-TARMIZ(KM)': '📍 Distance TARMIZ'
            }
            
            for old_name, new_name in rename_dict.items():
                if old_name in df_show.columns:
                    df_show = df_show.rename(columns={old_name: new_name})
            
            st.markdown("### 📊 Voici vos optiques :")
            st.dataframe(df_show, use_container_width=True, height=400)
            
            # Résumé simple du tableau
            st.markdown(f"""
            <div class="explanation-box info-news">
                <span class="emoji-big">📋</span>
                <strong>Vous regardez {len(df_show):,} optiques</strong> sur un total de {len(df):,} dans la base de données.
                <br><br>
                💾 Vous pouvez télécharger ces données en cliquant sur le bouton ci-dessous.
            </div>
            """, unsafe_allow_html=True)
        
        # Téléchargement simple
        st.markdown("### 💾 Télécharger les données")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Préparer les données pour l'export
            df_export = df_filtered.copy()
            
            # Nettoyer les données pour l'export
            if 'Note_Google' in df_export.columns:
                df_export['Note_Google'] = df_export['Note_Google'].round(2)
            
            csv_data = df_export.to_csv(index=False, encoding='utf-8-sig')  # utf-8-sig pour Excel
            
            st.download_button(
                label="📊 Télécharger toutes les données filtrées",
                data=csv_data,
                file_name=f"optiques_data_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                help="Fichier CSV compatible avec Excel"
            )
        
        with col2:
            if afficher_colonnes:
                csv_selection = df_show.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📋 Télécharger seulement les colonnes choisies",
                    data=csv_selection,
                    file_name=f"optiques_selection_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    help="Seulement les colonnes que vous avez sélectionnées"
                )
    
    # ----------------------------
    # RÉSUMÉ FINAL TRÈS SIMPLE
    # ----------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">🎯 Résumé : Ce qu\'il faut retenir</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🏆 Performance")
        if 'Note_Google' in df.columns:
            avg_note = df_filtered['Note_Google'].mean()
            excellent_count = len(df_filtered[df_filtered['Note_Google'] >= 4.5])
            excellent_pct = (excellent_count / len(df_filtered)) * 100
            
            if avg_note >= 4.2:
                performance_msg = "Les optiques ont d'excellentes notes !"
                performance_emoji = "🎉"
                performance_color = "good-news"
            elif avg_note >= 3.8:
                performance_msg = "Les notes sont correctes dans l'ensemble."
                performance_emoji = "👍"
                performance_color = "info-news"
            else:
                performance_msg = "Les notes pourraient être améliorées."
                performance_emoji = "⚠️"
                performance_color = "warning-news"
            
            st.markdown(f"""
            <div class="explanation-box {performance_color}">
                <span class="emoji-big">{performance_emoji}</span>
                <strong>{performance_msg}</strong>
                <br><br>
                📊 Note moyenne : <strong>{avg_note:.1f}/5</strong><br>
                🌟 {excellent_count} optiques excellentes ({excellent_pct:.0f}%)
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("### 💻 Présence Internet")
        web_analysis = ""
        web_emoji = "💻"
        web_color = "info-news"
        
        if 'Site web' in df.columns:
            web_count = df_filtered['Site web'].notna().sum()
            web_pct = (web_count / len(df_filtered)) * 100
            
            if web_pct >= 70:
                web_analysis = f"Très bien ! {web_pct:.0f}% ont un site web."
                web_emoji = "🚀"
                web_color = "good-news"
            elif web_pct >= 40:
                web_analysis = f"Moyen : {web_pct:.0f}% ont un site web."
                web_emoji = "📈"
                web_color = "info-news"
            else:
                web_analysis = f"Faible : seulement {web_pct:.0f}% ont un site web."
                web_emoji = "⚠️"
                web_color = "warning-news"
        
        if 'Score_Presence_Digitale' in df.columns:
            avg_digital = df_filtered['Score_Presence_Digitale'].mean()
            web_analysis += f"<br>📱 Score digital moyen : <strong>{avg_digital:.0f}/100</strong>"
        
        st.markdown(f"""
        <div class="explanation-box {web_color}">
            <span class="emoji-big">{web_emoji}</span>
            <strong>Présence sur internet :</strong>
            <br><br>
            {web_analysis}
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("### 🌍 Répartition")
        total_cities = df_filtered['Ville'].nunique()
        top_city = df_filtered['Ville'].value_counts().index[0]
        top_city_count = df_filtered['Ville'].value_counts().iloc[0]
        concentration = (top_city_count / len(df_filtered)) * 100
        
        if concentration > 50:
            geo_msg = f"Très concentré dans {top_city}"
            geo_emoji = "🎯"
            geo_color = "warning-news"
        elif concentration > 30:
            geo_msg = f"Assez concentré dans {top_city}"
            geo_emoji = "📍"
            geo_color = "info-news"
        else:
            geo_msg = "Bien réparti géographiquement"
            geo_emoji = "🌍"
            geo_color = "good-news"
        
        st.markdown(f"""
        <div class="explanation-box {geo_color}">
            <span class="emoji-big">{geo_emoji}</span>
            <strong>{geo_msg}</strong>
            <br><br>
            🏙️ <strong>{total_cities}</strong> villes différentes<br>
            🏆 Leader : <strong>{top_city}</strong> ({top_city_count} optiques)
        </div>
        """, unsafe_allow_html=True)
    
    # Conseils pratiques
    st.markdown("### 💡 Conseils pratiques")
    
    advice_col1, advice_col2 = st.columns(2)
    
    with advice_col1:
        st.markdown("""
        <div class="explanation-box info-news">
            <span class="emoji-big">🎯</span>
            <strong>Pour choisir une bonne optique :</strong>
            <br><br>
            ✅ Regardez la note Google (4.0 minimum)<br>
            ✅ Vérifiez le nombre d'avis (plus il y en a, mieux c'est)<br>
            ✅ Choisissez près de chez vous<br>
            ✅ Regardez si elle a un site web moderne
        </div>
        """, unsafe_allow_html=True)
    
    with advice_col2:
        st.markdown("""
        <div class="explanation-box good-news">
            <span class="emoji-big">🚀</span>
            <strong>Si vous êtes opticien :</strong>
            <br><br>
            📈 Demandez plus d'avis à vos clients contents<br>
            🌐 Créez un site web professionnel<br>
            📱 Soyez présent sur les réseaux sociaux<br>
            ⭐ Répondez aux avis clients (bons et mauvais)
        </div>
        """, unsafe_allow_html=True)
    
    # Footer simple
    st.markdown("---")
    
    # Informations sur les filtres actifs
    filters_text = []
    if selected_city != 'Toutes les villes':
        filters_text.append(f"🏙️ Ville : {selected_city}")
    if 'note_choice' in locals() and note_choice != "Toutes les notes":
        filters_text.append(f"⭐ {note_choice.lower()}")
    
    if filters_text:
        filters_display = " • ".join(filters_text)
        st.markdown(f"**🔍 Filtres actifs :** {filters_display}")
    
    st.markdown(f"""
    <div style="text-align: center; color: #666; margin-top: 2rem; padding: 1rem;">
        <strong>📊 Dashboard Optiques Simple</strong><br>
        Dernière mise à jour : {datetime.now().strftime('%d/%m/%Y à %H:%M')}<br>
        📈 Données analysées : {len(df_filtered):,} optiques sur {len(df):,} total
    </div>
    """, unsafe_allow_html=True)

else:
    # Page d'erreur très simple
    st.markdown("""
    <div style="background: #ffebee; border: 3px solid #f44336; border-radius: 15px; padding: 3rem; text-align: center; margin: 2rem 0;">
        <h1 style="color: #d32f2f; font-size: 3rem;">❌ Oups !</h1>
        <p style="font-size: 1.5rem; color: #666;">Je ne trouve pas le fichier avec les données des optiques.</p>
        <p style="font-size: 1.2rem; color: #666;">Le fichier <strong>OPTIQUESS.xlsx</strong> doit être dans le même dossier que ce programme.</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    ### 🔧 Comment résoudre le problème :
    
    1. **Vérifiez le nom du fichier** : il doit s'appeler exactement `OPTIQUESS.xlsx`
    2. **Vérifiez l'emplacement** : le fichier doit être dans le même dossier que ce programme
    3. **Vérifiez que le fichier fonctionne** : ouvrez-le avec Excel pour voir s'il n'est pas corrompu
    4. **Redémarrez** : fermez et relancez le programme
    
    Si le problème persiste, contactez l'administrateur du système.
    """)
