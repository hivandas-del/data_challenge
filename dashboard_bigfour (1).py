import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Big Four — Analyse Employés",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLORS = {
    "deloitte": "#86BC25",
    "ey":       "#FFD700",
    "kpmg":     "#00338D",
    "pwc":      "#D04A02",
}
LABELS = {
    "deloitte": "Deloitte",
    "ey":       "EY",
    "kpmg":     "KPMG",
    "pwc":      "PwC",
}
COLOR_MAP = {v: COLORS[k] for k, v in LABELS.items()}

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.2rem; padding-bottom: 1rem; }
    h1  { font-size: 1.7rem !important; }
    h2  { font-size: 1.1rem !important; border-bottom: 2px solid #e8eaf0; padding-bottom: 4px; margin-top: 0.6rem; }
    .kpi-card {
        background: #f7f9fc;
        border-radius: 10px;
        padding: 14px 18px;
        border-left: 5px solid #ccc;
        margin-bottom: 6px;
    }
    .kpi-label { font-size: 0.72rem; color: #888; text-transform: uppercase; font-weight: 700; letter-spacing: 0.04em; }
    .kpi-value { font-size: 1.8rem; font-weight: 800; color: #1a1a2e; line-height: 1.2; }
    .kpi-sub   { font-size: 0.78rem; color: #666; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# COLONNES THEMES (pré-annotées dans le fichier Excel)
# ─────────────────────────────────────────────────────────────────────────────
THEME_PRO_COLS = {
    "Contenu & exécution":  "theme_pro: Contenu et exécution du travail",
    "Environnement":        "theme_pro: Environnement de travail",
    "Aménagement":          "theme_pro: Aménagement du travail",
    "Aspect financier":     "theme_pro: Aspect financier",
}
THEME_CON_COLS = {
    "Contenu & exécution":  "theme_con: Contenu et exécution du travail",
    "Environnement":        "theme_con: Environnement de travail",
    "Aménagement":          "theme_con: Aménagement du travail",
    "Aspect financier":     "theme_con: Aspect financier",
}
SOUS_THEME_PRO_COLS = {
    "Autonomie":           "sous_theme_pro: Autonomie et libert\u00e9 d\u2019action / confiance",
    "Travail stimulant":   "sous_theme_pro: Travail stimulant/ satisfaction dans le travail/ sens du travail / motivation",
    "D\u00e9v. professionnel":  "sous_theme_pro: Opportunit\u00e9 de d\u00e9veloppement professionnel",
    "Ambiance \u00e9quipe":     "sous_theme_pro: Ambiance et collaboration en \u00e9quipe / appartenance \u00e0 l\u2019entreprise",
    "Bienveillance mgt":   "sous_theme_pro: Bienveillance mana\u00e9riale et reconnaissance / qualit\u00e9 de l\u2019encadrement",
    "Outils":              "sous_theme_pro: Outils et infrastructures",
    "Flexibilit\u00e9":        "sous_theme_pro: Flexibilit\u00e9",
    "\u00c9quilibre vie pro":  "sous_theme_pro: Equilibre vie priv\u00e9e - vie professionnelle",
    "Charge travail":      "sous_theme_pro: Charge de travail / travail stressant",
    "Aspect financier":    "sous_theme_pro: Aspect financier",
}
SOUS_THEME_CON_COLS = {
    "Autonomie":           "sous_theme_con: Autonomie et libert\u00e9 d\u2019action / confiance",
    "Travail stimulant":   "sous_theme_con: Travail stimulant/ satisfaction dans le travail/ sens du travail / motivation",
    "D\u00e9v. professionnel":  "sous_theme_con: Opportunit\u00e9 de d\u00e9veloppement professionnel",
    "Ambiance \u00e9quipe":     "sous_theme_con: Ambiance et collaboration en \u00e9quipe / appartenance \u00e0 l\u2019entreprise",
    "Bienveillance mgt":   "sous_theme_con: Bienveillance mana\u00e9riale et reconnaissance / qualit\u00e9 de l\u2019encadrement",
    "Outils":              "sous_theme_con: Outils et infrastructures",
    "Flexibilit\u00e9":        "sous_theme_con: Flexibilit\u00e9",
    "\u00c9quilibre vie pro":  "sous_theme_con: Equilibre vie priv\u00e9e - vie professionnelle",
    "Charge travail":      "sous_theme_con: Charge de travail / travail stressant",
    "Aspect financier":    "sous_theme_con: Aspect financier",
}

STOPWORDS = {
    "the","and","to","of","a","in","is","for","are","with","that","this","it","as","on",
    "have","at","be","by","or","an","not","but","our","we","i","you","your","my","from",
    "can","was","were","has","had","been","will","would","very","also","its","their",
    "there","they","which","more","some","when","if","no","so","do","up","out","about",
    "good","work","company","great","les","des","de","la","le","un","une","et","est",
    "en","du","au","qui","que","pas","sur","plus","dans","pour","avec","par","ce","se",
    "ne","on","il","je","nous","vous","ils","elles","mais","ou","donc","or","ni","car",
    "us","all","get","make","time","team","people","management","really","much","many",
    "lot","place","job","well","other","staff","firm","big","new","one","way","often",
    "what","even","here","most","than","just","every","long","hours","year","years",
    "each","overall","culture","working","entre","très","plus","cette","tout","comme",
    "pour","dans","qui","que","une","les","des","sur","par","pas","est","avec","sont",
    "être","avoir","faire","cette","cela","bien","peu","même","dont","quand","selon",
    "ainsi","sans","sous","lors","toute","après","avant","depuis","toujours","encore",
    "jamais","chez","leur","leurs","mon","ton","son","nos","vos","mes","ses",
    "entreprise","cabinet","travail","poste","équipe",
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Chargement des données…")
def load_data():
    df = pd.read_excel("big_four.xlsx")
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["Annee"] = df["date"].dt.year
    df["label"] = df["source"].map(LABELS)
    df["color"] = df["source"].map(COLORS)
    df["employee_type"] = df["employee_type"].str.replace("&nbsp;", " ", regex=False).str.strip()
    df["statut"] = df["employee_type"].apply(lambda x:
        "Actuel" if "actuel" in str(x).lower() else
        "Stagiaire" if "stagiaire" in str(x).lower() else "Ancien")
    return df

df_all = load_data()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILTRES
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔍 Filtres")

    entreprises = st.multiselect(
        "Cabinet(s)",
        options=["deloitte", "ey", "kpmg", "pwc"],
        default=["deloitte", "ey", "kpmg", "pwc"],
        format_func=lambda x: LABELS[x],
    )

    years_available = sorted(df_all["Annee"].dropna().unique().astype(int))
    year_range = st.slider(
        "Période",
        min_value=int(min(years_available)),
        max_value=int(max(years_available)),
        value=(2018, 2024),
    )

    note_range = st.slider("Note globale", 1, 5, (1, 5), step=1)

    grades = ["Tous"] + sorted(df_all["grade"].dropna().unique().tolist())
    grade_filter = st.selectbox("Grade", grades)

    hier = ["Tous"] + sorted(df_all["grade_hierarchical"].dropna().unique().tolist())
    hier_filter = st.selectbox("Niveau hiérarchique", hier)

    statuts = ["Tous", "Actuel", "Ancien", "Stagiaire"]
    statut_filter = st.selectbox("Statut employé", statuts)

    locations = ["Toutes"] + list(df_all["location"].value_counts().head(15).index)
    location_filter = st.selectbox("Ville", locations)

    st.markdown("---")
    st.markdown(
        "<small>📌 Data Challenge — Paris 1 Panthéon-Sorbonne<br>Master IMC & DS</small>",
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# FILTRAGE
# ─────────────────────────────────────────────────────────────────────────────
df = df_all[
    df_all["source"].isin(entreprises)
    & df_all["Annee"].between(year_range[0], year_range[1])
    & df_all["rating"].between(note_range[0], note_range[1])
].copy()

if grade_filter != "Tous":
    df = df[df["grade"] == grade_filter]
if hier_filter != "Tous":
    df = df[df["grade_hierarchical"] == hier_filter]
if statut_filter != "Tous":
    df = df[df["statut"] == statut_filter]
if location_filter != "Toutes":
    df = df[df["location"] == location_filter]

if df.empty:
    st.warning("Aucune donnée ne correspond aux filtres sélectionnés.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.title("📊 Big Four — Analyse des Avis Employés (France)")
st.caption(
    f"**{len(df):,}** avis · {year_range[0]}–{year_range[1]} · "
    f"Cabinets : {', '.join(LABELS[e] for e in entreprises)}"
)

# ─────────────────────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
def kpi(label, value, sub="", color="#2563EB"):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{color}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    kpi("Avis total", f"{len(df):,}", color="#2563EB")
with k2:
    kpi("Note moyenne", f"{df['rating'].mean():.2f} / 5", color="#16A34A")
with k3:
    kpi("Avis positifs ≥ 4★", f"{(df['rating'] >= 4).mean()*100:.1f}%", color="#F59E0B")
with k4:
    kpi("Avis négatifs ≤ 2★", f"{(df['rating'] <= 2).mean()*100:.1f}%", color="#DC2626")
with k5:
    kpi("Villes représentées", f"{df['location'].nunique()}", color="#7C3AED")
with k6:
    best = df.groupby("source")["rating"].mean().idxmax()
    kpi("Cabinet leader", LABELS[best], f"note {df[df['source']==best]['rating'].mean():.2f}", color=COLORS[best])

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 1 — Note moyenne + Distribution
# ─────────────────────────────────────────────────────────────────────────────
col_a, col_b = st.columns([1, 1.4])

with col_a:
    st.markdown("## Note moyenne par cabinet")
    note_co = (
        df.groupby("source")["rating"]
        .agg(["mean", "count"]).reset_index()
        .sort_values("mean", ascending=True)
    )
    note_co["label"] = note_co["source"].map(LABELS)
    note_co["color"] = note_co["source"].map(COLORS)

    fig = go.Figure()
    for _, row in note_co.iterrows():
        fig.add_trace(go.Bar(
            x=[row["mean"]], y=[row["label"]], orientation="h",
            marker_color=row["color"], name=row["label"],
            text=f"  {row['mean']:.2f}  ({int(row['count'])} avis)",
            textposition="outside", showlegend=False,
        ))
    fig.update_layout(
        height=240, margin=dict(l=10, r=90, t=10, b=10),
        xaxis=dict(range=[1, 5.8], title="Note / 5"),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True)

with col_b:
    st.markdown("## Distribution des notes")
    dist = df.groupby(["source", "rating"]).size().reset_index(name="count")
    dist["label"] = dist["source"].map(LABELS)
    dist["pct"] = dist.groupby("source")["count"].transform(lambda x: x / x.sum() * 100)

    fig = px.bar(
        dist, x="rating", y="pct", color="label", barmode="group",
        color_discrete_map=COLOR_MAP,
        labels={"pct": "% des avis", "rating": "Note ★", "label": ""},
        height=240,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.3),
        xaxis=dict(tickvals=[1,2,3,4,5], ticktext=["1★","2★","3★","4★","5★"]),
    )
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 2 — Évolution temporelle + Donut
# ─────────────────────────────────────────────────────────────────────────────
col_c, col_d = st.columns([1.7, 1])

with col_c:
    st.markdown("## Évolution de la note moyenne (annuelle)")
    trend = df.groupby(["Annee", "source"])["rating"].mean().reset_index()
    trend["label"] = trend["source"].map(LABELS)

    fig = px.line(
        trend, x="Annee", y="rating", color="label", markers=True,
        color_discrete_map=COLOR_MAP,
        labels={"rating": "Note / 5", "Annee": "Année", "label": ""},
        height=270,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.25), yaxis=dict(range=[1, 5.5]),
    )
    st.plotly_chart(fig, use_container_width=True)

with col_d:
    st.markdown("## Répartition des avis")
    vol = df["source"].value_counts().reset_index()
    vol.columns = ["source", "count"]
    vol["label"] = vol["source"].map(LABELS)
    vol["color"] = vol["source"].map(COLORS)

    fig = go.Figure(go.Pie(
        labels=vol["label"], values=vol["count"],
        marker_colors=vol["color"].tolist(),
        hole=0.45, textinfo="label+percent",
    ))
    fig.update_layout(
        height=270, margin=dict(l=10, r=10, t=10, b=10),
        showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 3 — Note par grade + Note par statut employé
# ─────────────────────────────────────────────────────────────────────────────
col_e, col_f = st.columns(2)

with col_e:
    st.markdown("## Note par grade")
    grade_note = (
        df[df["grade"] != "Unknown"]
        .groupby(["grade", "source"])["rating"].mean().reset_index()
    )
    grade_note["label"] = grade_note["source"].map(LABELS)
    grade_avg = grade_note.groupby("grade")["rating"].mean().sort_values(ascending=True)

    fig = px.bar(
        grade_note, x="rating", y="grade", color="label",
        orientation="h", barmode="group", color_discrete_map=COLOR_MAP,
        category_orders={"grade": grade_avg.index.tolist()},
        labels={"rating": "Note moy.", "grade": "", "label": ""},
        height=300,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.25), xaxis=dict(range=[0, 5.5]),
    )
    st.plotly_chart(fig, use_container_width=True)

with col_f:
    st.markdown("## Note selon le statut employé")
    statut_note = df.groupby(["statut", "source"])["rating"].mean().reset_index()
    statut_note["label"] = statut_note["source"].map(LABELS)

    fig = px.bar(
        statut_note, x="statut", y="rating", color="label",
        barmode="group", color_discrete_map=COLOR_MAP,
        labels={"rating": "Note moy.", "statut": "Statut", "label": ""},
        height=300,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.25), yaxis=dict(range=[0, 5.5]),
    )
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 4 — Analyse thématique NLP pré-annotée
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 🗂️ Analyse thématique — Thèmes annotés")
st.caption("Taux de mention (%) par thème identifié dans les verbatims")

def compute_themes(theme_dict, df_filtered):
    records = []
    for ent in entreprises:
        sub = df_filtered[df_filtered["source"] == ent]
        for theme_label, col in theme_dict.items():
            if col in sub.columns:
                records.append({
                    "Cabinet": LABELS[ent],
                    "Thème": theme_label,
                    "Taux mention (%)": (sub[col] > 0).mean() * 100,
                })
    return pd.DataFrame(records)

tab_pro, tab_con = st.tabs(["✅ Points positifs", "❌ Points négatifs"])

with tab_pro:
    p1, p2 = st.columns(2)
    with p1:
        st.markdown("### Thèmes principaux")
        df_tp = compute_themes(THEME_PRO_COLS, df)
        fig = px.bar(df_tp, x="Thème", y="Taux mention (%)", color="Cabinet",
                     barmode="group", color_discrete_map=COLOR_MAP, height=280)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10),
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                          legend=dict(orientation="h", y=-0.35), xaxis_tickangle=-10)
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        st.markdown("### Sous-thèmes détaillés")
        df_stp = compute_themes(SOUS_THEME_PRO_COLS, df)
        fig = px.bar(df_stp, x="Thème", y="Taux mention (%)", color="Cabinet",
                     barmode="group", color_discrete_map=COLOR_MAP, height=280)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10),
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                          legend=dict(orientation="h", y=-0.4), xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)

with tab_con:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Thèmes principaux")
        df_tc = compute_themes(THEME_CON_COLS, df)
        fig = px.bar(df_tc, x="Thème", y="Taux mention (%)", color="Cabinet",
                     barmode="group", color_discrete_map=COLOR_MAP, height=280)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10),
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                          legend=dict(orientation="h", y=-0.35), xaxis_tickangle=-10)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown("### Sous-thèmes détaillés")
        df_stc = compute_themes(SOUS_THEME_CON_COLS, df)
        fig = px.bar(df_stc, x="Thème", y="Taux mention (%)", color="Cabinet",
                     barmode="group", color_discrete_map=COLOR_MAP, height=280)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10),
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                          legend=dict(orientation="h", y=-0.4), xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 5 — NLP verbatims bruts
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 🔤 Analyse lexicale des verbatims")

def extract_keywords(series, top_n=15):
    text = " ".join(series.dropna().astype(str).str.lower())
    words = re.findall(r"[a-zA-ZÀ-ÿ]{4,}", text)
    freq = Counter(w for w in words if w not in STOPWORDS)
    return pd.DataFrame(freq.most_common(top_n), columns=["mot", "count"])

tab_np, tab_nc = st.tabs(["✅ Verbatims Pros", "❌ Verbatims Cons"])

for tab_nlp, col_text in [(tab_np, "pros"), (tab_nc, "cons")]:
    with tab_nlp:
        cols_nlp = st.columns(len(entreprises))
        for idx, ent in enumerate(entreprises):
            with cols_nlp[idx]:
                kw = extract_keywords(df[df["source"] == ent][col_text])
                fig = px.bar(kw, x="count", y="mot", orientation="h",
                             color_discrete_sequence=[COLORS[ent]],
                             labels={"count": "", "mot": ""},
                             title=LABELS[ent], height=340)
                fig.update_layout(
                    margin=dict(l=10, r=10, t=36, b=10),
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    yaxis=dict(autorange="reversed"), showlegend=False,
                )
                st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 6 — Top villes + Durée d'emploi vs note
# ─────────────────────────────────────────────────────────────────────────────
col_g, col_h = st.columns(2)

with col_g:
    st.markdown("## 📍 Top 10 villes")
    top_loc = (
        df.dropna(subset=["location"])
        .groupby(["location", "source"]).size().reset_index(name="count")
    )
    top10_loc = df["location"].value_counts().head(10).index.tolist()
    top_loc = top_loc[top_loc["location"].isin(top10_loc)]
    top_loc["label"] = top_loc["source"].map(LABELS)
    order = top_loc.groupby("location")["count"].sum().sort_values(ascending=True).index.tolist()

    fig = px.bar(
        top_loc, x="count", y="location", color="label",
        orientation="h", color_discrete_map=COLOR_MAP,
        category_orders={"location": order},
        labels={"count": "Nb d'avis", "location": "", "label": ""},
        height=320,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.2),
    )
    st.plotly_chart(fig, use_container_width=True)

with col_h:
    st.markdown("## ⏱️ Note selon l'ancienneté")
    dur_map = {0.5: "< 1 an", 2.0: "1–3 ans", 4.0: "3–5 ans",
               6.5: "5–8 ans", 9.0: "8–11 ans"}
    df_dur = df.copy()
    df_dur["durée"] = df_dur["employment_duration"].map(dur_map)
    df_dur = df_dur.dropna(subset=["durée"])
    dur_note = df_dur.groupby(["durée", "source"])["rating"].mean().reset_index()
    dur_note["label"] = dur_note["source"].map(LABELS)
    dur_order = ["< 1 an", "1–3 ans", "3–5 ans", "5–8 ans", "8–11 ans"]

    fig = px.line(
        dur_note, x="durée", y="rating", color="label", markers=True,
        color_discrete_map=COLOR_MAP,
        category_orders={"durée": dur_order},
        labels={"rating": "Note moy.", "durée": "Ancienneté", "label": ""},
        height=320,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.2), yaxis=dict(range=[1, 5.5]),
    )
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 7 — Tableau synthétique
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 📋 Tableau comparatif synthétique")

summary_rows = []
for ent in entreprises:
    sub = df[df["source"] == ent]
    top_loc_val = sub["location"].value_counts().index[0] if sub["location"].notna().any() else "—"
    top_g = sub[sub["grade"] != "Unknown"]["grade"].value_counts()
    top_grade_val = top_g.index[0] if len(top_g) > 0 else "—"
    summary_rows.append({
        "Cabinet":         LABELS[ent],
        "Nb avis":         f"{len(sub):,}",
        "Note moy.":       f"{sub['rating'].mean():.2f} ⭐",
        "Avis ≥ 4★":      f"{(sub['rating'] >= 4).mean()*100:.1f}%",
        "Avis ≤ 2★":      f"{(sub['rating'] <= 2).mean()*100:.1f}%",
        "Top ville":       top_loc_val,
        "Grade dominant":  top_grade_val,
    })

st.dataframe(
    pd.DataFrame(summary_rows).set_index("Cabinet"),
    use_container_width=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# ROW 8 — Avis bruts
# ─────────────────────────────────────────────────────────────────────────────
with st.expander("🔎 Explorer les avis bruts", expanded=False):
    cols_show = ["label", "rating", "grade", "grade_hierarchical",
                 "statut", "location", "title", "pros", "cons", "date"]
    sample_df = (
        df[cols_show]
        .rename(columns={
            "label": "Cabinet", "rating": "Note", "grade": "Grade",
            "grade_hierarchical": "Niveau", "statut": "Statut",
            "location": "Ville", "title": "Titre", "date": "Date",
        })
        .sample(min(300, len(df)))
        .reset_index(drop=True)
    )
    st.dataframe(sample_df, height=400, use_container_width=True)
    st.caption(f"Échantillon aléatoire de {len(sample_df)} avis parmi {len(df):,}")
