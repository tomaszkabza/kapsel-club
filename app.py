import streamlit as st
import pandas as pd
import numpy as np

# Konfiguracja pod telefon
st.set_page_config(page_title="Kapsel Club Browar", layout="centered")

# Barwy klubowe
st.markdown("""
    <style>
    .main { background-color: #FFFFFF; }
    h1, h2, h3 { color: #2E7D32 !important; font-family: 'Segoe UI', sans-serif; }
    .stButton>button { background-color: #2E7D32; color: white; border-radius: 5px; }
    .stButton>button:hover { background-color: #1B5E20; color: white; }
    div[data-testid="stDataFrame"] { background-color: #FFF9C4; }
    </style>
""", unsafe_allow_html=True)

st.title("🏆 Kapsel Club Browar")
st.subheader("Oficjalny Panel Live • Puchar Lata 2026")

# Inicjalizacja bazy i historii
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.players = ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR']
    
    # Oficjalne punkty turniejowe za R1 i R2
    st.session_state.history = {
        "DAN": [20, 20], "RDX": [18, 14], "SIW": [12, 16], "BĄB": [14, 12],
        "JAC": [16, 9],  "KRO": [0, 18],  "PAW": [7, 10],  "PYR": [9, 6],
        "SZP": [10, 3],  "DOM": [8, 0],   "CYG": [0, 8],   "DAR": [0, 7]
    }
    
    # BAZA HISTORII BIEGÓW (Dane szczegółowe wyścigów dla poprzednich rund)
    st.session_state.heats_archive = {
        1: [
            {"Bieg": "Bieg 1", "1. Miejsce": "DAN", "2. Miejsce": "JAC", "3. Miejsce": "BĄB", "4. Miejsce": "SIW"},
            {"Bieg": "Bieg 2", "1. Miejsce": "RDX", "2. Miejsce": "DAN", "3. Miejsce": "SZP", "4. Miejsce": "JAC"},
            {"Bieg": "Bieg 3", "1. Miejsce": "DAN", "2. Miejsce": "RDX", "3. Miejsce": "BĄB", "4. Miejsce": "DOM"},
            {"Bieg": "Bieg 4", "1. Miejsce": "JAC", "2. Miejsce": "SIW", "3. Miejsce": "DAN", "4. Miejsce": "PAW"},
            {"Bieg": "Bieg 5", "1. Miejsce": "DAN", "2. Miejsce": "RDX", "3. Miejsce": "SIW", "4. Miejsce": "BĄB"},
        ],
        2: [
            {"Bieg": "Bieg 1", "1. Miejsce": "KRO", "2. Miejsce": "SIW", "3. Miejsce": "DAN", "4. Miejsce": "RDX"},
            {"Bieg": "Bieg 2", "1. Miejsce": "DAN", "2. Miejsce": "KRO", "3. Miejsce": "RDX", "4. Miejsce": "PAW"},
            {"Bieg": "Bieg 3", "1. Miejsce": "SIW", "2. Miejsce": "DAN", "3. Miejsce": "BĄB", "4. Miejsce": "JAC"},
            {"Bieg": "Bieg 4", "1. Miejsce": "KRO", "2. Miejsce": "RDX", "3. Miejsce": "SIW", "4. Miejsce": "CYG"},
            {"Bieg": "Bieg 5", "1. Miejsce": "DAN", "2. Miejsce": "KRO", "3. Miejsce": "BĄB", "4. Miejsce": "PAW"},
        ]
    }

def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

# DOKŁADNIE TAKIE ZAKŁADKI JAK NA ZDJĘCIU
tab1, tab2 = st.tabs(["🏠 STRONA GŁÓWNA (LIVE & GENERALNA)", "📚 HISTORIA RUND (1-12)"])

# --- TAB 1: STRONA GŁÓWNA ---
with tab1:
    st.header("⚡ Aktualna Runda na Żywo")
    
    st.write("### ➕ Dopisz nowego zawodnika z palca:")
    quick_col1, quick_col2 = st.columns([2, 1])
    with quick_col1:
        new_p_name = st.text_input("Inicjały nowego (np. TOM):").upper().strip()
    with quick_col2:
        st.write("##") 
        if st.button("Dodaj do listy"):
            if new_p_name and new_p_name not in st.session_state.players:
                st.session_state.players.append(new_p_name)
                ile_rund = len(list(st.session_state.history.values())[0])
                st.session_state.history[new_p_name] = [0] * ile_rund
                st.success(f"Dodano {new_p_name}!")
                st.rerun()

    st.write("---")
    obecna_ilosc_rund = len(list(st.session_state.history.values())[0])
    nr_rundy = st.number_input("Numer rozgrywanej rundy", min_value=1, max_value=12, value=obecna_ilosc_rund + 1)
    
    st.write("**Zaznacz zawodników startujących dzisiaj:**")
    active_today = []
    cols = st.columns(4)
    for idx, p in enumerate(st.session_state.players):
        with cols[idx % 4]:
            if st.checkbox(p, key=f"active_{p}", value=True):
                active_today.append(p)
                
    if len(active_today) > 0:
        st.write("---")
        st.write("### Wpisz wyniki biegów (0 - 10):")
        
        scores = {}
        for p in active_today:
            st.write(f"**Zawodnik: {p}**")
            p_cols = st.columns(5)
            p_scores = []
            for b in range(5):
                with p_cols[b]:
                    val = st.number_input(f"Bieg {b+1}", min_value=0, max_value=10, value=0, key=f"score_{p}_{b}")
                    p_scores.append(val)
            scores[p] = p_scores
            
        st.write("---")
        st.write("### 📊 Wyniki Rundy na Żywo")
        
        live_rows = []
        for p, b_vals in scores.items():
            suma = sum(b_vals)
            srednia = round(np.mean(b_vals), 1)
            live_rows.append({"Zawodnik": p, "Suma": suma, "Średnia": srednia})
            
        df_live = pd.DataFrame(live_rows)
        df_live = df_live.sort_values(by="Suma", ascending=False).reset_index(drop=True)
        df_live.index += 1
        df_live.insert(0, 'Miejsce', df_live.index)
        df_live["Pkt Turniejowe"] = df_live["Miejsce"].apply(get_tournament_points)
        
        st.dataframe(df_live, use_container_width=True)
        
        # Zestawienie biegów na żywo dla aktualnego piątku
        st.write("### 🏁 Zestawienie biegów tej rundy (Na Żywo):")
        heat_summary = []
        for b in range(5):
            sorted_players = sorted(scores.keys(), key=lambda x: scores[x][b], reverse=True)
            row_data = {"Bieg": f"Bieg {b+1}"}
            for pos, pl in enumerate(sorted_players[:6]):
                row_data[f"{pos+1}. Miejsce"] = f"{pl} ({scores[pl][b]} pkt)"
            heat_summary.append(row_data)
        
        df_heats = pd.DataFrame(heat_summary)
        st.dataframe(df_heats, use_container_width=True)
        
        if st.button("💾 ZAPISZ OFICJALNE WYNIKI RUNDY"):
            # Zapis punktów generalnych
            for p in st.session_state.players:
                if p in df_live["Zawodnik"].values:
                    wywalczone = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                    st.session_state.history[p].append(wywalczone)
                else:
                    st.session_state.history[p].append(0)
            
            # Zapis szczegółowego przebiegu biegów do archiwum
            st.session_state.heats_archive[nr_rundy] = heat_summary
            st.success(f"Pomyślnie zapisano i zamknięto Rundę {nr_rundy}!")
            st.rerun()

    st.write("---")
    st.header("🏆 Oficjalna Klasyfikacja Generalna Pucharu")
    
    gen_rows = []
    for p, rounds in st.session_state.history.items():
        total_suma = sum(rounds)
        row_dict = {"Zawodnik": p}
        for r_idx, r_pts in enumerate(rounds):
            if r_pts == 0 and p in ['KRO','CYG','DAR','DOM'] and r_idx < 2:
                row_dict[f"R{r_idx+1}"] = "-"
            elif r_pts == 0 and r_idx >= 2:
                row_dict[f"R{r_idx+1}"] = "-"
            else:
                row_dict[f"R{r_idx+1}"] = r_pts
        row_dict["SUMA PUNKTÓW"] = total_suma
        gen_rows.append(row_dict)
        
    df_gen = pd.DataFrame(gen_rows)
    df_gen = df_gen.sort_values(by="SUMA PUNKTÓW", ascending=False).reset_index(drop=True)
    df_gen.index += 1
    df_gen.insert(0, 'Poz.', df_gen.index)
    
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA PUNKTÓW']), use_container_width=True)

# --- TAB 2: HISTORIA RUND (1-12) ---
with tab2:
    st.header("📚 Archiwum Przebiegu Poszczególnych Rund")
    st.write("Wybierz rundę, aby zobaczyć szczegółowe wyniki poszczególnych wyścigów:")
    
    # Wybór rundy do podejrzenia historii
    wybrana_runda = st.selectbox("Wybierz numer rundy:", options=sorted(list(st.session_state.heats_archive.keys())), format_func=lambda x: f"Runda {x}")
    
    if wybrana_runda in st.session_state.heats_archive:
        st.write(f"### 🏁 Przebieg biegów – Runda {wybrana_runda}")
        df_arch_heats = pd.DataFrame(st.session_state.heats_archive[wybrana_runda])
        st.dataframe(df_arch_heats, use_container_width=True)
    else:
        st.info("Brak szczegółowych danych dla tej rundy.")
