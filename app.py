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

# Inicjalizacja oficjalnych wyników
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.players = ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR']
    
    # Generalka (stan faktyczny)
    st.session_state.history = {
        "DAN": [20, 20], "RDX": [18, 14], "SIW": [12, 16], "BĄB": [14, 12],
        "JAC": [16, 9],  "KRO": [0, 18],  "PAW": [7, 10],  "PYR": [9, 6],
        "SZP": [10, 3],  "DOM": [8, 0],   "CYG": [0, 8],   "DAR": [0, 7]
    }
    
    # PEŁNE DANE ARCHIWALNE BIEGÓW (Odwzorowanie 1:1 z Excela)
    st.session_state.heats_archive = {
        1: pd.DataFrame({
            "Bieg": ["Bieg 1", "Bieg 2", "Bieg 3", "Bieg 4", "Bieg 5"],
            "JAC": [5, 4, 3, 2, 2], "DAR": [0, 0, 0, 0, 0], "CYG": [0, 0, 0, 0, 0],
            "PAW": [2, 1, 2, 1, 1], "RDX": [8, 10, 8, 6, 7], "KRO": [0, 0, 0, 0, 0],
            "SIW": [4, 3, 4, 3, 4], "BĄB": [6, 7, 5, 4, 5], "SZP": [3, 2, 6, 5, 3],
            "PYR": [1, 5, 1, 7, 6], "DAN": [10, 8, 10, 10, 10], "DOM": [7, 6, 7, 0, 0]
        }),
        2: pd.DataFrame({
            "Bieg": ["Bieg 1", "Bieg 2", "Bieg 3", "Bieg 4", "Bieg 5"],
            "JAC": [1, 1, 9, 10, 3], "DAR": [7, 5, 5, 3, 0], "CYG": [4, 4, 3, 9, 1],
            "PAW": [6, 0, 6, 5, 8],  "RDX": [10, 7, 2, 0, 10], "KRO": [8, 8, 7, 8, 2],
            "SIW": [3, 6, 10, 6, 6], "BĄB": [0, 10, 1, 7, 7], "SZP": [2, 2, 4, 1, 4],
            "PYR": [5, 3, 0, 2, 5],  "DAN": [9, 9, 8, 4, 9]
        })
    }

def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

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
        
        # Wyświetlanie czystej tabeli bez indeksów
        st.dataframe(df_live, use_container_width=True, hide_index=True)
        
        if st.button("💾 ZAPISZ OFICJALNE WYNIKI RUNDY"):
            for p in st.session_state.players:
                if p in df_live["Zawodnik"].values:
                    wywalczone = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                    st.session_state.history[p].append(wywalczone)
                else:
                    st.session_state.history[p].append(0)
            
            # Zapis pełnej macierzy do archiwum
            live_matrix = {"Bieg": ["Bieg 1", "Bieg 2", "Bieg 3", "Bieg 4", "Bieg 5"]}
            for p in active_today:
                live_matrix[p] = scores[p]
            st.session_state.heats_archive[nr_rundy] = pd.DataFrame(live_matrix)
            
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
    
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA PUNKTÓW']), use_container_width=True, hide_index=True)

# --- TAB 2: HISTORIA RUND (1-12) ---
with tab2:
    st.header("📚 Archiwum Przebiegu Poszczególnych Rund")
    
    wybrana_runda = st.selectbox("Wybierz numer rundy:", options=sorted(list(st.session_state.heats_archive.keys())), format_func=lambda x: f"Runda {x}")
    
    if wybrana_runda in st.session_state.heats_archive:
        st.write(f"### 🏁 Pełna tabela punktowa – Runda {wybrana_runda}")
        
        # Pobieramy bazową tabelę biegów
        df_arch = st.session_state.heats_archive[wybrana_runda].copy()
        
        # Dynamicznie wyliczamy podsumowanie dokładnie tak jak w Excelu!
        player_cols = [c for c in df_arch.columns if c != "Bieg"]
        
        sums = ["Suma punktów"]
        averages = ["Średnia na bieg"]
        ranks = ["Miejsce"]
        t_points = ["Punkty Turniejowe"]
        
        # Wyliczenia dla każdego gracza
        player_totals = {}
        for p in player_cols:
            s_val = df_arch[p].sum()
            player_totals[p] = s_val
            sums.append(s_val)
            averages.append(round(df_arch[p].mean(), 1))
            
        # Wyliczenie miejsc (od najwyższej sumy)
        sorted_players_by_sum = sorted(player_totals.items(), key=lambda x: x[1], reverse=True)
        player_ranks = {}
        for rank_idx, (p, _) in enumerate(sorted_players_by_sum):
            player_ranks[p] = rank_idx + 1
            
        for p in player_cols:
            rk = player_ranks[p]
            ranks.append(rk)
            t_points.append(get_tournament_points(rk))
            
        # Dołączamy wiersze podsumowujące na dół tabeli
        df_extra = pd.DataFrame(columns=df_arch.columns)
        df_extra.loc[len(df_extra)] = sums
        df_extra.loc[len(df_extra)] = averages
        df_extra.loc[len(df_extra)] = ranks
        df_extra.loc[len(df_extra)] = t_points
        
        df_full_display = pd.concat([df_arch, df_extra], ignore_index=True)
        
        # Wyświetlamy kompletną, niepoobcinaną tabelę bez szarych indeksów
        st.dataframe(df_full_display, use_container_width=True, hide_index=True)
