import streamlit as st
import pandas as pd
import numpy as np
import openpyxl

# Konfiguracja strony pod smartfona
st.set_page_config(page_title="Kapsel Club Browar", layout="centered")

# Stylizacja CSS - Barwy klubowe (Żółto-Biało-Zielone)
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

# FUNKCJA: Inteligentne wczytywanie bazy z pliku Excel (szukanie nagłówka w pionie)
def load_data_from_excel(filename="Puchar_Lata_2026_Browar_Korekta_R1.xlsx"):
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
        ws = wb["Puchar Lata 2026"]
        
        # Szukamy wiersza z Klasyfikacją Generalną w Kolumnie B
        gen_row_idx = None
        for row in range(1, 200):
            val = ws.cell(row=row, column=2).value
            if val and "KLASYFIKACJA GENERALNA" in str(val):
                gen_row_idx = row
                break
        
        if not gen_row_idx:
            # Awaryjny stan początkowy, jeśli plik byłby pusty
            return ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR'], {}

        # Czytamy zawodników i ich dotychczasowe punkty (zaczynamy 2 wiersze pod nagłówkiem)
        players_list = []
        history_dict = {}
        
        # Ustalamy ile rund już rozegrano na podstawie nagłówków (Kolumny D, E, F...)
        header_row = gen_row_idx + 1
        round_count = 0
        for col in range(4, 16):
            h_val = ws.cell(row=header_row, column=col).value
            if h_val and str(h_val).startswith("R"):
                round_count += 1
        
        # Wczytujemy wiersze zawodników
        for r in range(header_row + 1, header_row + 20):
            p_name = ws.cell(row=r, column=3).value
            if p_name and str(p_name).strip() != "":
                p_name = str(p_name).strip()
                players_list.append(p_name)
                
                # Pobieramy punkty z rozegranych rund
                p_rounds = []
                for c in range(4, 4 + round_count):
                    pts = ws.cell(row=r, column=c).value
                    if pts is None or pts == "-":
                        p_rounds.append(0)
                    else:
                        try:
                            p_rounds.append(int(pts))
                        except:
                            p_rounds.append(0)
                history_dict[p_name] = p_rounds
                
        return players_list, history_dict
    except Exception as e:
        # Awaryjny powrót w przypadku braku dostępu do pliku
        return ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR'], {}

# Inicjalizacja bazy danych w pamięci sesji na podstawie Excela
if "initialized" not in st.session_state:
    st.session_state.players, st.session_state.history = load_data_from_excel()
    if not st.session_state.history:
        # Rezerwowy twardy zapis na wypadek pierwszego ładowania
        st.session_state.history = {
            "DAN": [20, 20], "RDX": [18, 14], "SIW": [12, 16], "BĄB": [14, 12],
            "JAC": [16, 9],  "KRO": [0, 18],  "PAW": [7, 10],  "PYR": [9, 6],
            "SZP": [10, 3],  "DOM": [8, 0],   "CYG": [0, 8],   "DAR": [0, 7]
        }
    st.session_state.initialized = True

def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

# Zakładki menu aplikacji
tab1, tab2 = st.tabs(["🎮 RUNDA LIVE (PIĄTEK)", "📊 KLASYFIKACJA GENERALNA"])

# --- TAB 1: RUNDA LIVE ---
with tab1:
    st.header("Zapisy i Wyniki Bieżącej Rundy")
    
    st.write("### ➕ Ktoś przyszedł ekstra? Dopisz go:")
    quick_col1, quick_col2 = st.columns([2, 1])
    with quick_col1:
        new_p_name = st.text_input("Wpisz inicjały nowego (np. TOM):", key="quick_add_input").upper().strip()
    with quick_col2:
        st.write("##") 
        if st.button("Dodaj do listy"):
            if new_p_name and new_p_name not in st.session_state.players:
                st.session_state.players.append(new_p_name)
                ile_rund = len(list(st.session_state.history.values())[0]) if st.session_state.history else 2
                st.session_state.history[new_p_name] = [0] * ile_rund
                st.success(f"Dodano {new_p_name}!")
                st.rerun()

    st.write("---")
    # Automatycznie podpowiada kolejny numer rundy
    obecna_ilosc_rund = len(list(st.session_state.history.values())[0]) if st.session_state.history else 2
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
        st.write("### Wyniki Rundy na Żywo")
        
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
        
        # --- NOWOŚĆ: Szczegółowe zestawienie jak odbywały się biegi ---
        st.write("### 🏁 Szczegółowy przebieg wyścigów w tej rundzie:")
        heat_summary = []
        for b in range(5):
            sorted_players = sorted(scores.keys(), key=lambda x: scores[x][b], reverse=True)
            row_data = {"Bieg": f"Bieg {b+1}"}
            for pos, pl in enumerate(sorted_players[:6]):  # Pokazujemy TOP 6 w danym biegu
                row_data[f"Miejsce {pos+1}"] = f"{pl} ({scores[pl][b]} pkt)"
            heat_summary.append(row_data)
        
        df_heats = pd.DataFrame(heat_summary)
        st.dataframe(df_heats, use_container_width=True)
        
        if st.button("💾 ZAPISZ OFICJALNE WYNIKI RUNDY"):
            for p in st.session_state.players:
                if p in df_live["Zawodnik"].values:
                    wywalczone = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                    st.session_state.history[p].append(wywalczone)
                else:
                    st.session_state.history[p].append(0)
            st.success(f"Pomyślnie zapisano i zamknięto Rundę {nr_rundy}!")
            st.rerun()

# --- TAB 2: KLASYFIKACJA GENERALNA ---
with tab2:
    st.header("Oficjalna Klasyfikacja Generalna")
    
    gen_rows = []
    for p, rounds in st.session_state.history.items():
        total_suma = sum(rounds)
        row_dict = {"Zawodnik": p}
        for r_idx, r_pts in enumerate(rounds):
            row_dict[f"R{r_idx+1}"] = r_pts if r_pts > 0 or p in ['KRO','CYG','DAR','DOM'] else "-" 
        row_dict["SUMA PUNKTÓW"] = total_suma
        gen_rows.append(row_dict)
        
    df_gen = pd.DataFrame(gen_rows)
    df_gen = df_gen.fillna("-")
    df_gen = df_gen.sort_values(by="SUMA PUNKTÓW", ascending=False).reset_index(drop=True)
    df_gen.index += 1
    df_gen.insert(0, 'Poz.', df_gen.index)
    
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA PUNKTÓW']), use_container_width=True)
