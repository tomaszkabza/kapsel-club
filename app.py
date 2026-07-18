import streamlit as st
import pandas as pd
import numpy as np

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

# Inicjalizacja bazy danych w pamięci sesji (stan początkowy z pliku Excel)
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    # Lista aktualnych zawodników z Twojego pliku
    st.session_state.players = ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR']
    
    # Historia punktów turniejowych z R1 i R2 (stan faktyczny po korekcie)
    st.session_state.history = {
        "DAN": [20, 20], "RDX": [18, 14], "SIW": [12, 16], "BĄB": [14, 12],
        "JAC": [16, 9],  "KRO": [0, 18],  "PAW": [7, 10],  "PYR": [9, 6],
        "SZP": [10, 3],  "DOM": [8, 0],   "CYG": [0, 8],   "DAR": [0, 7]
    }

# Funkcja przyznawania punktów turniejowych (odpowiednik CHOOSE z Excela)
def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

# Zakładki menu aplikacji
tab1, tab2, tab3 = st.tabs(["🎮 RUNDA LIVE (PIĄTEK)", "📊 KLASYFIKACJA GENERALNA", "➕ ZARZĄDZANIE"])

# --- TAB 1: RUNDA LIVE ---
with tab1:
    st.header("Zapisy i Wyniki Bieżącej Rundy")
    
    # Wybór numeru rundy
    nr_rundy = st.number_input("Numer rozgrywanej rundy (np. 3)", min_value=3, max_value=12, value=3)
    
    # Wybór obecnych w ten piątek
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
        
        # Tworzenie siatki do wprowadzania danych z palca
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
            
        # Kalkulacja na Żywo
        st.write("---")
        st.write("### Wyniki Rundy na Żywo")
        
        live_rows = []
        for p, b_vals in scores.items():
            suma = sum(b_vals)
            srednia = round(np.mean(b_vals), 1)
            live_rows.append({"Zawodnik": p, "Suma": suma, "Średnia": srednia})
            
        df_live = pd.DataFrame(live_rows)
        # Wyliczanie miejsc na podstawie sumy punktów
        df_live = df_live.sort_values(by="Suma", ascending=False).reset_index(drop=True)
        df_live.index += 1
        df_live.insert(0, 'Miejsce', df_live.index)
        
        # Przypisanie Punktów Turniejowych na żywo
        df_live["Pkt Turniejowe"] = df_live["Miejsce"].apply(get_tournament_points)
        
        # Wyświetlenie tabeli live w kolorach
        st.dataframe(df_live.style.background_gradient(cmap="Greens", subset=["Suma"]), use_container_width=True)
        
        # Zapis wyników do bazy
        if st.button("💾 ZAPISZ OFICJALNE WYNIKI RUNDY"):
            for p in st.session_state.players:
                # Jeśli gracz startował - dopisz mu wywalczone pkt, jeśli nie - dopisz 0
                if p in df_live["Zawodnik"].values:
                    wywalczone = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                    st.session_state.history[p].append(wywalczone)
                else:
                    st.session_state.history[p].append(0)
            st.success(f"Pomyślnie zapisano i zamknięto Rundę {nr_rundy}!")
    else:
        st.warning("Zaznacz przynajmniej jednego zawodnika, aby rozpocząć wpisywanie rzutów.")

# --- TAB 2: KLASYFIKACJA GENERALNA ---
with tab2:
    st.header("Oficjalna Klasyfikacja Generalna")
    st.write("Zlicza automatycznie wszystkie zamknięte rundy.")
    
    gen_rows = []
    for p, rounds in st.session_state.history.items():
        total_suma = sum(rounds)
        row_dict = {"Zawodnik": p}
        # Dodaj historyczne rundy
        for r_idx, r_pts in enumerate(rounds):
            row_dict[f"R{r_idx+1}"] = r_pts if r_pts > 0 or p in ['KRO','CYG','DAR','DOM'] else "-" 
        row_dict["SUMA PUNKTÓW"] = total_suma
        gen_rows.append(row_dict)
        
    df_gen = pd.DataFrame(gen_rows)
    # Zastąp ewentualne braki kresek znakami pustymi dla estetyki
    df_gen = df_gen.fillna("-")
    
    # Sortowanie generalne od lidera
    df_gen = df_gen.sort_values(by="SUMA PUNKTÓW", ascending=False).reset_index(drop=True)
    df_gen.index += 1
    df_gen.insert(0, 'Poz.', df_gen.index)
    
    # Wyświetlanie tabeli w barwach klubowych (Żółta kolumna SUMA)
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA PUNKTÓW']), use_container_width=True)

# --- TAB 3: ZARZĄDZANIE ---
with tab3:
    st.header("Panel Administratora")
    new_player = st.text_input("Imię/Inicjały nowego zawodnika (np. TOM):").upper().strip()
    if st.button("➕ DODAJ NOWEGO ZAWODNIKA DO KLUBU"):
        if new_player and new_player not in st.session_state.players:
            st.session_state.players.append(new_player)
            # Dopisujemy mu historyczne zera, żeby zgadzała się liczba rund
            ile_rund_bylo = len(list(st.session_state.history.values())[0])
            st.session_state.history[new_player] = [0] * ile_rund_bylo
            st.success(f"Dodano gracza {new_player} do listy startowej!")
        else:
            st.error("Taki zawodnik już istnieje lub pole jest puste.")
