import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

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
st.subheader("Oficjalny Panel • Puchar Lata 2026")

EXCEL_FILE = "Puchar_Lata_2026_Browar.xlsx"

def znajdz_wiersz_po_tekscie(ws, tekst):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=2).value
        if val and tekst in str(val):
            return r
    return None

# Ładowanie danych z pliku Excel
@st.cache_data(ttl=10)
def wczytaj_baze_graczy():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Puchar Lata 2026"]
        idx_gen = znajdz_wiersz_po_tekscie(ws, "KLASYFIKACJA GENERALNA")
        
        gracze = []
        historia = {}
        wiersz_nag = idx_gen + 1 if idx_gen else 30
        
        for r in range(wiersz_nag + 1, wiersz_nag + 16):
            zawodnik = ws.cell(row=r, column=3).value
            if zawodnik and str(zawodnik).strip() != "":
                n_gracza = str(zawodnik).strip()
                gracze.append(n_gracza)
                
                # Odczyt historycznych punktów z pierwszych rund
                r1_val = ws.cell(row=r, column=4).value
                r2_val = ws.cell(row=r, column=5).value
                
                r1 = int(r1_val) if r1_val and str(r1_val) != "-" else 0
                r2 = int(r2_val) if r2_val and str(r2_val) != "-" else 0
                historia[n_gracza] = [r1, r2]
                
        return gracze, historia
    except:
        lista_awaryjna = ['DAN', 'RDX', 'SIW', 'BĄB', 'JAC', 'KRO', 'PAW', 'PYR', 'SZP', 'DOM', 'CYG', 'DAR']
        historia_awaryjna = {g: [20, 20] if g=='DAN' else [0,0] for g in lista_awaryjna} # placeholder
        return lista_awaryjna, historia_awaryjna

players_list, history_dict = wczytaj_baze_graczy()

if "players" not in st.session_state:
    st.session_state.players = players_list
    st.session_state.history = history_dict

def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

# --- NOWY UKŁAD GŁÓWNYCH ZAKŁADEK ---
tab_glowna, tab_historia = st.tabs(["🏠 STRONA GŁÓWNA (LIVE & GENERALNA)", "📚 HISTORIA RUND (1-12)"])

# ==========================================
# 🏠 ZAKŁADKA GŁÓWNA: AKTUALNA RUNDA + GENERALNA
# ==========================================
with tab_glowna:
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
                st.session_state.history[new_p_name] = [0, 0]
                st.success(f"Dodano {new_p_name}!")
                st.rerun()

    st.write("---")
    nr_rundy = st.number_input("Numer rozgrywanej rundy", min_value=1, max_value=100, value=3)
    
    st.write("**Zaznacz zawodników startujących dzisiaj:**")
    active_today = []
    cols = st.columns(4)
    for idx, p in enumerate(st.session_state.players):
        with cols[idx % 4]:
            if st.checkbox(p, key=f"active_{p}", value=True):
                active_today.append(p)
                
    if len(active_today) > 0:
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
            
        st.write("### 📊 Wyniki Bieżącej Rundy na Żywo")
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
        
        if st.button("💾 GENERUJ UZUPEŁNIONY PLIK EXCEL"):
            wb_mod = openpyxl.load_workbook(EXCEL_FILE)
            ws_mod = wb_mod["Puchar Lata 2026"]
            
            start_r = 48 + ((nr_rundy - 3) * 12)
            
            CLUB_GREEN_DARK = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            CLUB_YELLOW_LIGHT = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
            CLUB_GREEN_LIGHT = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
            WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            GRAY_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
            FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            FONT_BODY_BLACK = Font(name="Segoe UI", size=10, color="000000")
            FONT_BOLD_BLACK = Font(name="Segoe UI", size=10, bold=True, color="000000")
            THIN_BORDER = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            
            ws_mod.cell(row=start_r, column=2, value=f"Runda {nr_rundy} • Wyniki Live").font = Font(name="Segoe UI", size=11, bold=True)
            
            header_row = start_r + 1
            ws_mod.cell(row=header_row, column=2, value="Bieg").font = FONT_HEADER
            ws_mod.cell(row=header_row, column=2).fill = CLUB_GREEN_DARK
            
            for c_idx, p in enumerate(active_today, start=3):
                cell = ws_mod.cell(row=header_row, column=c_idx, value=p)
                cell.font = FONT_HEADER
                cell.fill = CLUB_GREEN_DARK
                cell.border = THIN_BORDER
                
            for b_idx in range(5):
                curr_r = header_row + 1 + b_idx
                is_zebra = (curr_r % 2 == 0)
                f_col = CLUB_GREEN_LIGHT if is_zebra else WHITE_FILL
                c_run = ws_mod.cell(row=curr_r, column=2, value=f"Bieg {b_idx+1}")
                c_run.font = FONT_BOLD_BLACK
                c_run.fill = f_col
                c_run.border = THIN_BORDER
                
                for c_idx, p in enumerate(active_today, start=3):
                    val = scores[p][b_idx]
                    c_sc = ws_mod.cell(row=curr_r, column=c_idx, value=val)
                    c_sc.font = FONT_BODY_BLACK
                    c_sc.fill = f_col
                    c_sc.border = THIN_BORDER
                    
            r_sum = header_row + 6
            r_pts = header_row + 9
            
            ws_mod.cell(row=r_sum, column=2, value="Suma punktów").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_sum, column=2).fill = CLUB_YELLOW_LIGHT
            ws_mod.cell(row=r_pts, column=2, value="Punkty Turniejowe").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_pts, column=2).fill = CLUB_YELLOW_LIGHT
            
            for c_idx, p in enumerate(active_today, start=3):
                col_let = get_column_letter(c_idx)
                ws_mod.cell(row=r_sum, column=c_idx, value=f"=SUM({col_let}{header_row+1}:{col_let}{header_row+5})").font = FONT_BOLD_BLACK
                ws_mod.cell(row=r_sum, column=c_idx).fill = CLUB_YELLOW_LIGHT
                ws_mod.cell(row=r_sum, column=c_idx).border = THIN_BORDER
                
                pt_live = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                ws_mod.cell(row=r_pts, column=c_idx, value=pt_live).font = FONT_BOLD_BLACK
                ws_mod.cell(row=r_pts, column=c_idx).fill = CLUB_YELLOW_LIGHT
                ws_mod.cell(row=r_pts, column=c_idx).border = THIN_BORDER

            idx_gen_live = znajdz_wiersz_po_tekscie(ws_mod, "KLASYFIKACJA GENERALNA")
            w_nag_live = idx_gen_live + 1 if idx_gen_live else 30
            kolumna_rundy = 3 + nr_rundy
            
            for r in range(w_nag_live + 1, w_nag_live + 16):
                z_name = ws_mod.cell(row=r, column=3).value
                if z_name and str(z_name).strip() in df_live["Zawodnik"].values:
                    p_name = str(z_name).strip()
                    p_turniejowe = int(df_live[df_live["Zawodnik"] == p_name]["Pkt Turniejowe"].values[0])
                    ws_mod.cell(row=r, column=kolumna_rundy, value=p_turniejowe)
                elif z_name:
                    ws_mod.cell(row=r, column=kolumna_rundy, value="-")

            buffer = io.BytesIO()
            wb_mod.save(buffer)
            buffer.seek(0)
            
            st.success(f"🔥 Wygenerowano Excel dla Rundy {nr_rundy}!")
            st.download_button(
                label="📥 POBIERZ UZUPEŁNIONY PLIK EXCEL",
                data=buffer,
                file_name=f"Puchar_Lata_2026_Gotowy_R{nr_rundy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.write("---")
    st.header("👑 Bieżąca Klasyfikacja Generalna Sezonu")
    gen_rows = []
    for p, rounds in st.session_state.history.items():
        total_suma = sum(rounds)
        row_dict = {"Zawodnik": p, "SUMA": total_suma}
        gen_rows.append(row_dict)
        
    df_gen = pd.DataFrame(gen_rows)
    df_gen = df_gen.sort_values(by="SUMA", ascending=False).reset_index(drop=True)
    df_gen.index += 1
    df_gen.insert(0, 'Poz.', df_gen.index)
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA']), use_container_width=True)

# ==========================================
# 📚 ZAKŁADKA HISTORII: RUNDY 1 DO 12 (SUB-TABS)
# ==========================================
with tab_historia:
    st.header("🗂️ Archiwum Rozegranych Rund")
    st.write("Wybierz interesującą Cię rundę z zakładek poniżej, aby podejrzeć oficjalne punkty turniejowe:")
    
    # Tworzymy 12 podzakładek wewnątrz historii
    sub_tabs = st.tabs([f"Runda {i}" for i in range(1, 13)])
    
    for idx, s_tab in enumerate(sub_tabs, start=1):
        with s_tab:
            st.subheader(f"📋 Wyniki: Runda {idx}")
            
            hist_rows = []
            for p, rounds in st.session_state.history.items():
                # Sprawdzamy, czy mamy dane dla tej rundy w historii sesji
                if idx <= len(rounds):
                    punkty_z_rundy = rounds[idx-1]
                    hist_rows.append({
                        "Zawodnik": p, 
                        "Zdobyte Punkty Turniejowe": punkty_z_rundy if punkty_z_rundy > 0 else "-"
                    })
                else:
                    hist_rows.append({"Zawodnik": p, "Zdobyte Punkty Turniejowe": "-"})
                    
            df_hist = pd.DataFrame(hist_rows)
            # Pokazujemy tylko tych, którzy faktycznie dostali punkty w danej rundzie, dla czytelności
            df_hist_filtered = df_hist[df_hist["Zdobyte Punkty Turniejowe"] != "-"].sort_values(by="Zdobyte Punkty Turniejowe", ascending=False).reset_index(drop=True)
            
            if not df_hist_filtered.empty:
                df_hist_filtered.index += 1
                st.dataframe(df_hist_filtered, use_container_width=True)
            else:
                st.info(f"Runda {idx} nie została jeszcze rozegrana lub brak zapisanych danych.")
