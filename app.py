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
st.subheader("Inteligentny Panel Live • Puchar Lata 2026")

# Nazwa Twojego pliku na GitHubie (obsługuje .xlsx)
EXCEL_FILE = "Puchar_Lata_2026_Browar.xlsx"

# Funkcja pomocnicza do skanowania arkusza w dół i szukania sekcji
def znajdz_wiersz_po_tekscie(ws, tekst):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=2).value
        if val and tekst in str(val):
            return r
    return None

# Wczytywanie aktualnego stanu z pliku Excel
@st.cache_data(ttl=60) # Odświeżaj z pliku co minutę
def wczytaj_dane_z_excela():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Puchar Lata 2026"]
        
        # 1. Szukamy gdzie zaczyna się Klasyfikacja Generalna
        idx_gen = znajdz_wiersz_po_tekscie(ws, "KLASYFIKACJA GENERALNA")
        if not idx_gen:
            return None, "Nie znaleziono sekcji Klasyfikacji Generalnej w pliku."
            
        # 2. Wyciągamy listę graczy i ich dotychczasowe punkty
        gracze = []
        historia = {}
        
        # Nagłówki rund (R1, R2 itd.) są w wierszu idx_gen + 1, sprawdzamy ile rund już rozegrano
        wiersz_naglowkow = idx_gen + 1
        aktywne_kolumny_rund = []
        for c in range(4, 16): # od kolumny D (R1) do O (R12)
            h_val = ws.cell(row=wiersz_naglowkow, column=c).value
            if h_val:
                aktywne_kolumny_rund.append(c)
                
        # Czytamy wiersze zawodników poniżej nagłówka (max 15 zawodników)
        for r in range(wiersz_naglowkow + 1, wiersz_naglowkow + 16):
            zawodnik = ws.cell(row=r, column=3).value
            if zawodnik and str(zawodnik).strip() != "":
                nazwa_gracza = str(zawodnik).strip()
                gracze.append(nazwa_gracza)
                
                # Zbieramy punkty z rund, które już fizycznie mają wpisane dane w Excelu
                punkty_rund = []
                for c_idx in range(4, 4 + len(aktywne_kolumny_rund)):
                    # Sprawdzamy prawdziwe formuły/wartości z pliku (używamy bazy bez data_only do zapisu)
                    val_cell = ws.cell(row=r, column=c_idx).value
                    if val_cell is None or val_cell == "-":
                        punkty_rund.append(0)
                    else:
                        try:
                            punkty_rund.append(int(val_cell))
                        except:
                            punkty_rund.append(0)
                historia[nazwa_gracza] = punkty_rund
                
        return {"players": gracze, "history": historia, "gen_start_row": idx_gen}, None
    except Exception as e:
        return None, f"Błąd wczytywania pliku Excel: {str(e)}"

# Pobranie danych startowych z Excela
db, err = wczytaj_dane_z_excela()
if err:
    st.error(err)
    st.info("Upewnij się, że plik 'Puchar_Lata_2026_Browar.xlsx' został poprawnie dodany na GitHubie obok app.py.")
    st.stop()

# Zachowanie listy graczy w sesji
if "players" not in st.session_state:
    st.session_state.players = db["players"]
    st.session_state.history = db["history"]

# Funkcja punktów turniejowych (CHOOSE)
def get_tournament_points(rank):
    pts_map = {1:20, 2:18, 3:16, 4:14, 5:12, 6:10, 7:9, 8:8, 9:7, 10:6, 11:5, 12:4, 13:3, 14:2, 15:1}
    return pts_map.get(rank, 0)

# Zakładki
tab1, tab2 = st.tabs(["🎮 RUNDA LIVE (PIĄTEK)", "📊 KLASYFIKACJA GENERALNA"])

# --- TAB 1: RUNDA LIVE ---
with tab1:
    st.header("Zapisy i Wyniki Bieżącej Rundy")
    
    st.write("### ➕ Dopisz nowego zawodnika z palca:")
    quick_col1, quick_col2 = st.columns([2, 1])
    with quick_col1:
        new_p_name = st.text_input("Inicjały nowego (np. TOM):", key="quick_add_input").upper().strip()
    with quick_col2:
        st.write("##")
        if st.button("Dodaj do listy"):
            if new_p_name and new_p_name not in st.session_state.players:
                st.session_state.players.append(new_p_name)
                ile_rund_bylo = len(list(st.session_state.history.values())[0]) if st.session_state.history else 2
                st.session_state.history[new_p_name] = [0] * ile_rund_bylo
                st.success(f"Dodano {new_p_name}!")
                st.rerun()

    st.write("---")
    # Ustalamy numer kolejnej rundy automatycznie na podstawie danych z Excela
    ile_rund_w_pliku = len(list(st.session_state.history.values())[0]) if st.session_state.history else 2
    r_rundy = st.number_input("Numer rozgrywanej rundy", min_value=1, max_value=100, value=3)
    
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
        
        st.dataframe(df_live.style.background_gradient(cmap="Greens", subset=["Suma"]), use_container_width=True)
        
        # --- ZAPIS I GENEROWANIE DYNAMICZNEGO EXCELA ---
        if st.button("💾 GENERUJ UZUPEŁNIONY PLIK EXCEL (PIONOWO)"):
            # Otwieramy oryginalny skoroszyt (z zachowaniem styli)
            wb_mod = openpyxl.load_workbook(EXCEL_FILE)
            ws_mod = wb_mod["Puchar Lata 2026"]
            
            # 1. Znajdujemy gdzie jest generalka PRZED modyfikacją
            idx_gen_old = znajdz_wiersz_po_tekscie(ws_mod, "KLASYFIKACJA GENERALNA")
            
            # Kopiujemy całą tabelę klasyfikacji generalnej do pamięci, aby przesunąć ją w dół
            wiersze_generalki = []
            for r in range(idx_gen_old, ws_mod.max_row + 1):
                row_cells = []
                for c in range(1, 20):
                    cell = ws_mod.cell(row=r, column=c)
                    row_cells.append({
                        "col": c, "val": cell.value, "font": cell.font, 
                        "fill": cell.fill, "align": cell.alignment, "border": cell.border
                    })
                wiersze_generalki.append(row_cells)
                
            # Czyszczenie starego miejsca po recall klasyfikacji
            for r in range(idx_gen_old, ws_mod.max_row + 1):
                for c in range(1, 20):
                    ws_mod.cell(row=r, column=c, value=None).fill = openpyxl.styles.PatternFill(fill_type=None)
                    ws_mod.cell(row=r, column=c).border = openpyxl.styles.Border()
            
            # 2. Wstawiamy nową tabelę rundy dokładnie w miejsce starego nagłówka generalki
            start_r = idx_gen_old
            CLUB_GREEN_DARK = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            CLUB_YELLOW_LIGHT = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
            CLUB_GREEN_LIGHT = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
            WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            GRAY_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
            
            FONT_SUBTITLE = Font(name="Segoe UI", size=10, italic=True, color="595959")
            FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            FONT_BODY_BLACK = Font(name="Segoe UI", size=10, color="000000")
            FONT_BOLD_BLACK = Font(name="Segoe UI", size=10, bold=True, color="000000")
            THIN_BORDER = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            
            ws_mod.cell(row=start_r, column=2, value=f"Runda {nr_rundy} • Wyniki Live z Aplikacji").font = FONT_SUBTITLE
            
            header_row = start_r + 1
            ws_mod.cell(row=header_row, column=2, value="Bieg").font = FONT_HEADER
            ws_mod.cell(row=header_row, column=2).fill = CLUB_GREEN_DARK
            ws_mod.cell(row=header_row, column=2).alignment = Alignment(horizontal="center")
            
            # Nagłówki graczy aktywnych w rundzie
            for c_idx, p in enumerate(active_today, start=3):
                cell = ws_mod.cell(row=header_row, column=c_idx, value=p)
                cell.font = FONT_HEADER
                cell.fill = CLUB_GREEN_DARK
                cell.alignment = Alignment(horizontal="center")
                cell.border = THIN_BORDER
                
            # Wpisywanie wyników 5 biegów
            for b_idx in range(5):
                curr_r = header_row + 1 + b_idx
                is_zebra = (curr_r % 2 == 0)
                fill_color = CLUB_GREEN_LIGHT if is_zebra else WHITE_FILL
                
                c_run = ws_mod.cell(row=curr_r, column=2, value=f"Bieg {b_idx+1}")
                c_run.font = FONT_BOLD_BLACK
                c_run.fill = fill_color
                c_run.alignment = Alignment(horizontal="center")
                c_run.border = THIN_BORDER
                
                for c_idx, p in enumerate(active_today, start=3):
                    val = scores[p][b_idx]
                    c_sc = ws_mod.cell(row=curr_r, column=c_idx, value=val)
                    c_sc.font = FONT_BODY_BLACK
                    c_sc.fill = fill_color
                    c_sc.alignment = Alignment(horizontal="center")
                    c_sc.border = THIN_BORDER
                    
            # Sumy, średnie, miejsca i punkty turniejowe w Excelu
            r_sum = header_row + 6
            r_avg = header_row + 7
            r_rnk = header_row + 8
            r_pts = header_row + 9
            
            ws_mod.cell(row=r_sum, column=2, value="Suma punktów").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_sum, column=2).fill = CLUB_YELLOW_LIGHT
            ws_mod.cell(row=r_sum, column=2).border = THIN_BORDER
            
            ws_mod.cell(row=r_avg, column=2, value="Średnia na bieg").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_avg, column=2).fill = CLUB_YELLOW_LIGHT
            ws_mod.cell(row=r_avg, column=2).border = THIN_BORDER
            
            ws_mod.cell(row=r_rnk, column=2, value="Miejsce").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_rnk, column=2).fill = GRAY_FILL
            ws_mod.cell(row=r_rnk, column=2).border = THIN_BORDER
            
            ws_mod.cell(row=r_pts, column=2, value="Punkty Turniejowe").font = FONT_BOLD_BLACK
            ws_mod.cell(row=r_pts, column=2).fill = CLUB_YELLOW_LIGHT
            ws_mod.cell(row=r_pts, column=2).border = THIN_BORDER
            
            for c_idx, p in enumerate(active_today, start=3):
                col_let = get_column_letter(c_idx)
                
                ws_mod.cell(row=r_sum, column=c_idx, value=f"=SUM({col_let}{header_row+1}:{col_let}{header_row+5})").font = FONT_BOLD_BLACK
                ws_mod.cell(row=r_sum, column=c_idx).fill = CLUB_YELLOW_LIGHT
                ws_mod.cell(row=r_sum, column=c_idx).border = THIN_BORDER
                
                ca = ws_mod.cell(row=r_avg, column=c_idx, value=f"=AVERAGE({col_let}{header_row+1}:{col_let}{header_row+5})")
                ca.font = FONT_BOLD_BLACK
                ca.fill = CLUB_YELLOW_LIGHT
                ca.number_format = "0.0"
                ca.border = THIN_BORDER
                
                rk_live = int(df_live[df_live["Zawodnik"] == p]["Miejsce"].values[0])
                ws_mod.cell(row=r_rnk, column=c_idx, value=rk_live).font = FONT_BOLD_BLACK
                ws_mod.cell(row=r_rnk, column=c_idx).fill = GRAY_FILL
                ws_mod.cell(row=r_rnk, column=c_idx).border = THIN_BORDER
                
                pt_live = int(df_live[df_live["Zawodnik"] == p]["Pkt Turniejowe"].values[0])
                ws_mod.cell(row=r_pts, column=c_idx, value=f"=CHOOSE({col_let}{r_rnk},20,18,16,14,12,10,9,8,7,6,5,4,3,2,1,0)").font = FONT_BOLD_BLACK
                ws_mod.cell(row=r_pts, column=c_idx).fill = CLUB_YELLOW_LIGHT
                ws_mod.cell(row=r_pts, column=c_idx).border = THIN_BORDER
                
            # 3. Rysujemy skopiowaną Klasyfikację Generalną w nowym miejscu (12 wierszy niżej)
            nowy_start_gen = r_pts + 3
            for r_data in wiersze_generalki:
                for c_data in r_data:
                    nowe_r = nowy_start_gen + (idx_gen_old - db["gen_start_row"])
                    # Przepisujemy dane ze starej struktury do przesuniętych wierszy
                    cell_new = ws_mod.cell(row=nowy_start_gen, column=c_data["col"], value=c_data["val"])
                    if c_data["font"]: cell_new.font = c_data["font"]
                    if c_data["fill"]: cell_new.fill = c_data["fill"]
                    if c_data["align"]: cell_new.alignment = c_data["align"]
                    if c_data["border"]: cell_new.border = c_data["border"]
                nowy_start_gen += 1
                
            # 4. Wpisujemy nowe punkty bezpośrednio do tabeli generalnej w odpowiednią kolumnę rundy
            kolumna_nowej_rundy = 3 + nr_rundy # R1=D(4), R2=E(5), R3=F(6) itd.
            idx_gen_new = znajdz_wiersz_po_tekscie(ws_mod, "KLASYFIKACJA GENERALNA")
            wiersz_nag_new = idx_gen_new + 1
            
            for r in range(wiersz_nag_new + 1, wiersz_nag_new + 16):
                z_name = ws_mod.cell(row=r, column=3).value
                if z_name and str(z_name).strip() in df_live["Zawodnik"].values:
                    p_name = str(z_name).strip()
                    p_turniejowe = int(df_live[df_live["Zawodnik"] == p_name]["Pkt Turniejowe"].values[0])
                    ws_mod.cell(row=r, column=kolumna_nowej_rundy, value=p_turniejowe).font = FONT_BODY_BLACK
                    ws_mod.cell(row=r, column=kolumna_nowej_rundy).alignment = Alignment(horizontal="center")
                elif z_name:
                    ws_mod.cell(row=r, column=kolumna_nowej_rundy, value="-").font = FONT_BODY_BLACK
                    ws_mod.cell(row=r, column=kolumna_nowej_rundy).alignment = Alignment(horizontal="center")

            # Zapis do strumienia pamięci, aby użytkownik mógł pobrać plik na telefon/komputer
            buffer = io.BytesIO()
            wb_mod.save(buffer)
            buffer.seek(0)
            
            st.success(f"🔥 Sukces! Wygenerowano strukturę pionową dla Rundy {nr_rundy}!")
            st.download_button(
                label="📥 POBIERZ ZAKTUALIZOWANY PLIK EXCEL",
                data=buffer,
                file_name=f"Puchar_Lata_2026_Runda_{nr_rundy}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- TAB 2: KLASYFIKACJA GENERALNA ---
with tab2:
    st.header("Oficjalna Klasyfikacja Generalna")
    st.write("Wczytana bezpośrednio z aktualnego pliku Excel.")
    
    gen_rows = []
    for p, rounds in st.session_state.history.items():
        total_suma = sum(rounds)
        row_dict = {"Zawodnik": p}
        for r_idx, r_pts in enumerate(rounds):
            row_dict[f"R{r_idx+1}"] = r_pts if r_pts > 0 or p in ['KRO','CYG','DAR','DOM'] else "-" 
        row_dict["SUMA PUNKTÓW"] = total_suma
        gen_rows.append(row_dict)
        
    df_gen = pd.DataFrame(gen_rows)
    df_gen = df_gen.sort_values(by="SUMA PUNKTÓW", ascending=False).reset_index(drop=True)
    df_gen.index += 1
    df_gen.insert(0, 'Poz.', df_gen.index)
    
    st.dataframe(df_gen.style.set_properties(**{'background-color': '#FFF9C4'}, subset=['SUMA PUNKTÓW']), use_container_width=True)
