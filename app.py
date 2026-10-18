import streamlit as st
import calendar
import os
import json
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ReportLab per i PDF
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Gestione Turni Web", page_icon="📅", layout="wide")

# --- CSS PERSONALIZZATO PER OTTIMIZZARE GLI SPAZI E LA GRANDEZZA DEL TESTO ---
st.markdown("""
    <style>
        /* Riduce il padding complessivo delle colonne per far stare tutto sullo schermo */
        [data-testid="column"] {
            padding: 2px !important;
        }
        /* Riduce la dimensione del testo e il padding dei menu a tendina Selectbox */
        div[data-baseweb="select"] > div {
            font-size: 11px !important;
            padding-left: 2px !important;
            padding-right: 2px !important;
            min-height: 28px !important;
        }
        /* Stile per i badge dei giorni */
        .day-header {
            background-color: #1F497D;
            color: white;
            text-align: center;
            font-weight: bold;
            font-size: 12px;
            padding: 4px;
            border-radius: 4px;
            margin-bottom: 5px;
        }
    </style>
""", unsafe_allow_html=True)

ITALIAN_MONTHS = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
]
ITALIAN_DAYS = ["LUN", "MAR", "MER", "GIO", "VEN", "SAB", "DOM"]

ALL_HOURS = ["--:--"]
for h in range(24):
    for m in [0, 30]:
        ALL_HOURS.append(f"{h:02d}:{m:02d}")

def time_to_hours(time_str):
    if not time_str or ":" not in time_str or time_str == "--:--":
        return 0.0
    try:
        parts = time_str.split(":")
        return int(parts[0]) + int(parts[1]) / 60.0
    except Exception:
        return 0.0

def calculate_shift_hours(in_time, out_time):
    in_h = time_to_hours(in_time)
    out_h = time_to_hours(out_time)
    diff = out_h - in_h
    if diff < 0:
        diff += 24.0
    return diff if (in_time and out_time and in_time != "--:--" and out_time != "--:--") else 0.0

# --- INIZIALIZZAZIONE STATE ---
if "operators" not in st.session_state:
    st.session_state.operators = []
if "clients" not in st.session_state:
    st.session_state.clients = []
if "schedule_data" not in st.session_state:
    st.session_state.schedule_data = {}

ANAG_FILE = "anagrafica.json"
SCHED_FILE = "turni_salvati.json"

def load_data_from_disk():
    if os.path.exists(ANAG_FILE):
        try:
            with open(ANAG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                st.session_state.operators = data.get("operators", [])
                st.session_state.clients = data.get("clients", [])
        except Exception:
            pass

    if os.path.exists(SCHED_FILE):
        try:
            with open(SCHED_FILE, "r", encoding="utf-8") as f:
                raw_data = json.load(f)
                st.session_state.schedule_data = {}
                for k, v in raw_data.items():
                    parts = k.split("|")
                    if len(parts) == 5:
                        client, y, m, d, off = parts[0], int(parts[1]), int(parts[2]), int(parts[3]), int(parts[4])
                        st.session_state.schedule_data[(client, y, m, d, off)] = v
        except Exception:
            pass

def save_data_to_disk():
    try:
        with open(ANAG_FILE, "w", encoding="utf-8") as f:
            json.dump({"operators": st.session_state.operators, "clients": st.session_state.clients}, f, ensure_ascii=False, indent=4)
        
        raw_data = {}
        for k, v in st.session_state.schedule_data.items():
            client, y, m, d, off = k
            raw_data[f"{client}|{y}|{m}|{d}|{off}"] = v
        with open(SCHED_FILE, "w", encoding="utf-8") as f:
            json.dump(raw_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"Errore di salvataggio: {e}")

if not st.session_state.operators and not st.session_state.clients:
    load_data_from_disk()

# --- FUNZIONI ESPORTAZIONE ---
def generate_excel_bytes(year, month_idx):
    month_name = ITALIAN_MONTHS[month_idx - 1]
    wb = openpyxl.Workbook()
    ws_data = wb.active
    ws_data.title = "Anagrafica"
    ws_sched = wb.create_sheet(title="Pianificazione Mensile")

    primary_fill = PatternFill(start_color="366092", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    side_thin = Side(style='thin', color='B0C4DE')
    thin_border = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

    ws_data['A1'] = "Gestione Anagrafica Operatori e Clienti"
    ws_data['A4'] = "Operatore"
    ws_data['A4'].fill = primary_fill
    ws_data['A4'].font = header_font

    for i, op in enumerate(st.session_state.operators):
        c = ws_data.cell(row=5+i, column=1, value=op)
        c.font = regular_font; c.border = thin_border

    ws_data['C4'] = "Cliente"
    ws_data['C4'].fill = primary_fill
    ws_data['C4'].font = header_font

    for i, cl in enumerate(st.session_state.clients):
        c = ws_data.cell(row=5+i, column=3, value=cl)
        c.font = regular_font; c.border = thin_border

    num_days = calendar.monthrange(year, month_idx)[1]
    ws_sched['A1'] = f"Programmazione Orari - {month_name} {year}"
    ws_sched.cell(row=3, column=1, value="Cliente").fill = primary_fill
    ws_sched.cell(row=3, column=1).font = header_font

    for day in range(1, num_days + 1):
        start_col = 2 + (day - 1) * 3
        wday_idx = calendar.weekday(year, month_idx, day)
        ws_sched.cell(row=3, column=start_col, value=f"{day} {month_name[:3]} ({ITALIAN_DAYS[wday_idx]})").fill = primary_fill
        ws_sched.cell(row=3, column=start_col).font = header_font

    for client_idx, client_name in enumerate(st.session_state.clients):
        client_start_row = 5 + (client_idx * 3)
        ws_sched.cell(row=client_start_row, column=1, value=client_name).font = bold_font
        for r_offset in range(3):
            r = client_start_row + r_offset
            for day in range(1, num_days + 1):
                col_op = 2 + (day - 1) * 3
                val_op, val_in, val_out = st.session_state.schedule_data.get((client_name, year, month_idx, day, r_offset), ("", "", ""))
                ws_sched.cell(row=r, column=col_op, value=val_op if val_op != "Nessuno" else "")
                ws_sched.cell(row=r, column=col_op+1, value=val_in if val_in != "--:--" else "")
                ws_sched.cell(row=r, column=col_op+2, value=val_out if val_out != "--:--" else "")

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def generate_pdf_bytes(year, month_idx):
    month_str = ITALIAN_MONTHS[month_idx - 1]
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#1F497D'), alignment=1)
    header_style = ParagraphStyle('Th', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('Td', parent=styles['Normal'], fontName='Helvetica', fontSize=8, alignment=1)
    client_style = ParagraphStyle('Tc', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#1F497D'))

    elements = [Paragraph(f"PANORAMICA TURNI MENSILE — {month_str.upper()} {year}", title_style), Spacer(1, 10)]
    num_days = calendar.monthrange(year, month_idx)[1]
    active_clients = [c for c in st.session_state.clients if c]

    part1_days = list(range(1, min(16, num_days + 1)))

    table_data = []
    header_row = [Paragraph("Cliente", header_style)]
    for day in part1_days:
        wday_idx = calendar.weekday(year, month_idx, day)
        header_row.append(Paragraph(f"{day}<br/>{ITALIAN_DAYS[wday_idx]}", header_style))
    table_data.append(header_row)

    for client_name in active_clients:
        row = [Paragraph(client_name, client_style)]
        for day in part1_days:
            day_shifts = []
            for r_offset in range(3):
                op_val, in_val, out_val = st.session_state.schedule_data.get((client_name, year, month_idx, day, r_offset), ("", "", ""))
                if op_val and in_val and op_val != "Nessuno" and in_val != "--:--":
                    day_shifts.append(f"<b>{op_val}</b> ({in_val}-{out_val})")
            cell_txt = "<br/>".join(day_shifts) if day_shifts else "-"
            row.append(Paragraph(cell_txt, cell_style))
        table_data.append(row)

    col_widths = [160] + [(1140-160)/len(part1_days)] * len(part1_days)
    t = Table(table_data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#366092')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#B0C4DE')),
    ]))
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

# --- INTERFACCIA WEB ---
st.sidebar.title("📌 Menu Gestionale")
menu = st.sidebar.radio("Seleziona Sezione", ["Pianificazione Settimanale", "Anagrafica", "Riepilogo Mensile", "Esportazioni & Stampa"])

# ================= 1. ANAGRAFICA =================
if menu == "Anagrafica":
    st.title("👥 Gestione Anagrafica Operatori e Clienti")
    st.info("Inserisci i nomi e clicca **Salva Anagrafica**.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Elenco Operatori (max 15)")
        new_ops = []
        for i in range(15):
            val = st.session_state.operators[i] if i < len(st.session_state.operators) else ""
            op = st.text_input(f"Operatore {i+1}", value=val, key=f"anag_op_{i}")
            if op.strip():
                new_ops.append(op.strip())

    with col2:
        st.subheader("Elenco Clienti (max 15)")
        new_cl = []
        for i in range(15):
            val = st.session_state.clients[i] if i < len(st.session_state.clients) else ""
            cl = st.text_input(f"Cliente {i+1}", value=val, key=f"anag_cl_{i}")
            if cl.strip():
                new_cl.append(cl.strip())

    if st.button("💾 Salva Anagrafica", type="primary"):
        st.session_state.operators = new_ops
        st.session_state.clients = new_cl
        save_data_to_disk()
        st.success("✅ Anagrafica salvata!")

# ================= 2. PIANIFICAZIONE SETTIMANALE OTTIMIZZATA =================
elif menu == "Pianificazione Settimanale":
    st.title("📅 Pianificazione Turni Settimanale")

    active_clients = [c for c in st.session_state.clients if c]
    active_ops = ["Nessuno"] + [o for o in st.session_state.operators if o]

    if not active_clients:
        st.warning("⚠️ Nessun cliente salvato. Vai prima nella sezione **Anagrafica**.")
    else:
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            year = st.selectbox("Anno", [2026, 2027, 2028, 2029, 2030], index=0)
        with c2:
            month_str = st.selectbox("Mese", ITALIAN_MONTHS, index=9)
            month_idx = ITALIAN_MONTHS.index(month_str) + 1
        with c3:
            num_days_total = calendar.monthrange(year, month_idx)[1]
            week_choice = st.selectbox("Seleziona Settimana", [
                "1ª Settimana (1 - 7)",
                "2ª Settimana (8 - 14)",
                "3ª Settimana (15 - 21)",
                "4ª Settimana (22 - 28)",
                "5ª Settimana (29 - fine)"
            ])

        if "1ª" in week_choice: days_to_show = list(range(1, min(8, num_days_total + 1)))
        elif "2ª" in week_choice: days_to_show = list(range(8, min(15, num_days_total + 1)))
        elif "3ª" in week_choice: days_to_show = list(range(15, min(22, num_days_total + 1)))
        elif "4ª" in week_choice: days_to_show = list(range(22, min(29, num_days_total + 1)))
        else: days_to_show = list(range(29, num_days_total + 1))

        st.markdown(f"### 🗓️ Settimana: **{days_to_show[0]} — {days_to_show[-1]} {month_str} {year}**")

        with st.form("weekly_schedule_form"):
            for client_name in active_clients:
                st.markdown(f"#### 👤 Cliente: `{client_name}`")
                
                # Griglia dei giorni
                cols = st.columns(len(days_to_show))
                
                for d_i, day in enumerate(days_to_show):
                    w_i = calendar.weekday(year, month_idx, day)
                    with cols[d_i]:
                        st.markdown(f"<div class='day-header'>{day} {ITALIAN_DAYS[w_i]}</div>", unsafe_allow_html=True)
                        
                        for r_offset in range(3):
                            s_op, s_in, s_out = st.session_state.schedule_data.get((client_name, year, month_idx, day, r_offset), ("Nessuno", "--:--", "--:--"))
                            init_op = s_op if s_op in active_ops else "Nessuno"
                            init_in = s_in if s_in in ALL_HOURS else "--:--"
                            init_out = s_out if s_out in ALL_HOURS else "--:--"

                            st.caption(f"T{r_offset+1}")
                            st.selectbox(f"Op T{r_offset+1}", active_ops, index=active_ops.index(init_op), key=f"op_{client_name}_{day}_{r_offset}", label_visibility="collapsed")
                            
                            c_in, c_out = st.columns(2)
                            with c_in:
                                st.selectbox(f"In T{r_offset+1}", ALL_HOURS, index=ALL_HOURS.index(init_in), key=f"in_{client_name}_{day}_{r_offset}", label_visibility="collapsed")
                            with c_out:
                                st.selectbox(f"Out T{r_offset+1}", ALL_HOURS, index=ALL_HOURS.index(init_out), key=f"out_{client_name}_{day}_{r_offset}", label_visibility="collapsed")
                st.markdown("---")

            save_btn = st.form_submit_button("💾 Salva Modifiche Settimana", type="primary")
            if save_btn:
                for client_name in active_clients:
                    for day in days_to_show:
                        for r_offset in range(3):
                            v_op = st.session_state[f"op_{client_name}_{day}_{r_offset}"]
                            v_in = st.session_state[f"in_{client_name}_{day}_{r_offset}"]
                            v_out = st.session_state[f"out_{client_name}_{day}_{r_offset}"]
                            
                            st.session_state.schedule_data[(client_name, year, month_idx, day, r_offset)] = (
                                "" if v_op == "Nessuno" else v_op,
                                "" if v_in == "--:--" else v_in,
                                "" if v_out == "--:--" else v_out
                            )
                save_data_to_disk()
                st.success("✅ Turni salvati con successo!")

# ================= 3. RIEPILOGO MENSILE =================
elif menu == "Riepilogo Mensile":
    st.title("📊 Riepilogo Ore Fine Mese")

    c1, c2, c3 = st.columns(3)
    with c1:
        sum_type = st.selectbox("Soggetto", ["Clienti", "Operatori"])
    with c2:
        sum_year = st.selectbox("Anno", [2026, 2027, 2028, 2029, 2030], index=0)
    with c3:
        sum_month_str = st.selectbox("Mese", ITALIAN_MONTHS, index=9)
        sum_month_idx = ITALIAN_MONTHS.index(sum_month_str) + 1

    num_days = calendar.monthrange(sum_year, sum_month_idx)[1]
    targets = st.session_state.clients if sum_type == "Clienti" else st.session_state.operators
    active_targets = [t for t in targets if t]

    results = []
    grand_total = 0.0

    for target in active_targets:
        tot_h = 0.0
        for day in range(1, num_days + 1):
            if sum_type == "Clienti":
                for off in range(3):
                    _, inv, outv = st.session_state.schedule_data.get((target, sum_year, sum_month_idx, day, off), ("", "", ""))
                    tot_h += calculate_shift_hours(inv, outv)
            else:
                for c_name in st.session_state.clients:
                    for off in range(3):
                        opv, inv, outv = st.session_state.schedule_data.get((c_name, sum_year, sum_month_idx, day, off), ("", "", ""))
                        if opv == target:
                            tot_h += calculate_shift_hours(inv, outv)
        grand_total += tot_h
        results.append({"Nome": target, "Totale Ore": f"{tot_h:.2f} h"})

    if results:
        st.table(pd.DataFrame(results))
        st.metric("TOTALE GENERALE COMPLESSIVO", f"{grand_total:.2f} Ore")
    else:
        st.info("Nessun dato presente.")

# ================= 4. ESPORTAZIONI & STAMPA =================
elif menu == "Esportazioni & Stampa":
    st.title("📥 Esportazione File Excel e PDF")

    c1, c2 = st.columns(2)
    with c1:
        exp_year = st.selectbox("Anno", [2026, 2027, 2028, 2029, 2030], index=0)
    with c2:
        exp_month_str = st.selectbox("Mese", ITALIAN_MONTHS, index=9)
        exp_month_idx = ITALIAN_MONTHS.index(exp_month_str) + 1

    st.markdown("---")

    col_ex, col_pdf = st.columns(2)
    with col_ex:
        st.subheader("📗 File Excel")
        excel_data = generate_excel_bytes(exp_year, exp_month_idx)
        st.download_button(
            label="📥 Scarica Excel (.xlsx)",
            data=excel_data,
            file_name=f"Gestione_Orari_{exp_month_str}_{exp_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_pdf:
        st.subheader("📄 Stampa PDF A3")
        pdf_data = generate_pdf_bytes(exp_year, exp_month_idx)
        st.download_button(
            label="🖨️ Scarica PDF Mensile A3 (.pdf)",
            data=pdf_data,
            file_name=f"Programma_Mensile_{exp_month_str}_{exp_year}.pdf",
            mime="application/pdf"
        )
