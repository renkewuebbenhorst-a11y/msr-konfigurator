import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import io

st.set_page_config(page_title="E-MSR Konfigurator", layout="wide")

st.title("⚡ Kabel- & Material-Konfigurator")
st.write("Erfassung der Anlagenkomponenten zur automatischen Generierung der Bestell-Stückliste.")

st.divider()

# ==========================================
# 1. STAMMDATEN DER ANLAGENKOMPONENTEN
# ==========================================

# Pumpen / Motoren (Basis: 1 Stück)
# Format: (Typ, Bezeichnung/Spezifikation, Artikelnummer, Menge_einzeln, Einheit, Bemerkung)
MOTOR_TEMPLATES = {
    "4.0 mm²": [
        ("Elektrik / Kabel", "LKSM-HF 3x4 Flex 0,6/1kV black", "8702804", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "LKSM-HF 3x4 Flex 0,6/1kV black", "8702804", 5, "m", "Vom Schaltschrank zum FU"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 5, "m", "Von Pumpe zum Schaltschrank"),
        ("Ringkabelschuhe", "KLAU Cu-Kerbkabelschuh 2R8 / 10QMM-M8", "", 4, "Stk", "Klemmenkasten Pumpe"),
        ("Aderendhülsen", "RED Aderendhülsen isoliert 18x10,00qmm", "7189680", 12, "Stk", "8x FU (In/Out), 4x Schaltschrank"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 0,75-10mm, grau", "121755", 4, "Stk", "2x Pumpe PTC, 2x Schaltschrank"),
        ("Kabelverschraubung", "M25", "1878116", 2, "Stk", "1x Pumpe , 1x FU-Gehäuse"),
        ("Kabelverschraubung", "M16", "1878109", 2, "Stk", "1x Pumpe PTC, 1x Schaltschrank"),
        ("Kabelverschraubung-Reduzierung", "WISKA Reduzierung MRM50/25", "1878665", 1, "Stk", "Reduzierung für Pumpengehäuse"),
        ("Elektrik / Kabel", "Industrial Ethernet (grün, 2x2xAWG22)", "2815977", 10, "m", "Profinetkabel"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 180", "103836", 1, "Stk", "Profinetanschluss für Frequenzumrichter"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 90", "103835", 1, "Stk", "Profinetanschluss im Schaltschrank")
    ],
    "10.0 mm²": [
        ("Elektrik / Kabel", "LKSM-HF 3x10 Flex 0,6/1kV black", "91000062", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "LKSM-HF 3x10 Flex 0,6/1kV black", "91000062", 5, "m", "Vom Schaltschrank zum FU"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 5, "m", "Von Pumpe zum Schaltschrank"),
        ("Ringkabelschuhe", "KLAU Cu-Kerbkabelschuh 2R8 / 10QMM-M8", "", 4, "Stk", "Klemmenkasten Pumpe"),
        ("Aderendhülsen", "RED Aderendhülsen isoliert 18x10,00qmm", "7189680", 12, "Stk", "8x FU (In/Out), 4x Schaltschrank"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 0,75-10mm, grau", "121755", 4, "Stk", "2x Pumpe PTC, 2x Schaltschrank"),
        ("Kabelverschraubung", "M32", "1878120", 2, "Stk", "1x Pumpe , 1x FU-Gehäuse"),
        ("Kabelverschraubung", "M16", "1878109", 2, "Stk", "1x Pumpe PTC, 1x Schaltschrank"),
        ("Kabelverschraubung-Reduzierung", "WISKA Reduzierung MRM50/32", "1878666", 1, "Stk", "Reduzierung für Pumpengehäuse"),
        ("Elektrik / Kabel", "Industrial Ethernet (grün, 2x2xAWG22)", "2815977", 10, "m", "Profinetkabel"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 180", "103836", 1, "Stk", "Profinetanschluss für Frequenzumrichter"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 90", "103835", 1, "Stk", "Profinetanschluss im Schaltschrank")
    ],
    "16.0 mm²": [
        ("Elektrik / Kabel", "LKSM-HF 3x16 Flex 0,6/1kV black", "8870747", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "LKSM-HF 3x16 Flex 0,6/1kV black", "8870747", 5, "m", "Vom Schaltschrank zum FU"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 30, "m", "Vom FU zur Pumpe"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 5, "m", "Von Pumpe zum Schaltschrank"),
        ("Ringkabelschuhe", "KLAU Cu-Kerbkabelschuh 2R8 / 16QMM-M8", "", 4, "Stk", "Klemmenkasten Pumpe"),
        ("Aderendhülsen", "RED Aderendhülsen isoliert 18x10,00qmm", "7189680", 12, "Stk", "8x FU (In/Out), 4x Schaltschrank"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 0,75-10mm, grau", "121755", 4, "Stk", "2x Pumpe PTC, 2x Schaltschrank"),
        ("Kabelverschraubung", "M40", "1878090", 2, "Stk", "1x Pumpe , 1x FU-Gehäuse"),
        ("Kabelverschraubung", "M16", "1878109", 2, "Stk", "1x Pumpe PTC, 1x Schaltschrank"),
        ("Kabelverschraubung-Reduzierung", "WISKA Reduzierung MRM50/40", "1878667", 1, "Stk", "Reduzierung für Pumpengehäuse"),
        ("Elektrik / Kabel", "Industrial Ethernet (grün, 2x2xAWG22)", "2815977", 10, "m", "Profinetkabel"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 180", "103836", 1, "Stk", "Profinetanschluss für Frequenzumrichter"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 90", "103835", 1, "Stk", "Profinetanschluss im Schaltschrank")
    ]
}

# Sensorik, Ventile, Dosierung & Flowmeter (Basis: 1 Stück)
KOMPONENTEN_TEMPLATES = {
    "Analogwert Einfach": [
        ("Elektrik / Kabel", "IFM Kabeldose ifm electronic M12 EVC003", "2495606", 1, "Stk", "20m"),
        ("Aderendhülsen", "RED Aderendhülsen 0,34-8mm, türkis", "7189675", 2, "Stk", "Schaltschrankseite")
    ],
    "Flowmeter": [
        ("Elektrik / Kabel", "LKM-HF 3G1,5 0,6/1kV black", "8878095", 20, "m", "Vom Schaltschrank zu FIT1500"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 1,5-10mm, schwarz", "121762", 6, "Stk", "3x FIT1500, 3x Schaltschrank"),
        ("Kabelverschraubung", "M20", "1878087", 2, "Stk", "1x FIT1500, 1x Schaltschrank"),
        ("Elektrik / Kabel", "Metz Ethernet D-kodiert M12 St. gerade-RJ45 St. gerade PU", "4834961", 20, "m", "Von FIT1500 zum Switch"),
        ("Profinet Stecker", "Siemens IE FC RJ45 Plug 180", "103836", 1, "Stk", "Schaltschrank- / Switch-Anschluss")
    ],
    "Ventil Einfach": [
        ("Elektrik / Kabel", "IFM Kabeldose ifm electronic M12 EVC003", "2495606", 1, "Stk", "20m"),
        ("Aderendhülsen", "RED Aderendhülsen 0,34-8mm, türkis", "7189675", 2, "Stk", "Schaltschrankseite")
    ],
    "Dosierpumpe": [
        ("Elektrik / Kabel", "LKM-HF 3G1,5 0,6/1kV black", "8878095", 30, "m", "Von Abzweigdose zu Schaltschrank"),
        ("Elektrik / Kabel", "RFE-HF 1x2x0,75 250V grey", "6870203", 30, "m", "Von Abzweigdose zu Schaltschrank"),
        ("Terminal Box", "Hensel Reihenklemmenkasten RK 0207 T", "", 1, "Stk", "Abzweigdose in der Nähe der Dosierpumpe"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 1,5-10mm, schwarz", "121762", 9, "Stk", "6x Hensel Box (Rein/Raus), 3x Schaltschrank"),
        ("Aderendhülsen", "KLAUKE Aderendhülsen 1,5-10mm, schwarz", "121762", 6, "Stk", "4x Hensel Box (Rein/Raus), 2x Schaltschrank"),
        ("Kabelverschraubung", "M20", "1878087", 5, "Stk", "4x Hensel Box, 1x Schaltschrank")
    ]
}

VERFUEGBARE_QUERSCHNITTE = ["4.0 mm²", "10.0 mm²", "16.0 mm²"]

# ==========================================
# 2. BENUTZEROBERFLÄCHE (Eingaben)
# ==========================================
st.header("1. Anlagenkomponenten & Parameter")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Pumpen & ERD")
    
    # SeawaterPump
    anz_seawater = st.number_input("SeawaterPump (Anzahl)", min_value=0, value=1, step=1)
    qs_seawater = st.selectbox("SeawaterPump Kabelquerschnitt", options=VERFUEGBARE_QUERSCHNITTE, index=1)
    
    st.markdown("---")
    
    # HighPressurePump
    anz_hpp = st.number_input("HighPressurePump (Anzahl)", min_value=0, value=1, step=1)
    qs_hpp = st.selectbox("HighPressurePump Kabelquerschnitt", options=VERFUEGBARE_QUERSCHNITTE, index=2)
    
    st.markdown("---")
    
    # EnergyRecoveryDevice
    anz_erd = st.number_input("EnergyRecoveryDevice (Anzahl)", min_value=0, value=0, step=1)
    qs_erd = st.selectbox("EnergyRecoveryDevice Kabelquerschnitt", options=VERFUEGBARE_QUERSCHNITTE, index=0)

with col2:
    st.subheader("Sensorik, Ventile & Dosierung")
    
    anz_ventile = st.number_input("Ventile Einfach (Anzahl)", min_value=0, value=10, step=1)
    anz_analog = st.number_input("Analogwerte Einfach (Anzahl)", min_value=0, value=20, step=1)
    anz_flowmeter = st.number_input("Flowmeter (Anzahl)", min_value=0, value=5, step=1)
    anz_dosing = st.number_input("Dosierpumpen (Anzahl)", min_value=0, value=3, step=1)

st.divider()

# ==========================================
# 3. GENERIERUNG DER STÜCKLISTE (EXCEL)
# ==========================================
def erstelle_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stückliste"

    # Überschriften inkl. Artikelnummer & Bemerkung
    headers = ["Anlage / Komponente", "Verbrauchsmaterial", "Spezifikation", "Artikelnummer", "Menge pro Stk", "Einheit", "Gesamtmenge", "Bemerkung"]
    ws.append(headers)

    # Header-Design
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    # Hilfsfunktion für Pumpen mit variablen Querschnitten
    def verarbeite_pumpe(gruppen_name, anz, querschnitt):
        if anz > 0 and querschnitt in MOTOR_TEMPLATES:
            template = MOTOR_TEMPLATES[querschnitt]
            for typ, bez, art_nr, menge_einzeln, einheit, bemerkung in template:
                gesamt = menge_einzeln * anz
                ws.append([f"{gruppen_name} ({anz}x {querschnitt})", typ, bez, art_nr, menge_einzeln, einheit, gesamt, bemerkung])

    # Hilfsfunktion für Standard-Komponenten
    def verarbeite_komponente(gruppen_name, anz):
        if anz > 0 and gruppen_name in KOMPONENTEN_TEMPLATES:
            template = KOMPONENTEN_TEMPLATES[gruppen_name]
            for typ, bez, art_nr, menge_einzeln, einheit, bemerkung in template:
                gesamt = menge_einzeln * anz
                ws.append([f"{gruppen_name} ({anz}x)", typ, bez, art_nr, menge_einzeln, einheit, gesamt, bemerkung])

    # 1. Pumpen verarbeiten
    verarbeite_pumpe("SeawaterPump", anz_seawater, qs_seawater)
    verarbeite_pumpe("HighPressurePump", anz_hpp, qs_hpp)
    verarbeite_pumpe("EnergyRecoveryDevice", anz_erd, qs_erd)

    # 2. Übrige Komponenten verarbeiten
    verarbeite_komponente("Analogwert Einfach", anz_analog)
    verarbeite_komponente("Flowmeter", anz_flowmeter)
    verarbeite_komponente("Ventil Einfach", anz_ventile)
    verarbeite_komponente("Dosierpumpe", anz_dosing)

    # Spaltenbreiten optimieren
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

st.header("2. Stückliste generieren")

if st.button("🚀 Stückliste erstellen"):
    excel_data = erstelle_excel()
    st.success("Stückliste erfolgreich generiert!")
    
    st.download_button(
        label="📥 Excel-Datei herunterladen",
        data=excel_data,
        file_name="Bestellliste_Anlage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )