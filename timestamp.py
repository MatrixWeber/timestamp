import pandas as pd
import datetime
from openpyxl import Workbook
import calendar
import requests

TAGES_SOLL_STUNDEN = 8
PAUSE_IN_STUNDEN = 0.75

weekday_names = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']
feiertage_filter = ['08-08', '08-15', '11-19']

year = datetime.datetime.now().year
# URL der API
url = f"https://feiertage-api.de/api/?jahr={year}&nur_land=BY"

# API-Aufruf
response = requests.get(url)

# Antwort in JSON umwandeln
data = response.json()

# Leere Liste für die Feiertage
feiertage = []

holidays = response.json()

# Iteriere über alle Feiertage
for holiday, date in holidays.items():
    # Füge das Datum des Feiertags der Liste hinzu
    if date["datum"][5:] not in feiertage_filter:
        feiertage.append(date["datum"])

# Erstelle ein leeres Workbook
wb = Workbook()

# Benenne das Standard-Tab "Sheet" in "Übersicht" um
if "Sheet" in wb.sheetnames:
    wb["Sheet"].title = "Übersicht"

# Füge die Summen der Monate in die Übersicht ein
uebersicht_ws = wb["Übersicht"]

# Schreibe die Kopfzeile in die Übersicht (inklusive der neuen Spalte "Urlaub")
uebersicht_ws.append(["Monat", "Summe Soll", "Summe Arbeitszeit", "Summe Überstunden", "Urlaub"])

# Variable zur Speicherung der Überstunden des Vormonats
ueberstunden_vormonat = 0

# Iteriere über alle Monate des aktuellen Jahres
for month in range(1, 13):
    # Erstelle einen leeren DataFrame mit den Spalten "Datum", "Gekommen", "Gehzeit", "Pause", "Arbeitszeit", "Soll", "Überstunden"
    df = pd.DataFrame(columns=['Datum', 'Tag', 'Gekommen', 'Gehzeit', 'Pause', 'Arbeitszeit', 'Soll', 'Überstunden'])

    # Bestimme die Anzahl der Tage im aktuellen Monat
    num_days = calendar.monthrange(datetime.datetime.now().year, month)[1]

    # Iteriere über alle Tage des aktuellen Monats
    for day in pd.date_range(start=f'{datetime.datetime.now().year}-{month:02d}-01', end=f'{datetime.datetime.now().year}-{month:02d}-{num_days}'):
        day_is_holiday = False
        for feiertag in feiertage:
            if feiertag == day._date_repr:
                day_is_holiday = True
                break
        
        # Überprüfe, ob der aktuelle Tag ein Wochentag ist
        date = day.date()
        if not day_is_holiday and day.dayofweek < 5:
            # Setze die Gekommen-Zeit auf 7:30
            time = datetime.datetime.combine(date, datetime.time(7, 30))

            # Setze die Gehzeit auf 16:30
            gehzeit = datetime.datetime.combine(date, datetime.time(16, 30))

            # Berechne die Arbeitszeit zwischen der Uhrzeit und der Gehzeit in Minuten
            diff = (gehzeit - time).total_seconds() / 60

            # Wandle die Arbeitszeit in Stunden um und ziehe die Pause ab
            diff_hours = diff / 60
            arbeitszeit = diff_hours - PAUSE_IN_STUNDEN
            ueberstunden = arbeitszeit - TAGES_SOLL_STUNDEN * (len(df) + 1) # Berechne Überstunden

            # Erstelle eine neue Zeile als DataFrame
            new_row = pd.DataFrame({
                'Tag': weekday_names[day.dayofweek],
                'Datum': date.day,
                'Gekommen': '7:30',
                'Gehzeit': '16:30',
                'Pause': PAUSE_IN_STUNDEN,
                'Arbeitszeit': arbeitszeit,
                'Soll': TAGES_SOLL_STUNDEN,
                'Überstunden': ueberstunden
            }, index=[len(df)])

            # Füge die neue Zeile an den DataFrame an
            df = pd.concat([df, new_row])
            
    ueberstunden_vormonat = ueberstunden + ueberstunden_vormonat
    # Erstelle ein neues Arbeitsblatt für den aktuellen Monat
    ws = wb.create_sheet(title=calendar.month_name[month])

    # Schreibe die Spaltenüberschriften manuell in das Arbeitsblatt
    for col, value in enumerate(df.columns):
        ws.cell(row=1, column=col + 1).value = value

    # Schreibe den DataFrame manuell in das Arbeitsblatt
    for row_idx, row in enumerate(df.iterrows(), start=2):
        for col_idx, value in enumerate(row[1]):
            ws.cell(row=row_idx, column=col_idx + 1).value = value

    # Füge Formeln für Arbeitszeit, Summe Soll, Summe Ist und Überstunden hinzu
    for row_idx in range(2, len(df) + 2):
        ws[f'F{row_idx}'] = f'=IF(OR(C{row_idx}="U", C{row_idx}="K", D{row_idx}=""), {TAGES_SOLL_STUNDEN}, (D{row_idx}-C{row_idx})*24 - E{row_idx})' # Arbeitszeit
        ws[f'G{row_idx}'] = f'={TAGES_SOLL_STUNDEN}' # Soll
        ws[f'H{row_idx}'] = f'=IF(OR(C{row_idx}="U", C{row_idx}="K"), 0, F{row_idx} - G{row_idx})' # Überstunden
        if row_idx <= 2:
            if month != 1:
                prev_month_sheet = wb[calendar.month_name[month - 1]]
                last_row_prev_month = len(prev_month_sheet['A'])
                month_name = calendar.month_name[month - 1]

    # Füge Summen für Soll, Arbeitszeit und Überstunden in die letzte Zeile ein
    last_row = len(df) + 2
    ws[f'F{last_row}'] = f'=SUM(F2:F{last_row - 1})' # Summe Arbeitszeit
    ws[f'G{last_row}'] = f'=SUM(G2:G{last_row - 1})' # Summe Soll
    ws[f'H{last_row}'] = f'=SUM(H2:H{last_row - 1})' # Summe Überstunden

    # Berechne die Anzahl der Urlaubstage (Zellen in Spalte "Gekommen" mit "U")
    ws[f'C{last_row}'] = f'=COUNTIF(C2:C{last_row - 1}, "U")'
    # Optional: Beschriftung für die Summenzeile
    ws[f'B{last_row}'] = "Genommene Urlaubstage"
    # Optional: Beschriftung für die Summenzeile
    ws[f'E{last_row}'] = "Summen"
    
    
# Iteriere über alle Monats-Tabs und berechne die Summen
for month in range(1, 13):
    month_name = calendar.month_name[month]
    if month_name in wb.sheetnames:
        month_ws = wb[month_name]
        last_row = len(month_ws['A'])  # Letzte Zeile des Monats-Tabs

        # Füge die Summen und die Urlaubstage in die Übersicht ein
        uebersicht_ws.append([
            month_name,
            f'={month_name}!G{last_row}',  # Summe Soll
            f'={month_name}!F{last_row}',  # Summe Arbeitszeit
            f'={month_name}!H{last_row}',  # Summe Überstunden
            f'={month_name}!C{last_row}'  # Anzahl der Urlaubstage
        ])

# Füge die Summenzeile in die Übersicht ein
last_row_uebersicht = len(uebersicht_ws['A']) + 1  # Nächste freie Zeile in der Übersicht
uebersicht_ws[f'A{last_row_uebersicht}'] = "Summen"
uebersicht_ws[f'B{last_row_uebersicht}'] = f'=SUM(B2:B{last_row_uebersicht - 1})'  # Summe Soll
uebersicht_ws[f'C{last_row_uebersicht}'] = f'=SUM(C2:C{last_row_uebersicht - 1})'  # Summe Arbeitszeit
uebersicht_ws[f'D{last_row_uebersicht}'] = f'=SUM(D2:D{last_row_uebersicht - 1})'  # Summe Überstunden
uebersicht_ws[f'E{last_row_uebersicht}'] = f'=SUM(E2:E{last_row_uebersicht - 1})'  # Summe Urlaubstage

# Speichere das Workbook
wb.save("calendar.xlsx")