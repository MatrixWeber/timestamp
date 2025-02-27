import pandas as pd
import datetime
from openpyxl import Workbook
import calendar
import requests

TAGES_SOLL_STUNDEN = 8
PAUSE_IN_STUNDEN = 0.5

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

# Variable zur Speicherung der Überstunden des Vormonats
ueberstunden_vormonat = 0

# Iteriere über alle Monate des aktuellen Jahres
for month in range(1, 13):
    # Erstelle einen leeren DataFrame mit den Spalten "Datum", "Gekommen", "Gehzeit", "Pause", "Arbeitszeit", "Summe Soll", "Summe Ist", "Überstunden"
    df = pd.DataFrame(columns=['Datum', 'Tag', 'Gekommen', 'Gehzeit', 'Pause', 'Arbeitszeit', 'Summe Soll', 'Summe Ist', 'Überstunden'])

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
            diff_sum = df['Arbeitszeit'].sum() + arbeitszeit
            ueberstunden = diff_sum - TAGES_SOLL_STUNDEN * (len(df) + 1) # Berechne Überstunden

            # Erstelle eine neue Zeile als DataFrame
            new_row = pd.DataFrame({
                'Tag': weekday_names[day.dayofweek],
                'Datum': date.day,
                'Gekommen': '7:30',
                'Gehzeit': '16:30',
                'Pause': PAUSE_IN_STUNDEN,
                'Arbeitszeit': arbeitszeit,
                'Summe Soll': TAGES_SOLL_STUNDEN,
                'Summe Ist': diff_sum,
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
        ws[f'F{row_idx}'] = f'=IF(D{row_idx}="", 0, (D{row_idx}-C{row_idx})*24 - E{row_idx})' # Arbeitszeit
        ws[f'G{row_idx}'] = f'=IF(OR(B{row_idx}="Samstag", B{row_idx}="Sonntag", A{row_idx}=""), 0, {TAGES_SOLL_STUNDEN} * COUNTIFS(B$2:B{row_idx}, "<>Samstag", B$2:B{row_idx}, "<>Sonntag", B$2:B{row_idx}, "<>"))' # Summe Soll
        ws[f'H{row_idx}'] = f'=SUM(F$2:F{row_idx})' # Summe Ist
        if row_idx > 2:
            ws[f'I{row_idx}'] = f'=F{row_idx} - {TAGES_SOLL_STUNDEN} + I{row_idx - 1}' # Überstunden
        else:
            if month == 1:
                ws[f'I{row_idx}'] = f'=F{row_idx} - {TAGES_SOLL_STUNDEN}' # Überstunden für Januar
            else:
                prev_month_sheet = wb[calendar.month_name[month - 1]]
                last_row_prev_month = len(prev_month_sheet['A'])
                month_name = calendar.month_name[month - 1]
                ws[f'I{row_idx}'] = f'=F{row_idx} - {TAGES_SOLL_STUNDEN} + {month_name}!I{last_row_prev_month}' # Überstunden für andere Monate

# Speichere das Workbook
wb.save("calendar.xlsx")