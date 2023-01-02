import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Color
import calendar
import requests

weekday_names = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']

year = datetime.datetime.now().year
# URL der API
url = f"https://feiertage-api.de/api/?jahr=${year}&nur_land=BY"

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
    feiertage.append(date["datum"])

for feiertag in feiertage:
    print(feiertag)



# Erstelle ein leeres Workbook
wb = Workbook()

# Iteriere über alle Monate des aktuellen Jahres
for month in range(1, 13):
    # Erstelle einen leeren DataFrame mit den Spalten "Datum", "Zeit", "Gehzeit" und "Differenz"
    df = pd.DataFrame(columns=['Datum', 'Zeit', 'Gehzeit', 'Differenz'])

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
        if not day_is_holiday:
            if day.dayofweek < 5:  # 5 entspricht Samstag
                # Speichere das aktuelle Datum und die aktuelle Uhrzeit in separaten Variablen
                time = datetime.datetime.now()

                # Setze die Gehzeit auf 5 Stunden nach der aktuellen Zeit
                gehzeit = (time + datetime.timedelta(hours=8)).time()

                # Wandle Zeit und Gehzeit in Datetime-Objekte um
                time = datetime.datetime.combine(date, time.time())

                gehzeit = datetime.datetime.combine(date, gehzeit)

                # Berechne die Differenz zwischen der Uhrzeit und der Gehzeit in Minuten
                diff = (gehzeit - time).total_seconds() / 60

                # Wandle die Differenz in Stunden um
                diff_hours = diff / 60
                diff_sum = df['Differenz'].sum()
                # Erstelle eine neue Zeile als DataFrame
                new_row = pd.DataFrame({'Tag': weekday_names[day.dayofweek], 'Datum': date, 'Zeit': time.strftime("%H:%M"), 'Gehzeit': gehzeit.strftime("%H:%M"), 'Differenz': diff_hours, 'Summe': diff_sum}, index=[len(df)])

                # Füge die neue Zeile an den DataFrame an
                df = pd.concat([df, new_row])
        else:
            new_row = pd.DataFrame({'Tag': weekday_names[day.dayofweek], 'Datum': date, 'Zeit': '', 'Gehzeit': '', 'Differenz': 0.0, 'Summe': 0.0}, index=[len(df)])
            # Füge die neue Zeile an den DataFrame an
            df = pd.concat([df, new_row])
            
            red_fill = PatternFill(patternType='solid', fgColor=Color(rgb='00FF0000'))
            ws = wb.active
            for cell in ws[len(df) + 1]:
                # Füge dem Style-Objekt den Wert der Zelle hinzu
                cell.fill = red_fill




    # Erstelle ein neues Arbeitsblatt für den aktuellen Monat
    ws = wb.create_sheet(title=calendar.month_name[month])

    # Schreibe den DataFrame manuell in das Arbeitsblatt
    for row in df.iterrows():
        ws.append(list(row[1]))

    # Schreibe die Spaltenüberschriften manuell in das Arbeitsblatt
    for col, value in enumerate(df.columns):
        ws.cell(row=1, column=col + 1).value = value

# Speichere das Workbook
wb.save("calendar.xlsx")
