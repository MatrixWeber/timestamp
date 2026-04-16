"""Rich-basierte TUI-Ansichten für die Zeiterfassung."""

from datetime import date, timedelta
import calendar

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box

from stamp.db import Stamp
from stamp.service import (
    get_today, get_week, get_month, get_year,
    get_overtime_total, get_vacation_balance, get_missing_days,
    is_workday,
)
from stamp.holidays import is_holiday, get_holidays_for_year

console = Console()

WEEKDAYS_DE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
MONTHS_DE = [
    "", "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

TYPE_LABELS = {
    "WORK": "Arbeit",
    "VACATION": "🌴 Urlaub",
    "SICK": "🤒 Krank",
    "FLEX": "⚡ Gleittag",
    "TRAVEL": "✈️  Dienstreise",
}

TYPE_COLORS = {
    "WORK": "white",
    "VACATION": "cyan",
    "SICK": "yellow",
    "FLEX": "green",
    "TRAVEL": "blue",
}


def _format_time(t) -> str:
    if t is None:
        return "—"
    return t.strftime("%H:%M")


def _format_hours(h) -> str:
    if h is None:
        return "—"
    return f"{h:+.2f}" if h < 0 or h > 0 else f"{h:.2f}"


def _overtime_style(h) -> str:
    if h is None:
        return "white"
    return "green" if h >= 0 else "red"


def show_stamp_result(entry: Stamp, action: str):
    """Zeigt das Ergebnis einer Stempel-Aktion."""
    wd = WEEKDAYS_DE[entry.date.weekday()]
    day_str = f"{wd}, {entry.date.strftime('%d.%m.%Y')}"

    if action == "in":
        console.print(
            Panel(
                f"[bold green]✓ Eingestempelt[/bold green]\n"
                f"📅 {day_str}\n"
                f"🕐 Kommen: [bold]{_format_time(entry.stamp_in)}[/bold]\n"
                f"⏸️  Pause: {entry.pause * 60:.0f} Min",
                title="Stempel",
                border_style="green",
            )
        )
    elif action == "out":
        ot_style = _overtime_style(entry.overtime)
        console.print(
            Panel(
                f"[bold green]✓ Ausgestempelt[/bold green]\n"
                f"📅 {day_str}\n"
                f"🕐 Kommen: {_format_time(entry.stamp_in)}  →  "
                f"Gehen: [bold]{_format_time(entry.stamp_out)}[/bold]\n"
                f"⏸️  Pause: {entry.pause * 60:.0f} Min\n"
                f"⏱️  Arbeitszeit: [bold]{entry.work_hours:.2f}h[/bold]\n"
                f"📊 Überstunden: [{ot_style}][bold]{_format_hours(entry.overtime)}h[/bold][/{ot_style}]",
                title="Stempel",
                border_style="green",
            )
        )


def show_today():
    """Zeigt den heutigen Status."""
    entry = get_today()
    today = date.today()
    wd = WEEKDAYS_DE[today.weekday()]
    day_str = f"{wd}, {today.strftime('%d.%m.%Y')}"

    holiday, holiday_name = is_holiday(today)
    if holiday:
        console.print(Panel(f"📅 {day_str}\n🎉 Feiertag: {holiday_name}", title="Heute", border_style="magenta"))
        return

    if today.weekday() >= 5:
        console.print(Panel(f"📅 {day_str}\n🛋️  Wochenende!", title="Heute", border_style="dim"))
        return

    if not entry:
        console.print(
            Panel(f"📅 {day_str}\n⚠️  [yellow]Noch nicht gestempelt![/yellow]\n\nNutze: [bold]stamp in[/bold]",
                  title="Heute", border_style="yellow")
        )
        return

    if entry.type != "WORK":
        label = TYPE_LABELS.get(entry.type, entry.type)
        color = TYPE_COLORS.get(entry.type, "white")
        console.print(Panel(f"📅 {day_str}\n[{color}]{label}[/{color}]", title="Heute", border_style=color))
        return

    lines = [f"📅 {day_str}"]
    lines.append(f"🕐 Kommen: [bold]{_format_time(entry.stamp_in)}[/bold]")

    if entry.stamp_out:
        lines.append(f"🕐 Gehen:  [bold]{_format_time(entry.stamp_out)}[/bold]")
        lines.append(f"⏸️  Pause:  {entry.pause * 60:.0f} Min")
        lines.append(f"⏱️  Arbeitszeit: [bold]{entry.work_hours:.2f}h[/bold]")
        ot_style = _overtime_style(entry.overtime)
        lines.append(f"📊 Überstunden:  [{ot_style}][bold]{_format_hours(entry.overtime)}h[/bold][/{ot_style}]")
        border = "green"
    else:
        lines.append(f"⏸️  Pause:  {entry.pause * 60:.0f} Min")
        lines.append(f"\n[yellow]⏳ Noch nicht ausgestempelt[/yellow]")
        border = "blue"

    console.print(Panel("\n".join(lines), title="Heute", border_style=border))


def show_week(ref_date: date | None = None):
    """Zeigt die Wochenübersicht."""
    if ref_date is None:
        ref_date = date.today()

    monday = ref_date - timedelta(days=ref_date.weekday())
    friday = monday + timedelta(days=4)

    table = Table(
        title=f"📅 KW {ref_date.isocalendar()[1]} ({monday.strftime('%d.%m.')} – {friday.strftime('%d.%m.%Y')})",
        box=box.ROUNDED,
        show_lines=True,
    )
    table.add_column("Tag", style="bold", width=12)
    table.add_column("Kommen", justify="center", width=8)
    table.add_column("Gehen", justify="center", width=8)
    table.add_column("Pause", justify="center", width=7)
    table.add_column("AZ", justify="center", width=7)
    table.add_column("ÜS", justify="center", width=7)
    table.add_column("Typ", width=14)

    entries = get_week(ref_date)
    entry_map = {e.date: e for e in entries}

    total_work = 0.0
    total_overtime = 0.0

    for i in range(5):  # Mo-Fr
        day = monday + timedelta(days=i)
        wd = WEEKDAYS_DE[i]
        day_str = f"{wd} {day.strftime('%d.%m.')}"

        holiday, hname = is_holiday(day)
        if holiday:
            table.add_row(day_str, "", "", "", "", "", f"[magenta]🎉 {hname}[/magenta]")
            continue

        entry = entry_map.get(day)
        if not entry:
            if day <= date.today():
                style = "yellow" if day < date.today() else "dim"
                table.add_row(f"[{style}]{day_str}[/{style}]", "—", "—", "—", "—", "—", f"[{style}]fehlt[/{style}]")
            else:
                table.add_row(f"[dim]{day_str}[/dim]", "", "", "", "", "", "")
            continue

        if entry.type != "WORK":
            label = TYPE_LABELS.get(entry.type, entry.type)
            color = TYPE_COLORS.get(entry.type, "white")
            table.add_row(day_str, "", "", "", "", "", f"[{color}]{label}[/{color}]")
            continue

        ot_style = _overtime_style(entry.overtime)
        work_str = f"{entry.work_hours:.2f}" if entry.work_hours else "—"
        ot_str = f"[{ot_style}]{_format_hours(entry.overtime)}[/{ot_style}]" if entry.overtime is not None else "—"

        if entry.work_hours:
            total_work += entry.work_hours
        if entry.overtime:
            total_overtime += entry.overtime

        table.add_row(
            day_str,
            _format_time(entry.stamp_in),
            _format_time(entry.stamp_out),
            f"{entry.pause * 60:.0f}m",
            work_str,
            ot_str,
            "",
        )

    # Summenzeile
    ot_total_style = _overtime_style(total_overtime)
    table.add_row(
        "[bold]Summe[/bold]", "", "", "",
        f"[bold]{total_work:.2f}[/bold]",
        f"[bold][{ot_total_style}]{_format_hours(total_overtime)}[/{ot_total_style}][/bold]",
        "",
    )

    console.print(table)


def show_month(month: int | None = None, year: int | None = None):
    """Zeigt die Monatsübersicht."""
    if month is None:
        month = date.today().month
    if year is None:
        year = date.today().year

    month_name = MONTHS_DE[month]
    num_days = calendar.monthrange(year, month)[1]

    table = Table(
        title=f"📅 {month_name} {year}",
        box=box.ROUNDED,
        show_lines=True,
    )
    table.add_column("Datum", style="bold", width=14)
    table.add_column("Kommen", justify="center", width=8)
    table.add_column("Gehen", justify="center", width=8)
    table.add_column("Pause", justify="center", width=7)
    table.add_column("AZ", justify="center", width=7)
    table.add_column("Soll", justify="center", width=6)
    table.add_column("ÜS", justify="center", width=7)
    table.add_column("Typ", width=14)

    entries = get_month(month, year)
    entry_map = {e.date: e for e in entries}
    holidays = get_holidays_for_year(year)
    target = 8.0  # Will be read from config in service

    total_work = 0.0
    total_target = 0.0
    total_overtime = 0.0

    for day_num in range(1, num_days + 1):
        day = date(year, month, day_num)
        wd = WEEKDAYS_DE[day.weekday()]
        day_str = f"{wd} {day.strftime('%d.%m.')}"

        # Wochenende
        if day.weekday() >= 5:
            table.add_row(f"[dim]{day_str}[/dim]", "", "", "", "", "", "", "[dim]WE[/dim]")
            continue

        # Feiertag
        if day in holidays:
            table.add_row(day_str, "", "", "", "", "", "", f"[magenta]🎉 {holidays[day]}[/magenta]")
            continue

        entry = entry_map.get(day)
        if not entry:
            if day < date.today():
                table.add_row(f"[yellow]{day_str}[/yellow]", "—", "—", "—", "—", f"{target:.0f}", "—", "[yellow]fehlt[/yellow]")
                total_target += target
            else:
                table.add_row(f"[dim]{day_str}[/dim]", "", "", "", "", f"{target:.0f}", "", "")
                if day == date.today():
                    total_target += target
            continue

        if entry.type != "WORK":
            label = TYPE_LABELS.get(entry.type, entry.type)
            color = TYPE_COLORS.get(entry.type, "white")
            table.add_row(day_str, "", "", "", f"{target:.0f}", f"{target:.0f}", "0.00", f"[{color}]{label}[/{color}]")
            total_work += target
            total_target += target
            continue

        ot_style = _overtime_style(entry.overtime)
        work_str = f"{entry.work_hours:.2f}" if entry.work_hours else "—"
        ot_str = f"[{ot_style}]{_format_hours(entry.overtime)}[/{ot_style}]" if entry.overtime is not None else "—"

        if entry.work_hours:
            total_work += entry.work_hours
        if entry.overtime:
            total_overtime += entry.overtime
        total_target += target

        table.add_row(
            day_str,
            _format_time(entry.stamp_in),
            _format_time(entry.stamp_out),
            f"{entry.pause * 60:.0f}m" if entry.pause else "—",
            work_str,
            f"{target:.0f}",
            ot_str,
            "",
        )

    # Summen
    ot_total_style = _overtime_style(total_overtime)
    table.add_row(
        "[bold]Summe[/bold]", "", "", "",
        f"[bold]{total_work:.2f}[/bold]",
        f"[bold]{total_target:.0f}[/bold]",
        f"[bold][{ot_total_style}]{_format_hours(total_overtime)}[/{ot_total_style}][/bold]",
        "",
    )

    console.print(table)


def show_overtime():
    """Zeigt den Überstunden-Überblick."""
    ot = get_overtime_total()
    table = Table(title="📊 Überstunden", box=box.ROUNDED)
    table.add_column("Zeitraum", style="bold")
    table.add_column("Stunden", justify="right")

    for label, key in [
        ("Heute", "today"),
        ("Diese Woche", "week"),
        ("Dieser Monat", "month"),
        ("Dieses Jahr", "year"),
        ("Übertrag VJ", "carryover"),
        ("─── Gesamt", "total"),
    ]:
        val = ot[key]
        style = _overtime_style(val)
        bold = "[bold]" if key == "total" else ""
        table.add_row(f"{bold}{label}", f"[{style}]{bold}{_format_hours(val)}h")

    console.print(table)


def show_vacation():
    """Zeigt den Resturlaub."""
    vac = get_vacation_balance()
    table = Table(title="🌴 Urlaubsübersicht", box=box.ROUNDED)
    table.add_column("", style="bold")
    table.add_column("Tage", justify="right")

    table.add_row("Jahresanspruch", f"{vac['total']}")
    table.add_row("Übertrag VJ", f"{vac['carryover']}")
    table.add_row("Genommen", f"[cyan]{vac['taken']}[/cyan]")

    remaining_style = "green" if vac["remaining"] > 5 else ("yellow" if vac["remaining"] > 0 else "red")
    table.add_row("[bold]Verbleibend", f"[bold][{remaining_style}]{vac['remaining']}[/{remaining_style}]")
    table.add_row("", "")
    table.add_row("Gleittage", f"{vac['flex_taken']}")
    table.add_row("Kranktage", f"{vac['sick_days']}")
    table.add_row("Dienstreisen", f"{vac['travel_days']}")

    console.print(table)

    if vac["remaining"] <= 5:
        console.print(f"\n[yellow]⚠️  Achtung: Nur noch {vac['remaining']} Urlaubstage übrig![/yellow]")


def show_missing_days(year: int | None = None):
    """Zeigt fehlende Tage an."""
    missing = get_missing_days(year)
    if not missing:
        console.print("[green]✓ Keine fehlenden Einträge![/green]")
        return

    table = Table(title="⚠️  Fehlende Einträge", box=box.ROUNDED)
    table.add_column("Datum", style="bold")
    table.add_column("Tag")

    for day in missing:
        wd = WEEKDAYS_DE[day.weekday()]
        table.add_row(day.strftime("%d.%m.%Y"), wd)

    console.print(table)
    console.print(f"\n[yellow]{len(missing)} Arbeitstage ohne Stempel![/yellow]")


def show_absence_result(entries: list[Stamp], absence_type: str):
    """Zeigt das Ergebnis einer Abwesenheits-Eintragung."""
    label = TYPE_LABELS.get(absence_type, absence_type)
    color = TYPE_COLORS.get(absence_type, "white")

    if len(entries) == 1:
        e = entries[0]
        wd = WEEKDAYS_DE[e.date.weekday()]
        console.print(f"[{color}]✓ {label}[/{color}] eingetragen: {wd}, {e.date.strftime('%d.%m.%Y')}")
    else:
        console.print(f"[{color}]✓ {label}[/{color}] eingetragen: {len(entries)} Tage")
        for e in entries:
            wd = WEEKDAYS_DE[e.date.weekday()]
            console.print(f"  • {wd}, {e.date.strftime('%d.%m.%Y')}")
