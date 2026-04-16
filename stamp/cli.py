"""CLI-Interface für das Zeiterfassungs-Tool."""

from datetime import date, time, datetime
from typing import Optional

import typer
from rich.console import Console

from stamp.db import init_db, get_config, set_config, get_session, Config
from stamp import service, tui

app = typer.Typer(
    name="stamp",
    help="⏱️  Interaktives Zeiterfassungs-Tool",
    no_args_is_help=False,
    invoke_without_command=True,
)
console = Console()


def _parse_time(time_str: str) -> time:
    """Parst HH:MM zu time-Objekt."""
    try:
        return datetime.strptime(time_str, "%H:%M").time()
    except ValueError:
        raise typer.BadParameter(f"Ungültiges Zeitformat: '{time_str}'. Erwartet: HH:MM")


def _parse_date(date_str: str) -> date:
    """Parst DD.MM.YYYY oder YYYY-MM-DD zu date-Objekt."""
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m."):
        try:
            parsed = datetime.strptime(date_str, fmt)
            if fmt == "%d.%m.":
                parsed = parsed.replace(year=date.today().year)
            return parsed.date()
        except ValueError:
            continue
    raise typer.BadParameter(f"Ungültiges Datumsformat: '{date_str}'. Erwartet: DD.MM.YYYY oder YYYY-MM-DD")


@app.callback()
def main(ctx: typer.Context):
    """⏱️  stamp — Zeiterfassung leicht gemacht."""
    init_db()
    if ctx.invoked_subcommand is None:
        tui.show_today()


# --- Stempel-Commands ---

@app.command("in")
def cmd_in(
    zeit: Optional[str] = typer.Option(None, "--time", "-t", help="Manuelle Zeit (HH:MM)"),
):
    """Einstempeln (Kommen)."""
    stamp_time = _parse_time(zeit) if zeit else None
    try:
        entry = service.stamp_in(stamp_time)
        tui.show_stamp_result(entry, "in")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command("out")
def cmd_out(
    zeit: Optional[str] = typer.Option(None, "--time", "-t", help="Manuelle Zeit (HH:MM)"),
):
    """Ausstempeln (Gehen)."""
    stamp_time = _parse_time(zeit) if zeit else None
    try:
        entry = service.stamp_out(stamp_time)
        tui.show_stamp_result(entry, "out")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command()
def status():
    """Heutigen Status anzeigen."""
    tui.show_today()


@app.command()
def today():
    """Tagesdetails anzeigen."""
    tui.show_today()


@app.command()
def week():
    """Wochenübersicht anzeigen."""
    tui.show_week()


@app.command()
def month(
    monat: Optional[int] = typer.Argument(None, help="Monat (1-12)"),
    year: Optional[int] = typer.Option(None, "--year", "-y", help="Jahr"),
):
    """Monatsübersicht anzeigen."""
    tui.show_month(monat, year)


@app.command()
def overtime():
    """Überstunden-Stand anzeigen."""
    tui.show_overtime()


# --- Abwesenheiten ---

@app.command()
def vacation(
    datum: str = typer.Argument(..., help="Startdatum (DD.MM.YYYY)"),
    bis: Optional[str] = typer.Option(None, "--to", "-t", help="Enddatum (DD.MM.YYYY)"),
    note: Optional[str] = typer.Option(None, "--note", "-n", help="Notiz"),
):
    """Urlaub eintragen."""
    start = _parse_date(datum)
    end = _parse_date(bis) if bis else None
    try:
        entries = service.add_absence("VACATION", start, end, note)
        tui.show_absence_result(entries, "VACATION")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command()
def sick(
    datum: str = typer.Argument(..., help="Startdatum (DD.MM.YYYY)"),
    bis: Optional[str] = typer.Option(None, "--to", "-t", help="Enddatum (DD.MM.YYYY)"),
    note: Optional[str] = typer.Option(None, "--note", "-n", help="Notiz"),
):
    """Kranktag eintragen."""
    start = _parse_date(datum)
    end = _parse_date(bis) if bis else None
    try:
        entries = service.add_absence("SICK", start, end, note)
        tui.show_absence_result(entries, "SICK")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command()
def flex(
    datum: str = typer.Argument(..., help="Datum (DD.MM.YYYY)"),
):
    """Gleittag eintragen."""
    day = _parse_date(datum)
    try:
        entries = service.add_absence("FLEX", day)
        tui.show_absence_result(entries, "FLEX")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command()
def travel(
    datum: str = typer.Argument(..., help="Startdatum (DD.MM.YYYY)"),
    bis: Optional[str] = typer.Option(None, "--to", "-t", help="Enddatum (DD.MM.YYYY)"),
    note: Optional[str] = typer.Option(None, "--note", "-n", help="Notiz"),
):
    """Dienstreise eintragen."""
    start = _parse_date(datum)
    end = _parse_date(bis) if bis else None
    try:
        entries = service.add_absence("TRAVEL", start, end, note)
        tui.show_absence_result(entries, "TRAVEL")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command()
def cancel(
    datum: str = typer.Argument(..., help="Datum (DD.MM.YYYY)"),
):
    """Eintrag löschen/stornieren."""
    day = _parse_date(datum)
    if service.cancel_entry(day):
        console.print(f"[green]✓ Eintrag für {day.strftime('%d.%m.%Y')} gelöscht.[/green]")
    else:
        console.print(f"[yellow]Kein Eintrag für {day.strftime('%d.%m.%Y')} gefunden.[/yellow]")


@app.command()
def edit(
    datum: str = typer.Argument(..., help="Datum (DD.MM.YYYY)"),
    ein: Optional[str] = typer.Option(None, "--in", "-i", help="Kommen-Zeit (HH:MM)"),
    aus: Optional[str] = typer.Option(None, "--out", "-o", help="Gehen-Zeit (HH:MM)"),
    pause: Optional[float] = typer.Option(None, "--pause", "-p", help="Pause in Minuten"),
    typ: Optional[str] = typer.Option(None, "--type", help="Typ (WORK/VACATION/SICK/FLEX/TRAVEL)"),
    note: Optional[str] = typer.Option(None, "--note", "-n", help="Notiz"),
):
    """Tag bearbeiten."""
    day = _parse_date(datum)
    try:
        entry = service.edit_stamp(
            day,
            stamp_in_time=_parse_time(ein) if ein else None,
            stamp_out_time=_parse_time(aus) if aus else None,
            pause=pause / 60 if pause is not None else None,
            entry_type=typ,
            note=note,
        )
        console.print(f"[green]✓ Eintrag für {day.strftime('%d.%m.%Y')} aktualisiert.[/green]")
        if entry.stamp_in and entry.stamp_out:
            tui.show_stamp_result(entry, "out")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


@app.command("pause")
def cmd_pause(
    minuten: float = typer.Argument(..., help="Pause in Minuten"),
):
    """Pause für heute setzen."""
    try:
        entry = service.set_pause(minuten)
        console.print(f"[green]✓ Pause auf {minuten:.0f} Minuten gesetzt.[/green]")
        if entry.stamp_in and entry.stamp_out:
            tui.show_stamp_result(entry, "out")
    except ValueError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


# --- Übersichten ---

@app.command("vacation-left")
def vacation_left():
    """Resturlaub anzeigen."""
    tui.show_vacation()


@app.command()
def check():
    """Fehlende Einträge prüfen."""
    tui.show_missing_days()


@app.command("import")
def cmd_import(
    datei: str = typer.Argument(..., help="Pfad zur Excel-Datei"),
    year: Optional[int] = typer.Option(None, "--year", "-y", help="Jahr"),
):
    """Bestehende Excel-Datei importieren."""
    from stamp.data_migration import import_excel
    try:
        stats = import_excel(datei, year)
        console.print(f"[green]✓ Import abgeschlossen![/green]")
        console.print(f"  📥 {stats['imported']} Einträge importiert")
        console.print(f"  🔄 {stats.get('updated', 0)} aktualisiert")
        console.print(f"  ⏭️  {stats['skipped']} übersprungen")
        if stats['errors']:
            console.print(f"  ⚠️  {stats['errors']} Fehler")
        console.print(f"  📅 {stats['months']} Monate verarbeitet")
    except FileNotFoundError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)


# --- Konfiguration ---

@app.command()
def export(
    year: Optional[int] = typer.Option(None, "--year", "-y", help="Jahr"),
    file: Optional[str] = typer.Option(None, "--file", "-f", help="Ausgabedatei"),
):
    """Excel-Export (Stundenschreibung)."""
    from stamp.excel_export import export_excel
    y = year or date.today().year
    output = export_excel(y, file)
    console.print(f"[green]✓ Excel exportiert: {output}[/green]")


config_app = typer.Typer(help="⚙️  Konfiguration verwalten")
app.add_typer(config_app, name="config")


@config_app.command("show")
def config_show():
    """Aktuelle Konfiguration anzeigen."""
    from rich.table import Table
    table = Table(title="⚙️  Konfiguration", show_lines=True)
    table.add_column("Schlüssel", style="bold")
    table.add_column("Wert")

    with get_session() as session:
        for entry in session.query(Config).order_by(Config.key).all():
            table.add_row(entry.key, entry.value)

    console.print(table)


@config_app.command("set")
def config_set(
    key: str = typer.Argument(..., help="Schlüssel"),
    value: str = typer.Argument(..., help="Wert"),
):
    """Konfigurationswert setzen."""
    set_config(key, value)
    console.print(f"[green]✓ {key} = {value}[/green]")


@app.command()
def serve(
    port: int = typer.Option(8000, "--port", "-p", help="Port"),
    host: str = typer.Option("0.0.0.0", "--host", help="Host"),
):
    """Web-Dashboard starten."""
    import uvicorn
    console.print(f"[bold green]🚀 stamp Web-Dashboard[/bold green]")
    console.print(f"   http://localhost:{port}")
    console.print(f"   API-Docs: http://localhost:{port}/docs")
    console.print(f"   [dim]Strg+C zum Beenden[/dim]\n")
    uvicorn.run("stamp.api.main:app", host=host, port=port, reload=False)


if __name__ == "__main__":
    app()
