import { useEffect, useRef, useState } from 'react';
import type { FormEvent } from 'react';
import { api, ApiError, ENTRY_TYPES } from '../api/client';
import type { BatchPreview, BatchRequest, BatchResult, EntryType } from '../api/client';

function localDate(): string {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

function displayDate(date: string): string {
  return date.split('-').reverse().join('.');
}

export default function BatchEntryModal({ dates, onCancel, onSaved }: {
  dates: string[] | null;
  onCancel: () => void;
  onSaved: (result: BatchResult) => void;
}) {
  const [selected, setSelected] = useState(dates ?? []);
  const [start, setStart] = useState(localDate);
  const [end, setEnd] = useState(localDate);
  const [type, setType] = useState<EntryType>('VACATION');
  const [inTime, setInTime] = useState('');
  const [outTime, setOutTime] = useState('');
  const [pause, setPause] = useState('');
  const [note, setNote] = useState('');
  const [preview, setPreview] = useState<BatchPreview | null>(null);
  const [overwrite, setOverwrite] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [configReady, setConfigReady] = useState(false);
  const [configAttempt, setConfigAttempt] = useState(0);
  const dialog = useRef<HTMLDivElement>(null);
  const previewHeading = useRef<HTMLHeadingElement>(null);

  useEffect(() => {
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    dialog.current?.querySelector<HTMLElement>('input, select, button')?.focus();
    return () => {
      if (previous?.isConnected) previous.focus();
      else document.querySelector<HTMLButtonElement>('.header-actions button')?.focus();
    };
  }, []);

  useEffect(() => {
    let active = true;
    api.getConfig().then(config => {
      if (!active) return;
      const values = Object.fromEntries(config.map(item => [item.key, item.value]));
      if (!values.default_start || !values.default_end || !Number.isFinite(Number(values.default_pause))) {
        throw new Error('Standardzeiten oder Pause fehlen in der Konfiguration.');
      }
      setInTime(values.default_start);
      setOutTime(values.default_end);
      setPause(String(Math.round(Number(values.default_pause) * 60)));
      setConfigReady(true);
      setError(null);
    }).catch((err: unknown) => {
      if (active) setError(err instanceof Error ? err.message : 'Konfiguration konnte nicht geladen werden.');
    });
    return () => { active = false; };
  }, [configAttempt]);

  function invalidate() {
    setPreview(null);
    setOverwrite(false);
    setError(null);
  }

  useEffect(() => {
    function handleKeys(event: KeyboardEvent) {
      if (event.key === 'Escape' && !busy) {
        event.preventDefault();
        onCancel();
      }
      if (event.key !== 'Tab') return;
      const elements = Array.from(dialog.current?.querySelectorAll<HTMLElement>(
        'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]'
      ) ?? []);
      const first = elements[0], last = elements[elements.length - 1];
      const active = document.activeElement;
      const outside = !(active instanceof HTMLElement) || !elements.includes(active);
      if (event.shiftKey && (active === first || outside)) {
        event.preventDefault();
        last?.focus();
      } else if (!event.shiftKey && (active === last || outside)) {
        event.preventDefault();
        first?.focus();
      }
    }
    document.addEventListener('keydown', handleKeys);
    return () => document.removeEventListener('keydown', handleKeys);
  }, [busy, onCancel]);

  async function submit(event: FormEvent) {
    event.preventDefault();
    if (busy) return;
    if (dates !== null && selected.length === 0) {
      setError('Bitte mindestens einen Tag auswählen.');
      return;
    }
    if (dates === null && end < start) {
      setError('Das Bis-Datum darf nicht vor dem Von-Datum liegen.');
      return;
    }
    const data: BatchRequest = {
      type, note,
      ...(dates === null ? { start_date: start, end_date: end } : { dates: selected }),
      ...(type === 'WORK' ? { stamp_in: inTime, stamp_out: outTime, pause: Number(pause) } : {}),
    };
    setBusy(true);
    setError(null);
    try {
      if (preview) {
        const result = await api.saveBatch({ ...data, preview_token: preview.preview_token, overwrite });
        onSaved(result);
      } else {
        setPreview(await api.previewBatch(data));
        setOverwrite(false);
        requestAnimationFrame(() => previewHeading.current?.focus());
      }
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Die Anfrage ist fehlgeschlagen.');
      if (err instanceof ApiError && err.status === 409) {
        setPreview(null);
        setOverwrite(false);
      }
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="modal-overlay" onClick={() => { if (!busy) onCancel(); }}>
      <div ref={dialog} className="modal batch-modal" role="dialog" aria-modal="true"
           aria-labelledby="batch-title" onClick={event => event.stopPropagation()}>
        <h2 className="modal-title" id="batch-title">Mehrere Tage erfassen</h2>
        <p className="batch-hint">Wochenenden und Feiertage werden automatisch ausgelassen.</p>
        <form onSubmit={submit}>
          <fieldset className="modal-fields batch-fields" disabled={busy}>
            {dates === null ? (
              <div className="batch-date-range">
                <label><span>Von (einschließlich)</span>
                  <input autoFocus type="date" required value={start} onChange={event => { setStart(event.target.value); invalidate(); }} />
                </label>
                <label><span>Bis (einschließlich)</span>
                  <input type="date" required value={end} min={start} onChange={event => { setEnd(event.target.value); invalidate(); }} />
                </label>
              </div>
            ) : (
              <div>
                <div className="batch-hint">{selected.length} Tage ausgewählt</div>
                <ul className="batch-date-list">
                  {selected.map(day => (
                    <li key={day}>
                      <span>{displayDate(day)}</span>
                      <button type="button" className="btn" aria-label={`${displayDate(day)} aus Auswahl entfernen`}
                              onClick={() => {
                                setSelected(selected.filter(item => item !== day));
                                invalidate();
                                requestAnimationFrame(() => dialog.current?.querySelector('select')?.focus());
                              }}>Entfernen</button>
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <label><span>Tagesart</span>
              <select value={type} onChange={event => { setType(event.target.value as EntryType); invalidate(); }}>
                {ENTRY_TYPES.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}
              </select>
            </label>
            {type === 'WORK' && (
              <>
                {!configReady && <button type="button" className="btn" onClick={() => setConfigAttempt(configAttempt + 1)}>Standardwerte erneut laden</button>}
                <div className="batch-date-range">
                  <label><span>Kommen</span>
                    <input type="time" required value={inTime} onChange={event => { setInTime(event.target.value); invalidate(); }} />
                  </label>
                  <label><span>Gehen</span>
                    <input type="time" required value={outTime} onChange={event => { setOutTime(event.target.value); invalidate(); }} />
                  </label>
                </div>
                <label><span>Pause (Minuten)</span>
                  <input type="number" required min="0" step="any" value={pause} onChange={event => { setPause(event.target.value); invalidate(); }} />
                </label>
                <p className="batch-hint">Zeiten und Pause gelten für alle ausgewählten Arbeitstage.</p>
              </>
            )}
            <label><span>Gemeinsamer Hinweis</span>
              <input type="text" value={note} onChange={event => { setNote(event.target.value); invalidate(); }} placeholder="Optional" />
            </label>
            <p className="batch-hint">Beim Überschreiben werden bestehende Tageswerte vollständig ersetzt.
              Ein leerer Hinweis entfernt vorhandene Hinweise.</p>
            {preview && (
              <section className="batch-preview" aria-labelledby="batch-preview-title">
                <h3 ref={previewHeading} tabIndex={-1} id="batch-preview-title">Vorschau: {preview.dates.length} Arbeitstage</h3>
                <details>
                  <summary>Betroffene Tage anzeigen</summary>
                  <ul className="batch-date-list">{preview.dates.map(day => <li key={day}>{displayDate(day)}</li>)}</ul>
                </details>
                {preview.skipped.length > 0 && (
                  <details>
                    <summary>{preview.skipped.length} Tage werden ausgelassen</summary>
                    <ul className="batch-date-list">{preview.skipped.map(day => <li key={day.date}>{displayDate(day.date)}: {day.reason}</li>)}</ul>
                  </details>
                )}
                {preview.conflicts.length > 0 && (
                  <div className="batch-conflicts">
                    <p><strong>{preview.conflicts.length} vorhandene Einträge werden ersetzt:</strong></p>
                    <ul className="batch-date-list">
                      {preview.conflicts.map(entry => (
                        <li key={entry.date}>
                          {displayDate(entry.date)} · {ENTRY_TYPES.find(option => option.value === entry.type)?.label ?? entry.type}
                          {entry.stamp_in && ` · ${entry.stamp_in.slice(0, 5)} – ${entry.stamp_out?.slice(0, 5) ?? 'offen'}`}
                          {entry.note && ` · ${entry.note}`}
                        </li>
                      ))}
                    </ul>
                    <label className="batch-confirm">
                      <input type="checkbox" checked={overwrite} onChange={event => setOverwrite(event.target.checked)} />
                      Vorhandene Einträge überschreiben
                    </label>
                  </div>
                )}
              </section>
            )}
          </fieldset>
          {error && <p className="batch-error" role="alert">{error}</p>}
          <div className="modal-actions">
            <button type="button" className="btn" disabled={busy} onClick={onCancel}>Abbrechen</button>
            <button type="submit" className="btn btn-accent"
                    disabled={busy || (dates !== null && selected.length === 0) || (type === 'WORK' && !configReady)
                      || (!!preview?.conflicts.length && !overwrite)}>
              {busy ? 'Bitte warten…' : preview ? 'Jetzt speichern' : 'Vorschau anzeigen'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
