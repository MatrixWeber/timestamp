import { useState, useEffect, useCallback, useRef } from 'react';
import { api, DayInfo, OvertimeData, VacationData, MonthData } from './api/client';

type View = 'dashboard' | 'month' | 'year';

const WEEKDAYS_SHORT = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
const MONTHS_DE = ['', 'Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];
const MONTHS_FULL = ['', 'Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];

function fmtTime(t: string | null): string {
  if (!t) return '—';
  return t.substring(0, 5);
}

function fmtHours(h: number | null): string {
  if (h === null || h === undefined) return '—';
  return h >= 0 ? `+${h.toFixed(2)}` : h.toFixed(2);
}

function fmtDate(d: string): string {
  const parts = d.split('-');
  return `${parts[2]}.${parts[1]}.`;
}

function isToday(d: string): boolean {
  return d === new Date().toISOString().split('T')[0];
}

function typeBadge(type: string): { label: string; cls: string } {
  switch (type) {
    case 'VACATION': return { label: '🌴 Urlaub', cls: 'badge-vacation' };
    case 'SICK': return { label: '🤒 Krank', cls: 'badge-sick' };
    case 'FLEX': return { label: '⚡ Gleittag', cls: 'badge-flex' };
    case 'TRAVEL': return { label: '✈️ Dienstreise', cls: 'badge-travel' };
    default: return { label: '', cls: '' };
  }
}

// ==================== CLOCK ====================
function useClock() {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  return now;
}

// ==================== EDIT MODAL ====================
function EditModal({ stamp, onSave, onCancel }: {
  stamp: { date: string; stamp_in: string | null; stamp_out: string | null; pause: number; note: string | null };
  onSave: (data: { stamp_in?: string; stamp_out?: string; pause?: number; note?: string }) => void;
  onCancel: () => void;
}) {
  const [inTime, setInTime] = useState(stamp.stamp_in?.substring(0, 5) || '');
  const [outTime, setOutTime] = useState(stamp.stamp_out?.substring(0, 5) || '');
  const [pause, setPause] = useState(String(Math.round(stamp.pause * 60)));
  const [note, setNote] = useState(stamp.note || '');

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-title">✏️ Tag bearbeiten — {fmtDate(stamp.date)}</div>
        <div className="modal-fields">
          <label>
            <span>Kommen</span>
            <input type="time" value={inTime} onChange={e => setInTime(e.target.value)} />
          </label>
          <label>
            <span>Gehen</span>
            <input type="time" value={outTime} onChange={e => setOutTime(e.target.value)} />
          </label>
          <label>
            <span>Pause (Min)</span>
            <input type="number" value={pause} onChange={e => setPause(e.target.value)} min="0" max="120" />
          </label>
          <label>
            <span>Hinweis</span>
            <input type="text" value={note} onChange={e => setNote(e.target.value)} placeholder="z.B. Homeoffice, Meeting..." />
          </label>
        </div>
        <div className="modal-actions">
          <button className="btn" onClick={onCancel}>Abbrechen</button>
          <button className="btn btn-accent" onClick={() => onSave({
            stamp_in: inTime || undefined,
            stamp_out: outTime || undefined,
            pause: pause ? Number(pause) : undefined,
            note: note,
          })}>Speichern</button>
        </div>
      </div>
    </div>
  );
}

// ==================== STAMP BUTTON ====================
function StampButton({ today, onStamp, onEdit }: {
  today: DayInfo | null; onStamp: () => void; onEdit: () => void;
}) {
  const clock = useClock();
  const hh = String(clock.getHours()).padStart(2, '0');
  const mm = String(clock.getMinutes()).padStart(2, '0');
  const ss = String(clock.getSeconds()).padStart(2, '0');

  const stamp = today?.stamp;
  const hasIn = stamp?.stamp_in != null;
  const hasOut = stamp?.stamp_out != null;

  let btnClass = 'stamp-btn';
  let btnLabel = 'STAMP IN';
  let onClick = onStamp;
  if (hasIn && !hasOut) {
    btnClass += ' stamped-in';
    btnLabel = 'STAMP OUT';
  } else if (hasOut) {
    btnClass += ' stamped-done';
    btnLabel = '✓ FERTIG';
    onClick = onEdit;
  }

  let disabled = false;
  if (today?.is_weekend) { btnLabel = 'WOCHENENDE'; disabled = true; }
  if (today?.is_holiday) { btnLabel = 'FEIERTAG'; disabled = true; }

  return (
    <div className="stamp-section">
      <div className="stamp-clock">
        {hh}:{mm}<span className="stamp-clock-seconds">:{ss}</span>
      </div>
      <button className={btnClass} onClick={disabled ? undefined : onClick} disabled={disabled}>
        {btnLabel}
      </button>
      <div className="stamp-status">
        {stamp?.stamp_in && (
          <>Kommen: <strong>{fmtTime(stamp.stamp_in)}</strong></>
        )}
        {stamp?.stamp_out && (
          <> → Gehen: <strong>{fmtTime(stamp.stamp_out)}</strong></>
        )}
        {stamp?.work_hours != null && (
          <> · <strong>{stamp.work_hours.toFixed(2)}h</strong></>
        )}
      </div>
      {hasOut && (
        <div className="stamp-edit-hint">Klick zum Bearbeiten</div>
      )}
      {stamp?.note && (
        <div className="stamp-note">📝 {stamp.note}</div>
      )}
    </div>
  );
}

// ==================== STAT CARD ====================
function StatCard({ label, value, sub, colorClass, delay }: {
  label: string; value: string; sub?: string; colorClass?: string; delay?: number;
}) {
  return (
    <div className={`card fade-in fade-in-${delay || 1}`}>
      <div className="card-label">{label}</div>
      <div className={`card-value ${colorClass || 'neutral'}`}>{value}</div>
      {sub && <div className="card-sub">{sub}</div>}
    </div>
  );
}

// ==================== WEEK TABLE ====================
function WeekTable({ days, totalWork, totalOt, onEditDay, onDeleteDay }: {
  days: DayInfo[]; totalWork: number; totalOt: number;
  onEditDay?: (date: string, data: Record<string, unknown>) => void;
  onDeleteDay?: (date: string) => void;
}) {
  const [editingDay, setEditingDay] = useState<DayInfo | null>(null);

  return (
    <>
    <div className="card table-card fade-in fade-in-4">
      <div className="card-header">
        <span className="card-title">Diese Woche</span>
        <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>Klick zum Bearbeiten</span>
      </div>
      <table>
        <thead>
          <tr>
            <th>Tag</th>
            <th>Kommen</th>
            <th>Gehen</th>
            <th>Pause</th>
            <th>AZ</th>
            <th>ÜS</th>
            <th>Typ</th>
            <th>Hinweis</th>
          </tr>
        </thead>
        <tbody>
          {days.map((d) => {
            const s = d.stamp;
            let rowClass = '';
            if (d.is_weekend) rowClass = 'row-weekend';
            else if (d.is_holiday) rowClass = 'row-holiday';
            else if (s?.type === 'VACATION') rowClass = 'row-vacation';
            else if (s?.type === 'SICK') rowClass = 'row-sick';
            else if (s?.type === 'FLEX') rowClass = 'row-flex';
            else if (s?.type === 'TRAVEL') rowClass = 'row-travel';
            else if (!s && !d.is_weekend && !d.is_holiday && new Date(d.date) < new Date()) rowClass = 'row-missing';
            if (isToday(d.date)) rowClass += ' row-today';

            const badge = s && s.type !== 'WORK' ? typeBadge(s.type) : null;
            const clickable = !d.is_weekend && !d.is_holiday && !!onEditDay;

            return (
              <tr key={d.date} className={`${rowClass} ${clickable ? 'row-clickable' : ''}`}
                  onClick={clickable ? () => setEditingDay(d) : undefined}>
                <td>{d.weekday.substring(0, 2)} {fmtDate(d.date)}</td>
                <td>{s?.type === 'WORK' ? fmtTime(s.stamp_in) : ''}</td>
                <td>{s?.type === 'WORK' ? fmtTime(s.stamp_out) : ''}</td>
                <td>{s?.type === 'WORK' && s.pause ? `${(s.pause * 60).toFixed(0)}m` : ''}</td>
                <td>{s?.type === 'WORK' && s.work_hours != null ? s.work_hours.toFixed(2) : ''}</td>
                <td className={s?.overtime != null ? (s.overtime >= 0 ? 'positive' : 'negative') : ''}>
                  {s?.type === 'WORK' && s.overtime != null ? fmtHours(s.overtime) : ''}
                </td>
                <td>
                  {d.is_holiday && <span className="badge badge-holiday">🎉 {d.holiday_name}</span>}
                  {badge && <span className={`badge ${badge.cls}`}>{badge.label}</span>}
                  {!s && !d.is_weekend && !d.is_holiday && new Date(d.date) < new Date() && (
                    <span className="badge badge-missing">fehlt</span>
                  )}
                </td>
                <td className="note-cell">{s?.note || ''}</td>
              </tr>
            );
          })}
          <tr className="sum-row">
            <td>Summe</td>
            <td></td><td></td><td></td>
            <td>{totalWork.toFixed(2)}</td>
            <td className={totalOt >= 0 ? 'positive' : 'negative'}>{fmtHours(totalOt)}</td>
            <td></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>

    {editingDay && onEditDay && (
      <RowEditModal
        day={editingDay}
        onSave={(date, data) => { onEditDay(date, data); setEditingDay(null); }}
        onCancel={() => setEditingDay(null)}
        onDelete={onDeleteDay ? (date) => { onDeleteDay(date); setEditingDay(null); } : () => setEditingDay(null)}
      />
    )}
    </>
  );
}

// ==================== ROW EDIT MODAL ====================
function RowEditModal({ day, onSave, onCancel, onDelete }: {
  day: DayInfo;
  onSave: (date: string, data: Record<string, unknown>) => void;
  onCancel: () => void;
  onDelete: (date: string) => void;
}) {
  const s = day.stamp;
  const [inTime, setInTime] = useState(s?.stamp_in?.substring(0, 5) || '');
  const [outTime, setOutTime] = useState(s?.stamp_out?.substring(0, 5) || '');
  const [pause, setPause] = useState(s ? String(Math.round(s.pause * 60)) : '45');
  const [note, setNote] = useState(s?.note || '');
  const [entryType, setEntryType] = useState(s?.type || 'WORK');

  const typeOptions = [
    { value: 'WORK', label: '💼 Arbeit' },
    { value: 'VACATION', label: '🌴 Urlaub' },
    { value: 'SICK', label: '🤒 Krank' },
    { value: 'FLEX', label: '⚡ Gleittag' },
    { value: 'TRAVEL', label: '✈️ Dienstreise' },
  ];

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-title">
          ✏️ {day.weekday}, {fmtDate(day.date)}
        </div>
        <div className="modal-fields">
          <label>
            <span>Typ</span>
            <select value={entryType} onChange={e => setEntryType(e.target.value)}>
              {typeOptions.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
            </select>
          </label>
          {entryType === 'WORK' && (
            <>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                <label>
                  <span>Kommen</span>
                  <input type="time" value={inTime} onChange={e => setInTime(e.target.value)} />
                </label>
                <label>
                  <span>Gehen</span>
                  <input type="time" value={outTime} onChange={e => setOutTime(e.target.value)} />
                </label>
              </div>
              <label>
                <span>Pause (Min)</span>
                <input type="number" value={pause} onChange={e => setPause(e.target.value)} min="0" max="120" />
              </label>
            </>
          )}
          <label>
            <span>Hinweis</span>
            <input type="text" value={note} onChange={e => setNote(e.target.value)} placeholder="z.B. Homeoffice, Meeting, AUBIL..." />
          </label>
        </div>
        <div className="modal-actions">
          {s && (
            <button className="btn btn-danger" onClick={() => { if (confirm('Eintrag wirklich löschen?')) onDelete(day.date); }}
                    style={{ marginRight: 'auto' }}>
              🗑 Löschen
            </button>
          )}
          <button className="btn" onClick={onCancel}>Abbrechen</button>
          <button className="btn btn-accent" onClick={() => {
            if (entryType !== 'WORK') {
              onSave(day.date, { type: entryType, note: note });
            } else {
              onSave(day.date, {
                stamp_in: inTime || undefined,
                stamp_out: outTime || undefined,
                pause: pause ? Number(pause) : undefined,
                note: note,
                type: 'WORK',
              });
            }
          }}>Speichern</button>
        </div>
      </div>
    </div>
  );
}

// ==================== MONTH TABLE ====================
function MonthTable({ data, onEditDay, onDeleteDay }: {
  data: MonthData;
  onEditDay: (date: string, data: Record<string, unknown>) => void;
  onDeleteDay: (date: string) => void;
}) {
  const [editingDay, setEditingDay] = useState<DayInfo | null>(null);

  return (
    <>
    <div className="card table-card fade-in fade-in-2">
      <div className="card-header">
        <span className="card-title">{MONTHS_FULL[data.month]} {data.year}</span>
        <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>Klick auf Zeile zum Bearbeiten</span>
      </div>
      <table>
        <thead>
          <tr>
            <th>Datum</th>
            <th>Kommen</th>
            <th>Gehen</th>
            <th>Pause</th>
            <th>AZ</th>
            <th>Soll</th>
            <th>ÜS</th>
            <th>Typ</th>
            <th>Hinweis</th>
          </tr>
        </thead>
        <tbody>
          {data.days.map((d) => {
            const s = d.stamp;
            let rowClass = '';
            if (d.is_weekend) rowClass = 'row-weekend';
            else if (d.is_holiday) rowClass = 'row-holiday';
            else if (s?.type === 'VACATION') rowClass = 'row-vacation';
            else if (s?.type === 'SICK') rowClass = 'row-sick';
            else if (s?.type === 'FLEX') rowClass = 'row-flex';
            else if (s?.type === 'TRAVEL') rowClass = 'row-travel';
            else if (!s && !d.is_weekend && !d.is_holiday && new Date(d.date) < new Date()) rowClass = 'row-missing';
            if (isToday(d.date)) rowClass += ' row-today';

            const clickable = !d.is_weekend && !d.is_holiday;
            const badge = s && s.type !== 'WORK' ? typeBadge(s.type) : null;

            return (
              <tr key={d.date} className={`${rowClass} ${clickable ? 'row-clickable' : ''}`}
                  onClick={clickable ? () => setEditingDay(d) : undefined}>
                <td>{d.weekday.substring(0, 2)} {fmtDate(d.date)}</td>
                <td>{s?.type === 'WORK' ? fmtTime(s.stamp_in) : ''}</td>
                <td>{s?.type === 'WORK' ? fmtTime(s.stamp_out) : ''}</td>
                <td>{s?.type === 'WORK' && s.pause ? `${(s.pause * 60).toFixed(0)}m` : ''}</td>
                <td>{s?.work_hours != null ? s.work_hours.toFixed(2) : ''}</td>
                <td>{!d.is_weekend && !d.is_holiday ? '8' : ''}</td>
                <td className={s?.overtime != null ? (s.overtime >= 0 ? 'positive' : 'negative') : ''}>
                  {s?.overtime != null ? fmtHours(s.overtime) : ''}
                </td>
                <td>
                  {d.is_holiday && <span className="badge badge-holiday">🎉 {d.holiday_name}</span>}
                  {d.is_weekend && <span className="badge badge-weekend">WE</span>}
                  {badge && <span className={`badge ${badge.cls}`}>{badge.label}</span>}
                  {!s && !d.is_weekend && !d.is_holiday && new Date(d.date) < new Date() && (
                    <span className="badge badge-missing">fehlt</span>
                  )}
                </td>
                <td className="note-cell">{s?.note || ''}</td>
              </tr>
            );
          })}
          <tr className="sum-row">
            <td>Summe</td>
            <td></td><td></td><td></td>
            <td>{data.total_work.toFixed(2)}</td>
            <td></td>
            <td className={data.total_overtime >= 0 ? 'positive' : 'negative'}>{fmtHours(data.total_overtime)}</td>
            <td></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>

    {editingDay && (
      <RowEditModal
        day={editingDay}
        onSave={(date, data) => {
          onEditDay(date, data);
          setEditingDay(null);
        }}
        onCancel={() => setEditingDay(null)}
        onDelete={(date) => {
          onDeleteDay(date);
          setEditingDay(null);
        }}
      />
    )}
    </>
  );
}

// ==================== PIE CHART ====================
function PieChart({ segments, size = 160 }: {
  segments: { value: number; color: string; label: string }[];
  size?: number;
}) {
  const total = segments.reduce((s, seg) => s + seg.value, 0);
  if (total === 0) {
    return (
      <div style={{ width: size, height: size, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto' }}>
        <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
          <circle cx={size / 2} cy={size / 2} r={size / 2 - 8} fill="none" stroke="var(--border)" strokeWidth={16} opacity={0.3} />
          <text x="50%" y="50%" textAnchor="middle" dy="0.35em" fill="var(--text-muted)" fontSize={14}>Keine Daten</text>
        </svg>
      </div>
    );
  }

  const cx = size / 2, cy = size / 2, r = size / 2 - 8;
  let cumAngle = -90;

  const arcs = segments.filter(s => s.value > 0).map((seg) => {
    const angle = (seg.value / total) * 360;
    const startAngle = cumAngle;
    const endAngle = cumAngle + angle;
    cumAngle = endAngle;

    const startRad = (startAngle * Math.PI) / 180;
    const endRad = (endAngle * Math.PI) / 180;
    const x1 = cx + r * Math.cos(startRad);
    const y1 = cy + r * Math.sin(startRad);
    const x2 = cx + r * Math.cos(endRad);
    const y2 = cy + r * Math.sin(endRad);
    const largeArc = angle > 180 ? 1 : 0;

    const d = angle >= 359.9
      ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy}`
      : `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;

    return { ...seg, d, angle };
  });

  return (
    <div style={{ textAlign: 'center' }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        {arcs.map((arc, i) => (
          <path key={i} d={arc.d} fill={arc.color} opacity={0.85}>
            <title>{arc.label}: {arc.value}</title>
          </path>
        ))}
        <circle cx={cx} cy={cy} r={r * 0.55} fill="var(--bg-card)" />
        <text x="50%" y="46%" textAnchor="middle" fill="var(--text-primary)" fontSize={22} fontWeight={700}>{total}</text>
        <text x="50%" y="62%" textAnchor="middle" fill="var(--text-muted)" fontSize={10}>Tage</text>
      </svg>
      <div style={{ display: 'flex', gap: 12, justifyContent: 'center', flexWrap: 'wrap', marginTop: 8 }}>
        {arcs.map((seg, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, color: 'var(--text-secondary)' }}>
            <span style={{ width: 10, height: 10, borderRadius: '50%', background: seg.color, display: 'inline-block' }} />
            {seg.label} ({seg.value})
          </div>
        ))}
      </div>
    </div>
  );
}

// ==================== SIDEBAR ====================
function Sidebar({ overtime, vacation, onExport, onImport, importStatus }: {
  overtime: OvertimeData | null;
  vacation: VacationData | null;
  onExport: () => void;
  onImport: (file: File) => void;
  importStatus: string | null;
}) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="sidebar">
      {/* Overtime Panel */}
      <div className="card fade-in fade-in-3">
        <div className="card-label">📊 Überstunden</div>
        {overtime && (
          <div>
            <div className="stat-bar">
              <span className="stat-icon">📅</span>
              <span className="stat-bar-label">Heute</span>
              <span className={`stat-bar-value ${overtime.today >= 0 ? 'positive' : 'negative'}`}>
                {fmtHours(overtime.today)}h
              </span>
            </div>
            <div className="stat-bar">
              <span className="stat-icon">📆</span>
              <span className="stat-bar-label">Diese Woche</span>
              <span className={`stat-bar-value ${overtime.week >= 0 ? 'positive' : 'negative'}`}>
                {fmtHours(overtime.week)}h
              </span>
            </div>
            <div className="stat-bar">
              <span className="stat-icon">🗓️</span>
              <span className="stat-bar-label">Dieser Monat</span>
              <span className={`stat-bar-value ${overtime.month >= 0 ? 'positive' : 'negative'}`}>
                {fmtHours(overtime.month)}h
              </span>
            </div>
            <div className="stat-bar">
              <span className="stat-icon">📊</span>
              <span className="stat-bar-label">Dieses Jahr</span>
              <span className={`stat-bar-value ${overtime.year >= 0 ? 'positive' : 'negative'}`}>
                {fmtHours(overtime.year)}h
              </span>
            </div>
            <div className="stat-bar" style={{ borderTop: '2px solid var(--border-accent)', paddingTop: 14 }}>
              <span className="stat-icon">🏦</span>
              <span className="stat-bar-label">Übertrag VJ</span>
              <span className="stat-bar-value neutral">{fmtHours(overtime.carryover)}h</span>
            </div>
            <div className="stat-bar" style={{ borderBottom: 'none' }}>
              <span className="stat-icon" style={{ fontSize: 20 }}>Σ</span>
              <span className="stat-bar-label" style={{ fontWeight: 700, color: 'var(--text-primary)' }}>Gesamt (bis heute)</span>
              <span className={`stat-bar-value ${overtime.total >= 0 ? 'positive' : 'negative'}`}
                    style={{ fontSize: 18 }}>
                {fmtHours(overtime.total)}h
              </span>
            </div>
          </div>
        )}
      </div>

      {/* Vacation Panel — with taken/planned split */}
      <div className="card fade-in fade-in-4">
        <div className="card-label">🌴 Urlaub & Abwesenheit</div>
        {vacation && (
          <div>
            <div className="stat-bar">
              <span className="stat-icon">📋</span>
              <span className="stat-bar-label">Anspruch</span>
              <span className="stat-bar-value neutral">{vacation.total} Tage</span>
            </div>
            <div className="stat-bar">
              <span className="stat-icon">🏦</span>
              <span className="stat-bar-label">Übertrag VJ</span>
              <span className="stat-bar-value neutral">{vacation.carryover} Tage</span>
            </div>

            {/* Vacation rows */}
            <table className="absence-table">
              <thead>
                <tr>
                  <th></th>
                  <th style={{ fontSize: 10, color: 'var(--text-muted)' }}>Genommen</th>
                  <th style={{ fontSize: 10, color: 'var(--text-muted)' }}>Geplant</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>🌴 Urlaub</td>
                  <td style={{ color: 'var(--cyan)' }}>{vacation.taken}</td>
                  <td style={{ color: 'var(--purple)' }}>{vacation.planned}</td>
                </tr>
                <tr>
                  <td>⚡ Gleittage</td>
                  <td style={{ color: 'var(--cyan)' }}>{vacation.flex_taken}</td>
                  <td style={{ color: 'var(--purple)' }}>{vacation.flex_planned}</td>
                </tr>
                <tr>
                  <td>✈️ Dienstreise</td>
                  <td style={{ color: 'var(--cyan)' }}>{vacation.travel_taken}</td>
                  <td style={{ color: 'var(--purple)' }}>{vacation.travel_planned}</td>
                </tr>
                <tr>
                  <td>🤒 Krank</td>
                  <td colSpan={2} style={{ color: 'var(--red)', textAlign: 'center' }}>{vacation.sick_days} Tage</td>
                </tr>
              </tbody>
            </table>

            <div className="stat-bar" style={{ marginTop: 8 }}>
              <span className="stat-icon">📊</span>
              <span className="stat-bar-label">Verfügbar</span>
              <span className="stat-bar-value neutral">
                {vacation.remaining + vacation.planned} Tage
              </span>
            </div>
            <div className="stat-bar" style={{ fontSize: 11, paddingLeft: 36 }}>
              <span className="stat-bar-label" style={{ color: 'var(--text-muted)' }}>davon geplant</span>
              <span className="stat-bar-value" style={{ color: 'var(--purple)', fontSize: 13 }}>
                {vacation.planned} Tage
              </span>
            </div>
            <div className="stat-bar" style={{ borderBottom: 'none' }}>
              <span className="stat-icon" style={{ fontSize: 20 }}>Σ</span>
              <span className="stat-bar-label" style={{ fontWeight: 700, color: 'var(--text-primary)' }}>Noch frei</span>
              <span className="stat-bar-value" style={{
                fontSize: 18,
                color: vacation.remaining > 5 ? 'var(--green)' : vacation.remaining > 0 ? 'var(--yellow)' : 'var(--red)'
              }}>
                {vacation.remaining} Tage
              </span>
            </div>
          </div>
        )}
      </div>

      {/* Export / Import */}
      <div className="card fade-in fade-in-5">
        <div className="card-label">📥 Export / Import</div>
        {importStatus && (
          <div style={{ padding: '8px 0', fontSize: 13 }}>{importStatus}</div>
        )}
        <button className="btn btn-accent" onClick={onExport} style={{ width: '100%', marginTop: 8 }}>
          📥 Excel herunterladen
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onImport(file);
            e.target.value = '';
          }}
        />
        <button className="btn" onClick={() => fileInputRef.current?.click()} style={{ width: '100%', marginTop: 8 }}>
          📤 Excel importieren
        </button>
      </div>
    </div>
  );
}

// ==================== MONTH SIDEBAR ====================
function MonthSidebar({ monthData, onExport, onImport, importStatus }: {
  monthData: MonthData;
  onExport: () => void;
  onImport: (file: File) => void;
  importStatus: string | null;
}) {
  const stats = monthData.stats;
  const fileInputRef = useRef<HTMLInputElement>(null);

  const segments = [
    { value: stats.work_days, color: '#60a5fa', label: '💼 Arbeit' },
    { value: stats.vacation_days, color: '#4ade80', label: '🌴 Urlaub' },
    { value: stats.flex_days, color: '#a78bfa', label: '⚡ Gleittag' },
    { value: stats.sick_days, color: '#f87171', label: '🤒 Krank' },
    { value: stats.travel_days, color: '#22d3ee', label: '✈️ Reise' },
  ];

  return (
    <div className="sidebar">
      {/* Month Overtime */}
      <div className="card fade-in fade-in-3">
        <div className="card-label">📊 Überstunden {MONTHS_FULL[monthData.month]}</div>
        <div>
          <div className="stat-bar">
            <span className="stat-icon">🗓️</span>
            <span className="stat-bar-label">Gesamtstunden</span>
            <span className="stat-bar-value neutral">{monthData.total_work.toFixed(1)}h</span>
          </div>
          <div className="stat-bar" style={{ borderBottom: 'none' }}>
            <span className="stat-icon" style={{ fontSize: 20 }}>Σ</span>
            <span className="stat-bar-label" style={{ fontWeight: 700, color: 'var(--text-primary)' }}>Überstunden</span>
            <span className={`stat-bar-value ${monthData.total_overtime >= 0 ? 'positive' : 'negative'}`}
                  style={{ fontSize: 18 }}>
              {fmtHours(monthData.total_overtime)}h
            </span>
          </div>
        </div>
      </div>

      {/* Pie Chart */}
      <div className="card fade-in fade-in-4">
        <div className="card-label">📅 Tagesverteilung {MONTHS_FULL[monthData.month]}</div>
        <div style={{ padding: '12px 0' }}>
          <PieChart segments={segments} size={170} />
        </div>
      </div>

      {/* Export / Import */}
      <div className="card fade-in fade-in-5">
        <div className="card-label">📥 Export / Import</div>
        {importStatus && (
          <div style={{ padding: '8px 0', fontSize: 13 }}>{importStatus}</div>
        )}
        <button className="btn btn-accent" onClick={onExport} style={{ width: '100%', marginTop: 8 }}>
          📥 Excel herunterladen
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onImport(file);
            e.target.value = '';
          }}
        />
        <button className="btn" onClick={() => fileInputRef.current?.click()} style={{ width: '100%', marginTop: 8 }}>
          📤 Excel importieren
        </button>
      </div>
    </div>
  );
}

// ==================== APP ====================
export default function App() {
  const [view, setView] = useState<View>('dashboard');
  const [today, setToday] = useState<DayInfo | null>(null);
  const [weekDays, setWeekDays] = useState<DayInfo[]>([]);
  const [weekTotals, setWeekTotals] = useState({ work: 0, ot: 0 });
  const [overtime, setOvertime] = useState<OvertimeData | null>(null);
  const [vacation, setVacation] = useState<VacationData | null>(null);
  const [monthData, setMonthData] = useState<MonthData | null>(null);
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editModal, setEditModal] = useState(false);

  const clock = useClock();
  const dateStr = clock.toLocaleDateString('de-DE', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });

  // Track the current date string to detect day changes
  const currentDateRef = useRef(new Date().toISOString().split('T')[0]);

  const loadDashboard = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const [t, w, ot, vac] = await Promise.all([
        api.getToday(),
        api.getWeek(),
        api.getOvertime(),
        api.getVacation(),
      ]);
      setToday(t);
      setWeekDays(w.days);
      setWeekTotals({ work: w.total_work, ot: w.total_overtime });
      setOvertime(ot);
      setVacation(vac);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const loadMonth = useCallback(async (m: number) => {
    try {
      setLoading(true);
      const data = await api.getMonth(m);
      setMonthData(data);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  // Initial load
  useEffect(() => {
    loadDashboard();
  }, [loadDashboard]);

  // Auto-refresh every 60 seconds + detect day change
  useEffect(() => {
    const interval = setInterval(() => {
      const nowDate = new Date().toISOString().split('T')[0];
      if (nowDate !== currentDateRef.current) {
        currentDateRef.current = nowDate;
        // Day changed — full reload
        loadDashboard();
        if (view === 'month') loadMonth(selectedMonth);
        return;
      }
      // Regular refresh (silent — no loading spinner)
      Promise.all([
        api.getToday(),
        api.getOvertime(),
        api.getVacation(),
      ]).then(([t, ot, vac]) => {
        setToday(t);
        setOvertime(ot);
        setVacation(vac);
      }).catch(() => {});
    }, 60_000);
    return () => clearInterval(interval);
  }, [view, selectedMonth, loadDashboard, loadMonth]);

  // Also refresh when window regains focus (e.g., after being minimized overnight)
  useEffect(() => {
    const onFocus = () => {
      const nowDate = new Date().toISOString().split('T')[0];
      if (nowDate !== currentDateRef.current) {
        currentDateRef.current = nowDate;
        loadDashboard();
        if (view === 'month') loadMonth(selectedMonth);
      } else {
        // Quick silent refresh on focus
        Promise.all([api.getToday(), api.getOvertime(), api.getVacation()])
          .then(([t, ot, vac]) => { setToday(t); setOvertime(ot); setVacation(vac); })
          .catch(() => {});
      }
    };
    window.addEventListener('focus', onFocus);
    return () => window.removeEventListener('focus', onFocus);
  }, [view, selectedMonth, loadDashboard, loadMonth]);

  useEffect(() => {
    if (view === 'month') loadMonth(selectedMonth);
  }, [view, selectedMonth, loadMonth]);

  const handleStamp = async () => {
    try {
      const stamp = today?.stamp;
      if (!stamp?.stamp_in) {
        await api.stampIn();
      } else if (!stamp?.stamp_out) {
        await api.stampOut();
      }
      await loadDashboard();
    } catch (e: any) {
      setError(e.message);
    }
  };

  const handleEdit = async (data: { stamp_in?: string; stamp_out?: string; pause?: number; note?: string }) => {
    try {
      const todayDate = new Date().toISOString().split('T')[0];
      await api.editStamp(todayDate, data);
      setEditModal(false);
      await loadDashboard();
    } catch (e: any) {
      setError(e.message);
    }
  };

  const handleEditDay = async (date: string, data: Record<string, unknown>) => {
    try {
      if (data.type && data.type !== 'WORK') {
        await api.addAbsence(data.type as string, date, undefined, data.note as string | undefined);
      } else {
        await api.editStamp(date, data);
      }
      await loadMonth(selectedMonth);
      // Refresh dashboard data too
      const [ot, vac] = await Promise.all([api.getOvertime(), api.getVacation()]);
      setOvertime(ot);
      setVacation(vac);
    } catch (e: any) {
      setError(e.message);
    }
  };

  const handleDeleteDay = async (date: string) => {
    try {
      await api.deleteStamp(date);
      await loadMonth(selectedMonth);
      const [ot, vac] = await Promise.all([api.getOvertime(), api.getVacation()]);
      setOvertime(ot);
      setVacation(vac);
    } catch (e: any) {
      setError(e.message);
    }
  };

  // Same as handleEditDay but reloads the week/dashboard instead of month
  const handleEditWeekDay = async (date: string, data: Record<string, unknown>) => {
    try {
      if (data.type && data.type !== 'WORK') {
        await api.addAbsence(data.type as string, date, undefined, data.note as string | undefined);
      } else {
        await api.editStamp(date, data);
      }
      await loadDashboard();
    } catch (e: any) {
      setError(e.message);
    }
  };

  const handleDeleteWeekDay = async (date: string) => {
    try {
      await api.deleteStamp(date);
      await loadDashboard();
    } catch (e: any) {
      setError(e.message);
    }
  };

  const [importStatus, setImportStatus] = useState<string | null>(null);
  const handleImport = async (file: File) => {
    try {
      setImportStatus('⏳ Importiere...');
      const result = await api.importExcel(file);
      setImportStatus(`✅ ${result.imported} neu, ${result.updated || 0} aktualisiert, ${result.skipped} übersprungen`);
      await loadMonth(selectedMonth);
      await loadDashboard();
      setTimeout(() => setImportStatus(null), 5000);
    } catch (e: any) {
      setImportStatus(`❌ ${e.message}`);
      setTimeout(() => setImportStatus(null), 8000);
    }
  };

  if (error) {
    return (
      <div className="app">
        <div className="card" style={{ textAlign: 'center', padding: 40 }}>
          <div style={{ fontSize: 24, marginBottom: 8 }}>⚠️</div>
          <div style={{ color: 'var(--red)', marginBottom: 16 }}>{error}</div>
          <button className="btn" onClick={() => { setError(null); loadDashboard(); }}>Erneut versuchen</button>
        </div>
      </div>
    );
  }

  return (
    <div className="app">
      {/* Header */}
      <header className="header">
        <div className="header-left">
          <div className="logo">stamp<span className="logo-dot"></span></div>
          <div className="header-date">{dateStr}</div>
        </div>
      </header>

      {/* Nav */}
      <nav className="nav">
        {(['dashboard', 'month'] as View[]).map((v) => (
          <button key={v} className={`nav-btn ${view === v ? 'active' : ''}`}
                  onClick={() => setView(v)}>
            {v === 'dashboard' ? '📊 Dashboard' : '📅 Monatsansicht'}
          </button>
        ))}
      </nav>

      {loading && <div className="loading-center"><div className="spinner"></div></div>}

      {/* Dashboard View */}
      {!loading && view === 'dashboard' && (
        <>
          {/* Stat Cards */}
          <div className="grid grid-stats">
            <StatCard
              label="Arbeitszeit heute"
              value={today?.stamp?.work_hours != null ? `${today.stamp.work_hours.toFixed(2)}h` : '—'}
              sub={today?.stamp?.stamp_in ? `${fmtTime(today.stamp.stamp_in)} – ${fmtTime(today.stamp.stamp_out)}` : 'Noch nicht gestempelt'}
              delay={1}
            />
            <StatCard
              label="Überstunden heute"
              value={today?.stamp?.overtime != null ? `${fmtHours(today.stamp.overtime)}h` : '0.00h'}
              colorClass={today?.stamp?.overtime != null ? (today.stamp.overtime >= 0 ? 'positive' : 'negative') : 'neutral'}
              delay={2}
            />
            <StatCard
              label="Überstunden gesamt"
              value={overtime ? `${fmtHours(overtime.total)}h` : '—'}
              colorClass={overtime ? (overtime.total >= 0 ? 'positive' : 'negative') : 'neutral'}
              sub={overtime ? `davon ${fmtHours(overtime.carryover)}h Übertrag VJ` : ''}
              delay={3}
            />
            <StatCard
              label="Urlaub frei"
              value={vacation ? `${vacation.remaining} Tage` : '—'}
              colorClass={vacation ? (vacation.remaining > 5 ? 'positive' : vacation.remaining > 0 ? 'neutral' : 'negative') : 'neutral'}
              sub={vacation ? `${vacation.taken} genommen · ${vacation.planned} geplant` : ''}
              delay={4}
            />
          </div>

          {/* Main Content */}
          <div className="grid grid-main">
            <div>
              {/* Stamp Button */}
              <div className="card fade-in fade-in-2" style={{ marginBottom: 20 }}>
                <StampButton today={today} onStamp={handleStamp} onEdit={() => setEditModal(true)} />
              </div>
              {/* Week Table */}
              <WeekTable days={weekDays} totalWork={weekTotals.work} totalOt={weekTotals.ot} onEditDay={handleEditWeekDay} onDeleteDay={handleDeleteWeekDay} />
            </div>
            <Sidebar overtime={overtime} vacation={vacation} onExport={() => api.exportExcel()} onImport={handleImport} importStatus={importStatus} />
          </div>
        </>
      )}

      {/* Month View */}
      {!loading && view === 'month' && (
        <>
          <div style={{ display: 'flex', gap: 8, marginBottom: 20, flexWrap: 'wrap' }}>
            {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
              <button key={m} className={`nav-btn ${selectedMonth === m ? 'active' : ''}`}
                      onClick={() => setSelectedMonth(m)} style={{ padding: '6px 14px', fontSize: 12 }}>
                {MONTHS_DE[m]}
              </button>
            ))}
          </div>
          {monthData && (
            <div className="grid grid-main">
              <MonthTable data={monthData} onEditDay={handleEditDay} onDeleteDay={handleDeleteDay} />
              <MonthSidebar
                monthData={monthData}
                onExport={() => api.exportExcel()}
                onImport={handleImport}
                importStatus={importStatus}
              />
            </div>
          )}
        </>
      )}

      {/* Edit Modal */}
      {editModal && today?.stamp && (
        <EditModal
          stamp={{
            date: today.stamp.date,
            stamp_in: today.stamp.stamp_in,
            stamp_out: today.stamp.stamp_out,
            pause: today.stamp.pause,
            note: today.stamp.note,
          }}
          onSave={handleEdit}
          onCancel={() => setEditModal(false)}
        />
      )}
    </div>
  );
}
