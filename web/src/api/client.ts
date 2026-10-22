const BASE = '/api';

export class ApiError extends Error {
  constructor(message: string, public status: number, public detail: unknown) {
    super(message);
  }
}

async function request<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    headers: { 'Content-Type': 'application/json' },
    ...options,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    const detail: unknown = err.detail;
    const message = typeof detail === 'string' ? detail
      : detail && typeof detail === 'object' && 'message' in detail ? String(detail.message)
      : Array.isArray(detail) ? detail.map(item => `${item.loc?.slice(1).join('.')}: ${item.msg}`).join('; ')
      : 'Anfrage fehlgeschlagen';
    throw new ApiError(message, res.status, detail);
  }
  return res.json();
}

export interface StampEntry {
  id: number;
  date: string;
  stamp_in: string | null;
  stamp_out: string | null;
  pause: number;
  work_hours: number | null;
  overtime: number | null;
  type: string;
  note: string | null;
}

export interface DayInfo {
  date: string;
  weekday: string;
  is_holiday: boolean;
  holiday_name: string | null;
  is_weekend: boolean;
  stamp: StampEntry | null;
}

export interface WeekData {
  kw: number;
  start: string;
  end: string;
  days: DayInfo[];
  total_work: number;
  total_overtime: number;
}

export interface MonthData {
  month: number;
  year: number;
  days: DayInfo[];
  total_work: number;
  total_overtime: number;
  stats: {
    work_days: number;
    vacation_days: number;
    flex_days: number;
    sick_days: number;
    travel_days: number;
  };
}

export interface OvertimeData {
  today: number;
  week: number;
  month: number;
  year: number;
  carryover: number;
  total: number;
}

export interface VacationData {
  total: number;
  carryover: number;
  taken: number;
  planned: number;
  remaining: number;
  flex_taken: number;
  flex_planned: number;
  sick_days: number;
  travel_taken: number;
  travel_planned: number;
}

export interface ConfigEntry {
  key: string;
  value: string;
}

export const ENTRY_TYPES = [
  { value: 'WORK', label: '💼 Arbeit' },
  { value: 'VACATION', label: '🌴 Urlaub' },
  { value: 'SICK', label: '🤒 Krank' },
  { value: 'FLEX', label: '⚡ Gleittag' },
  { value: 'TRAVEL', label: '✈️ Dienstreise' },
] as const;

export type EntryType = typeof ENTRY_TYPES[number]['value'];

export interface BatchRequest {
  type: EntryType;
  start_date?: string;
  end_date?: string;
  dates?: string[];
  stamp_in?: string;
  stamp_out?: string;
  pause?: number;
  note: string;
}

export interface SkippedDay {
  date: string;
  reason: string;
}

export interface BatchPreview {
  dates: string[];
  skipped: SkippedDay[];
  conflicts: StampEntry[];
  preview_token: string;
}

export interface BatchResult {
  created: number;
  replaced: number;
  skipped: SkippedDay[];
  entries: StampEntry[];
}

export const api = {
  getToday: () => request<DayInfo>('/today'),
  getWeek: () => request<WeekData>('/week'),
  getMonth: (month: number, year?: number) =>
    request<MonthData>(`/month/${month}${year ? `?year=${year}` : ''}`),
  getOvertime: () => request<OvertimeData>('/overtime'),
  getVacation: () => request<VacationData>('/vacation'),
  getMissing: () => request<{ date: string; weekday: string }[]>('/missing'),
  getConfig: () => request<ConfigEntry[]>('/config'),
  previewBatch: (data: BatchRequest) => request<BatchPreview>('/stamps/batch/preview', {
    method: 'POST', body: JSON.stringify(data),
  }),
  saveBatch: (data: BatchRequest & { preview_token: string; overwrite: boolean }) =>
    request<BatchResult>('/stamps/batch', { method: 'POST', body: JSON.stringify(data) }),

  stampIn: (time?: string) => request<StampEntry>('/stamp/in', {
    method: 'POST', body: JSON.stringify({ time }),
  }),
  stampOut: (time?: string) => request<StampEntry>('/stamp/out', {
    method: 'POST', body: JSON.stringify({ time }),
  }),

  addAbsence: (type: string, start_date: string, end_date?: string, note?: string) =>
    request<{ count: number; entries: StampEntry[] }>('/absence', {
      method: 'POST',
      body: JSON.stringify({ type, start_date, end_date, note }),
    }),

  editStamp: (date: string, data: Record<string, unknown>) =>
    request<StampEntry>(`/stamp/${date}`, {
      method: 'PUT', body: JSON.stringify(data),
    }),

  deleteStamp: (date: string) =>
    request<{ deleted: boolean }>(`/stamp/${date}`, { method: 'DELETE' }),

  exportExcel: (year?: number) => {
    const url = year ? `${BASE}/export/excel?year=${year}` : `${BASE}/export/excel`;
    window.open(url, '_blank');
  },

  importExcel: async (file: File, year?: number): Promise<{ success: boolean; imported: number; updated?: number; skipped: number }> => {
    const form = new FormData();
    form.append('file', file);
    const url = year ? `${BASE}/import/excel?year=${year}` : `${BASE}/import/excel`;
    const res = await fetch(url, { method: 'POST', body: form });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(err.detail || 'Import failed');
    }
    return res.json();
  },
};
