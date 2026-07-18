import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Car, Wrench, CalendarClock, BarChart3, Plus, ChevronRight, ChevronLeft,
  Gauge, Fuel, Palette, StickyNote, Trash2, Check, Search, MoreVertical,
  FileSpreadsheet, X as XIcon, Armchair, ClipboardList, Layers, AlertTriangle,
  MapPin, User, Calendar, Hash, Home, Wrench as WrenchHub,
} from 'lucide-react';

/* Adattamento per l'uso fuori da Claude: window.storage (solo per gli artifact)
   e' sostituito con un piccolo wrapper sopra il localStorage del browser,
   con la stessa identica interfaccia (get/set restituiscono {value}). */
const storage = {
  get: async (key) => {
    const v = localStorage.getItem('manutenzione:' + key);
    if (v === null) throw new Error('not found');
    return { value: v };
  },
  set: async (key, value) => {
    localStorage.setItem('manutenzione:' + key, value);
    return { value };
  },
};


/* =========================================================================
   SEZIONE CONDIVISA — usata da tutti i moduli (Mezzi, Carrozzine, futuri)
   ========================================================================= */

const fmtDate = (iso) => { if (!iso) return '—'; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };

function Pill({ style, children, sz }) {
  const s = style || { bg: '#EEE', fg: '#666' };
  return (
    <span style={{ background: s.bg, color: s.fg, fontSize: sz || 11, fontWeight: 700, padding: '3px 9px', borderRadius: 999, whiteSpace: 'nowrap', display: 'inline-block' }}>
      {children}
    </span>
  );
}

function Card({ theme, children, onClick, style }) {
  return (
    <div onClick={onClick} style={{ background: theme.surface, borderRadius: 14, border: `1px solid ${theme.line}`, padding: 14, cursor: onClick ? 'pointer' : 'default', ...style }}>
      {children}
    </div>
  );
}

function Empty({ theme, icon: Icon, text }) {
  return (
    <div style={{ textAlign: 'center', padding: '60px 24px', color: theme.muted }}>
      <Icon size={30} style={{ opacity: 0.35, marginBottom: 10 }} />
      <div style={{ fontSize: 14 }}>{text}</div>
    </div>
  );
}

function SectionLabel({ theme, children }) {
  return <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 700, fontSize: 12.5, textTransform: 'uppercase', letterSpacing: '0.05em', color: theme.primary, margin: '18px 2px 8px' }}>{children}</div>;
}

function StatCard({ theme, label, value, accent }) {
  return (
    <div style={{ background: theme.surface, border: `1px solid ${theme.line}`, borderRadius: 14, padding: '13px 14px', borderLeft: `4px solid ${accent}` }}>
      <div style={{ fontSize: 10.5, color: theme.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 4 }}>{label}</div>
      <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 19, color: theme.ink }}>{value}</div>
    </div>
  );
}

function InfoRow({ theme, icon: Icon, label, value }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <Icon size={13} color={theme.muted} style={{ flexShrink: 0 }} />
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 10, color: theme.muted, textTransform: 'uppercase', letterSpacing: '0.03em' }}>{label}</div>
        <div style={{ fontWeight: 600, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{value || '—'}</div>
      </div>
    </div>
  );
}

function TopBar({ theme, title, subtitle, onBack, backIcon: BackIcon = ChevronLeft, right }) {
  return (
    <div style={{ background: theme.primary, padding: '18px 18px 16px', color: '#fff', position: 'sticky', top: 0, zIndex: 10 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        {onBack && (
          <button onClick={onBack} style={{ background: 'rgba(255,255,255,0.12)', border: 'none', borderRadius: 999, width: 32, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', flexShrink: 0 }}>
            <BackIcon size={18} />
          </button>
        )}
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 20, lineHeight: 1.15, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{title}</div>
          {subtitle && <div style={{ fontSize: 12.5, opacity: 0.75, marginTop: 2 }}>{subtitle}</div>}
        </div>
        {right}
      </div>
    </div>
  );
}

function MenuButton({ onClick }) {
  return (
    <button onClick={onClick} style={{ background: 'rgba(255,255,255,0.12)', border: 'none', borderRadius: 999, width: 34, height: 34, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', flexShrink: 0 }}>
      <MoreVertical size={19} />
    </button>
  );
}

function MenuSheet({ theme, onClose, onExport, exportTitle, exportSub }) {
  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 40, display: 'flex', alignItems: 'flex-end' }}>
      <div onClick={onClose} style={{ position: 'absolute', inset: 0, background: 'rgba(20,24,26,0.45)' }} />
      <div style={{ position: 'relative', width: '100%', maxWidth: 480, margin: '0 auto', background: '#fff', borderRadius: '18px 18px 0 0', padding: '10px 16px calc(20px + env(safe-area-inset-bottom))', boxShadow: '0 -8px 24px rgba(0,0,0,0.18)' }}>
        <div style={{ width: 36, height: 4, background: theme.line, borderRadius: 999, margin: '4px auto 14px' }} />
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <span style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 16 }}>Menu</span>
          <button onClick={onClose} style={{ background: theme.bg, border: 'none', borderRadius: 999, width: 30, height: 30, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <XIcon size={16} color={theme.muted} />
          </button>
        </div>
        <button onClick={onExport} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 12, background: theme.bg, border: `1px solid ${theme.line}`, borderRadius: 12, padding: '13px 14px', textAlign: 'left' }}>
          <div style={{ width: 36, height: 36, borderRadius: 10, background: '#DCEEE3', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <FileSpreadsheet size={18} color={theme.ok} />
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>{exportTitle || 'Esporta in Excel'}</div>
            <div style={{ fontSize: 12, color: theme.muted, marginTop: 1 }}>{exportSub || 'Scarica i dati in .xlsx'}</div>
          </div>
        </button>
      </div>
    </div>
  );
}

function BottomNav({ theme, tab, setTab, items }) {
  return (
    <div style={{ position: 'fixed', bottom: 0, left: 0, right: 0, background: '#fff', borderTop: `1px solid ${theme.line}`, display: 'flex', paddingBottom: 'env(safe-area-inset-bottom)', zIndex: 20 }}>
      {items.map(([key, Icon, label]) => {
        const active = tab === key;
        return (
          <button key={key} onClick={() => setTab(key)} style={{ flex: 1, background: 'none', border: 'none', padding: '9px 0 8px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3, color: active ? theme.primary : theme.muted }}>
            <Icon size={21} strokeWidth={active ? 2.4 : 2} />
            <span style={{ fontSize: 10.5, fontWeight: active ? 700 : 500 }}>{label}</span>
          </button>
        );
      })}
    </div>
  );
}

const GLOBAL_FONTS = `
  @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
  * { box-sizing: border-box; }
`;


/* ---------- design tokens ---------- */
const MEZZI_COLORS = {
  bg: '#F4F1E8',
  surface: '#FFFFFF',
  ink: '#20262B',
  muted: '#6B7280',
  line: '#E4DFD1',
  primary: '#25454F',
  primaryDeep: '#152C33',
  amber: '#D98E04',
  danger: '#B23A2E',
  ok: '#3C7A5C',
  plateBand: '#1A4D8F',
};

const TIPI_MANUTENZIONE = [
  'Tagliando Ordinario', 'Cambio Pneumatici', 'Revisione', 'Cambio Freni', 'Cambio Olio',
  'Riparazione Motore', 'Sostituzione Filtri', 'Controllo Climatizzatore', 'Pulizia Interna', 'Autolavaggio',
];
const STATI_MANUTENZIONE = ['Completato', 'In Garanzia', 'Programmato', 'In Attesa Ricambi'];
const TIPI_VEICOLO = ['Auto', 'Furgone', 'Pulmino'];
const CARBURANTI = ['Benzina', 'Diesel', 'GPL', 'Metano', 'Elettrico', 'Ibrido'];

const SEED_VEHICLES = [
  { id: 'v1', targa: 'FV-460-MV', marca: 'Fiat', modello: 'Panda', tipo: 'Auto', anno: 2019, km: 87500, carburante: 'Benzina', colore: 'Bianco', assicurazione: '2026-04-23', revisione: '2026-07-07', bollo: '2026-09-25', note: '' },
  { id: 'v2', targa: 'EV-388-FY', marca: 'Volkswagen', modello: 'Polo', tipo: 'Auto', anno: 2020, km: 177651, carburante: 'Diesel', colore: 'Bianco', assicurazione: '2026-09-05', revisione: '2028-07-31', bollo: '2026-04-08', note: 'Mezzo di rappresentanza' },
  { id: 'v3', targa: 'FG-762-VS', marca: 'Fiat', modello: 'Panda Van', tipo: 'Auto', anno: 2018, km: 125000, carburante: '', colore: 'Bianco', assicurazione: '2027-01-03', revisione: '2026-03-24', bollo: '2026-06-07', note: 'Officina mobile' },
  { id: 'v4', targa: 'EY-466-NZ', marca: 'Renault', modello: 'Kangoo', tipo: 'Pulmino', anno: 2021, km: 42000, carburante: '', colore: 'Bianco', assicurazione: '2026-05-08', revisione: '2027-04-13', bollo: '2026-08-06', note: '' },
  { id: 'v5', targa: '', marca: 'Renault', modello: 'Kangoo', tipo: 'Auto', anno: 2022, km: 28000, carburante: '', colore: 'Grigio', assicurazione: '2026-10-15', revisione: '2027-07-22', bollo: '2026-05-23', note: 'Auto direzione' },
  { id: 'v6', targa: 'EY-005-BS', marca: 'Renault', modello: 'Clio', tipo: 'Auto', anno: null, km: 211917, carburante: '', colore: 'Rosso', assicurazione: '', revisione: '2028-07-31', bollo: '', note: '' },
  { id: 'v7', targa: 'GL-936-ZA', marca: 'Fiat', modello: 'Panda', tipo: 'Auto', anno: null, km: null, carburante: '', colore: 'Bianco', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v8', targa: 'GX-928-NA', marca: 'Fiat', modello: 'Panda', tipo: 'Auto', anno: null, km: null, carburante: '', colore: 'Bianco', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v9', targa: 'DZ-', marca: 'Fiat', modello: 'Punto', tipo: 'Auto', anno: null, km: null, carburante: '', colore: 'Bianco', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v10', targa: 'HA-579-NV', marca: 'Fiat', modello: 'Ducato', tipo: 'Furgone', anno: null, km: null, carburante: '', colore: '', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v11', targa: 'EY-524-BT', marca: 'Fiat', modello: 'Scudo', tipo: 'Furgone', anno: null, km: null, carburante: '', colore: '', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v12', targa: 'FE-859-EJ', marca: 'Volkswagen', modello: 'Crafter', tipo: 'Furgone', anno: null, km: 124378, carburante: '', colore: '', assicurazione: '', revisione: '2028-07-31', bollo: '', note: '' },
  { id: 'v13', targa: 'GR-191-TH', marca: 'Dacia', modello: 'Duster', tipo: 'Auto', anno: null, km: null, carburante: '', colore: '', assicurazione: '', revisione: '', bollo: '', note: '' },
  { id: 'v14', targa: 'EN-', marca: 'Fiat', modello: 'Panda', tipo: 'Auto', anno: 2017, km: 198000, carburante: 'Benzina', colore: 'Bianco', assicurazione: '2026-03-19', revisione: '2026-05-28', bollo: '2027-01-13', note: 'Trasporto merci pesanti' },
];

const SEED_MAINTS = [
  { id: 'm1', targa: 'GR-191-TH', data: '2026-05-07', km: null, tipo: 'Cambio Pneumatici', descrizione: 'Pneumatici Estivi', officina: '', costo: null, stato: 'Programmato' },
  { id: 'm2', targa: 'HA-579-NV', data: '2026-05-04', km: null, tipo: 'Cambio Pneumatici', descrizione: 'Pneumatici Estivi', officina: 'ENI Via Torino', costo: null, stato: 'Completato' },
  { id: 'm3', targa: 'FV-460-MV', data: '2026-04-23', km: null, tipo: 'Pulizia Interna', descrizione: 'Aspirazione abitacolo', officina: 'Melo', costo: null, stato: 'Completato' },
  { id: 'm4', targa: 'EV-388-FY', data: '2026-04-23', km: null, tipo: 'Pulizia Interna', descrizione: 'Cambio pneumatici stagionale', officina: 'Melo', costo: null, stato: 'Completato' },
  { id: 'm5', targa: 'EV-388-FY', data: '2026-07-14', km: null, tipo: 'Revisione', descrizione: '', officina: 'Riparazioni Dario', costo: 64.75, stato: 'Completato' },
];

const DEFAULT_PARAMS = { ivaRate: 22, urgentDays: 30, mediumDays: 90 };
const MEZZI_NAV_ITEMS = [
  ['veicoli', Car, 'Veicoli'],
  ['manutenzioni', Wrench, 'Interventi'],
  ['scadenze', CalendarClock, 'Scadenze'],
  ['dashboard', BarChart3, 'Riepilogo'],
];

/* ---------- helpers ---------- */
const uid = () => Math.random().toString(36).slice(2, 10);
const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtEuro = (n) => n == null || n === '' ? '—' : Number(n).toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
const fmtKm = (n) => n == null || n === '' ? '—' : `${Number(n).toLocaleString('it-IT')} km`;
const daysUntil = (iso) => { if (!iso) return null; const d = new Date(iso + 'T00:00:00'); const t = new Date(); t.setHours(0, 0, 0, 0); return Math.round((d - t) / 86400000); };
const vehicleLabel = (v) => v ? `${v.marca} ${v.modello}${v.targa ? ' · ' + v.targa : ''}` : '—';
const urgencyOf = (days, params) => {
  if (days == null) return null;
  if (days <= params.urgentDays) return 'urgent';
  if (days <= params.mediumDays) return 'medium';
  return 'ok';
};
const URGENCY_STYLE = {
  urgent: { bg: '#FBEAE7', fg: MEZZI_COLORS.danger, label: 'Urgente' },
  medium: { bg: '#FBF1DC', fg: '#9A6B03', label: 'Attenzione' },
  ok: { bg: '#E7F3EC', fg: MEZZI_COLORS.ok, label: 'OK' },
};

/* ---------- small UI atoms ---------- */
function Plate({ targa, size = 'md' }) {
  const has = targa && targa.trim() && !targa.trim().endsWith('-');
  const dims = size === 'sm' ? { h: 22, fs: 11, band: 7 } : { h: 28, fs: 14, band: 9 };
  return (
    <span
      style={{
        display: 'inline-flex', alignItems: 'stretch', height: dims.h,
        borderRadius: 5, overflow: 'hidden', border: '1.5px solid #0F2F52',
        boxShadow: '0 1px 0 rgba(0,0,0,0.08)', flexShrink: 0,
      }}
    >
      <span style={{ width: dims.band, background: MEZZI_COLORS.plateBand, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <span style={{ width: 4, height: 4, borderRadius: 999, background: '#FFD23F' }} />
      </span>
      <span
        style={{
          fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: dims.fs,
          letterSpacing: '0.03em', background: '#FBFBF6', color: has ? '#12181C' : '#B0AB9C',
          padding: '0 8px', display: 'flex', alignItems: 'center', whiteSpace: 'nowrap',
        }}
      >
        {has ? targa : 'senza targa'}
      </span>
    </span>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: 'block', marginBottom: 14 }}>
      <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: MEZZI_COLORS.muted, marginBottom: 5, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  width: '100%', boxSizing: 'border-box', padding: '11px 12px', borderRadius: 10,
  border: `1.5px solid ${MEZZI_COLORS.line}`, fontSize: 15, fontFamily: 'Inter, sans-serif',
  background: '#FCFBF7', color: MEZZI_COLORS.ink, outline: 'none',
};

function FAB({ onClick, label }) {
  return (
    <button
      onClick={onClick}
      style={{
        position: 'fixed', right: 18, bottom: 84, background: MEZZI_COLORS.amber, color: '#241900',
        border: 'none', borderRadius: 999, height: 52, padding: '0 20px 0 16px',
        display: 'flex', alignItems: 'center', gap: 7, fontWeight: 700, fontSize: 14.5,
        boxShadow: '0 6px 16px rgba(217,142,4,0.4)', zIndex: 20,
      }}
    >
      <Plus size={20} strokeWidth={2.6} /> {label}
    </button>
  );
}

/* ---------- Vehicles ---------- */
function VeicoliScreen({ vehicles, maints, params, onOpen, onAdd, onMenu, onHome }) {
  const [q, setQ] = useState('');
  const filtered = vehicles.filter(v => `${v.marca} ${v.modello} ${v.targa}`.toLowerCase().includes(q.toLowerCase()));

  const worstUrgency = (v) => {
    const days = [v.assicurazione, v.revisione, v.bollo].map(daysUntil).filter(d => d != null);
    if (!days.length) return null;
    return urgencyOf(Math.min(...days), params);
  };

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Veicoli" subtitle={`${vehicles.length} in flotta`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ position: 'relative', marginBottom: 12 }}>
          <Search size={16} style={{ position: 'absolute', left: 12, top: 12, color: MEZZI_COLORS.muted }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Cerca targa, marca, modello…"
            style={{ ...inputStyle, paddingLeft: 34 }} />
        </div>
        {filtered.length === 0 && <Empty theme={MEZZI_COLORS} icon={Car} text="Nessun veicolo trovato." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {filtered.map(v => {
            const u = worstUrgency(v);
            const nInterventi = maints.filter(m => m.targa && m.targa === v.targa).length;
            return (
              <Card theme={MEZZI_COLORS} key={v.id} onClick={() => onOpen(v)}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 15.5, color: MEZZI_COLORS.ink, marginBottom: 6 }}>{v.marca} {v.modello}</div>
                    <Plate targa={v.targa} size="sm" />
                  </div>
                  <div style={{ textAlign: 'right', flexShrink: 0 }}>
                    {u && <div style={{ marginBottom: 6 }}><Pill style={URGENCY_STYLE[u]}>{URGENCY_STYLE[u].label}</Pill></div>}
                    <div style={{ fontSize: 11.5, color: MEZZI_COLORS.muted }}>{nInterventi} interventi</div>
                  </div>
                  <ChevronRight size={18} color={MEZZI_COLORS.muted} />
                </div>
              </Card>
            );
          })}
        </div>
      </div>
      <FAB onClick={onAdd} label="Veicolo" />
    </>
  );
}

function VeicoloDetail({ vehicle, maints, params, onBack, onEdit, onDelete }) {
  const deadlines = [
    ['Assicurazione', vehicle.assicurazione], ['Revisione', vehicle.revisione], ['Bollo', vehicle.bollo],
  ].filter(([, d]) => d);
  const own = maints.filter(m => m.targa === vehicle.targa).sort((a, b) => (b.data || '').localeCompare(a.data || ''));
  const totale = own.reduce((s, m) => s + (Number(m.costo) || 0) * 1.22, 0);

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={`${vehicle.marca} ${vehicle.modello}`} subtitle={vehicle.tipo} onBack={onBack} />
      <div style={{ padding: 14 }}>
        <Card theme={MEZZI_COLORS} style={{ marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <Plate targa={vehicle.targa} />
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => onEdit(vehicle)} style={{ border: `1.5px solid ${MEZZI_COLORS.line}`, background: '#fff', borderRadius: 9, padding: '7px 12px', fontSize: 13, fontWeight: 600, color: MEZZI_COLORS.primary }}>Modifica</button>
              <button onClick={() => onDelete(vehicle)} style={{ border: `1.5px solid ${MEZZI_COLORS.line}`, background: '#fff', borderRadius: 9, padding: '7px 9px' }}><Trash2 size={15} color={MEZZI_COLORS.danger} /></button>
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, fontSize: 13.5 }}>
            <InfoRow theme={MEZZI_COLORS} icon={Gauge} label="Km attuali" value={fmtKm(vehicle.km)} />
            <InfoRow theme={MEZZI_COLORS} icon={Fuel} label="Carburante" value={vehicle.carburante || '—'} />
            <InfoRow theme={MEZZI_COLORS} icon={Palette} label="Colore" value={vehicle.colore || '—'} />
            <InfoRow theme={MEZZI_COLORS} icon={Car} label="Anno" value={vehicle.anno || '—'} />
          </div>
          {vehicle.note && (
            <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${MEZZI_COLORS.line}`, display: 'flex', gap: 7, fontSize: 13, color: MEZZI_COLORS.muted }}>
              <StickyNote size={14} style={{ marginTop: 1, flexShrink: 0 }} /> {vehicle.note}
            </div>
          )}
        </Card>

        {deadlines.length > 0 && (
          <>
            <SectionLabel theme={MEZZI_COLORS}>Scadenze</SectionLabel>
            <Card theme={MEZZI_COLORS} style={{ marginBottom: 12 }}>
              {deadlines.map(([label, d], i) => {
                const days = daysUntil(d);
                const u = urgencyOf(days, params);
                return (
                  <div key={label} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '9px 0', borderTop: i ? `1px solid ${MEZZI_COLORS.line}` : 'none' }}>
                    <span style={{ fontSize: 13.5, fontWeight: 600 }}>{label}</span>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ fontSize: 13, color: MEZZI_COLORS.muted }}>{fmtDate(d)}</span>
                      <Pill style={URGENCY_STYLE[u]}>{days < 0 ? `scaduta` : `${days}gg`}</Pill>
                    </div>
                  </div>
                );
              })}
            </Card>
          </>
        )}

        <SectionLabel theme={MEZZI_COLORS}>Storico interventi · {fmtEuro(totale)}</SectionLabel>
        {own.length === 0 && <Empty theme={MEZZI_COLORS} icon={Wrench} text="Nessun intervento registrato." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {own.map(m => (
            <Card theme={MEZZI_COLORS} key={m.id}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{m.tipo}</span>
                <span style={{ fontSize: 13, color: MEZZI_COLORS.muted }}>{fmtDate(m.data)}</span>
              </div>
              {m.descrizione && <div style={{ fontSize: 12.5, color: MEZZI_COLORS.muted, marginBottom: 4 }}>{m.descrizione}</div>}
              <div style={{ fontSize: 12, color: MEZZI_COLORS.muted }}>{m.officina || '—'} · {m.costo ? fmtEuro(m.costo * 1.22) : 'costo n.d.'}</div>
            </Card>
          ))}
        </div>
      </div>
    </>
  );
}

/* ---------- Vehicle form ---------- */
function VehicleForm({ initial, onSave, onCancel }) {
  const [f, setF] = useState(initial || { targa: '', marca: '', modello: '', tipo: 'Auto', anno: '', km: '', carburante: '', colore: '', assicurazione: '', revisione: '', bollo: '', note: '' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const valid = f.marca.trim() && f.modello.trim();

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={initial ? 'Modifica veicolo' : 'Nuovo veicolo'} onBack={onCancel} />
      <div style={{ padding: 16 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label="Marca *"><input style={inputStyle} value={f.marca} onChange={set('marca')} placeholder="Fiat" /></Field>
          <Field label="Modello *"><input style={inputStyle} value={f.modello} onChange={set('modello')} placeholder="Panda" /></Field>
        </div>
        <Field label="Targa"><input style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace" }} value={f.targa} onChange={set('targa')} placeholder="AA-000-AA" /></Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label="Tipo"><select style={inputStyle} value={f.tipo} onChange={set('tipo')}>{TIPI_VEICOLO.map(t => <option key={t}>{t}</option>)}</select></Field>
          <Field label="Anno"><input type="number" style={inputStyle} value={f.anno} onChange={set('anno')} /></Field>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label="Km attuali"><input type="number" style={inputStyle} value={f.km} onChange={set('km')} /></Field>
          <Field label="Carburante"><select style={inputStyle} value={f.carburante} onChange={set('carburante')}><option value="">—</option>{CARBURANTI.map(t => <option key={t}>{t}</option>)}</select></Field>
        </div>
        <Field label="Colore"><input style={inputStyle} value={f.colore} onChange={set('colore')} /></Field>
        <SectionLabel theme={MEZZI_COLORS}>Scadenze</SectionLabel>
        <Field label="Assicurazione"><input type="date" style={inputStyle} value={f.assicurazione} onChange={set('assicurazione')} /></Field>
        <Field label="Revisione"><input type="date" style={inputStyle} value={f.revisione} onChange={set('revisione')} /></Field>
        <Field label="Bollo"><input type="date" style={inputStyle} value={f.bollo} onChange={set('bollo')} /></Field>
        <Field label="Note"><textarea style={{ ...inputStyle, minHeight: 60, resize: 'vertical' }} value={f.note} onChange={set('note')} /></Field>

        <button
          disabled={!valid}
          onClick={() => onSave({ ...f, id: f.id || uid(), anno: f.anno ? Number(f.anno) : null, km: f.km ? Number(f.km) : null })}
          style={{ width: '100%', background: valid ? MEZZI_COLORS.primary : '#B7C0C2', color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 6 }}
        >
          Salva veicolo
        </button>
      </div>
    </>
  );
}

/* ---------- Maintenance ---------- */
function ManutenzioniScreen({ maints, vehicles, onOpen, onAdd, onMenu, onHome }) {
  const sorted = [...maints].sort((a, b) => (b.data || '').localeCompare(a.data || ''));
  const vOf = (targa) => vehicles.find(v => v.targa === targa);
  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Manutenzioni" subtitle={`${maints.length} interventi registrati`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        {sorted.length === 0 && <Empty theme={MEZZI_COLORS} icon={Wrench} text="Nessun intervento ancora. Aggiungine uno." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {sorted.map(m => (
            <Card theme={MEZZI_COLORS} key={m.id} onClick={() => onOpen(m)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 4 }}>{m.tipo}</div>
                  <div style={{ marginBottom: 6 }}><Plate targa={m.targa} size="sm" /></div>
                  <div style={{ fontSize: 12, color: MEZZI_COLORS.muted }}>{vOf(m.targa) ? `${vOf(m.targa).marca} ${vOf(m.targa).modello} · ` : ''}{fmtDate(m.data)}</div>
                </div>
                <div style={{ textAlign: 'right', flexShrink: 0, marginLeft: 8 }}>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>{m.costo ? fmtEuro(m.costo * 1.22) : '—'}</div>
                  <div style={{ fontSize: 11, color: MEZZI_COLORS.muted, marginTop: 4 }}>{m.stato}</div>
                </div>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <FAB onClick={onAdd} label="Intervento" />
    </>
  );
}

function MaintForm({ initial, vehicles, params, onSave, onCancel, onDelete }) {
  const [f, setF] = useState(initial || { targa: vehicles[0]?.targa || '', data: todayISO(), km: '', tipo: TIPI_MANUTENZIONE[0], descrizione: '', officina: '', costo: '', stato: 'Programmato' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const costo = Number(f.costo) || 0;
  const iva = costo * (params.ivaRate / 100);
  const totale = costo + iva;

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={initial ? 'Modifica intervento' : 'Nuovo intervento'} onBack={onCancel} />
      <div style={{ padding: 16 }}>
        <Field label="Veicolo *">
          <select style={inputStyle} value={f.targa} onChange={set('targa')}>
            {vehicles.map(v => <option key={v.id} value={v.targa}>{v.marca} {v.modello}{v.targa ? ' · ' + v.targa : ' · senza targa'}</option>)}
          </select>
        </Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label="Data"><input type="date" style={inputStyle} value={f.data} onChange={set('data')} /></Field>
          <Field label="Km al momento"><input type="number" style={inputStyle} value={f.km} onChange={set('km')} /></Field>
        </div>
        <Field label="Tipo manutenzione">
          <select style={inputStyle} value={f.tipo} onChange={set('tipo')}>{TIPI_MANUTENZIONE.map(t => <option key={t}>{t}</option>)}</select>
        </Field>
        <Field label="Descrizione"><input style={inputStyle} value={f.descrizione} onChange={set('descrizione')} placeholder="Dettagli dell'intervento" /></Field>
        <Field label="Officina / Fornitore"><input style={inputStyle} value={f.officina} onChange={set('officina')} /></Field>
        <Field label="Costo netto (€)"><input type="number" step="0.01" style={inputStyle} value={f.costo} onChange={set('costo')} placeholder="0.00" /></Field>
        {costo > 0 && (
          <div style={{ background: '#EFF3EE', borderRadius: 10, padding: 12, marginBottom: 14, fontSize: 13 }}>
            <Row label={`IVA (${params.ivaRate}%)`} value={fmtEuro(iva)} />
            <Row label="Totale" value={fmtEuro(totale)} bold />
          </div>
        )}
        <Field label="Stato"><select style={inputStyle} value={f.stato} onChange={set('stato')}>{STATI_MANUTENZIONE.map(t => <option key={t}>{t}</option>)}</select></Field>

        <button
          onClick={() => onSave({ ...f, id: f.id || uid(), km: f.km ? Number(f.km) : null, costo: f.costo ? Number(f.costo) : null })}
          style={{ width: '100%', background: MEZZI_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}
        >
          Salva intervento
        </button>
        {initial && (
          <button onClick={() => onDelete(initial)} style={{ width: '100%', background: 'none', border: 'none', color: MEZZI_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
            Elimina intervento
          </button>
        )}
      </div>
    </>
  );
}

function Row({ label, value, bold }) {
  return (
    <div style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0', fontWeight: bold ? 700 : 500, fontSize: bold ? 14 : 13 }}>
      <span style={{ color: bold ? MEZZI_COLORS.ink : MEZZI_COLORS.muted }}>{label}</span><span>{value}</span>
    </div>
  );
}

/* ---------- Scadenze ---------- */
function ScadenzeScreen({ vehicles, params, onMenu, onHome }) {
  const rows = useMemo(() => {
    const out = [];
    vehicles.forEach(v => {
      [['Assicurazione', v.assicurazione], ['Revisione', v.revisione], ['Bollo', v.bollo]].forEach(([label, d]) => {
        if (!d) return;
        const days = daysUntil(d);
        out.push({ key: v.id + label, v, label, d, days, u: urgencyOf(days, params) });
      });
    });
    return out.sort((a, b) => a.days - b.days);
  }, [vehicles, params]);

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Scadenze" subtitle={`${rows.length} scadenze in archivio`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        {rows.length === 0 && <Empty theme={MEZZI_COLORS} icon={CalendarClock} text="Nessuna scadenza registrata." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {rows.map(r => (
            <Card theme={MEZZI_COLORS} key={r.key}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 3 }}>{r.label}</div>
                  <div style={{ fontSize: 12.5, color: MEZZI_COLORS.muted, marginBottom: 6 }}>{r.v.marca} {r.v.modello}</div>
                  <Plate targa={r.v.targa} size="sm" />
                </div>
                <div style={{ textAlign: 'right', flexShrink: 0 }}>
                  <Pill style={URGENCY_STYLE[r.u]}>{r.days < 0 ? 'Scaduta' : `${r.days} giorni`}</Pill>
                  <div style={{ fontSize: 12, color: MEZZI_COLORS.muted, marginTop: 6 }}>{fmtDate(r.d)}</div>
                </div>
              </div>
            </Card>
          ))}
        </div>
      </div>
    </>
  );
}

/* ---------- Dashboard ---------- */
function DashboardScreen({ vehicles, maints, params, onMenu, onHome }) {
  const withCost = maints.filter(m => m.costo);
  const totale = withCost.reduce((s, m) => s + m.costo * (1 + params.ivaRate / 100), 0);
  const perVeicolo = useMemo(() => {
    const map = {};
    maints.forEach(m => {
      if (!m.targa) return;
      map[m.targa] = map[m.targa] || { n: 0, costo: 0 };
      map[m.targa].n += 1;
      map[m.targa].costo += (m.costo || 0) * (1 + params.ivaRate / 100);
    });
    return Object.entries(map).map(([targa, v]) => ({ targa, ...v, veh: vehicles.find(x => x.targa === targa) }))
      .sort((a, b) => b.costo - a.costo);
  }, [maints, vehicles, params]);
  const perTipo = useMemo(() => {
    const map = {};
    maints.forEach(m => { map[m.tipo] = (map[m.tipo] || 0) + 1; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [maints]);
  const maxTipo = perTipo.length ? perTipo[0][1] : 1;
  const urgenti = vehicles.filter(v => [v.assicurazione, v.revisione, v.bollo].some(d => urgencyOf(daysUntil(d), params) === 'urgent')).length;

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Dashboard" subtitle="Riepilogo costi e flotta" onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 6 }}>
          <StatCard theme={MEZZI_COLORS} label="Spesa totale" value={fmtEuro(totale)} accent={MEZZI_COLORS.primary} />
          <StatCard theme={MEZZI_COLORS} label="Interventi" value={maints.length} accent={MEZZI_COLORS.primary} />
          <StatCard theme={MEZZI_COLORS} label="Veicoli in flotta" value={vehicles.length} accent={MEZZI_COLORS.primary} />
          <StatCard theme={MEZZI_COLORS} label="Scadenze urgenti" value={urgenti} accent={urgenti ? MEZZI_COLORS.danger : MEZZI_COLORS.ok} />
        </div>

        <SectionLabel theme={MEZZI_COLORS}>Costi per veicolo</SectionLabel>
        {perVeicolo.length === 0 && <Empty theme={MEZZI_COLORS} icon={BarChart3} text="Nessun costo registrato ancora." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 4 }}>
          {perVeicolo.map(p => (
            <Card theme={MEZZI_COLORS} key={p.targa}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                <span style={{ fontWeight: 600, fontSize: 13.5 }}>{p.veh ? `${p.veh.marca} ${p.veh.modello}` : p.targa}</span>
                <span style={{ fontWeight: 700, fontSize: 13.5 }}>{fmtEuro(p.costo)}</span>
              </div>
              <div style={{ fontSize: 11.5, color: MEZZI_COLORS.muted }}>{p.n} interventi</div>
            </Card>
          ))}
        </div>

        <SectionLabel theme={MEZZI_COLORS}>Interventi per tipo</SectionLabel>
        <Card theme={MEZZI_COLORS}>
          {perTipo.map(([tipo, n], i) => (
            <div key={tipo} style={{ marginTop: i ? 10 : 0 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12.5, marginBottom: 4 }}>
                <span>{tipo}</span><span style={{ color: MEZZI_COLORS.muted }}>{n}</span>
              </div>
              <div style={{ height: 7, background: '#EEEAE0', borderRadius: 999 }}>
                <div style={{ height: '100%', width: `${(n / maxTipo) * 100}%`, background: MEZZI_COLORS.amber, borderRadius: 999 }} />
              </div>
            </div>
          ))}
        </Card>
      </div>
    </>
  );
}


/* ---------- Root ---------- */
function MezziModule({ onHome }) {
  const [ready, setReady] = useState(false);
  const [vehicles, setVehicles] = useState([]);
  const [maints, setMaints] = useState([]);
  const [params] = useState(DEFAULT_PARAMS);
  const [tab, setTab] = useState('veicoli');
  const [view, setView] = useState({ name: 'list' });
  const [toast, setToast] = useState('');
  const [showMenu, setShowMenu] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        let v, m;
        try { v = JSON.parse((await storage.get('vehicles')).value); } catch { v = null; }
        try { m = JSON.parse((await storage.get('maintenances')).value); } catch { m = null; }
        setVehicles(v && v.length ? v : SEED_VEHICLES);
        setMaints(m && m.length ? m : SEED_MAINTS);
      } catch (e) {
        setVehicles(SEED_VEHICLES); setMaints(SEED_MAINTS);
      }
      setReady(true);
    })();
  }, []);

  useEffect(() => { if (ready) storage.set('vehicles', JSON.stringify(vehicles)).catch(() => {}); }, [vehicles, ready]);
  useEffect(() => { if (ready) storage.set('maintenances', JSON.stringify(maints)).catch(() => {}); }, [maints, ready]);

  const exportToExcel = () => {
    const dt = (iso) => (iso ? new Date(iso + 'T00:00:00') : '');

    // Foglio "Anagrafica Veicoli" - stesse intestazioni e ordine colonne del file originale
    const anagraficaRows = vehicles.map(v => ({
      'Targa': v.targa || '', 'Marca': v.marca, 'Modello': v.modello, 'Tipo': v.tipo,
      'Anno Immatricolazione': v.anno || '', 'Data Acquisto': '', 'Km Attuali': v.km || '',
      'Carburante': v.carburante || '', 'Colore': v.colore || '', 'Telaio': '',
      'Assicurazione Scad.': dt(v.assicurazione), 'Revisione Scad.': dt(v.revisione), 'Bollo Scad.': dt(v.bollo),
      'Note': v.note || '',
    }));
    const wsAna = XLSX.utils.json_to_sheet(anagraficaRows, { cellDates: true });
    wsAna['!cols'] = [11, 9, 12, 9, 10, 12, 10, 11, 9, 18, 13, 13, 11, 26].map(w => ({ wch: w }));

    // Foglio "Registro Manutenzione" - IVA e Totale ricalcolati con l'aliquota corrente
    const registroRows = maints.map(m => {
      const v = vehicles.find(x => x.targa === m.targa);
      const costo = m.costo != null ? Number(m.costo) : '';
      const iva = costo !== '' ? Math.round(costo * (params.ivaRate / 100) * 100) / 100 : '';
      const totale = costo !== '' ? Math.round((costo + iva) * 100) / 100 : '';
      return {
        'ID': m.id, 'Targa Veicolo': m.targa || '', 'Marca / Modello': v ? `${v.marca} ${v.modello}` : '',
        'Tipo Veicolo': v ? v.tipo : '', 'Data Intervento': dt(m.data), 'Km al Momento': m.km || '',
        'Tipo Manutenzione': m.tipo, 'Descrizione Intervento': m.descrizione || '', 'Officina / Fornitore': m.officina || '',
        'Costo (€)': costo, 'IVA (€)': iva, 'Totale (€)': totale, 'Stato': m.stato,
      };
    });
    const wsReg = XLSX.utils.json_to_sheet(registroRows, { cellDates: true });
    wsReg['!cols'] = [10, 13, 18, 11, 13, 12, 16, 28, 16, 10, 9, 10, 16].map(w => ({ wch: w }));

    // Foglio "Parametri" - le stesse soglie/aliquote usate nell'app
    const parametriRows = [
      { 'Parametro': 'Giorni preavviso scadenza URGENTE', 'Valore': params.urgentDays, 'Unità': 'giorni' },
      { 'Parametro': 'Giorni preavviso scadenza MEDIA', 'Valore': params.mediumDays, 'Unità': 'giorni' },
      { 'Parametro': 'Aliquota IVA applicata', 'Valore': params.ivaRate, 'Unità': '%' },
    ];
    const wsPar = XLSX.utils.json_to_sheet(parametriRows);
    wsPar['!cols'] = [30, 10, 10].map(w => ({ wch: w }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsAna, 'Anagrafica Veicoli');
    XLSX.utils.book_append_sheet(wb, wsReg, 'Registro Manutenzione');
    XLSX.utils.book_append_sheet(wb, wsPar, 'Parametri');

    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `registro_manutenzione_veicoli_${stamp}.xlsx`);
    setShowMenu(false);
    flash('File Excel scaricato');
  };

  useEffect(() => { setView({ name: 'list' }); }, [tab]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 1800); };

  const saveVehicle = (v) => {
    setVehicles(prev => prev.some(x => x.id === v.id) ? prev.map(x => x.id === v.id ? v : x) : [...prev, v]);
    flash('Veicolo salvato');
    setView({ name: 'list' });
  };
  const deleteVehicle = (v) => {
    setVehicles(prev => prev.filter(x => x.id !== v.id));
    flash('Veicolo eliminato');
    setView({ name: 'list' });
  };
  const saveMaint = (m) => {
    setMaints(prev => prev.some(x => x.id === m.id) ? prev.map(x => x.id === m.id ? m : x) : [...prev, m]);
    flash('Intervento salvato');
    setView({ name: 'list' });
  };
  const deleteMaint = (m) => {
    setMaints(prev => prev.filter(x => x.id !== m.id));
    flash('Intervento eliminato');
    setView({ name: 'list' });
  };

  if (!ready) {
    return <div style={{ minHeight: '100vh', background: MEZZI_COLORS.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', color: MEZZI_COLORS.muted }}>Caricamento…</div>;
  }

  let content;
  if (tab === 'veicoli') {
    if (view.name === 'detail') content = <VeicoloDetail vehicle={vehicles.find(v => v.id === view.id)} maints={maints} params={params} onBack={() => setView({ name: 'list' })} onEdit={(v) => setView({ name: 'edit', v })} onDelete={deleteVehicle} />;
    else if (view.name === 'add') content = <VehicleForm onSave={saveVehicle} onCancel={() => setView({ name: 'list' })} />;
    else if (view.name === 'edit') content = <VehicleForm initial={view.v} onSave={saveVehicle} onCancel={() => setView({ name: 'detail', id: view.v.id })} />;
    else content = <VeicoliScreen vehicles={vehicles} maints={maints} params={params} onOpen={(v) => setView({ name: 'detail', id: v.id })} onAdd={() => setView({ name: 'add' })} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  } else if (tab === 'manutenzioni') {
    if (view.name === 'add') content = <MaintForm vehicles={vehicles} params={params} onSave={saveMaint} onCancel={() => setView({ name: 'list' })} />;
    else if (view.name === 'edit') content = <MaintForm initial={view.m} vehicles={vehicles} params={params} onSave={saveMaint} onCancel={() => setView({ name: 'list' })} onDelete={deleteMaint} />;
    else content = <ManutenzioniScreen maints={maints} vehicles={vehicles} onOpen={(m) => setView({ name: 'edit', m })} onAdd={() => setView({ name: 'add' })} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  } else if (tab === 'scadenze') {
    content = <ScadenzeScreen vehicles={vehicles} params={params} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  } else {
    content = <DashboardScreen vehicles={vehicles} maints={maints} params={params} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  }

  return (
    <div style={{ minHeight: '100vh', background: MEZZI_COLORS.bg, fontFamily: 'Inter, sans-serif', color: MEZZI_COLORS.ink, maxWidth: 480, margin: '0 auto', position: 'relative' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: ${MEZZI_COLORS.primary} !important; }
      `}</style>
      <div style={{ paddingBottom: 78 }}>{content}</div>
      {view.name === 'list' && <BottomNav theme={MEZZI_COLORS} tab={tab} setTab={setTab} items={MEZZI_NAV_ITEMS} />}
      {showMenu && <MenuSheet theme={MEZZI_COLORS} onClose={() => setShowMenu(false)} onExport={exportToExcel} exportSub="Scarica anagrafica e registro in .xlsx" />}
      {toast && (
        <div style={{ position: 'fixed', bottom: view.name === 'list' ? 92 : 20, left: '50%', transform: 'translateX(-50%)', background: MEZZI_COLORS.primaryDeep, color: '#fff', padding: '10px 18px', borderRadius: 999, fontSize: 13, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 7, zIndex: 30, maxWidth: 400 }}>
          <Check size={15} /> {toast}
        </div>
      )}
    </div>
  );
}

/* ---------- design tokens ---------- */
const CARROZZINE_COLORS = {
  bg: '#F4F1E9',
  surface: '#FFFFFF',
  ink: '#212823',
  muted: '#6E7972',
  line: '#E2E0D2',
  primary: '#33594E',
  primaryDeep: '#1E3830',
  amber: '#C6790C',
  danger: '#B23A2E',
  ok: '#3C7A5C',
  info: '#2C5C82',
};

const NUCLEO_COLORS = {
  'Rosso': { bg: '#F9D9D6', fg: '#8A2A20', dot: '#D8392C' },
  'Grigio': { bg: '#DEDEDE', fg: '#3A3A3A', dot: '#6B6B6B' },
  'Giallo': { bg: '#FBEBBE', fg: '#7A5A00', dot: '#EFB81C' },
  'Verde': { bg: '#CFE9D1', fg: '#215A2A', dot: '#3EA24C' },
  'Bianco': { bg: '#F1F0EC', fg: '#4A4A46', dot: '#C7C4B8' },
  'Blu': { bg: '#D2DEEF', fg: '#1E3D66', dot: '#3363A8' },
  'IKEA': { bg: '#E3DCF1', fg: '#4C3B75', dot: '#8E7CC3' },
  'Terra': { bg: '#F6E7C4', fg: '#7A5A18', dot: '#C99B3E' },
  'Palestra FKT': { bg: '#D9F0E1', fg: '#1F6B45', dot: '#4FAE79' },
  'Primo Piano': { bg: '#E4DEF0', fg: '#4C3B75', dot: '#9A87C9' },
};
const NUCLEI = Object.keys(NUCLEO_COLORS);

const STATO_STYLE = {
  'Disponibile': { bg: '#DCEEE3', fg: '#1F6B45' },
  'In uso': { bg: '#DCE8F0', fg: '#1E4D73' },
  'Manutenzione': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Fuori uso': { bg: '#F7DCD9', fg: '#A3352A' },
};
const STATI_CARROZZINA = Object.keys(STATO_STYLE);

const COND_STYLE = {
  'OK': { bg: '#DCEEE3', fg: '#1F6B45' },
  'Sostituito': { bg: '#DCE8F0', fg: '#1E4D73' },
  'Da sistemare': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Mancante': { bg: '#F7DCD9', fg: '#A3352A' },
};
const CONDIZIONI = ['OK', 'Da sistemare', 'Sostituito', 'Mancante'];
const TIPOLOGIE = ['Standard', 'Reclinabile', 'Bascula', 'Transito'];
const FORNITORI = ['Melo', 'Ats', 'Ospite'];

const COMPONENTI = [
  ['gomme', 'Gomme'], ['mozzi', 'Mozzi'], ['freni', 'Freni'], ['pedalini', 'Pedalini'],
  ['braccioli', 'Braccioli'], ['portaborraccia', 'Portaborraccia'], ['tavolino', 'Tavolino'],
  ['manopole', 'Manopole'], ['seduta', 'Seduta'], ['poggiatesta', 'Poggiatesta'], ['pulizia', 'Pulizia'],
];
const CARROZZINE_NAV_ITEMS = [
  ['carrozzine', Armchair, 'Carrozzine'],
  ['controlli', ClipboardList, 'Controlli'],
  ['nuclei', Layers, 'Nuclei'],
  ['riepilogo', BarChart3, 'Riepilogo'],
];

/* ---------- helpers ---------- */
const needsAttention = (w) => COMPONENTI.some(([k]) => w.c[k] === 'Da sistemare' || w.c[k] === 'Mancante');
const attentionCount = (w) => COMPONENTI.filter(([k]) => w.c[k] === 'Da sistemare' || w.c[k] === 'Mancante').length;
const labelOf = (w) => [w.marca, w.modello].filter(Boolean).join(' ') || 'Carrozzina senza marca';

/* ---------- small UI atoms ---------- */
function NucleoTag({ nucleo }) {
  if (!nucleo) return <span style={{ fontSize: 12, color: CARROZZINE_COLORS.muted }}>Nucleo n.d.</span>;
  const s = NUCLEO_COLORS[nucleo] || { bg: '#EEE', fg: CARROZZINE_COLORS.muted, dot: '#999' };
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: s.bg, color: s.fg, fontSize: 11, fontWeight: 700, padding: '3px 9px 3px 7px', borderRadius: 999 }}>
      <span style={{ width: 7, height: 7, borderRadius: 999, background: s.dot, flexShrink: 0 }} />
      {nucleo}
    </span>
  );
}

const selectStyle = (s) => ({
  width: '100%', boxSizing: 'border-box', padding: '8px 10px', borderRadius: 9, border: 'none',
  fontSize: 12.5, fontWeight: 700, fontFamily: 'Inter, sans-serif', background: s.bg, color: s.fg,
  outline: 'none', appearance: 'none', WebkitAppearance: 'none',
});

/* ---------- Carrozzine ---------- */
function CarrozzineScreen({ items, onOpen, onMenu, filterNucleo, setFilterNucleo, onHome }) {
  const [q, setQ] = useState('');
  const filtered = items.filter(w => {
    if (filterNucleo && w.nucleo !== filterNucleo) return false;
    const hay = `${w.marca} ${w.modello} ${w.seriale} ${w.ospite} ${w.nucleo}`.toLowerCase();
    return hay.includes(q.toLowerCase());
  });

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Carrozzine" subtitle={`${items.length} in archivio`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ position: 'relative', marginBottom: 10 }}>
          <Search size={16} style={{ position: 'absolute', left: 12, top: 12, color: CARROZZINE_COLORS.muted }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Cerca marca, modello, ospite…"
            style={{ width: '100%', boxSizing: 'border-box', padding: '11px 12px 11px 34px', borderRadius: 10, border: `1.5px solid ${CARROZZINE_COLORS.line}`, fontSize: 15, background: '#FCFBF7', outline: 'none' }} />
        </div>
        {filterNucleo && (
          <div style={{ marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, color: CARROZZINE_COLORS.muted }}>Filtro:</span>
            <NucleoTag nucleo={filterNucleo} />
            <button onClick={() => setFilterNucleo(null)} style={{ background: 'none', border: 'none', color: CARROZZINE_COLORS.muted, fontSize: 12, textDecoration: 'underline' }}>rimuovi</button>
          </div>
        )}
        {filtered.length === 0 && <Empty theme={CARROZZINE_COLORS} icon={Armchair} text="Nessuna carrozzina trovata." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {filtered.map(w => {
            const att = attentionCount(w);
            return (
              <Card theme={CARROZZINE_COLORS} key={w.id} onClick={() => onOpen(w)}>
                <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 6 }}>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: CARROZZINE_COLORS.muted }}>#{w.id}</span>
                      <span style={{ fontWeight: 700, fontSize: 14.5 }}>{labelOf(w)}</span>
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: w.ospite ? 6 : 0 }}>
                      <NucleoTag nucleo={w.nucleo} />
                      {w.stato && <Pill style={STATO_STYLE[w.stato]}>{w.stato}</Pill>}
                      {att > 0 && <Pill style={COND_STYLE['Mancante']}><AlertTriangle size={10} style={{ display: 'inline', marginRight: 3, marginBottom: -1 }} />{att}</Pill>}
                    </div>
                    {w.ospite && <div style={{ fontSize: 12, color: CARROZZINE_COLORS.muted, display: 'flex', alignItems: 'center', gap: 4 }}><User size={11} />{w.ospite}</div>}
                  </div>
                  <ChevronRight size={18} color={CARROZZINE_COLORS.muted} style={{ marginTop: 4, flexShrink: 0 }} />
                </div>
              </Card>
            );
          })}
        </div>
      </div>
    </>
  );
}

function CarrozzinaDetail({ w, onBack, onUpdate }) {
  const set = (patch) => onUpdate({ ...w, ...patch });
  const setC = (key, val) => onUpdate({ ...w, c: { ...w.c, [key]: val } });

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title={labelOf(w)} subtitle={`ID ${w.id}${w.seriale ? ' · ' + w.seriale : ''}`} onBack={onBack} />
      <div style={{ padding: 14 }}>
        <Card theme={CARROZZINE_COLORS} style={{ marginBottom: 12 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
            <InfoRow theme={CARROZZINE_COLORS} icon={Calendar} label="Data" value={fmtDate(w.data)} />
            <InfoRow theme={CARROZZINE_COLORS} icon={Hash} label="Tipologia" value={w.tipologia} />
            <InfoRow theme={CARROZZINE_COLORS} icon={MapPin} label="Stanza" value={w.stanza} />
            <InfoRow theme={CARROZZINE_COLORS} icon={User} label="Ospite" value={w.ospite} />
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, paddingTop: 10, borderTop: `1px solid ${CARROZZINE_COLORS.line}` }}>
            <div>
              <div style={{ fontSize: 10, color: CARROZZINE_COLORS.muted, textTransform: 'uppercase', marginBottom: 4 }}>Nucleo</div>
              <select value={w.nucleo || ''} onChange={e => set({ nucleo: e.target.value })} style={selectStyle(NUCLEO_COLORS[w.nucleo] || { bg: '#EEE', fg: CARROZZINE_COLORS.ink })}>
                <option value="">—</option>
                {NUCLEI.map(n => <option key={n} value={n}>{n}</option>)}
              </select>
            </div>
            <div>
              <div style={{ fontSize: 10, color: CARROZZINE_COLORS.muted, textTransform: 'uppercase', marginBottom: 4 }}>Stato</div>
              <select value={w.stato || ''} onChange={e => set({ stato: e.target.value })} style={selectStyle(STATO_STYLE[w.stato] || { bg: '#EEE', fg: CARROZZINE_COLORS.ink })}>
                <option value="">—</option>
                {STATI_CARROZZINA.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
        </Card>

        <SectionLabel theme={CARROZZINE_COLORS}>Checklist componenti</SectionLabel>
        <Card theme={CARROZZINE_COLORS} style={{ marginBottom: 12 }}>
          {COMPONENTI.map(([key, label], i) => {
            const val = w.c[key];
            const isFreeText = val && !CONDIZIONI.includes(val);
            return (
              <div key={key} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, padding: '9px 0', borderTop: i ? `1px solid ${CARROZZINE_COLORS.line}` : 'none' }}>
                <span style={{ fontSize: 13.5, fontWeight: 600, flexShrink: 0 }}>{label}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                  {isFreeText && <span style={{ fontSize: 11, color: CARROZZINE_COLORS.muted, fontStyle: 'italic', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 110 }}>{val}</span>}
                  <select
                    value={CONDIZIONI.includes(val) ? val : ''}
                    onChange={e => setC(key, e.target.value)}
                    style={{ ...selectStyle(COND_STYLE[val] || { bg: '#F1EFE6', fg: CARROZZINE_COLORS.muted }), width: 118 }}
                  >
                    <option value="">{isFreeText ? 'nota storica' : '—'}</option>
                    {CONDIZIONI.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>
            );
          })}
        </Card>

        <SectionLabel theme={CARROZZINE_COLORS}>Note</SectionLabel>
        <Card theme={CARROZZINE_COLORS}>
          <textarea value={w.note || ''} onChange={e => set({ note: e.target.value })} placeholder="Aggiungi una nota…"
            style={{ width: '100%', boxSizing: 'border-box', border: 'none', outline: 'none', resize: 'vertical', minHeight: 70, fontSize: 13.5, fontFamily: 'Inter, sans-serif', background: 'transparent' }} />
        </Card>
      </div>
    </>
  );
}

/* ---------- Controlli ---------- */
function ControlliScreen({ items, onOpen, onMenu, onHome }) {
  const rows = items.filter(needsAttention);
  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Controlli" subtitle={`${rows.length} carrozzine da verificare`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        {rows.length === 0 && <Empty theme={CARROZZINE_COLORS} icon={Check} text="Nessuna carrozzina segnalata. Tutto a posto." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {rows.map(w => (
            <Card theme={CARROZZINE_COLORS} key={w.id} onClick={() => onOpen(w)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 8 }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 4 }}>{labelOf(w)}</div>
                  <NucleoTag nucleo={w.nucleo} />
                </div>
                <ChevronRight size={18} color={CARROZZINE_COLORS.muted} />
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {COMPONENTI.filter(([k]) => w.c[k] === 'Da sistemare' || w.c[k] === 'Mancante').map(([k, label]) => (
                  <Pill key={k} style={COND_STYLE[w.c[k]]}>{label}: {w.c[k]}</Pill>
                ))}
              </div>
            </Card>
          ))}
        </div>
      </div>
    </>
  );
}

/* ---------- Nuclei ---------- */
function NucleiScreen({ items, onFilter, onMenu, onHome }) {
  const groups = useMemo(() => {
    const map = {};
    items.forEach(w => {
      const key = w.nucleo || '—';
      map[key] = map[key] || { tot: 0, disp: 0, uso: 0, man: 0, fuori: 0, att: 0 };
      map[key].tot++;
      if (w.stato === 'Disponibile') map[key].disp++;
      if (w.stato === 'In uso') map[key].uso++;
      if (w.stato === 'Manutenzione') map[key].man++;
      if (w.stato === 'Fuori uso') map[key].fuori++;
      if (needsAttention(w)) map[key].att++;
    });
    return Object.entries(map).sort((a, b) => b[1].tot - a[1].tot);
  }, [items]);

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Nuclei" subtitle={`${groups.length} nuclei attivi`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {groups.map(([nucleo, g]) => (
            <Card theme={CARROZZINE_COLORS} key={nucleo} onClick={() => onFilter(nucleo)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <NucleoTag nucleo={nucleo === '—' ? null : nucleo} />
                <span style={{ fontWeight: 800, fontFamily: "'Archivo', sans-serif", fontSize: 16 }}>{g.tot}</span>
              </div>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {g.disp > 0 && <Pill style={STATO_STYLE['Disponibile']}>{g.disp} disponibili</Pill>}
                {g.uso > 0 && <Pill style={STATO_STYLE['In uso']}>{g.uso} in uso</Pill>}
                {g.man > 0 && <Pill style={STATO_STYLE['Manutenzione']}>{g.man} manutenzione</Pill>}
                {g.fuori > 0 && <Pill style={STATO_STYLE['Fuori uso']}>{g.fuori} fuori uso</Pill>}
                {g.att > 0 && <Pill style={COND_STYLE['Mancante']}>{g.att} da controllare</Pill>}
              </div>
            </Card>
          ))}
        </div>
      </div>
    </>
  );
}

/* ---------- Riepilogo ---------- */
function RiepilogoScreen({ items, onMenu, onHome }) {
  const tot = items.length;
  const attCount = items.filter(needsAttention).length;
  const perStato = STATI_CARROZZINA.map(s => [s, items.filter(w => w.stato === s).length]);
  const senzaStato = items.filter(w => !w.stato).length;
  const perMarca = useMemo(() => {
    const map = {};
    items.forEach(w => { if (w.marca) map[w.marca] = (map[w.marca] || 0) + 1; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, 8);
  }, [items]);
  const maxMarca = perMarca.length ? perMarca[0][1] : 1;

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Riepilogo" subtitle="Vista d'insieme della flotta" onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 6 }}>
          <StatCard theme={CARROZZINE_COLORS} label="Carrozzine totali" value={tot} accent={CARROZZINE_COLORS.primary} />
          <StatCard theme={CARROZZINE_COLORS} label="Da controllare" value={attCount} accent={attCount ? CARROZZINE_COLORS.danger : CARROZZINE_COLORS.ok} />
        </div>

        <SectionLabel theme={CARROZZINE_COLORS}>Per stato</SectionLabel>
        <Card theme={CARROZZINE_COLORS} style={{ marginBottom: 4 }}>
          {perStato.map(([s, n], i) => (
            <div key={s} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderTop: i ? `1px solid ${CARROZZINE_COLORS.line}` : 'none' }}>
              <Pill style={STATO_STYLE[s]}>{s}</Pill>
              <span style={{ fontWeight: 700, fontSize: 14 }}>{n}</span>
            </div>
          ))}
          {senzaStato > 0 && (
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderTop: `1px solid ${CARROZZINE_COLORS.line}` }}>
              <span style={{ fontSize: 12.5, color: CARROZZINE_COLORS.muted }}>Senza stato assegnato</span>
              <span style={{ fontWeight: 700, fontSize: 14, color: CARROZZINE_COLORS.muted }}>{senzaStato}</span>
            </div>
          )}
        </Card>

        <SectionLabel theme={CARROZZINE_COLORS}>Marche più diffuse</SectionLabel>
        <Card theme={CARROZZINE_COLORS}>
          {perMarca.map(([m, n], i) => (
            <div key={m} style={{ marginTop: i ? 10 : 0 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12.5, marginBottom: 4 }}>
                <span>{m}</span><span style={{ color: CARROZZINE_COLORS.muted }}>{n}</span>
              </div>
              <div style={{ height: 7, background: '#EEEAE0', borderRadius: 999 }}>
                <div style={{ height: '100%', width: `${(n / maxMarca) * 100}%`, background: CARROZZINE_COLORS.amber, borderRadius: 999 }} />
              </div>
            </div>
          ))}
        </Card>
      </div>
    </>
  );
}

/* ---------- Menu / Nav ---------- */
const SEED = [
{id:1,data:"2026-03-31",marca:"Nuova Blandino",modello:"GR101",seriale:"SN 4309",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"-",manopole:"",seduta:"",poggiatesta:"-",pulizia:""},note:""},
{id:2,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Palestra FKT",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:3,data:"",marca:"OSD",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Terreni Costantina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:4,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Bertagnon Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:5,data:"",marca:"Nuova Blandino",modello:"",seriale:"SN 0860",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Comazzi Pietro",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:6,data:"",marca:"Vassilli",modello:"17.60XL50",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Rivanina Nunzia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:7,data:"",marca:"OSD",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Giordana Nilde",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:8,data:"",marca:"Vermeiren",modello:"R708TII",seriale:"N0010783",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Aliverti Giancarlo",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:9,data:"",marca:"Vermeiren",modello:"Inovys II",seriale:"5117450",tipologia:"",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Locarno Giovanni",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:10,data:"",marca:"Emineo",modello:"",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Lamoglie Luigi",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:11,data:"",marca:"",modello:"",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Perin Bruno",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:12,data:"",marca:"Surace",modello:"200 Classic",seriale:"",tipologia:"Transito",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Sinico Elda",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:13,data:"",marca:"Nuova Blandino",modello:"GR104F",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Pozzi Paola",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:14,data:"",marca:"Emineo",modello:"",seriale:"",tipologia:"Bascula",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Castano Luigia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:15,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Ziggiotto Eleonora",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"si",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:16,data:"",marca:"Sunrise Medical",modello:"Breezy 90",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Cirigliano Angela",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:17,data:"",marca:"Surace",modello:"Euro 600",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Bertani Teresa",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:18,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Galeazzi Bruna",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:19,data:"",marca:"Vermeiren",modello:"V500",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Napoli Michele",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:20,data:"",marca:"Vermeiren",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Guzzetti Bruno",stato:"In uso",c:{gomme:"ok",mozzi:"",freni:"",pedalini:"ok",braccioli:"",portaborraccia:"-",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:21,data:"",marca:"Sunrise Medical",modello:"Breezy 90",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Mariuz Evelina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:22,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Botterio Lanfranco",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:23,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Botterio Rosanna",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:24,data:"",marca:"Wimed",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"",ospite:"Postizzi Luciano",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:25,data:"",marca:"Nuova Blandino",modello:"GR 104F",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Creazzo Piera",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:26,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Pinciroli Vincenzina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:27,data:"",marca:"Vermeiren",modello:"R708 TII",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Croce Enrica",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:28,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Monguzzi Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:29,data:"",marca:"Moretti",modello:"Ardea One",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Schiesaro Simonetta",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:30,data:"",marca:"Sunrise Medical",modello:"Breezy 90",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Longhi Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:31,data:"",marca:"Thuasne",modello:"Classic EI+",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Blu",stanza:"",ospite:"Sartori Anna",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:32,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Bianco",stanza:"",ospite:"Costa Annamaria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:33,data:"",marca:"Suspa",modello:"Varilock",seriale:"",tipologia:"Bascula",fornitore:"Melo",nucleo:"Bianco",stanza:"",ospite:"Samperisi Luigia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:34,data:"",marca:"Nuova Blandino",modello:"GR 104F",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Bianco",stanza:"",ospite:"Baruch Miriam",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:35,data:"",marca:"Vermeiren",modello:"R750",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Bianco",stanza:"",ospite:"Riganti Carla",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:36,data:"",marca:"Thuasne",modello:"Classic EI+",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Bianco",stanza:"",ospite:"Baglioni Irma",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:37,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Primo Piano",stanza:"",ospite:"Scorta 4",stato:"In uso",c:{gomme:"sostituita camera d'aria dx",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:38,data:"",marca:"Vermeiren",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Primo Piano",stanza:"",ospite:"Scorta",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:39,data:"",marca:"Nuova Blandino",modello:"GR101",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"14 bis",ospite:"Beretta Adele",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:40,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Morello Livia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:41,data:"",marca:"Wimed",modello:"Winner TWO",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Cerello Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:42,data:"",marca:"Rusch SRL",modello:"854006",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Mainini Antonia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:43,data:"",marca:"Wimed",modello:"Winner One",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Dal Cin Laura",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:44,data:"",marca:"Nuova Blandino",modello:"GR 104F",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Antognoli Luigi Felice",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:45,data:"",marca:"Thuasne",modello:"Classic EI+",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Sesona Giovanni",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:46,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Galuppi Mario",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:47,data:"",marca:"Sunrise Medical",modello:"Breezy Rubix 2",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Ruggeri Armando",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:48,data:"",marca:"Vermeiren",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Bertini Roberto Eugenio",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:49,data:"",marca:"Termigea",modello:"M3",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"15",ospite:"Carnaghi Lidia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:50,data:"",marca:"Wimed",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Battaglia Elena",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:51,data:"",marca:"Nuova Blandino",modello:"GR101",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Riganti Carla",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:52,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"3",ospite:"Calderara Ausilia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:53,data:"",marca:"Surace",modello:"500",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"3 bis",ospite:"Manusardi Maria Pia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:54,data:"",marca:"Anonimo 1",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"3 bis",ospite:"Negri Emma",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:55,data:"",marca:"Sunrise Medical",modello:"Breezy Rubix 2",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"3",ospite:"Belloli Laura",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:56,data:"",marca:"Anonimo 1",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"Salone",ospite:"Battaglia Elena",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"sx manca molla",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:57,data:"",marca:"Wimed",modello:"Millennium III",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Peri Maurizio",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:58,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"CD",ospite:"Salemme Ida",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:59,data:"",marca:"Wimed",modello:"Winner TWO",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"Scorta",ospite:"Scorta GV",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:60,data:"",marca:"Vermeiren",modello:"708 Delight",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"Scorta",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:61,data:"",marca:"Wimed",modello:"Winner TWO",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"Scorta",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:62,data:"",marca:"Sunrise Medical",modello:"Breezy Unix 2",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"74",ospite:"Galli Marisa",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:63,data:"",marca:"Vermeiren",modello:"R708 TII",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"5 bis",ospite:"Montani Francesca",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:64,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Milani Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:65,data:"",marca:"OSD",modello:"Millennium II",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Checchi Daniele",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:66,data:"",marca:"Vermeiren",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"11",ospite:"Bonini Franca",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:67,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Giallo",stanza:"11",ospite:"Pivaro Alice",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:68,data:"",marca:"Wimed",modello:"Winner TWO",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"12",ospite:"Rovarotto Lidia",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:69,data:"",marca:"Surace",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"12",ospite:"Scoto Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:70,data:"",marca:"Nuova Blandino",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"13",ospite:"Albini Anna",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:71,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Giallo",stanza:"10",ospite:"Allemagna Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:72,data:"",marca:"Nuova Blandino",modello:"GR104F",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Giallo",stanza:"9",ospite:"Benetti Marcello",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:73,data:"",marca:"Surace",modello:"",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Giallo",stanza:"9",ospite:"Poto Michelangelo",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:74,data:"",marca:"Anonimo",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"14",ospite:"Donatelli Carmela",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:75,data:"",marca:"Wimed",modello:"Light Plus",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"14",ospite:"Bollini Maria Angela",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:76,data:"",marca:"Ardea",modello:"One",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Giallo",stanza:"11",ospite:"Calloni Onorina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:77,data:"",marca:"Nuova Blandino",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"4",ospite:"Tasca Romana",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:78,data:"",marca:"Wimed",modello:"Winner Excel TWO",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"2",ospite:"Canali Anna Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:79,data:"",marca:"Mediland SRL",modello:"854726",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"1",ospite:"Pozzi Maria Rosa",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:80,data:"",marca:"Wimed",modello:"Easy Light",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"1",ospite:"Mettifogo Laura",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:81,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"Verde",stanza:"5",ospite:"Bergantino Francesca",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:82,data:"",marca:"Media Reha",modello:"",seriale:"",tipologia:"Bascula",fornitore:"Melo",nucleo:"Verde",stanza:"3",ospite:"Milani Giuseppina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"2026-05-06 00:00:00",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:83,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Scorta 01 G/V",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:84,data:"",marca:"Nuova Blandino",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Scorta 02 G/V",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:85,data:"",marca:"Vermeiren",modello:"R708TII",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Basoni Maria - Scorta 03 G/V",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:86,data:"",marca:"Wimed",modello:"Winner TWO Transit",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Scorta 04 G/V",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:87,data:"",marca:"Vermeiren",modello:"R708TII",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Scorta 05 G/V",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:88,data:"",marca:"Vermeiren",modello:"R708TII",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Grigio",stanza:"16",ospite:"De Bernardi Anna Maria",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:89,data:"",marca:"Termigea",modello:"M3",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Verde",stanza:"",ospite:"Scorta CDP",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:"3-Bis",data:"",marca:"Surace",modello:"200 Classic",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Rosso",stanza:"",ospite:"Terreni  Costantina",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:90,data:"",marca:"Wimed",modello:"Winner Excel Transit Plus",seriale:"",tipologia:"Transito",fornitore:"Melo",nucleo:"Palestra FKT",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:91,data:"",marca:"Netti",modello:"4 U CE Plus",seriale:"",tipologia:"Bascula",fornitore:"Melo",nucleo:"Palestra FKT",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:92,data:"",marca:"Practica",modello:"Easy",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"Terra",stanza:"",ospite:"Scorta",stato:"In uso",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:93,data:"",marca:"Sunrise Medical",modello:"Breezy Rubix 2",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:94,data:"",marca:"Nuova Blandino",modello:"",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:95,data:"",marca:"Nuova Blandino",modello:"",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:96,data:"",marca:"Surace",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:97,data:"",marca:"Nuova Blandino",modello:"GR105",seriale:"",tipologia:"Transito",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:98,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:99,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:100,data:"",marca:"Vassilli",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:101,data:"",marca:"Surace",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:102,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:103,data:"",marca:"Wimed",modello:"Millennium III",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:104,data:"",marca:"Mediland SRL",modello:"Kometa",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:105,data:"",marca:"Nuova Blandino",modello:"104",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:106,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:107,data:"",marca:"SunTec",modello:"4021T40",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:108,data:"",marca:"",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:109,data:"",marca:"Wimed",modello:"",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"mancano",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:110,data:"",marca:"Surace",modello:"Euro 600",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:111,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:112,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:113,data:"",marca:"Wimed",modello:"Reclining",seriale:"",tipologia:"Reclinabile",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""},
{id:114,data:"",marca:"OSD",modello:"Millennium II",seriale:"",tipologia:"Standard",fornitore:"Melo",nucleo:"IKEA",stanza:"",ospite:"",stato:"",c:{gomme:"",mozzi:"",freni:"",pedalini:"",braccioli:"",portaborraccia:"",tavolino:"",manopole:"",seduta:"",poggiatesta:"",pulizia:""},note:""}
];/* ---------- Root ---------- */
function CarrozzineModule({ onHome }) {
  const [ready, setReady] = useState(false);
  const [items, setItems] = useState([]);
  const [tab, setTab] = useState('carrozzine');
  const [openId, setOpenId] = useState(null);
  const [filterNucleo, setFilterNucleo] = useState(null);
  const [showMenu, setShowMenu] = useState(false);
  const [toast, setToast] = useState('');

  useEffect(() => {
    (async () => {
      let data = null;
      try { data = JSON.parse((await storage.get('carrozzine')).value); } catch { data = null; }
      setItems(data && data.length ? data : SEED);
      setReady(true);
    })();
  }, []);

  useEffect(() => { if (ready) storage.set('carrozzine', JSON.stringify(items)).catch(() => {}); }, [items, ready]);
  useEffect(() => { setOpenId(null); }, [tab]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 1800); };

  const updateItem = (w) => {
    setItems(prev => prev.map(x => x.id === w.id ? w : x));
  };

  const exportToExcel = () => {
    const rows = items.map(w => ({
      'ID': w.id, 'Data': w.data ? new Date(w.data + 'T00:00:00') : '', 'Marca': w.marca, 'Modello': w.modello,
      'Seriale': w.seriale, 'Tipologia': w.tipologia, 'Fornitore': w.fornitore, 'Nucleo': w.nucleo,
      'Stanza': w.stanza, 'Ospite': w.ospite, 'Stato': w.stato,
      'Gomme': w.c.gomme, 'Mozzi': w.c.mozzi, 'Freni': w.c.freni, 'Pedalini': w.c.pedalini,
      'Braccioli': w.c.braccioli, 'Portaborraccia': w.c.portaborraccia, 'Tavolino': w.c.tavolino,
      'Manopole': w.c.manopole, 'Seduta': w.c.seduta, 'Poggiatesta': w.c.poggiatesta, 'Pulizia': w.c.pulizia,
      'Note': w.note,
    }));
    const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true });
    ws['!cols'] = [6, 12, 17, 14, 12, 13, 11, 13, 10, 20, 13, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 24].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Totale');
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Carrozzine-Melo2_Totale_${stamp}.xlsx`);
    setShowMenu(false);
    flash('File Excel scaricato');
  };

  if (!ready) {
    return <div style={{ minHeight: '100vh', background: CARROZZINE_COLORS.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', color: CARROZZINE_COLORS.muted }}>Caricamento…</div>;
  }

  const openItem = openId != null ? items.find(w => w.id === openId) : null;
  const onMenu = () => setShowMenu(true);

  let content;
  if (openItem) {
    content = <CarrozzinaDetail w={openItem} onBack={() => setOpenId(null)} onUpdate={updateItem} />;
  } else if (tab === 'carrozzine') {
    content = <CarrozzineScreen items={items} onOpen={(w) => setOpenId(w.id)} onMenu={onMenu} filterNucleo={filterNucleo} setFilterNucleo={setFilterNucleo} onHome={onHome} />;
  } else if (tab === 'controlli') {
    content = <ControlliScreen items={items} onOpen={(w) => setOpenId(w.id)} onMenu={onMenu} onHome={onHome} />;
  } else if (tab === 'nuclei') {
    content = <NucleiScreen items={items} onMenu={onMenu} onFilter={(n) => { setFilterNucleo(n === '—' ? null : n); setTab('carrozzine'); }} onHome={onHome} />;
  } else {
    content = <RiepilogoScreen items={items} onMenu={onMenu} onHome={onHome} />;
  }

  return (
    <div style={{ minHeight: '100vh', background: CARROZZINE_COLORS.bg, fontFamily: 'Inter, sans-serif', color: CARROZZINE_COLORS.ink, maxWidth: 480, margin: '0 auto', position: 'relative' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        select { cursor: pointer; }
      `}</style>
      <div style={{ paddingBottom: 78 }}>{content}</div>
      {!openItem && <BottomNav theme={CARROZZINE_COLORS} tab={tab} setTab={setTab} items={CARROZZINE_NAV_ITEMS} />}
      {showMenu && <MenuSheet theme={CARROZZINE_COLORS} onClose={() => setShowMenu(false)} onExport={exportToExcel} exportSub="Scarica il foglio Totale in .xlsx" />}
      {toast && (
        <div style={{ position: 'fixed', bottom: !openItem ? 92 : 20, left: '50%', transform: 'translateX(-50%)', background: CARROZZINE_COLORS.primaryDeep, color: '#fff', padding: '10px 18px', borderRadius: 999, fontSize: 13, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 7, zIndex: 30, maxWidth: 400 }}>
          <Check size={15} /> {toast}
        </div>
      )}
    </div>
  );
}
/* =========================================================================
   HUB — schermata di ingresso con i moduli disponibili
   ========================================================================= */

const HUB_COLORS = {
  bg: '#EFEDE6',
  surface: '#FFFFFF',
  ink: '#20221F',
  muted: '#75786F',
  line: '#DEDBCE',
};

const MODULES = [
  {
    key: 'mezzi',
    name: 'Mezzi',
    desc: 'Veicoli aziendali, interventi e scadenze',
    icon: Car,
    color: '#25454F',
    colorSoft: '#DCE7E9',
    stat: (d) => `${d.vehicles.length} veicoli`,
  },
  {
    key: 'carrozzine',
    name: 'Carrozzine',
    desc: 'Inventario, nuclei e controlli componenti',
    icon: Armchair,
    color: '#33594E',
    colorSoft: '#DCEAE3',
    stat: (d) => `${d.carrozzine.length} carrozzine`,
  },
];

function HubScreen({ onOpen, counts }) {
  return (
    <div style={{ minHeight: '100vh', background: HUB_COLORS.bg, fontFamily: 'Inter, sans-serif', color: HUB_COLORS.ink, maxWidth: 480, margin: '0 auto' }}>
      <style>{GLOBAL_FONTS}</style>
      <div style={{ background: '#1C2321', padding: '34px 20px 30px', color: '#fff' }}>
        <div style={{ width: 46, height: 46, borderRadius: 13, background: 'rgba(255,255,255,0.12)', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 14 }}>
          <WrenchHub size={24} />
        </div>
        <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 26 }}>Manutenzione</div>
        <div style={{ fontSize: 13.5, opacity: 0.7, marginTop: 4 }}>Scegli un modulo per iniziare</div>
      </div>

      <div style={{ padding: 16, display: 'flex', flexDirection: 'column', gap: 12 }}>
        {MODULES.map(m => {
          const Icon = m.icon;
          return (
            <div
              key={m.key}
              onClick={() => onOpen(m.key)}
              style={{ background: HUB_COLORS.surface, border: `1px solid ${HUB_COLORS.line}`, borderRadius: 16, padding: 16, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}
            >
              <div style={{ width: 50, height: 50, borderRadius: 13, background: m.colorSoft, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                <Icon size={24} color={m.color} />
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 16.5, marginBottom: 2 }}>{m.name}</div>
                <div style={{ fontSize: 12.5, color: HUB_COLORS.muted, marginBottom: 6 }}>{m.desc}</div>
                <span style={{ fontSize: 11, fontWeight: 700, color: m.color, background: m.colorSoft, padding: '2.5px 8px', borderRadius: 999 }}>{m.stat(counts)}</span>
              </div>
              <ChevronRight size={19} color={HUB_COLORS.muted} />
            </div>
          );
        })}

        <div style={{ border: `1.5px dashed ${HUB_COLORS.line}`, borderRadius: 16, padding: '20px 16px', textAlign: 'center', color: HUB_COLORS.muted, fontSize: 12.5 }}>
          Altri moduli in arrivo
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   APP RADICE — passa dall'Hub ai moduli, mantenendo ciascuno indipendente
   ========================================================================= */

export default function ManutenzioneApp() {
  const [screen, setScreen] = useState('hub'); // 'hub' | 'mezzi' | 'carrozzine'
  const [counts, setCounts] = useState({ vehicles: SEED_VEHICLES, carrozzine: SEED });

  // Tengo aggiornati i conteggi mostrati sull'Hub leggendo lo storage al volo
  useEffect(() => {
    (async () => {
      let v, c;
      try { v = JSON.parse((await storage.get('vehicles')).value); } catch { v = null; }
      try { c = JSON.parse((await storage.get('carrozzine')).value); } catch { c = null; }
      setCounts({
        vehicles: v && v.length ? v : SEED_VEHICLES,
        carrozzine: c && c.length ? c : SEED,
      });
    })();
  }, [screen]);

  if (screen === 'mezzi') return <MezziModule onHome={() => setScreen('hub')} />;
  if (screen === 'carrozzine') return <CarrozzineModule onHome={() => setScreen('hub')} />;
  return <HubScreen onOpen={setScreen} counts={counts} />;
}
