import React, { useState, useEffect, useMemo, useRef, useContext, createContext } from 'react';
import * as XLSX from 'xlsx';
import { supabase, supabaseConfigured } from './supabaseClient.js';
import {
  Car, Wrench, CalendarClock, BarChart3, Plus, ChevronRight, ChevronLeft,
  Gauge, Fuel, Palette, StickyNote, Trash2, Check, Search, MoreVertical,
  FileSpreadsheet, X as XIcon, Armchair, ClipboardList, Layers, AlertTriangle,
  MapPin, User, Calendar, Hash, Home, Wrench as WrenchHub,
  BedDouble, Building2, Users, Wallet,
  LogOut, ShieldCheck, Mail, Lock, Eye, EyeOff, UserCog, Image as ImageIcon,
  BookOpen, RefreshCw, Package, ExternalLink, FileText,
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

// Riferimento condiviso per lo stato "lista": riusarlo (invece di creare
// ogni volta un nuovo oggetto {name:'list'}) evita che i reset automatici
// di vista generino voci fantasma nella cronologia del tasto indietro.
const LIST_VIEW = { name: 'list' };

/* =========================================================================
   TASTO INDIETRO DEL TELEFONO — invece di uscire dall'app, torna alla
   schermata precedente (schermata > tab > dettaglio > form).
   Ogni livello di navigazione (screen, tab, view, openId, ecc.) si registra
   con useBackable(valore, setValore): ad ogni cambiamento viene creata una
   voce nella cronologia del browser; quando l'utente preme "indietro" la
   voce piu' recente viene ripristinata senza uscire dalla pagina.
   Se non c'e' piu' nulla da annullare (si e' alla schermata Home), il tasto
   indietro esce normalmente dall'app, come ci si aspetta.
   ========================================================================= */
const __backStack = [];
const __skipPushFor = new Set();

function __pushBack(setValue, previous) {
  __backStack.push({ setValue, previous });
  try { window.history.pushState({ appNav: true, depth: __backStack.length }, ''); } catch (e) {}
}

function __popBack() {
  if (__backStack.length === 0) return false;
  const { setValue, previous } = __backStack.pop();
  __skipPushFor.add(setValue);
  setValue(previous);
  return true;
}

function useBackable(value, setValue) {
  const prevRef = useRef(value);
  const firstRun = useRef(true);
  useEffect(() => {
    if (firstRun.current) { firstRun.current = false; prevRef.current = value; return; }
    if (__skipPushFor.has(setValue)) {
      __skipPushFor.delete(setValue);
      prevRef.current = value;
      return;
    }
    if (value !== prevRef.current) {
      __pushBack(setValue, prevRef.current);
      prevRef.current = value;
    }
  }, [value, setValue]);
}

// Da usare nei pulsanti "Indietro"/"Annulla" dentro l'app: cosi' si comportano
// esattamente come il tasto indietro del telefono e restano sincronizzati
// con la cronologia del browser.
function goBack() {
  if (__backStack.length > 0) { window.history.back(); }
}

// Da usare quando un'azione deve saltare piu' di un livello in un colpo solo
// (es: eliminare un elemento dalla sua scheda di modifica raggiunta da un
// dettaglio — il dettaglio sottostante non avrebbe piu' senso una volta che
// l'elemento non esiste piu', quindi si salta dritti alla destinazione finale).
let __suppressPopstateCount = 0;
function goBackMulti(n, setValue, finalValue) {
  if (__backStack.length === 0) return;
  n = Math.min(n, __backStack.length);
  for (let i = 0; i < n; i++) __backStack.pop();
  __skipPushFor.add(setValue);
  setValue(finalValue);
  __suppressPopstateCount += n;
  try { window.history.go(-n); } catch (e) {}
}


function useHardwareBack() {
  useEffect(() => {
    const handler = () => {
      if (__suppressPopstateCount > 0) { __suppressPopstateCount--; return; }
      __popBack();
    };
    window.addEventListener('popstate', handler);
    return () => window.removeEventListener('popstate', handler);
  }, []);
}


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

function ConfirmDelete({ theme, message, onConfirm, onCancel }) {
  return (
    <div onClick={onCancel} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 90, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: '#fff', borderRadius: 16, padding: 20, maxWidth: 320, width: '100%' }}>
        <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 8, color: theme.ink }}>Confermi l'eliminazione?</div>
        <div style={{ fontSize: 13.5, color: theme.muted, marginBottom: 20 }}>{message || 'L\'azione non si puo\' annullare.'}</div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button onClick={onCancel} style={{ flex: 1, padding: '11px', borderRadius: 10, border: `1px solid ${theme.line}`, background: '#fff', color: theme.ink, fontWeight: 600, fontSize: 14 }}>Annulla</button>
          <button onClick={onConfirm} style={{ flex: 1, padding: '11px', borderRadius: 10, border: 'none', background: theme.danger, color: '#fff', fontWeight: 700, fontSize: 14 }}>Elimina</button>
        </div>
      </div>
    </div>
  );
}

function RenameDialog({ theme, label, currentName, onSave, onCancel }) {
  const [nuovoNome, setNuovoNome] = useState(currentName);
  const valido = nuovoNome.trim() && nuovoNome.trim() !== currentName;
  return (
    <div onClick={onCancel} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:90, display:'flex', alignItems:'center', justifyContent:'center', padding:24 }}>
      <div onClick={e => e.stopPropagation()} style={{ background:'#fff', borderRadius:16, padding:20, maxWidth:320, width:'100%' }}>
        <div style={{ fontWeight:700, fontSize:16, marginBottom:6, color:theme.ink }}>Rinomina {label}</div>
        <div style={{ fontSize:12.5, color:theme.muted, marginBottom:14 }}>Valore attuale: <b>{currentName}</b><br/>Tutti gli elementi con questo nome verranno aggiornati.</div>
        <input
          autoFocus
          value={nuovoNome}
          onChange={e => setNuovoNome(e.target.value)}
          style={{ width:'100%', boxSizing:'border-box', padding:'11px 12px', borderRadius:10, border:`1.5px solid ${theme.line}`, fontSize:15, outline:'none', color:theme.ink, marginBottom:16 }}
        />
        <div style={{ display:'flex', gap:10 }}>
          <button onClick={onCancel} style={{ flex:1, padding:'11px', borderRadius:10, border:`1px solid ${theme.line}`, background:'#fff', color:theme.ink, fontWeight:600, fontSize:14 }}>Annulla</button>
          <button onClick={() => valido && onSave(nuovoNome.trim())} style={{ flex:1, padding:'11px', borderRadius:10, border:'none', background:theme.primary, color:'#fff', fontWeight:700, fontSize:14, opacity:valido?1:0.4 }}>Salva</button>
        </div>
      </div>
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

function BottomNav({ theme, tab, setTab, items, badges }) {
  return (
    <div style={{ position: 'fixed', bottom: 0, left: 0, right: 0, background: '#fff', borderTop: `1px solid ${theme.line}`, display: 'flex', paddingBottom: 'env(safe-area-inset-bottom)', zIndex: 20 }}>
      {items.map(([key, Icon, label]) => {
        const active = tab === key;
        const badge = badges && badges[key];
        return (
          <button key={key} onClick={() => setTab(key)} style={{ flex: 1, background: 'none', border: 'none', padding: '9px 0 8px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3, color: active ? theme.primary : theme.muted }}>
            <span style={{ position: 'relative', display: 'inline-flex' }}>
              <Icon size={21} strokeWidth={active ? 2.4 : 2} />
              {!!badge && (
                <span style={{
                  position: 'absolute', top: -6, right: -8, background: '#C0392B', color: '#fff',
                  fontSize: 10, fontWeight: 700, lineHeight: 1, minWidth: 15, height: 15, borderRadius: 999,
                  display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 3px', border: '1.5px solid #fff',
                }}>
                  {badge > 99 ? '99+' : badge}
                </span>
              )}
            </span>
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
/* =========================================================================
   AUTENTICAZIONE E RUOLI (Supabase)
   ========================================================================= */
// ============================================================
//  SISTEMA NOTIFICHE
// ============================================================

const NOTIFICA_TIPI = {
  intervento_struttura:  { label: 'Interventi Struttura',  modulo: 'struttura',  emoji: '🏥' },
  anomalia_mezzi:        { label: 'Anomalie Mezzi',        modulo: 'mezzi',      emoji: '🚗' },
  segnalazione_carrozzine: { label: 'Segnalazione Carrozzine', modulo: 'carrozzine', emoji: '♿' },
};

async function inviaNotifica({ tipo, titolo, corpo, linkId, inviata_da }) {
  const id = `N-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  const modulo = NOTIFICA_TIPI[tipo]?.modulo || 'sistema';
  try {
    // 1. Salva nel database
    await supabase.from('notifiche').insert({ id, tipo, modulo, titolo, corpo, link_id: linkId, inviata_da });
    // 2. Chiama la Edge Function per spedire le push (anche app chiuse)
    const { data: sessionData } = await supabase.auth.getSession();
    const token = sessionData?.session?.access_token || import.meta.env.VITE_SUPABASE_ANON_KEY;
    await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/push-notification`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({ record: { id, tipo, titolo, corpo, link_id: linkId, inviata_da } })
    });
  } catch (e) {
    if (process.env.NODE_ENV === 'development') console.warn('inviaNotifica failed:', e);
  }
}

function useNotifications(userId, role) {
  const isAdmin = role === 'admin';
  const [notifiche, setNotifiche] = useState([]);
  const [lette, setLette] = useState(new Set());
  const [iscrizioni, setIscrizioni] = useState([]);

  useEffect(() => {
    if (!userId) return;
    let mounted = true;

    // Carica iscrizioni dell'utente corrente
    supabase.from('notifica_iscrizioni').select('tipo').eq('user_id', userId)
      .then(({ data }) => { if (mounted && data) setIscrizioni(data.map(r => r.tipo)); });

    // Carica ultime 50 notifiche
    supabase.from('notifiche').select('*').order('created_at', { ascending: false }).limit(50)
      .then(({ data }) => { if (mounted && data) setNotifiche(data); });

    // Carica quelle lette
    supabase.from('notifiche_lette').select('notifica_id').eq('user_id', userId)
      .then(({ data }) => { if (mounted && data) setLette(new Set(data.map(r => r.notifica_id))); });

    // Realtime: nuova notifica
    const ch = supabase.channel('notifiche-live')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'notifiche' }, (payload) => {
        if (!mounted) return;
        setNotifiche(prev => [payload.new, ...prev]);
        // NON mostrare notifica browser al mittente stesso
        if (payload.new.inviata_da === userId) return;
        playNotificationSound();
        // Notifica browser (admin vede sempre; altri solo se iscritti)
        const shouldNotify = isAdmin || iscrizioni.includes(payload.new.tipo);
        if (shouldNotify && typeof Notification !== 'undefined' && Notification.permission === 'granted') {
          new Notification(payload.new.titolo, {
            body: payload.new.corpo || '',
            icon: '/apple-touch-icon.png',
          });
        }
      })
      .subscribe();

    return () => { mounted = false; supabase.removeChannel(ch); };
  }, [userId, isAdmin]);

  async function markaLetta(notificaId) {
    if (lette.has(notificaId)) return;
    setLette(prev => new Set([...prev, notificaId]));
    try {
      await supabase.from('notifiche_lette').insert({ user_id: userId, notifica_id: notificaId });
    } catch (e) {
      if (process.env.NODE_ENV === 'development') console.warn('markaLetta failed:', e);
    }
  }

  async function markaTutte() {
    const target = notifiche.filter(
      n => (isAdmin || iscrizioni.includes(n.tipo)) && !lette.has(n.id)
    );
    if (!target.length) return;
    setLette(prev => new Set([...prev, ...target.map(n => n.id)]));
    try {
      await supabase.from('notifiche_lette').insert(
        target.map(n => ({ user_id: userId, notifica_id: n.id }))
      );
    } catch (e) {
      if (process.env.NODE_ENV === 'development') console.warn('markaTutte failed:', e);
    }
  }

  const mie = notifiche.filter(n => isAdmin || iscrizioni.includes(n.tipo));
  const nonLette = mie.filter(n => !lette.has(n.id)).length;

  return { notifiche: mie, lette, nonLette, iscrizioni, markaLetta, markaTutte, reload: () => {} };
}

const NotificheContext = createContext(null);
function useNotificheCtx() { return useContext(NotificheContext); }

function NotificationBell() {
  const ctx = useNotificheCtx();
  const [aperto, setAperto] = useState(false);

  async function richiediPermesso() {
    if (typeof Notification === 'undefined') return;
    let permission = Notification.permission;
    if (permission === 'default') {
      permission = await Notification.requestPermission();
    }
    if (permission === 'granted' && typeof registerPush === 'function') {
      await registerPush();
    }
  }

  if (!ctx) return null;
  const { notifiche, lette, nonLette, markaLetta, markaTutte } = ctx;

  return (
    <>
      <button
        onClick={() => { setAperto(true); richiediPermesso(); }}
        style={{ position: 'relative', background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: 10, width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', cursor: 'pointer' }}
      >
        🔔
        {nonLette > 0 && (
          <span style={{ position: 'absolute', top: -4, right: -4, background: '#C0392B', color: '#fff', fontSize: 10, fontWeight: 700, minWidth: 16, height: 16, borderRadius: 999, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 3px', border: '2px solid #1C2321' }}>
            {nonLette > 99 ? '99+' : nonLette}
          </span>
        )}
      </button>

      {aperto && (
        <div onClick={() => setAperto(false)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 80, display: 'flex', flexDirection: 'column', justifyContent: 'flex-end' }}>
          <div onClick={e => e.stopPropagation()} style={{ background: '#fff', color: '#111', borderRadius: '20px 20px 0 0', maxHeight: '75vh', display: 'flex', flexDirection: 'column', fontFamily: 'Inter, sans-serif' }}>
            <div style={{ padding: '16px 16px 8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #EEE' }}>
              <span style={{ fontWeight: 700, fontSize: 16 }}>Notifiche</span>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                {nonLette > 0 && <button onClick={markaTutte} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 12.5, fontWeight: 600 }}>Segna tutte lette</button>}
                <button onClick={() => setAperto(false)} style={{ background: 'none', border: 'none', fontSize: 18, color: '#fff' }}>✕</button>
              </div>
            </div>
            <div style={{ overflowY: 'auto', flex: 1, padding: '8px 0' }}>
              {notifiche.length === 0 && <div style={{ padding: '32px 16px', textAlign: 'center', color: '#888', fontSize: 14 }}>Nessuna notifica ricevuta.</div>}
              {notifiche.map(n => {
                const isLetta = lette.has(n.id);
                const info = NOTIFICA_TIPI[n.tipo] || { emoji: '📣' };
                return (
                  <div key={n.id} onClick={() => markaLetta(n.id)}
                    style={{ padding: '12px 16px', borderBottom: '1px solid #F5F5F5', background: isLetta ? '#fff' : '#F8F4FF', cursor: 'pointer', display: 'flex', gap: 12, alignItems: 'flex-start' }}>
                    <span style={{ fontSize: 22, flexShrink: 0 }}>{info.emoji}</span>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: isLetta ? 500 : 700, fontSize: 14, marginBottom: 2, color: '#111' }}>{n.titolo}</div>
                      {n.corpo && <div style={{ fontSize: 12.5, color: '#444', marginBottom: 4 }}>{n.corpo}</div>}
                      <div style={{ fontSize: 11.5, color: '#555' }}>
                        {n.inviata_da && `Da: ${n.inviata_da} · `}
                        {new Date(n.created_at).toLocaleString('it-IT', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })}
                      </div>
                    </div>
                    {!isLetta && <div style={{ width: 8, height: 8, borderRadius: 999, background: '#5E3A8A', flexShrink: 0, marginTop: 6 }} />}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function GestioneNotificheScreen({ onHome, myUserId }) {
  const [utenti, setUtenti] = useState([]);
  const [iscrizioni, setIscrizioni] = useState([]); // [{user_id, tipo}]
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    supabase.from('profiles').select('id,email,nome,cognome,role').order('email')
      .then(({ data }) => data && setUtenti(data));
    supabase.from('notifica_iscrizioni').select('*')
      .then(({ data }) => data && setIscrizioni(data));
  }, []);

  function isIscritto(userId, tipo) {
    return iscrizioni.some(r => r.user_id === userId && r.tipo === tipo);
  }

  async function toggleIscrizione(userId, tipo) {
    setBusy(true);
    if (isIscritto(userId, tipo)) {
      await supabase.from('notifica_iscrizioni').delete().eq('user_id', userId).eq('tipo', tipo);
      setIscrizioni(prev => prev.filter(r => !(r.user_id === userId && r.tipo === tipo)));
    } else {
      await supabase.from('notifica_iscrizioni').insert({ user_id: userId, tipo });
      setIscrizioni(prev => [...prev, { user_id: userId, tipo }]);
    }
    setBusy(false);
  }

  const tipi = Object.entries(NOTIFICA_TIPI);

  return (
    <div style={{ minHeight: '100vh', background: '#F2EEF8', fontFamily: 'Inter, sans-serif', color: '#1E1530', maxWidth: 480, margin: '0 auto' }}>
      <TopBar theme={{ primary: '#5E3A8A', line: '#DDD8EE', ink: '#1E1530', muted: '#7A6A9A', bg: '#F2EEF8' }}
        title="Gestione notifiche" subtitle="Chi riceve cosa" onBack={onHome} />
      <div style={{ padding: 16 }}>
        <div style={{ fontSize: 12.5, color: '#7A6A9A', marginBottom: 14, lineHeight: 1.5 }}>
          Seleziona per ogni utente quali tipologie di notifica riceverà. Le spunte attive ricevono notifiche push/in-app quando viene inviata una segnalazione.
        </div>
        {utenti.map(u => {
          const nome = (u.nome || u.cognome) ? `${u.nome || ''} ${u.cognome || ''}`.trim() : u.email;
          return (
            <div key={u.id} style={{ background: '#fff', border: '1px solid #DDD8EE', borderRadius: 14, padding: 14, marginBottom: 10 }}>
              <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>{nome}</div>
              <div style={{ fontSize: 12, color: '#7A6A9A', marginBottom: 10 }}>{u.email} · {u.role}</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {tipi.map(([tipo, info]) => (
                  <label key={tipo} style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={isIscritto(u.id, tipo)}
                      onChange={() => toggleIscrizione(u.id, tipo)}
                      disabled={busy}
                      style={{ width: 18, height: 18, accentColor: '#5E3A8A' }}
                    />
                    <span style={{ fontSize: 13.5 }}>{info.emoji} {info.label}</span>
                  </label>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const RoleContext = createContext({ role: 'lettore', email: '', nome: '', cognome: '' });
function useRole() { return useContext(RoleContext).role; }
function useUserEmail() { return useContext(RoleContext).email; }
function nomeVisualizzato({ nome, cognome, email }) {
  const n = `${nome || ''} ${cognome || ''}`.trim();
  return n || email;
}
function usePermessi() {
  const ctx = useContext(RoleContext);
  return {
    role: ctx.role,
    email: ctx.email,
    nome: ctx.nome,
    cognome: ctx.cognome,
    nomeVisualizzato: nomeVisualizzato(ctx),
    puoScrivere: true,
    puoEliminare: ctx.role === 'admin',
    isAdmin: ctx.role === 'admin',
  };
}

function traduciErroreAuth(msg) {
  const m = (msg || '').toLowerCase();
  if (m.includes('invalid login credentials')) return 'Email o password non corretti.';
  if (m.includes('already registered') || m.includes('already exists') || m.includes('user already')) return 'Esiste gia\' un account con questa email.';
  if (m.includes('password') && (m.includes('character') || m.includes('least'))) return 'La password deve avere almeno 6 caratteri.';
  if (m.includes('rate limit')) return 'Troppi tentativi: riprova tra qualche minuto.';
  if (m.includes('invalid') && m.includes('email')) return 'Indirizzo email non valido.';
  if (m.includes('network') || m.includes('fetch')) return 'Problema di connessione: controlla la rete e riprova.';
  return msg || 'Si e\' verificato un errore, riprova.';
}

function useAuth() {
  const [session, setSession] = useState(undefined); // undefined = non ancora verificato
  const [profile, setProfile] = useState(null);
  const [profileLoading, setProfileLoading] = useState(true);
  const [passwordRecovery, setPasswordRecovery] = useState(false);

  useEffect(() => {
    let mounted = true;
    supabase.auth.getSession().then(({ data }) => { if (mounted) setSession(data.session ?? null); });
    const { data: listener } = supabase.auth.onAuthStateChange((event, s) => {
      if (event === 'PASSWORD_RECOVERY') setPasswordRecovery(true);
      setSession(s ?? null);
    });
    return () => { mounted = false; listener.subscription.unsubscribe(); };
  }, []);

  useEffect(() => {
    let mounted = true;
    if (session === undefined) return undefined;
    if (session === null) { setProfile(null); setProfileLoading(false); return undefined; }
    setProfileLoading(true);
    supabase.from('profiles').select('*').eq('id', session.user.id).single()
      .then(({ data, error }) => { if (mounted) { setProfile(error ? null : data); setProfileLoading(false); } });
    return () => { mounted = false; };
  }, [session]);

  async function refreshProfile() {
    if (!session) return;
    const { data, error } = await supabase.from('profiles').select('*').eq('id', session.user.id).single();
    if (!error) setProfile(data);
  }

  return {
    session,
    profile,
    refreshProfile,
    passwordRecovery,
    clearPasswordRecovery: () => setPasswordRecovery(false),
    authLoading: session === undefined || (session !== null && profileLoading),
    signOut: () => supabase.auth.signOut(),
  };
}

function AuthScreen() {
  const [mode, setMode] = useState('login'); // 'login' | 'signup'
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');
  const [busy, setBusy] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (busy) return;
    setError(''); setInfo('');
    if (mode === 'reset') {
      if (!email.trim()) { setError('Inserisci la tua email.'); return; }
      setBusy(true);
      const { error } = await supabase.auth.resetPasswordForEmail(email.trim(), { redirectTo: window.location.href });
      if (error) setError(traduciErroreAuth(error.message));
      else setInfo('Se l\'indirizzo e\' registrato, ti abbiamo mandato un\'email con le istruzioni per scegliere una nuova password.');
      setBusy(false);
      return;
    }
    if (!email.trim() || !password) { setError('Inserisci email e password.'); return; }
    setBusy(true);
    if (mode === 'login') {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      if (error) setError(traduciErroreAuth(error.message));
    } else {
      const { data, error } = await supabase.auth.signUp({ email: email.trim(), password });
      if (error) setError(traduciErroreAuth(error.message));
      else if (!data.session) setInfo('Account creato: controlla la tua email per confermarlo, poi accedi. Un admin dovra\' assegnarti il ruolo giusto.');
      else setInfo('Account creato. Un admin dovra\' assegnarti il ruolo giusto prima che tu possa modificare i dati.');
    }
    setBusy(false);
  }

  const inputStyle = {
    width: '100%', boxSizing: 'border-box', padding: '12px 12px 12px 40px', borderRadius: 10,
    border: '1.5px solid #3A423F', background: '#222B27', color: '#fff', fontSize: 15, outline: 'none',
  };

  const titolo = mode === 'login' ? 'Accedi per continuare' : mode === 'signup' ? 'Crea un nuovo account' : 'Recupera la password';

  return (
    <div style={{ minHeight: '100vh', background: '#1C2321', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: 'Inter, sans-serif', color: '#fff' }}>
      <style>{GLOBAL_FONTS}</style>
      <div style={{ width: '100%', maxWidth: 360 }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: 28 }}>
          <div style={{ width: 52, height: 52, borderRadius: 14, background: 'rgba(255,255,255,0.12)', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 14 }}>
            <WrenchHub size={26} />
          </div>
          <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 24 }}>Manutenzione</div>
          <div style={{ fontSize: 13, opacity: 0.6, marginTop: 4 }}>{titolo}</div>
        </div>

        <form onSubmit={handleSubmit}>
          <div style={{ position: 'relative', marginBottom: 12 }}>
            <Mail size={16} style={{ position: 'absolute', left: 13, top: 14, opacity: 0.5 }} />
            <input type="email" autoComplete="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} style={inputStyle} />
          </div>
          {mode !== 'reset' && (
            <div style={{ position: 'relative', marginBottom: 8 }}>
              <Lock size={16} style={{ position: 'absolute', left: 13, top: 14, opacity: 0.5 }} />
              <input type={showPw ? 'text' : 'password'} autoComplete={mode === 'login' ? 'current-password' : 'new-password'} placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} style={{ ...inputStyle, paddingRight: 40 }} />
              <button type="button" onClick={() => setShowPw(!showPw)} style={{ position: 'absolute', right: 10, top: 10, background: 'none', border: 'none', color: '#fff', opacity: 0.6, padding: 4 }}>
                {showPw ? <EyeOff size={17} /> : <Eye size={17} />}
              </button>
            </div>
          )}

          {mode === 'login' && (
            <button
              type="button"
              onClick={() => { setMode('reset'); setError(''); setInfo(''); }}
              style={{ display: 'block', marginLeft: 'auto', background: 'none', border: 'none', color: '#8FA69D', fontSize: 12.5, padding: '2px 2px 14px' }}
            >
              Password dimenticata?
            </button>
          )}

          {error && (
            <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start', background: 'rgba(178,58,46,0.18)', color: '#F5A896', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>
              <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
            </div>
          )}
          {info && (
            <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start', background: 'rgba(76,122,79,0.2)', color: '#A9D6AC', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>
              <Check size={15} style={{ flexShrink: 0, marginTop: 1 }} /> {info}
            </div>
          )}

          <button type="submit" disabled={busy} style={{ width: '100%', background: '#fff', color: '#1C2321', border: 'none', borderRadius: 10, padding: '13px', fontWeight: 700, fontSize: 15, marginTop: 4, opacity: busy ? 0.6 : 1 }}>
            {busy ? 'Un attimo…' : mode === 'login' ? 'Accedi' : mode === 'signup' ? 'Crea account' : 'Invia email di ripristino'}
          </button>
        </form>

        {mode === 'reset' ? (
          <button
            onClick={() => { setMode('login'); setError(''); setInfo(''); }}
            style={{ width: '100%', background: 'none', border: 'none', color: '#B7C7C1', fontSize: 13, marginTop: 18, padding: 8 }}
          >
            Torna al login
          </button>
        ) : (
          <button
            onClick={() => { setMode(mode === 'login' ? 'signup' : 'login'); setError(''); setInfo(''); }}
            style={{ width: '100%', background: 'none', border: 'none', color: '#B7C7C1', fontSize: 13, marginTop: 18, padding: 8 }}
          >
            {mode === 'login' ? 'Non hai un account? Registrati' : 'Hai gia\' un account? Accedi'}
          </button>
        )}
      </div>
    </div>
  );
}

function NewPasswordScreen({ onDone }) {
  const [password, setPassword] = useState('');
  const [password2, setPassword2] = useState('');
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const inputStyle = {
    width: '100%', boxSizing: 'border-box', padding: '12px 40px 12px 40px', borderRadius: 10,
    border: '1.5px solid #3A423F', background: '#222B27', color: '#fff', fontSize: 15, outline: 'none',
  };

  async function handleSubmit(e) {
    e.preventDefault();
    if (busy) return;
    setError('');
    if (password.length < 6) { setError('La password deve avere almeno 6 caratteri.'); return; }
    if (password !== password2) { setError('Le due password non coincidono.'); return; }
    setBusy(true);
    const { error } = await supabase.auth.updateUser({ password });
    setBusy(false);
    if (error) { setError(traduciErroreAuth(error.message)); return; }
    onDone();
  }

  return (
    <div style={{ minHeight: '100vh', background: '#1C2321', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: 'Inter, sans-serif', color: '#fff' }}>
      <style>{GLOBAL_FONTS}</style>
      <div style={{ width: '100%', maxWidth: 360 }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: 28 }}>
          <div style={{ width: 52, height: 52, borderRadius: 14, background: 'rgba(255,255,255,0.12)', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 14 }}>
            <Lock size={24} />
          </div>
          <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 22, textAlign: 'center' }}>Scegli una nuova password</div>
        </div>
        <form onSubmit={handleSubmit}>
          <div style={{ position: 'relative', marginBottom: 12 }}>
            <Lock size={16} style={{ position: 'absolute', left: 13, top: 14, opacity: 0.5 }} />
            <input type={showPw ? 'text' : 'password'} autoComplete="new-password" placeholder="Nuova password" value={password} onChange={(e) => setPassword(e.target.value)} style={inputStyle} />
            <button type="button" onClick={() => setShowPw(!showPw)} style={{ position: 'absolute', right: 10, top: 10, background: 'none', border: 'none', color: '#fff', opacity: 0.6, padding: 4 }}>
              {showPw ? <EyeOff size={17} /> : <Eye size={17} />}
            </button>
          </div>
          <div style={{ position: 'relative', marginBottom: 12 }}>
            <Lock size={16} style={{ position: 'absolute', left: 13, top: 14, opacity: 0.5 }} />
            <input type={showPw ? 'text' : 'password'} autoComplete="new-password" placeholder="Ripeti la password" value={password2} onChange={(e) => setPassword2(e.target.value)} style={inputStyle} />
          </div>
          {error && (
            <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start', background: 'rgba(178,58,46,0.18)', color: '#F5A896', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>
              <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
            </div>
          )}
          <button type="submit" disabled={busy} style={{ width: '100%', background: '#fff', color: '#1C2321', border: 'none', borderRadius: 10, padding: '13px', fontWeight: 700, fontSize: 15, marginTop: 4, opacity: busy ? 0.6 : 1 }}>
            {busy ? 'Un attimo…' : 'Salva la nuova password'}
          </button>
        </form>
      </div>
    </div>
  );
}

const RUOLI = ['admin', 'operatore', 'lettore'];
const RUOLO_LABEL = { admin: 'Admin', operatore: 'Operatore', lettore: 'Sola lettura' };
const RUOLO_STYLE = {
  admin: { bg: '#F7DCD9', fg: '#A3352A' },
  operatore: { bg: '#FBEDD2', fg: '#8A5A00' },
  lettore: { bg: '#E8E5DC', fg: '#5B564A' },
};

function UtentiScreen({ onHome, myUserId }) {
  const [profiles, setProfiles] = useState(null);
  const [error, setError] = useState('');
  const [savingId, setSavingId] = useState(null);

  useEffect(() => {
    let mounted = true;
    supabase.from('profiles').select('*').order('email').then(({ data, error }) => {
      if (!mounted) return;
      if (error) setError(error.message); else setProfiles(data);
    });
    return () => { mounted = false; };
  }, []);

  async function cambiaRuolo(id, nuovoRuolo) {
    setError('');
    setSavingId(id);
    const prev = profiles;
    setProfiles((p) => p.map((u) => (u.id === id ? { ...u, role: nuovoRuolo } : u)));
    const { error } = await supabase.from('profiles').update({ role: nuovoRuolo }).eq('id', id);
    if (error) { setProfiles(prev); setError('Non sono riuscito a salvare: ' + error.message); }
    setSavingId(null);
  }

  return (
    <div style={{ minHeight: '100vh', background: HUB_COLORS.bg, fontFamily: 'Inter, sans-serif', color: HUB_COLORS.ink, maxWidth: 480, margin: '0 auto' }}>
      <style>{GLOBAL_FONTS}</style>
      <TopBar theme={HUB_COLORS} title="Gestione utenti" subtitle={profiles ? `${profiles.length} account` : 'Caricamento…'} onBack={onHome} backIcon={Home} />
      <div style={{ padding: 14 }}>
        <p style={{ fontSize: 12, color: HUB_COLORS.muted, marginTop: 0, marginBottom: 14 }}>
          Admin: accesso completo. Operatore: puo\' aggiungere e modificare, non eliminare. Sola lettura: puo\' solo consultare.
          Il tuo ruolo non e\' modificabile da qui: chiedi a un altro admin.
        </p>
        {error && <div style={{ background: '#F7DCD9', color: '#A3352A', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>{error}</div>}
        {!profiles && <div style={{ textAlign: 'center', color: HUB_COLORS.muted, padding: 30 }}>Caricamento…</div>}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {profiles && profiles.map((u) => {
            const isMe = u.id === myUserId;
            return (
              <div key={u.id} style={{ background: HUB_COLORS.surface, border: `1px solid ${HUB_COLORS.line}`, borderRadius: 14, padding: 14 }}>
                <div style={{ fontWeight: 700, fontSize: 14, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {(u.nome || u.cognome) ? `${u.nome || ''} ${u.cognome || ''}`.trim() : (u.email || '(email non disponibile)')}
                  {isMe && <span style={{ fontWeight: 500, color: HUB_COLORS.muted }}> (tu)</span>}
                </div>
                {(u.nome || u.cognome) && u.email && (
                  <div style={{ fontSize: 11.5, color: HUB_COLORS.muted, marginTop: 1 }}>{u.email}</div>
                )}
                <div style={{ display: 'flex', gap: 6, marginTop: 10, flexWrap: 'wrap' }}>
                  {RUOLI.map((r) => (
                    <button
                      key={r}
                      disabled={savingId === u.id || isMe}
                      onClick={() => cambiaRuolo(u.id, r)}
                      title={isMe ? 'Non puoi modificare il tuo stesso ruolo' : ''}
                      style={{
                        border: 'none', borderRadius: 999, padding: '6px 12px', fontSize: 12, fontWeight: 700,
                        cursor: isMe ? 'not-allowed' : 'pointer',
                        background: u.role === r ? RUOLO_STYLE[r].fg : RUOLO_STYLE[r].bg,
                        color: u.role === r ? '#fff' : RUOLO_STYLE[r].fg,
                        opacity: (savingId === u.id || isMe) ? 0.5 : 1,
                      }}
                    >
                      {RUOLO_LABEL[r]}
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ProfiloScreen({ onHome, profile, onSaved }) {
  const [nome, setNome] = useState(profile?.nome || '');
  const [cognome, setCognome] = useState(profile?.cognome || '');
  const [busy, setBusy] = useState(false);
  const [errore, setErrore] = useState('');
  const [salvato, setSalvato] = useState(false);

  async function salva() {
    setBusy(true); setErrore(''); setSalvato(false);
    const { error } = await supabase.rpc('update_my_profile', { p_nome: nome, p_cognome: cognome });
    setBusy(false);
    if (error) { setErrore(traduciErroreDati(error.message)); return; }
    setSalvato(true);
    onSaved();
  }

  return (
    <div style={{ minHeight: '100vh', background: HUB_COLORS.bg, fontFamily: 'Inter, sans-serif', color: HUB_COLORS.ink, maxWidth: 480, margin: '0 auto' }}>
      <style>{GLOBAL_FONTS}</style>
      <TopBar theme={HUB_COLORS} title="Il tuo profilo" subtitle={profile?.email} onBack={onHome} backIcon={Home} />
      <div style={{ padding: 16 }}>
        <p style={{ fontSize: 12.5, color: HUB_COLORS.muted, marginTop: 0, marginBottom: 18 }}>
          Nome e cognome vengono mostrati al posto dell'email nell'app (es. nelle segnalazioni). Il ruolo non si cambia da qui.
        </p>
        <label style={{ display: 'block', marginBottom: 14 }}>
          <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: HUB_COLORS.muted, marginBottom: 5, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Nome</span>
          <input
            value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Es. Mario"
            style={{ width: '100%', boxSizing: 'border-box', padding: '11px 12px', borderRadius: 10, border: `1.5px solid ${HUB_COLORS.line}`, fontSize: 15, background: '#FCFBF7', color: HUB_COLORS.ink, outline: 'none' }}
          />
        </label>
        <label style={{ display: 'block', marginBottom: 18 }}>
          <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: HUB_COLORS.muted, marginBottom: 5, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Cognome</span>
          <input
            value={cognome} onChange={(e) => setCognome(e.target.value)} placeholder="Es. Rossi"
            style={{ width: '100%', boxSizing: 'border-box', padding: '11px 12px', borderRadius: 10, border: `1.5px solid ${HUB_COLORS.line}`, fontSize: 15, background: '#FCFBF7', color: HUB_COLORS.ink, outline: 'none' }}
          />
        </label>
        {errore && <div style={{ background: '#F7DCD9', color: '#A3352A', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>{errore}</div>}
        {salvato && <div style={{ background: '#DCEEE3', color: '#1F6B45', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>Salvato.</div>}
        <button
          onClick={salva}
          disabled={busy}
          style={{ width: '100%', background: HUB_COLORS.ink || '#1C2321', color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, opacity: busy ? 0.6 : 1 }}
        >
          {busy ? 'Salvataggio…' : 'Salva'}
        </button>
      </div>
    </div>
  );
}

/* =========================================================================
   DATI CONDIVISI (Supabase) — un'unica hook riusabile per ogni tabella:
   carica le righe, le tiene sincronizzate in tempo reale con chi altro e'
   collegato, ed espone save/remove che scrivono sul database.
   ========================================================================= */
function useSupaTable(table, idKey, seed, dateFields = []) {
  const [rows, setRows] = useState(seed || []);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    let mounted = true;
    setReady(false);
    supabase.from(table).select('*').then(({ data, error }) => {
      if (!mounted) return;
      if (error) { setError(error.message); setRows(seed || []); }
      else { setRows(data || []); setError(''); }
      setReady(true);
    });
    // Se un'altra persona collegata modifica questa tabella, mi aggiorno da solo
    const channel = supabase
      .channel(`sync-${table}`)
      .on('postgres_changes', { event: '*', schema: 'public', table }, () => {
        supabase.from(table).select('*').then(({ data }) => { if (mounted && data) setRows(data); });
      })
      .subscribe();
    return () => { mounted = false; supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table]);

  async function save(record) {
    setError('');
    // Le colonne data non accettano stringa vuota: una data non compilata
    // deve arrivare come NULL, altrimenti Postgres la rifiuta.
    const pulito = { ...record };
    dateFields.forEach((f) => { if (pulito[f] === '') pulito[f] = null; });
    delete pulito.days;
    const { data, error } = await supabase.from(table).upsert(pulito).select().single();
    if (error) { const msg = traduciErroreDati(error.message); setError(msg); return { error: { message: msg } }; }
    setRows((prev) => (prev.some((r) => r[idKey] === data[idKey]) ? prev.map((r) => (r[idKey] === data[idKey] ? data : r)) : [...prev, data]));
    return { data };
  }

  async function remove(record) {
    setError('');
    const { error } = await supabase.from(table).delete().eq(idKey, record[idKey]);
    if (error) { const msg = traduciErroreDati(error.message); setError(msg); return { error: { message: msg } }; }
    setRows((prev) => prev.filter((r) => r[idKey] !== record[idKey]));
    return {};
  }

  async function reload() {
    const { data, error: err } = await supabase.from(table).select('*');
    if (!err && data) setRows(data);
  }

  return { rows, setRows, ready, error, setError, save, remove, reload };
}

function traduciErroreDati(msg) {
  const m = (msg || '').toLowerCase();
  if (m.includes('row-level security') || m.includes('permission denied')) return 'Non hai i permessi per questa operazione (il tuo ruolo non lo consente).';
  if (m.includes('duplicate key')) return 'Esiste gia\' un elemento con questo identificativo.';
  if (m.includes('network') || m.includes('fetch failed') || m.includes('failed to fetch')) return 'Problema di connessione: controlla la rete e riprova.';
  return msg || 'Si e\' verificato un errore, riprova.';
}

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

const MEZZI_FOTO_BUCKET = 'mezzi-foto';

async function caricaFotoMezzi(recordId, file) {
  const ext = (file.name.split('.').pop() || 'jpg').toLowerCase().replace(/[^a-z0-9]/g, '');
  const path = `${recordId}/${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
  const { error } = await supabase.storage.from(MEZZI_FOTO_BUCKET).upload(path, file, { upsert: false });
  if (error) throw error;
  return path;
}
async function eliminaFotoMezzi(paths) {
  if (!paths?.length) return;
  try { await supabase.storage.from(MEZZI_FOTO_BUCKET).remove(paths); } catch {}
}
async function urlFirmateMezzi(paths) {
  if (!paths?.length) return {};
  const { data } = await supabase.storage.from(MEZZI_FOTO_BUCKET).createSignedUrls(paths, 3600);
  const map = {};
  (data || []).forEach(d => { if (d.signedUrl && d.path) map[d.path] = d.signedUrl; });
  return map;
}

function MezziFotoPicker({ fotoEsistenti, onRemoveEsistente, fotoNuove, onAddNuove, onRemoveNuova }) {
  const [urls, setUrls] = useState({});
  const [zoom, setZoom] = useState(null);
  const fileInputRef = useRef(null);
  const anteprimeNuove = useMemo(() => fotoNuove.map(f => URL.createObjectURL(f)), [fotoNuove]);
  useEffect(() => {
    return () => { anteprimeNuove.forEach(url => URL.revokeObjectURL(url)); };
  }, [anteprimeNuove]);
  const chiave = (fotoEsistenti || []).join(',');
  useEffect(() => {
    let m = true;
    urlFirmateMezzi(fotoEsistenti || []).then(u => { if (m) setUrls(u); });
    return () => { m = false; };
  }, [chiave]);

  const thumbStyle = { position: 'relative', width: 72, height: 72, borderRadius: 10, overflow: 'hidden', background: MEZZI_COLORS.bg, border: `1px solid ${MEZZI_COLORS.line}`, flexShrink: 0, cursor: 'pointer' };
  const xStyle = { position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', border: 'none', borderRadius: 999, width: 20, height: 20, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center' };

  return (
    <div style={{ marginBottom: 14 }}>
      <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: MEZZI_COLORS.muted, marginBottom: 8, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Fotografie</span>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
        {(fotoEsistenti || []).map((p, i) => (
          <div key={i} style={thumbStyle} onClick={() => urls[p] && setZoom(urls[p])}>
            {urls[p] && <img src={urls[p]} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />}
            <button type="button" onClick={e => { e.stopPropagation(); onRemoveEsistente(p); }} style={xStyle}><XIcon size={12} /></button>
          </div>
        ))}
        {fotoNuove.map((file, i) => (
          <div key={`n${i}`} style={thumbStyle} onClick={() => setZoom(anteprimeNuove[i])}>
            <img src={anteprimeNuove[i]} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            <button type="button" onClick={e => { e.stopPropagation(); onRemoveNuova(i); }} style={xStyle}><XIcon size={12} /></button>
          </div>
        ))}
        <button type="button" onClick={() => fileInputRef.current?.click()}
          style={{ width: 72, height: 72, borderRadius: 10, border: `1.5px dashed ${MEZZI_COLORS.line}`, background: 'none', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: MEZZI_COLORS.muted, gap: 3, flexShrink: 0 }}>
          <Plus size={18} /><span style={{ fontSize: 10 }}>Aggiungi</span>
        </button>
        <input ref={fileInputRef} type="file" accept="image/*" multiple style={{ display: 'none' }}
          onChange={e => { onAddNuove(Array.from(e.target.files || [])); e.target.value = ''; }} />
      </div>
      {zoom && (
        <div onClick={() => setZoom(null)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.9)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
          <button onClick={() => setZoom(null)} style={{ position: 'absolute', top: 16, right: 16, background: 'rgba(255,255,255,0.15)', border: 'none', borderRadius: 999, width: 38, height: 38, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><XIcon size={18} /></button>
          <img src={zoom} alt="" style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} onClick={e => e.stopPropagation()} />
        </div>
      )}
    </div>
  );
}

const STATI_MANUTENZIONE = ['Completato', 'In Garanzia', 'Programmato', 'In Attesa Ricambi'];
const STATI_EFFETTUATI = ['Completato', 'In Garanzia'];
const STATI_PROGRAMMATI = ['Programmato', 'In Attesa Ricambi'];
const PRIORITA_MAINT = ['Alta', 'Media', 'Bassa'];
const PRIORITA_STYLE = {
  'Alta':  { bg: '#F7DCD9', fg: '#A3352A' },
  'Media': { bg: '#FDF3D0', fg: '#8A6000' },
  'Bassa': { bg: '#DCEEE3', fg: '#1F6B45' },
};
const TIPI_VEICOLO = ['Auto', 'Furgone', 'Pulmino'];
const CARBURANTI = ['Benzina', 'Diesel', 'GPL', 'Metano', 'Elettrico', 'Ibrido'];
const TIPI_GOMME = ['Invernali', 'Estive', 'Quattro stagioni'];

// Ritorna null | { tipo: 'invernali'|'estive', livello: 'giallo'|'rosso', testo: string }
function alertGomme(tipologiaGomme) {
  if (tipologiaGomme === 'Quattro stagioni') return null;
  const now = new Date();
  const mese = now.getMonth() + 1; // 1-12
  const giorno = now.getDate();

  // Calcolo "giorno dell'anno" semplificato: usiamo mese*100+giorno per comparazioni semplici
  const md = mese * 100 + giorno;

  // Periodi cambio:
  //   Estive (A -> verso invernali): giallo 1015-1115, rosso 1116-0414
  //   Invernali (B -> verso estive): giallo 0415-0515, rosso 0516-1014

  const gialloInvernali = (md >= 1015 && md <= 1115);
  const rossoInvernali  = (md >= 1116 || md <= 414); // novembre 16 -> aprile 14
  const gialloEstive    = (md >= 415 && md <= 515);
  const rossoEstive     = (md >= 516 && md <= 1014);

  if (tipologiaGomme === 'Invernali') {
    // Montate invernali: non avvisare di montare invernali (A)
    // Avvisare di montare estive (B)
    if (gialloEstive) return { tipo: 'estive', livello: 'giallo', testo: 'Montare gomme estive' };
    if (rossoEstive)  return { tipo: 'estive', livello: 'rosso',  testo: 'Montare gomme estive' };
    return null;
  }
  if (tipologiaGomme === 'Estive') {
    // Montate estive: non avvisare di montare estive (B)
    // Avvisare di montare invernali (A)
    if (gialloInvernali) return { tipo: 'invernali', livello: 'giallo', testo: 'Montare gomme invernali' };
    if (rossoInvernali)  return { tipo: 'invernali', livello: 'rosso',  testo: 'Montare gomme invernali' };
    return null;
  }
  // Tipologia non specificata: mostra tutti gli avvisi
  if (gialloInvernali) return { tipo: 'invernali', livello: 'giallo', testo: 'Montare gomme invernali' };
  if (rossoInvernali)  return { tipo: 'invernali', livello: 'rosso',  testo: 'Montare gomme invernali' };
  if (gialloEstive)    return { tipo: 'estive', livello: 'giallo', testo: 'Montare gomme estive' };
  if (rossoEstive)     return { tipo: 'estive', livello: 'rosso',  testo: 'Montare gomme estive' };
  return null;
}


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
  ['segnalazioni', AlertTriangle, 'Segnala'],
  ['scadenze', CalendarClock, 'Scadenze'],
  ['dashboard', BarChart3, 'Riepilogo'],
];

/* ---------- helpers ---------- */
const uid = () => Math.random().toString(36).slice(2, 10);
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; };
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

function FAB({ onClick, label, visible }) {
  const { isAdmin } = usePermessi();
  if (!isAdmin && !visible) return null;
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
  const { isAdmin } = usePermessi();
  const puoScrivere = isAdmin;
  const puoEliminare = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const deadlines = [
    ['Assicurazione', vehicle.assicurazione], ['Revisione', vehicle.revisione], ['Bollo', vehicle.bollo],
  ].filter(([, d]) => d);
  const own = maints.filter(m => m.targa === vehicle.targa).sort((a, b) => (b.data || '').localeCompare(a.data || ''));
  const totale = own.reduce((s, m) => s + (Number(m.costo) || 0) * 1.22, 0);
  const alertGomma = alertGomme(vehicle.tipologiaGomme || '');

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={`${vehicle.marca} ${vehicle.modello}`} subtitle={vehicle.tipo} onBack={onBack} />
      <div style={{ padding: 14 }}>
        {alertGomma && (
          <div style={{
            display: 'flex', alignItems: 'center', gap: 10, padding: '12px 14px', borderRadius: 12, marginBottom: 12,
            background: alertGomma.livello === 'rosso' ? '#F7DCD9' : '#FDF3D0',
            border: `1.5px solid ${alertGomma.livello === 'rosso' ? '#E8A09A' : '#F0D060'}`,
          }}>
            <AlertTriangle size={18} color={alertGomma.livello === 'rosso' ? '#A3352A' : '#8A6000'} style={{ flexShrink: 0 }} />
            <div>
              <div style={{ fontWeight: 700, fontSize: 13.5, color: alertGomma.livello === 'rosso' ? '#A3352A' : '#6B4A00' }}>
                {alertGomma.testo.toUpperCase()}
              </div>
              <div style={{ fontSize: 11.5, color: alertGomma.livello === 'rosso' ? '#C0504A' : '#8A6A00', marginTop: 2 }}>
                {alertGomma.livello === 'giallo' ? 'Periodo consigliato' : 'Periodo scaduto — cambio urgente'}
              </div>
            </div>
          </div>
        )}

        <Card theme={MEZZI_COLORS} style={{ marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <Plate targa={vehicle.targa} />
            {puoScrivere && (
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={() => onEdit(vehicle)} style={{ border: `1.5px solid ${MEZZI_COLORS.line}`, background: '#fff', borderRadius: 9, padding: '7px 12px', fontSize: 13, fontWeight: 600, color: MEZZI_COLORS.primary }}>Modifica</button>
                {puoEliminare && (
                  <button onClick={() => setConfermaElimina(true)} style={{ border: `1.5px solid ${MEZZI_COLORS.line}`, background: '#fff', borderRadius: 9, padding: '7px 9px' }}><Trash2 size={15} color={MEZZI_COLORS.danger} /></button>
                )}
              </div>
            )}
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

        <SectionLabel theme={MEZZI_COLORS}>Pneumatici</SectionLabel>
        <Card theme={MEZZI_COLORS} style={{ marginBottom: 12 }}>
          <InfoRow theme={MEZZI_COLORS} icon={Car} label="Modello pneumatici" value={vehicle.modelloGomme || '—'} />
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', paddingTop: 10, marginTop: 10, borderTop: `1px solid ${MEZZI_COLORS.line}` }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: MEZZI_COLORS.muted }}>Tipologia montata</span>
            <span style={{
              fontWeight: 700, fontSize: 13,
              color: vehicle.tipologiaGomme === 'Invernali' ? '#1E4D73' : vehicle.tipologiaGomme === 'Estive' ? '#1F6B45' : vehicle.tipologiaGomme === 'Quattro stagioni' ? '#4C3B75' : MEZZI_COLORS.muted
            }}>
              {vehicle.tipologiaGomme || '—'}
            </span>
          </div>
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
      {confermaElimina && (
        <ConfirmDelete
          theme={MEZZI_COLORS}
          message={`Eliminare ${vehicle.marca} ${vehicle.modello} (${vehicle.targa || 'senza targa'})?`}
          onConfirm={() => { setConfermaElimina(false); onDelete(vehicle); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Vehicle form ---------- */
function VehicleForm({ initial, onSave, onCancel }) {
  const { isAdmin } = usePermessi();
  const puoScrivere = isAdmin;
  const [f, setF] = useState(initial || { targa: '', marca: '', modello: '', tipo: 'Auto', anno: '', km: '', carburante: '', colore: '', assicurazione: '', revisione: '', bollo: '', modelloGomme: '', tipologiaGomme: '', note: '' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const valid = f.marca.trim() && f.modello.trim();

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={initial ? 'Modifica veicolo' : 'Nuovo veicolo'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
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
        <SectionLabel theme={MEZZI_COLORS}>Pneumatici</SectionLabel>
        <Field label="Modello pneumatici"><input style={inputStyle} value={f.modelloGomme || ''} onChange={set('modelloGomme')} placeholder="Es. Pirelli Cinturato P7 205/55 R16" /></Field>
        <Field label="Tipologia montata">
          <select style={inputStyle} value={f.tipologiaGomme || ''} onChange={set('tipologiaGomme')}>
            <option value="">— non specificata</option>
            {TIPI_GOMME.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <SectionLabel theme={MEZZI_COLORS}>Scadenze</SectionLabel>
        <Field label="Assicurazione"><input type="date" style={inputStyle} value={f.assicurazione} onChange={set('assicurazione')} /></Field>
        <Field label="Revisione"><input type="date" style={inputStyle} value={f.revisione} onChange={set('revisione')} /></Field>
        <Field label="Bollo"><input type="date" style={inputStyle} value={f.bollo} onChange={set('bollo')} /></Field>
        <Field label="Note"><textarea style={{ ...inputStyle, minHeight: 60, resize: 'vertical' }} value={f.note} onChange={set('note')} /></Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button
            disabled={!valid}
            onClick={() => onSave({ ...f, id: f.id || uid(), anno: f.anno ? Number(f.anno) : null, km: f.km ? Number(f.km) : null })}
            style={{ width: '100%', background: valid ? MEZZI_COLORS.primary : '#B7C0C2', color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 6 }}
          >
            Salva veicolo
          </button>
        </div>
      )}
    </>
  );
}

/* ---------- Maintenance ---------- */
function ManutenzioniScreen({ maints, vehicles, onOpen, onAdd, onMenu, onHome }) {
  const [tabM, setTabM] = useState('programmati');
  const vOf = (targa) => vehicles.find(v => v.targa === targa);

  const programmati = [...maints]
    .filter(m => STATI_PROGRAMMATI.includes(m.stato))
    .sort((a, b) => {
      const po = { 'Alta': 0, 'Media': 1, 'Bassa': 2 };
      const pa = po[a.priorita] ?? 1, pb = po[b.priorita] ?? 1;
      if (pa !== pb) return pa - pb;
      return (a.data || '').localeCompare(b.data || '');
    });

  const effettuati = [...maints]
    .filter(m => STATI_EFFETTUATI.includes(m.stato))
    .sort((a, b) => (b.data || '').localeCompare(a.data || ''));

  const lista = tabM === 'programmati' ? programmati : effettuati;

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Interventi" subtitle={`${programmati.length} programmati · ${effettuati.length} effettuati`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />

      {/* Tab switcher */}
      <div style={{ display: 'flex', borderBottom: `1px solid ${MEZZI_COLORS.line}`, background: '#fff' }}>
        {[['programmati', `Programmati (${programmati.length})`], ['effettuati', `Effettuati (${effettuati.length})`]].map(([key, label]) => (
          <button key={key} onClick={() => setTabM(key)}
            style={{ flex: 1, padding: '11px 4px', border: 'none', background: 'none', fontWeight: tabM === key ? 700 : 500, fontSize: 13,
              color: tabM === key ? MEZZI_COLORS.primary : MEZZI_COLORS.muted,
              borderBottom: tabM === key ? `2.5px solid ${MEZZI_COLORS.primary}` : '2.5px solid transparent' }}>
            {label}
          </button>
        ))}
      </div>

      <div style={{ padding: 14 }}>
        {lista.length === 0 && <Empty theme={MEZZI_COLORS} icon={Wrench} text={tabM === 'programmati' ? 'Nessun intervento programmato.' : 'Nessun intervento effettuato ancora.'} />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {lista.map(m => (
            <Card theme={MEZZI_COLORS} key={m.id} onClick={() => onOpen(m)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 4 }}>
                    <span style={{ fontWeight: 700, fontSize: 14.5 }}>{m.tipo}</span>
                    {m.priorita && tabM === 'programmati' && <Pill style={PRIORITA_STYLE[m.priorita] || {}}>{m.priorita}</Pill>}
                  </div>
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
      <FAB onClick={onAdd} label="Intervento" visible />
    </>
  );
}

function MaintForm({ initial, vehicles, maints, params, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { targa: vehicles[0]?.targa || '', data: todayISO(), km: '', tipo: TIPI_MANUTENZIONE[0], descrizione: '', officina: '', costo: '', stato: 'Programmato', priorita: 'Media' });
  const [nuovaTipologia, setNuovaTipologia] = useState(false);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const costo = Number(f.costo) || 0;
  const iva = costo * (params.ivaRate / 100);
  const totale = costo + iva;

  // Tipologie disponibili = default + quelle già usate nei record esistenti (escluse segnalazioni)
  const tipiDisponibili = useMemo(() => {
    const esistenti = (maints || [])
      .map(m => m.tipo)
      .filter(t => t && t !== 'Segnalazione anomalia' && !TIPI_MANUTENZIONE.includes(t));
    return [...TIPI_MANUTENZIONE, ...new Set(esistenti)];
  }, [maints]);

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title={initial ? 'Modifica intervento' : 'Nuovo intervento'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
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
          {!nuovaTipologia ? (
            <select
              style={inputStyle}
              value={tipiDisponibili.includes(f.tipo) ? f.tipo : ''}
              onChange={e => {
                if (e.target.value === '__nuova__') {
                  setNuovaTipologia(true);
                  setF({ ...f, tipo: '' });
                } else {
                  set('tipo')(e);
                }
              }}
            >
              {tipiDisponibili.map(t => <option key={t} value={t}>{t}</option>)}
              {isAdmin && <option value="__nuova__">➕ Aggiungi nuova tipologia…</option>}
            </select>
          ) : (
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input
                autoFocus
                style={{ ...inputStyle, flex: 1 }}
                value={f.tipo}
                onChange={set('tipo')}
                placeholder="Nuova tipologia…"
              />
              <button type="button" onClick={() => {
                if (!f.tipo.trim()) setF({ ...f, tipo: TIPI_MANUTENZIONE[0] });
                setNuovaTipologia(false);
              }} style={{ background: MEZZI_COLORS.bg, border: `1.5px solid ${MEZZI_COLORS.line}`, borderRadius: 8, padding: '10px 10px', fontSize: 12, color: MEZZI_COLORS.muted, whiteSpace: 'nowrap' }}>
                Annulla
              </button>
            </div>
          )}
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
        {STATI_PROGRAMMATI.includes(f.stato) && (
          <Field label="Priorità">
            <select style={inputStyle} value={f.priorita || 'Media'} onChange={set('priorita')}>
              {PRIORITA_MAINT.map(p => <option key={p}>{p}</option>)}
            </select>
          </Field>
        )}
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button
            onClick={() => onSave({ ...f, id: f.id || uid(), km: f.km ? Number(f.km) : null, costo: f.costo ? Number(f.costo) : null })}
            style={{ width: '100%', background: MEZZI_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}
          >
            Salva intervento
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: MEZZI_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina intervento
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={MEZZI_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
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
      // Scadenze classiche
      [['Assicurazione', v.assicurazione], ['Revisione', v.revisione], ['Bollo', v.bollo]].forEach(([label, d]) => {
        if (!d) return;
        const days = daysUntil(d);
        out.push({ key: v.id + label, v, label, d, days, u: urgencyOf(days, params), isGomme: false });
      });
      // Allarme gomme stagionali
      const alert = alertGomme(v.tipologiaGomme || '');
      if (alert) {
        out.push({
          key: v.id + '_gomme',
          v,
          label: alert.testo,
          d: null,
          days: alert.livello === 'rosso' ? -1 : 0, // rosso = già scaduto, giallo = imminente
          u: alert.livello === 'rosso' ? 'urgent' : 'soon',
          isGomme: true,
          livello: alert.livello,
        });
      }
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
                  {r.isGomme ? (
                    <Pill style={r.livello === 'rosso' ? URGENCY_STYLE.urgent : URGENCY_STYLE.soon}>
                      {r.livello === 'rosso' ? 'Urgente' : 'Consigliato'}
                    </Pill>
                  ) : (
                    <>
                      <Pill style={URGENCY_STYLE[r.u]}>{r.days < 0 ? 'Scaduta' : `${r.days} giorni`}</Pill>
                      <div style={{ fontSize: 12, color: MEZZI_COLORS.muted, marginTop: 6 }}>{fmtDate(r.d)}</div>
                    </>
                  )}
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


function SegnalazioneAnomaliaForm({ vehicles, onSave, onCancel }) {
  const { nomeVisualizzato } = usePermessi();
  const [f, setF] = useState({
    targa: vehicles[0]?.targa || '',
    data: todayISO(),
    tipo: 'Segnalazione anomalia',
    descrizione: '',
    priorita: 'Media',
    stato: 'Programmato',
    officina: '', km: '', costo: '',
    segnalatoDa: nomeVisualizzato,
    foto: [],
  });
  const [fotoNuove, setFotoNuove] = useState([]);
  const [salvataggio, setSalvataggio] = useState(false);
  const [errore, setErrore] = useState('');
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const valid = f.targa && f.descrizione.trim();

  async function handleSave() {
    if (!valid || salvataggio) return;
    setSalvataggio(true); setErrore('');
    const id = uid();
    try {
      const nuovePaths = [];
      for (const file of fotoNuove) nuovePaths.push(await caricaFotoMezzi(id, file));
      onSave({ ...f, id, foto: nuovePaths, km: f.km ? Number(f.km) : null, costo: f.costo ? Number(f.costo) : null });
    } catch (e) {
      setErrore('Errore caricamento foto: ' + (e.message || 'riprova.'));
      setSalvataggio(false);
    }
  }

  return (
    <>
      <TopBar theme={MEZZI_COLORS} title="Segnala anomalia" subtitle="Descrivi il problema riscontrato" onBack={onCancel} backIcon={Home} />
      <div style={{ padding: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 14, fontSize: 12.5, color: MEZZI_COLORS.muted, background: MEZZI_COLORS.bg, borderRadius: 10, padding: '8px 12px' }}>
          <User size={13} /> Segnalato da <b style={{ color: MEZZI_COLORS.ink }}>{f.segnalatoDa}</b>
        </div>
        <Field label="Veicolo *">
          <select style={inputStyle} value={f.targa} onChange={set('targa')}>
            {vehicles.map(v => <option key={v.id} value={v.targa}>{v.marca} {v.modello}{v.targa ? ' · ' + v.targa : ''}</option>)}
          </select>
        </Field>
        <Field label="Data rilevamento"><input type="date" style={inputStyle} value={f.data} onChange={set('data')} /></Field>
        <Field label="Descrizione anomalia *">
          <textarea style={{ ...inputStyle, minHeight: 80, resize: 'vertical' }}
            value={f.descrizione} onChange={set('descrizione')}
            placeholder="Descrivi il problema nel modo più preciso possibile…" />
        </Field>
        <MezziFotoPicker
          fotoEsistenti={[]}
          onRemoveEsistente={() => {}}
          fotoNuove={fotoNuove}
          onAddNuove={(files) => setFotoNuove(prev => [...prev, ...files])}
          onRemoveNuova={(i) => setFotoNuove(prev => prev.filter((_, idx) => idx !== i))}
        />
        <Field label="Priorità">
          <select style={inputStyle} value={f.priorita} onChange={set('priorita')}>
            {PRIORITA_MAINT.map(p => <option key={p}>{p}</option>)}
          </select>
        </Field>
        {errore && <div style={{ background: '#F7DCD9', color: '#A3352A', padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>{errore}</div>}
        <button disabled={!valid || salvataggio} onClick={handleSave}
          style={{ width: '100%', background: valid ? MEZZI_COLORS.primary : '#B7C0C2', color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 8, opacity: salvataggio ? 0.6 : 1 }}>
          {salvataggio ? 'Salvataggio…' : 'Invia segnalazione'}
        </button>
      </div>
    </>
  );
}

/* ---------- Root ---------- */
function MezziModule({ onHome }) {
  const vehiclesT = useSupaTable('vehicles', 'id', SEED_VEHICLES, ['assicurazione', 'revisione', 'bollo']);
  const maintsT = useSupaTable('maints', 'id', SEED_MAINTS, ['data']);
  const vehicles = vehiclesT.rows, maints = maintsT.rows;
  const ready = vehiclesT.ready && maintsT.ready;
  const dataError = vehiclesT.error || maintsT.error;
  const { nomeVisualizzato } = usePermessi();
  const [params] = useState(DEFAULT_PARAMS);
  const [tab, setTab] = useState('veicoli');
  const [view, setView] = useState(LIST_VIEW);
  const [toast, setToast] = useState('');
  const [showMenu, setShowMenu] = useState(false);
  useBackable(showMenu, setShowMenu);

  useBackable(tab, setTab);
  useBackable(view, setView);

  useEffect(() => { setView(LIST_VIEW); }, [tab]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 2400); };

  async function salvaOTorna(azione, msgOk) {
    const { error } = await azione();
    if (error) { flash(error.message); return; }
    flash(msgOk);
    goBack();
  }
  const saveVehicle = (v) => salvaOTorna(() => vehiclesT.save(v), 'Veicolo salvato');
  const deleteVehicle = (v) => salvaOTorna(() => vehiclesT.remove(v), 'Veicolo eliminato');
  const saveMaint = async (m) => {
    const { error } = await maintsT.save(m);
    if (error) { flash(error.message); return; }
    flash(m.tipo === 'Segnalazione anomalia' ? 'Segnalazione inviata' : 'Intervento salvato');
    if (m.tipo === 'Segnalazione anomalia') {
      const v = vehicles.find(x => x.targa === m.targa);
      inviaNotifica({
        tipo: 'anomalia_mezzi',
        titolo: `🚗 Mezzi: anomalia ${v ? v.marca + ' ' + v.modello : m.targa}`,
        corpo: m.descrizione || 'Anomalia segnalata',
        linkId: m.id,
        inviata_da: m.segnalatoDa || nomeVisualizzato,
      });
    }
    goBack();
  };
  const deleteMaint = (m) => salvaOTorna(() => maintsT.remove(m), 'Intervento eliminato');

  if (!ready) {
    return (
      <div style={{ minHeight: '100vh', background: MEZZI_COLORS.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', color: MEZZI_COLORS.muted, flexDirection: 'column', gap: 10, padding: 24, textAlign: 'center' }}>
        <span>Caricamento…</span>
        {dataError && <span style={{ color: MEZZI_COLORS.danger, fontSize: 13, maxWidth: 280 }}>{dataError}</span>}
      </div>
    );
  }
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
    goBack();
    flash('File Excel scaricato');
  };

  let content;
  if (tab === 'veicoli') {
    if (view.name === 'detail') content = <VeicoloDetail vehicle={vehicles.find(v => v.id === view.id)} maints={maints} params={params} onBack={() => goBack()} onEdit={(v) => setView({ name: 'edit', v })} onDelete={deleteVehicle} />;
    else if (view.name === 'add') content = <VehicleForm onSave={saveVehicle} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <VehicleForm initial={view.v} onSave={saveVehicle} onCancel={() => goBack()} />;
    else content = <VeicoliScreen vehicles={vehicles} maints={maints} params={params} onOpen={(v) => setView({ name: 'detail', id: v.id })} onAdd={() => setView({ name: 'add' })} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  } else if (tab === 'manutenzioni') {
    if (view.name === 'add') content = <MaintForm vehicles={vehicles} maints={maints} params={params} onSave={saveMaint} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <MaintForm initial={view.m} vehicles={vehicles} maints={maints} params={params} onSave={saveMaint} onCancel={() => goBack()} onDelete={deleteMaint} />;
    else content = <ManutenzioniScreen maints={maints} vehicles={vehicles} onOpen={(m) => setView({ name: 'edit', m })} onAdd={() => setView({ name: 'add' })} onMenu={() => setShowMenu(true)} onHome={onHome} />;
  } else if (tab === 'segnalazioni') {
    if (view.name === 'add') {
      content = <SegnalazioneAnomaliaForm vehicles={vehicles} onSave={(seg) => saveMaint(seg)} onCancel={() => goBack()} />;
    } else {
      const segnalazioni = [...maints]
        .filter(m => m.tipo === 'Segnalazione anomalia')
        .sort((a, b) => ({ Alta: 0, Media: 1, Bassa: 2 }[a.priorita] ?? 1) - ({ Alta: 0, Media: 1, Bassa: 2 }[b.priorita] ?? 1));
      content = (
        <div style={{ minHeight: '100vh', background: MEZZI_COLORS.bg }}>
          <TopBar theme={MEZZI_COLORS} title="Segnalazioni" subtitle={`${segnalazioni.length} anomalie segnalate`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={() => setShowMenu(true)} />} />
          <div style={{ padding: 14 }}>
            {segnalazioni.length === 0 && <Empty theme={MEZZI_COLORS} icon={AlertTriangle} text="Nessuna segnalazione ancora." />}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
              {segnalazioni.map(m => {
                const v = vehicles.find(x => x.targa === m.targa);
                return (
                  <Card theme={MEZZI_COLORS} key={m.id}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
                      <div style={{ minWidth: 0, flex: 1 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 5 }}>
                          <Pill style={PRIORITA_STYLE[m.priorita] || PRIORITA_STYLE.Media}>{m.priorita || 'Media'}</Pill>
                          <span style={{ fontSize: 11.5, color: MEZZI_COLORS.muted }}>{fmtDate(m.data)}</span>
                        </div>
                        <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 4 }}>{m.descrizione || '—'}</div>
                        <Plate targa={m.targa} size="sm" />
                        {v && <div style={{ fontSize: 11.5, color: MEZZI_COLORS.muted, marginTop: 3 }}>{v.marca} {v.modello}</div>}
                        {m.segnalatoDa && <div style={{ fontSize: 11, color: MEZZI_COLORS.muted, marginTop: 3 }}>Da: {m.segnalatoDa}</div>}
                      </div>
                      <Pill style={{ bg: '#FBEDD2', fg: '#8A5A00' }}>{m.stato}</Pill>
                    </div>
                  </Card>
                );
              })}
            </div>
          </div>
          <FAB onClick={() => setView({ name: 'add' })} label="Segnala" visible />
        </div>
      );
    }
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
      {showMenu && <MenuSheet theme={MEZZI_COLORS} onClose={() => goBack()} onExport={exportToExcel} exportSub="Scarica anagrafica e registro in .xlsx" />}
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
  'Terra': { bg: '#F7D6E8', fg: '#8A2159', dot: '#D93B8C' },
  'Palestra FKT': { bg: '#D9F0E1', fg: '#1F6B45', dot: '#4FAE79' },
  'Primo Piano': { bg: '#E4DEF0', fg: '#4C3B75', dot: '#9A87C9' },
};
const NUCLEI = Object.keys(NUCLEO_COLORS);

const STATO_STYLE = {
  'Disponibile': { bg: '#DCEEE3', fg: '#1F6B45' },
  'In uso': { bg: '#DCE8F0', fg: '#1E4D73' },
  'Manutenzione': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Fuori uso': { bg: '#F7DCD9', fg: '#A3352A' },
  'Scorta': { bg: '#E7E3F5', fg: '#4B3B87' },
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
  ['segnalazioni', AlertTriangle, 'Segnala'],
];

/* ---------- helpers ---------- */
const needsAttention = (w) => COMPONENTI.some(([k]) => w.c[k] === 'Da sistemare' || w.c[k] === 'Mancante');
const attentionCount = (w) => COMPONENTI.filter(([k]) => w.c[k] === 'Da sistemare' || w.c[k] === 'Mancante').length;
const labelOf = (w) => [w.marca, w.modello].filter(Boolean).join(' ') || 'Carrozzina senza marca';
const fmtCarrozzinaId = (id) => String(id).padStart(4, '0');

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
function CarrozzineScreen({ items, onOpen, onMenu, filtro, setFiltro, onHome }) {
  const [q, setQ] = useState('');
  const filtered = items.filter(w => {
    if (filtro?.tipo === 'nucleo' && w.nucleo !== filtro.valore) return false;
    if (filtro?.tipo === 'stato' && w.stato !== filtro.valore) return false;
    if (filtro?.tipo === 'attenzione' && !needsAttention(w)) return false;
    if (filtro?.tipo === 'marca' && w.marca !== filtro.valore) return false;
    const hay = `${w.marca} ${w.modello} ${w.seriale} ${w.ospite} ${w.nucleo} ${w.id} ${fmtCarrozzinaId(w.id)}`.toLowerCase();
    return hay.includes(q.toLowerCase());
  });

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Carrozzine" subtitle={`${items.length} in archivio`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ position: 'relative', marginBottom: 10 }}>
          <Search size={16} style={{ position: 'absolute', left: 12, top: 12, color: CARROZZINE_COLORS.muted }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Cerca numero, marca, modello, ospite…"
            style={{ width: '100%', boxSizing: 'border-box', padding: '11px 12px 11px 34px', borderRadius: 10, border: `1.5px solid ${CARROZZINE_COLORS.line}`, fontSize: 15, background: '#FCFBF7', outline: 'none' }} />
        </div>
        {filtro && (
          <div style={{ marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, color: CARROZZINE_COLORS.muted }}>Filtro:</span>
            {filtro.tipo === 'nucleo' && <NucleoTag nucleo={filtro.valore} />}
            {filtro.tipo === 'stato' && <Pill style={STATO_STYLE[filtro.valore] || {}}>{filtro.valore}</Pill>}
            {filtro.tipo === 'attenzione' && <Pill style={{ bg: '#F7DCD9', fg: '#A3352A' }}>Da controllare</Pill>}
            {filtro.tipo === 'marca' && <Pill style={{ bg: '#E8EEF5', fg: '#1E4D73' }}>{filtro.valore}</Pill>}
            <button onClick={() => setFiltro(null)} style={{ background: 'none', border: 'none', color: CARROZZINE_COLORS.muted, fontSize: 12, textDecoration: 'underline' }}>rimuovi</button>
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
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: CARROZZINE_COLORS.muted }}>#{fmtCarrozzinaId(w.id)}</span>
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
  const { isAdmin } = usePermessi();
  const puoScrivere = isAdmin;
  const set = (patch) => onUpdate({ ...w, ...patch });
  const setC = (key, val) => onUpdate({ ...w, c: { ...w.c, [key]: val } });

  const inputInlineStyle = {
    width: '100%', boxSizing: 'border-box', padding: '5px 8px', borderRadius: 7,
    border: `1.5px solid ${CARROZZINE_COLORS.line}`, fontSize: 13, background: '#FCFBF7',
    fontFamily: 'Inter, sans-serif', outline: 'none', color: CARROZZINE_COLORS.ink,
  };

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title={labelOf(w)} subtitle={`ID ${fmtCarrozzinaId(w.id)}${w.seriale ? ' · ' + w.seriale : ''}`} onBack={onBack} />
      <div style={{ padding: 14, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <Card theme={CARROZZINE_COLORS} style={{ marginBottom: 12 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
            <InfoRow theme={CARROZZINE_COLORS} icon={Calendar} label="Data" value={fmtDate(w.data)} />
            <InfoRow theme={CARROZZINE_COLORS} icon={Hash} label="Tipologia" value={w.tipologia} />
            <div>
              <div style={{ fontSize: 10, color: CARROZZINE_COLORS.muted, textTransform: 'uppercase', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 4 }}>
                <MapPin size={11} /> Stanza
              </div>
              <input value={w.stanza || ''} onChange={e => set({ stanza: e.target.value })} placeholder="—" style={inputInlineStyle} />
            </div>
            <div>
              <div style={{ fontSize: 10, color: CARROZZINE_COLORS.muted, textTransform: 'uppercase', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 4 }}>
                <User size={11} /> Ospite
              </div>
              <input value={w.ospite || ''} onChange={e => set({ ospite: e.target.value })} placeholder="—" style={inputInlineStyle} />
            </div>
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
            const dataSost = w.dataSostituzioni?.[key] || '';
            return (
              <div key={key} style={{ padding: '9px 0', borderTop: i ? `1px solid ${CARROZZINE_COLORS.line}` : 'none' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                  <span style={{ fontSize: 13.5, fontWeight: 600, flexShrink: 0 }}>{label}</span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                    {isFreeText && <span style={{ fontSize: 11, color: CARROZZINE_COLORS.muted, fontStyle: 'italic', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 110 }}>{val}</span>}
                    <select
                      value={CONDIZIONI.includes(val) ? val : ''}
                      onChange={e => {
                        const nuovoValore = e.target.value;
                        const nuovoDsSost = { ...(w.dataSostituzioni || {}) };
                        if (nuovoValore !== 'Sostituito') delete nuovoDsSost[key];
                        onUpdate({ ...w, c: { ...w.c, [key]: nuovoValore }, dataSostituzioni: nuovoDsSost });
                      }}
                      style={{ ...selectStyle(COND_STYLE[val] || { bg: '#F1EFE6', fg: CARROZZINE_COLORS.muted }), width: 118 }}
                    >
                      <option value="">{isFreeText ? 'nota storica' : '—'}</option>
                      {CONDIZIONI.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
                {val === 'Sostituito' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 6, paddingLeft: 4 }}>
                    <Calendar size={12} color={CARROZZINE_COLORS.muted} />
                    <span style={{ fontSize: 11.5, color: CARROZZINE_COLORS.muted, flexShrink: 0 }}>Data sostituzione:</span>
                    <input
                      type="date"
                      value={dataSost}
                      onChange={e => onUpdate({ ...w, dataSostituzioni: { ...(w.dataSostituzioni || {}), [key]: e.target.value } })}
                      style={{ fontSize: 12, border: `1px solid ${CARROZZINE_COLORS.line}`, borderRadius: 7, padding: '3px 7px', background: '#FCFBF7', color: CARROZZINE_COLORS.ink, outline: 'none' }}
                    />
                  </div>
                )}
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
function NucleiScreen({ items, onFilter, onMenu, onHome, onRinominaEnd }) {
  const { isAdmin } = usePermessi();
  const [rinomina, setRinomina] = useState(null); // nome da rinominare
  const [busy, setBusy] = useState(false);

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

  async function handleRinomina(vecchioNome, nuovoNome) {
    setBusy(true);
    await supabase.from('carrozzine').update({ nucleo: nuovoNome }).eq('nucleo', vecchioNome);
    setRinomina(null); setBusy(false);
    if (onRinominaEnd) onRinominaEnd();
  }

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Nuclei" subtitle={`${groups.length} nuclei attivi`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {groups.map(([nucleo, g]) => (
            <Card theme={CARROZZINE_COLORS} key={nucleo} onClick={() => onFilter(nucleo)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <NucleoTag nucleo={nucleo === '—' ? null : nucleo} />
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  {isAdmin && nucleo !== '—' && (
                    <button
                      onClick={e => { e.stopPropagation(); setRinomina(nucleo); }}
                      style={{ background: 'none', border: 'none', padding: 4, color: CARROZZINE_COLORS.muted }}
                      title="Rinomina nucleo"
                    >
                      ✎
                    </button>
                  )}
                  <span style={{ fontWeight: 800, fontFamily: "'Archivo', sans-serif", fontSize: 16 }}>{g.tot}</span>
                </div>
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
      {rinomina && (
        <RenameDialog
          theme={CARROZZINE_COLORS}
          label="nucleo"
          currentName={rinomina}
          onSave={nuovoNome => handleRinomina(rinomina, nuovoNome)}
          onCancel={() => setRinomina(null)}
        />
      )}
      {busy && <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.3)', zIndex:100, display:'flex', alignItems:'center', justifyContent:'center', color:'#fff', fontSize:14 }}>Aggiornamento…</div>}
    </>
  );
}

/* ---------- Riepilogo ---------- */
function RiepilogoScreen({ items, onMenu, onHome, onFilterStato, onFilterAttenzione, onFilterMarca }) {
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
          <div onClick={() => attCount > 0 && onFilterAttenzione()} style={{ cursor: attCount > 0 ? 'pointer' : 'default' }}>
            <StatCard theme={CARROZZINE_COLORS} label={attCount > 0 ? 'Da controllare →' : 'Da controllare'} value={attCount} accent={attCount ? CARROZZINE_COLORS.danger : CARROZZINE_COLORS.ok} />
          </div>
        </div>

        <SectionLabel theme={CARROZZINE_COLORS}>Per stato</SectionLabel>
        <Card theme={CARROZZINE_COLORS} style={{ marginBottom: 4 }}>
          {perStato.map(([s, n], i) => (
            <div
              key={s}
              onClick={() => n > 0 && onFilterStato(s)}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderTop: i ? `1px solid ${CARROZZINE_COLORS.line}` : 'none', cursor: n > 0 ? 'pointer' : 'default' }}
            >
              <Pill style={STATO_STYLE[s]}>{s}</Pill>
              <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{n}</span>
                {n > 0 && <ChevronRight size={14} color={CARROZZINE_COLORS.muted} />}
              </span>
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
            <div key={m} onClick={() => onFilterMarca(m)} style={{ marginTop: i ? 10 : 0, cursor: 'pointer' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 12.5, marginBottom: 4 }}>
                <span style={{ fontWeight: 600 }}>{m}</span>
                <span style={{ display: 'flex', alignItems: 'center', gap: 3, color: CARROZZINE_COLORS.muted }}>
                  {n} <ChevronRight size={13} />
                </span>
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

function SegnalazioneCarrozzinaForm({ items, onSave, onCancel }) {
  const { nomeVisualizzato } = usePermessi();
  const [f, setF] = useState({ carrozzinaId: items[0]?.id || '', descrizione: '', priorita: 'Media' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  return (
    <>
      <TopBar theme={CARROZZINE_COLORS} title="Segnala problema" onBack={onCancel} backIcon={Home} />
      <div style={{ padding: 14 }}>
        <Field label="Carrozzina">
          <select value={f.carrozzinaId} onChange={set('carrozzinaId')} style={{ ...inputStyle, background: '#FCFBF7' }}>
            {items.map(w => (
              <option key={w.id} value={w.id}>{labelOf(w)} ({fmtCarrozzinaId(w.id)})</option>
            ))}
          </select>
        </Field>
        <Field label="Descrizione problema">
          <textarea value={f.descrizione} onChange={set('descrizione')} rows={4} placeholder="Descrivi il problema..."
            style={{ ...inputStyle, resize: 'vertical', minHeight: 80 }} />
        </Field>
        <Field label="Priorita">
          <select value={f.priorita} onChange={set('priorita')} style={{ ...inputStyle, background: '#FCFBF7' }}>
            {['Alta', 'Media', 'Bassa'].map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </Field>
        <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
          <button onClick={onCancel} style={{ flex: 1, padding: 12, borderRadius: 10, border: `1.5px solid ${CARROZZINE_COLORS.line}`, background: '#fff', fontWeight: 600 }}>Annulla</button>
          <button onClick={() => { if (!f.descrizione.trim()) return; onSave({ carrozzinaId: f.carrozzinaId, descrizione: f.descrizione, priorita: f.priorita, segnalatoDa: nomeVisualizzato }); }} style={{ flex: 1, padding: 12, borderRadius: 10, border: 'none', background: CARROZZINE_COLORS.primary, color: '#fff', fontWeight: 700 }}>Invia segnalazione</button>
        </div>
      </div>
    </>
  );
}

function CarrozzineModule({ onHome }) {
  const { isAdmin } = usePermessi();
  const itemsT = useSupaTable('carrozzine', 'id', SEED, ['data']);
  const items = itemsT.rows;
  const ready = itemsT.ready;
  const dataError = itemsT.error;
  const [tab, setTab] = useState('carrozzine');
  const [openId, setOpenId] = useState(null);
  const [filtro, setFiltro] = useState(null); // { tipo: 'nucleo'|'stato', valore }
  const [showMenu, setShowMenu] = useState(false);
  const [view, setView] = useState(LIST_VIEW);
  const interventiT = useSupaTable('interventi', 'id', [], ['dataSegnalazione', 'dataChiusura']);
  const segnalazioniCarrozzine = interventiT.rows.filter(i => i.tipologia === 'carrozzina');
  const segnalazioniAperte = segnalazioniCarrozzine.filter(i => i.stato !== 'Chiuso');
  useBackable(showMenu, setShowMenu);
  useBackable(view, setView);
  const [toast, setToast] = useState('');

  useBackable(tab, setTab);
  useBackable(openId, setOpenId);

  useEffect(() => { setOpenId(null); }, [tab]);
  useEffect(() => { setView(LIST_VIEW); }, [tab]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 2400); };

  const updateItem = async (w) => {
    const { error } = await itemsT.save(w);
    if (error) flash(error.message);
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
    goBack();
    flash('File Excel scaricato');
  };

  if (!ready) {
    return (
      <div style={{ minHeight: '100vh', background: CARROZZINE_COLORS.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', color: CARROZZINE_COLORS.muted, flexDirection: 'column', gap: 10, padding: 24, textAlign: 'center' }}>
        <span>Caricamento…</span>
        {dataError && <span style={{ color: CARROZZINE_COLORS.danger, fontSize: 13, maxWidth: 280 }}>{dataError}</span>}
      </div>
    );
  }

  const openItem = openId != null ? items.find(w => w.id === openId) : null;
  const onMenu = () => setShowMenu(true);

  let content;
  if (openItem) {
    content = <CarrozzinaDetail w={openItem} onBack={() => goBack()} onUpdate={updateItem} />;
  } else if (tab === 'carrozzine') {
    content = <CarrozzineScreen items={items} onOpen={(w) => setOpenId(w.id)} onMenu={onMenu} filtro={filtro} setFiltro={setFiltro} onHome={onHome} />;
  } else if (tab === 'controlli') {
    content = <ControlliScreen items={items} onOpen={(w) => setOpenId(w.id)} onMenu={onMenu} onHome={onHome} />;
  } else if (tab === 'nuclei') {
    content = <NucleiScreen items={items} onMenu={onMenu} onFilter={(n) => { setFiltro(n === '—' ? null : { tipo: 'nucleo', valore: n }); setTab('carrozzine'); }} onHome={onHome} onRinominaEnd={() => itemsT.reload ? itemsT.reload() : window.location.reload()} />;
  } else if (tab === 'segnalazioni') {
    if (view.name === 'add') {
      content = <SegnalazioneCarrozzinaForm
        items={items}
        onSave={async (seg) => {
          const id = uid();
          const { error } = await interventiT.save({
            id,
            dataSegnalazione: todayISO(),
            cameraZona: '',
            descrizione: seg.descrizione,
            priorita: seg.priorita,
            tecnico: '',
            stato: 'Programmato',
            dataChiusura: null,
            costo: null,
            note: '',
            segnalatoDa: seg.segnalatoDa,
            foto: seg.foto || [],
            tipologia: 'carrozzina',
            carrozzinaId: seg.carrozzinaId,
          });
          if (error) {
            setToast('Errore: ' + error.message);
            return;
          }
          inviaNotifica({
            tipo: 'segnalazione_carrozzine',
            titolo: 'Problema carrozzina ' + labelOf(items.find(w => w.id === seg.carrozzinaId) || {}),
            corpo: seg.descrizione,
            linkId: id,
            inviata_da: seg.segnalatoDa,
          });
          setToast('Segnalazione inviata');
          goBack();
        }}
        onCancel={() => goBack()}
      />;
    } else if (view.name === 'detail' && view.segnalazione) {
      const s = view.segnalazione;
      const w = items.find(x => x.id === s.carrozzinaId);
      content = (
        <>
          <TopBar theme={CARROZZINE_COLORS} title="Segnalazione" onBack={() => goBack()} backIcon={Home} />
          <div style={{ padding: 14 }}>
            <Card theme={CARROZZINE_COLORS}>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>{s.descrizione || '—'}</div>
              <div style={{ fontSize: 12.5, color: CARROZZINE_COLORS.muted, marginBottom: 6 }}>
                {fmtDate(s.dataSegnalazione)} · Stato: <b>{s.stato}</b>
              </div>
              {w && <div style={{ fontSize: 12.5, color: CARROZZINE_COLORS.muted, marginBottom: 10 }}>Carrozzina: {labelOf(w)} ({fmtCarrozzinaId(w.id)})</div>}
              {s.segnalatoDa && <div style={{ fontSize: 11.5, color: CARROZZINE_COLORS.muted, marginBottom: 10 }}>Da: {s.segnalatoDa}</div>}
              {s.foto?.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 12 }}>
                  {s.foto.map((p, i) => (
                    <div key={i} style={{ width: 64, height: 64, borderRadius: 8, overflow: 'hidden', background: CARROZZINE_COLORS.bg, border: `1px solid ${CARROZZINE_COLORS.line}` }}>
                      <img src={p} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                    </div>
                  ))}
                </div>
              )}
              {s.stato !== 'Chiuso' && (
                <button onClick={async () => {
                  const { error } = await interventiT.save({ ...s, stato: 'Chiuso', dataChiusura: todayISO() });
                  if (error) { setToast('Errore: ' + error.message); return; }
                  setToast('Segnalazione chiusa');
                  goBack();
                }} style={{ width: '100%', padding: 12, borderRadius: 10, border: 'none', background: CARROZZINE_COLORS.ok, color: '#fff', fontWeight: 700 }}>
                  Chiudi segnalazione
                </button>
              )}
              {isAdmin && (
                <button onClick={async () => {
                  if (!window.confirm('Eliminare definitivamente questa segnalazione?')) return;
                  const { error } = await supabase.from('interventi').delete().eq('id', s.id);
                  if (error) { setToast('Errore: ' + error.message); return; }
                  setToast('Segnalazione eliminata');
                  goBack();
                }} style={{ width: '100%', padding: 12, borderRadius: 10, border: 'none', background: CARROZZINE_COLORS.danger, color: '#fff', fontWeight: 700, marginTop: 10 }}>
                  Elimina segnalazione
                </button>
              )}
            </Card>
          </div>
        </>
      );
    } else {
      content = (
        <>
          <TopBar theme={CARROZZINE_COLORS} title='Segnalazioni' subtitle={`${segnalazioniCarrozzine.length} segnalazioni · ${segnalazioniAperte.length} aperte`} onMenu={onMenu} onBack={onHome} backIcon={Home} />
          <div style={{ padding: 14 }}>
            {segnalazioniCarrozzine.length === 0 && <Empty theme={CARROZZINE_COLORS} icon={AlertTriangle} text="Nessuna segnalazione ancora." />}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
              {segnalazioniCarrozzine.map(i => {
                const w = items.find(x => x.id === i.carrozzinaId);
                const isAperta = i.stato !== 'Chiuso';
                return (
                  <Card theme={CARROZZINE_COLORS} key={i.id} onClick={() => setView({ name: 'detail', segnalazione: i })}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
                      <div style={{ minWidth: 0, flex: 1 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 5 }}>
                          <Pill style={PRIORITA_STYLE[i.priorita] || PRIORITA_STYLE.Media}>{i.priorita || 'Media'}</Pill>
                          <span style={{ fontSize: 11.5, color: CARROZZINE_COLORS.muted }}>{fmtDate(i.dataSegnalazione)}</span>
                          {!isAperta && <Pill style={{ bg: '#E7F3EC', fg: '#3C7A5C' }}>Chiuso</Pill>}
                        </div>
                        <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 4 }}>{i.descrizione || '—'}</div>
                        {w && <div style={{ fontSize: 11.5, color: CARROZZINE_COLORS.muted, marginTop: 3 }}>{labelOf(w)} ({fmtCarrozzinaId(w.id)})</div>}
                        {i.segnalatoDa && <div style={{ fontSize: 11, color: CARROZZINE_COLORS.muted, marginTop: 3 }}>Da: {i.segnalatoDa}</div>}
                        {i.foto?.length > 0 && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, fontSize: 11, color: CARROZZINE_COLORS.muted, marginTop: 4 }}>
                            <ImageIcon size={12} /> {i.foto.length} foto
                          </span>
                        )}
                      </div>
                      {isAperta && <div style={{ width: 8, height: 8, borderRadius: 999, background: CARROZZINE_COLORS.danger, flexShrink: 0 }} />}
                    </div>
                  </Card>
                );
              })}
            </div>
          </div>
          <FAB onClick={() => setView({ name: 'add' })} label="Segnala" visible />
        </>
      );
    }

  } else {
    content = <RiepilogoScreen items={items} onMenu={onMenu} onHome={onHome}
      onFilterStato={(s) => { setFiltro({ tipo: 'stato', valore: s }); setTab('carrozzine'); }}
      onFilterAttenzione={() => { setFiltro({ tipo: 'attenzione' }); setTab('carrozzine'); }}
      onFilterMarca={(m) => { setFiltro({ tipo: 'marca', valore: m }); setTab('carrozzine'); }}
    />;
  }

  return (
    <div style={{ minHeight: '100vh', background: CARROZZINE_COLORS.bg, fontFamily: 'Inter, sans-serif', color: CARROZZINE_COLORS.ink, maxWidth: 480, margin: '0 auto', position: 'relative' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        select { cursor: pointer; }
      `}</style>
      <div style={{ paddingBottom: 78 }}>{content}</div>
      {!openItem && <BottomNav theme={CARROZZINE_COLORS} tab={tab} setTab={setTab} items={CARROZZINE_NAV_ITEMS} badges={{ segnalazioni: segnalazioniAperte.length }} />}
      {showMenu && <MenuSheet theme={CARROZZINE_COLORS} onClose={() => goBack()} onExport={exportToExcel} exportSub="Scarica il foglio Totale in .xlsx" />}
      {toast && (
        <div style={{ position: 'fixed', bottom: !openItem ? 92 : 20, left: '50%', transform: 'translateX(-50%)', background: CARROZZINE_COLORS.primaryDeep, color: '#fff', padding: '10px 18px', borderRadius: 999, fontSize: 13, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 7, zIndex: 30, maxWidth: 400 }}>
          <Check size={15} /> {toast}
        </div>
      )}
    </div>
  );
}
/* =========================================================================
   STRUTTURA — camere, reparti, tecnici, interventi, manutenzioni, costi
   ========================================================================= */
const STR_COLORS = {
  bg: '#F3EEE4',
  surface: '#FFFFFF',
  ink: '#2A241E',
  muted: '#79715F',
  line: '#E4DCC9',
  primary: '#6E4A2E',
  primaryDeep: '#432C1A',
  amber: '#BE8A2E',
  danger: '#B23A2E',
  ok: '#3C7A5C',
  info: '#2C5C82',
};

const STR_NUCLEO_COLORS = {
  'Ala Nord': { bg: '#E4DEF0', fg: '#4C3B75', dot: '#9A87C9' },
  'Ala Sud': { bg: '#F6E7C4', fg: '#7A5A18', dot: '#C99B3E' },
  'Bianco': { bg: '#F1F0EC', fg: '#4A4A46', dot: '#C7C4B8' },
  'Blu': { bg: '#D2DEEF', fg: '#1E3D66', dot: '#3363A8' },
  'Grigio': { bg: '#DEDEDE', fg: '#3A3A3A', dot: '#6B6B6B' },
  'Rosso': { bg: '#F9D9D6', fg: '#8A2A20', dot: '#D8392C' },
  'Giallo': { bg: '#FBEBBE', fg: '#7A5A00', dot: '#EFB81C' },
  'Verde': { bg: '#CFE9D1', fg: '#215A2A', dot: '#3EA24C' },
};

const STR_STATO_CAMERA_STYLE = {
  'Attiva': { bg: '#DCEEE3', fg: '#1F6B45' },
  'In Manutenzione': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Fuori Servizio': { bg: '#F7DCD9', fg: '#A3352A' },
};
const STR_STATI_CAMERA = Object.keys(STR_STATO_CAMERA_STYLE);
const STR_TIPI_CAMERA = ['Singola', 'Doppia'];

const STR_CATEGORIE_REPARTO = ['Assistenziale', 'Amministrativo', 'Servizi'];
const STR_PIANI_REPARTO = ['Piano Terra', 'Primo Piano', 'Secondo Piano', 'Piano Interrato', 'Esterno'];
const STR_TIPI_TECNICO = ['Interno', 'Esterno', 'Fornitore'];

const STR_PRIORITA_STYLE = {
  'Bassa': { bg: '#DCEEE3', fg: '#1F6B45' },
  'Media': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Alta': { bg: '#F7DCD9', fg: '#A3352A' },
  'Urgente': { bg: '#F7DCD9', fg: '#A3352A' },
};
const STR_PRIORITA_LIST = Object.keys(STR_PRIORITA_STYLE);

const STR_TIPOLOGIE_INTERVENTO = [
  'Elettrico', 'Idraulico', 'Muratura', 'Falegnameria', 'Serramenti',
  'Climatizzazione', 'Verde/Giardino', 'Pulizie', 'Informatica', 'Altro',
];

const STR_STATO_INTERVENTO_STYLE = {
  'Aperto': { bg: '#F7DCD9', fg: '#A3352A' },
  'In corso': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Chiuso': { bg: '#DCEEE3', fg: '#1F6B45' },
  'Annullato': { bg: '#E8E5DC', fg: '#79715F' },
};
const STR_STATI_INTERVENTO = Object.keys(STR_STATO_INTERVENTO_STYLE);

const STR_TIPI_MANUTENZIONE = ['Ordinaria', 'Straordinaria', 'Preventiva', 'Correttiva', 'Verifica Normativa'];
const STR_FREQUENZE = ['Settimanale', 'Mensile', 'Trimestrale', 'Semestrale', 'Annuale', 'Una Tantum'];

const STR_ALERT_STYLE = {
  'SCADUTO': { bg: '#F7DCD9', fg: '#A3352A' },
  'IN SCADENZA': { bg: '#FBEDD2', fg: '#8A5A00' },
  'OK': { bg: '#DCEEE3', fg: '#1F6B45' },
};

const STR_TIPI_COSTO = ['Preventivo', 'Fattura', 'Costo Effettivo'];
const STR_STATO_PAGAMENTO_STYLE = {
  'Da pagare': { bg: '#FBEDD2', fg: '#8A5A00' },
  'Pagato': { bg: '#DCEEE3', fg: '#1F6B45' },
  'Parzialmente pagato': { bg: '#FBEDD2', fg: '#8A5A00' },
  'In contestazione': { bg: '#F7DCD9', fg: '#A3352A' },
};
const STR_STATI_PAGAMENTO = Object.keys(STR_STATO_PAGAMENTO_STYLE);

const STR_NAV_ITEMS = [
  ['camere', BedDouble, 'Camere'],
  ['interventi', ClipboardList, 'Interventi'],
  ['scadenze', CalendarClock, 'Scadenze'],
  ['reparti', Building2, 'Reparti'],
  ['riepilogo', BarChart3, 'Riepilogo'],
];

/* ---------- helpers ---------- */
const strAlertStatus = (days) => { if (days == null) return 'OK'; if (days < 0) return 'SCADUTO'; if (days <= 7) return 'IN SCADENZA'; return 'OK'; };

/* ---------- small UI atoms (tema Struttura) ---------- */
function STR_Field({ label, children }) {
  return (
    <label style={{ display: 'block', marginBottom: 14 }}>
      <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: STR_COLORS.muted, marginBottom: 5, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{label}</span>
      {children}
    </label>
  );
}
const strInputStyle = {
  width: '100%', boxSizing: 'border-box', padding: '11px 12px', borderRadius: 10,
  border: `1.5px solid ${STR_COLORS.line}`, fontSize: 15, fontFamily: 'Inter, sans-serif',
  background: '#FCFBF7', color: STR_COLORS.ink, outline: 'none',
};
function STR_FAB({ onClick, label, visible }) {
  const { isAdmin } = usePermessi();
  if (!isAdmin && !visible) return null;
  return (
    <button
      onClick={onClick}
      style={{
        position: 'fixed', right: 18, bottom: 84, background: STR_COLORS.amber, color: '#2A1B00',
        border: 'none', borderRadius: 999, height: 52, padding: '0 20px 0 16px',
        display: 'flex', alignItems: 'center', gap: 7, fontWeight: 700, fontSize: 14.5,
        boxShadow: '0 6px 16px rgba(190,138,46,0.4)', zIndex: 20,
      }}
    >
      <Plus size={20} strokeWidth={2.6} /> {label}
    </button>
  );
}
function STR_NucleoTag({ nucleo }) {
  if (!nucleo) return null;
  const s = STR_NUCLEO_COLORS[nucleo] || { bg: '#EEE', fg: STR_COLORS.muted, dot: '#999' };
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: s.bg, color: s.fg, fontSize: 11, fontWeight: 700, padding: '3px 9px 3px 7px', borderRadius: 999 }}>
      <span style={{ width: 7, height: 7, borderRadius: 999, background: s.dot, flexShrink: 0 }} />
      {nucleo}
    </span>
  );
}
function STR_Row({ label, value, bold }) {
  return (
    <div style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0', fontWeight: bold ? 700 : 500, fontSize: bold ? 14 : 13 }}>
      <span style={{ color: bold ? STR_COLORS.ink : STR_COLORS.muted }}>{label}</span><span>{value}</span>
    </div>
  );
}

/* ---------- dati di esempio (dal registro reale della struttura) ---------- */
const S_CAMERE = [{"codice": "001", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "002", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "003", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "004", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "005", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "006", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "007", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "008", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "009", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "010", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "011", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "012", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "013", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "014", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "015", "piano": "Piano Terra", "nucleo": "Ala Nord", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "016", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "017", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "018", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "019", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "020", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "021", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "022", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "023", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "024", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "025", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "026", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "027", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Doppia", "stato": "Fuori Servizio", "note": ""}, {"codice": "028", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "029", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "030", "piano": "Piano Terra", "nucleo": "Ala Sud", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "101", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "102", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "103", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "104", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "105", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "106", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "107", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "108", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "109", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "110", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "111", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "112", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "113", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "114", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "115", "piano": "Primo Piano", "nucleo": "Bianco", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "116", "piano": "Primo Piano", "nucleo": "Blu", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "117", "piano": "Primo Piano", "nucleo": "Blu", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "118", "piano": "Primo Piano", "nucleo": "Blu", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "119", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "120", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "121", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "122", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "123", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "124", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "125", "piano": "Primo Piano", "nucleo": "Grigio", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "126", "piano": "Primo Piano", "nucleo": "Rosso", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "127", "piano": "Primo Piano", "nucleo": "Rosso", "tipo": "Doppia", "stato": "Fuori Servizio", "note": ""}, {"codice": "128", "piano": "Primo Piano", "nucleo": "Rosso", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "129", "piano": "Primo Piano", "nucleo": "Rosso", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "130", "piano": "Primo Piano", "nucleo": "Rosso", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "201", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "202", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "203", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "204", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "205", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "206", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "207", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "208", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "209", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "210", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "211", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "212", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "213", "piano": "Secondo Piano", "nucleo": "Giallo", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "214", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "215", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "216", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "217", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "218", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "219", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "220", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "221", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "222", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "223", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "224", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Attiva", "note": ""}, {"codice": "225", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "226", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "In Manutenzione", "note": ""}, {"codice": "227", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Fuori Servizio", "note": ""}, {"codice": "228", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "229", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Singola", "stato": "Attiva", "note": ""}, {"codice": "230", "piano": "Secondo Piano", "nucleo": "Verde", "tipo": "Doppia", "stato": "Attiva", "note": ""}];
const S_REPARTI = [{"codice": "Z01", "nome": "Cucina", "categoria": "Servizi", "responsabile": "Responsabile Cucina", "note": ""}, {"codice": "Z02", "nome": "Sala da Pranzo", "categoria": "Servizi", "responsabile": "Responsabile Cucina", "note": ""}, {"codice": "Z03", "nome": "Uffici Amministrativi", "categoria": "Amministrativo", "responsabile": "Responsabile Amministrativo", "note": ""}, {"codice": "Z04", "nome": "Ufficio Direzione", "categoria": "Amministrativo", "responsabile": "Direttore di Struttura", "note": ""}, {"codice": "Z05", "nome": "Ufficio Coord. Infermieristico", "categoria": "Amministrativo", "responsabile": "Coordinatore Infermieristico", "note": ""}, {"codice": "Z06", "nome": "Lavanderia", "categoria": "Servizi", "responsabile": "Responsabile Lavanderia", "note": ""}, {"codice": "Z07", "nome": "Guardaroba", "categoria": "Servizi", "responsabile": "Responsabile Lavanderia", "note": ""}, {"codice": "Z08", "nome": "Bagno Comune Piano Terra", "categoria": "Assistenziale", "responsabile": "Coordinatore Assistenziale", "note": ""}, {"codice": "Z09", "nome": "Bagno Comune Primo Piano", "categoria": "Assistenziale", "responsabile": "Coordinatore Assistenziale", "note": ""}, {"codice": "Z10", "nome": "Bagno Comune Secondo Piano", "categoria": "Assistenziale", "responsabile": "Coordinatore Assistenziale", "note": ""}, {"codice": "Z11", "nome": "Parco Giochi", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z12", "nome": "Giardino animali", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z13", "nome": "Giardino d'Inverno", "categoria": "Assistenziale", "responsabile": "Animatore/Educatore", "note": ""}, {"codice": "Z14", "nome": "Palestra / Fisioterapia", "categoria": "Assistenziale", "responsabile": "Responsabile Fisioterapia", "note": ""}, {"codice": "Z15", "nome": "Cappella", "categoria": "Servizi", "responsabile": "Responsabile Struttura", "note": ""}, {"codice": "Z16", "nome": "Ingresso / Reception", "categoria": "Amministrativo", "responsabile": "Responsabile Reception", "note": ""}, {"codice": "Z17", "nome": "Ascensore Nord", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z18", "nome": "Ascensore Sud", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z19", "nome": "Ascensore Secondario", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z20", "nome": "Scala di Sicurezza Nord", "categoria": "Servizi", "responsabile": "Responsabile Sicurezza", "note": ""}, {"codice": "Z21", "nome": "Scala di Sicurezza Sud", "categoria": "Servizi", "responsabile": "Responsabile Sicurezza", "note": ""}, {"codice": "Z22", "nome": "Magazzino Generale", "categoria": "Servizi", "responsabile": "Simone Cabrelle", "note": ""}, {"codice": "Z23", "nome": "Centrale Termica", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z24", "nome": "Locale Tecnico Impianti Elettrici", "categoria": "Servizi", "responsabile": "Responsabile Manutenzione", "note": ""}, {"codice": "Z25", "nome": "Parcheggio", "categoria": "Servizi", "responsabile": "Responsabile Sicurezza", "note": ""}];
const S_TECNICI = [{"id": "T01", "nome": "Biagio", "tipo": "Interno", "specializzazione": "Manutentore Generico", "telefono": "329 000 0001", "email": "m.rossi@struttura-esempio.it", "note": ""}, {"id": "T02", "nome": "Simone Napolitano", "tipo": "Interno", "specializzazione": "Idraulico", "telefono": "329 000 0002", "email": "g.verdi@struttura-esempio.it", "note": ""}, {"id": "T03", "nome": "Raffaele", "tipo": "Interno", "specializzazione": "Manutentore Generico", "telefono": "329 000 0003", "email": "l.bianchi@struttura-esempio.it", "note": ""}, {"id": "T04", "nome": "ElettroService Snc", "tipo": "Esterno", "specializzazione": "Impianti Elettrici", "telefono": "0332 000 100", "email": "info@elettroservice-esempio.it", "note": ""}, {"id": "T05", "nome": "Idro Pronto Srl", "tipo": "Esterno", "specializzazione": "Impianti Idraulici", "telefono": "0332 000 101", "email": "info@idropronto-esempio.it", "note": ""}, {"id": "T06", "nome": "ClimaTech Impianti", "tipo": "Esterno", "specializzazione": "Climatizzazione / HVAC", "telefono": "0332 000 102", "email": "info@climatech-esempio.it", "note": ""}, {"id": "T07", "nome": "Ascensori Sicuri Srl", "tipo": "Esterno", "specializzazione": "Manutenzione Ascensori", "telefono": "0332 000 103", "email": "info@ascensorisicuri-esempio.it", "note": ""}, {"id": "T08", "nome": "AntincendioPiu Srl", "tipo": "Esterno", "specializzazione": "Sistemi Antincendio", "telefono": "0332 000 104", "email": "info@antincendiopiu-esempio.it", "note": ""}, {"id": "T09", "nome": "Falegnameria Colombo", "tipo": "Esterno", "specializzazione": "Falegnameria", "telefono": "0332 000 105", "email": "info@falegnameriacolombo-esempio.it", "note": ""}, {"id": "T10", "nome": "Verde Giardini Sas", "tipo": "Esterno", "specializzazione": "Manutenzione Giardino", "telefono": "0332 000 106", "email": "info@verdegiardini-esempio.it", "note": ""}, {"id": "T11", "nome": "Pulizie Splendor Srl", "tipo": "Esterno", "specializzazione": "Pulizie e Sanificazione", "telefono": "0332 000 107", "email": "info@puliziesplendor-esempio.it", "note": ""}, {"id": "T12", "nome": "DisinfestaService", "tipo": "Esterno", "specializzazione": "Disinfestazione / Derattizzazione", "telefono": "0332 000 108", "email": "info@disinfestaservice-esempio.it", "note": ""}, {"id": "T13", "nome": "InfoTech Solutions", "tipo": "Esterno", "specializzazione": "Assistenza Informatica", "telefono": "0332 000 109", "email": "info@infotech-esempio.it", "note": ""}, {"id": "T14", "nome": "Sicurtech Impianti", "tipo": "Esterno", "specializzazione": "Sicurezza / Videosorveglianza", "telefono": "0332 000 110", "email": "info@sicurtech-esempio.it", "note": ""}, {"id": "T15", "nome": "Edil Ripristino Srl", "tipo": "Esterno", "specializzazione": "Opere Edili / Muratura", "telefono": "0332 000 111", "email": "info@edilripristino-esempio.it", "note": ""}];
const S_INTERVENTI = [{"id": "SI001", "dataSegnalazione": "2026-04-03", "cameraZona": "003", "descrizione": "Perdita rubinetto bagno camera", "priorita": "Media", "tecnico": "Biagio", "stato": "Chiuso", "dataChiusura": "2026-04-04", "costo": 45.0, "note": ""}, {"id": "SI002", "dataSegnalazione": "2026-04-10", "cameraZona": "Cucina", "descrizione": "Guasto abbattitore di temperatura", "priorita": "Alta", "tecnico": "ClimaTech Impianti", "stato": "Chiuso", "dataChiusura": "2026-04-12", "costo": 320.0, "note": ""}, {"id": "SI003", "dataSegnalazione": "2026-04-15", "cameraZona": "112", "descrizione": "Presa elettrica non funzionante", "priorita": "Media", "tecnico": "Simone Napolitano", "stato": "Chiuso", "dataChiusura": "2026-04-15", "costo": 0.0, "note": ""}, {"id": "SI004", "dataSegnalazione": "2026-04-22", "cameraZona": "Giardino", "descrizione": "Irrigazione automatica bloccata", "priorita": "Bassa", "tecnico": "Verde Giardini Sas", "stato": "Chiuso", "dataChiusura": "2026-04-28", "costo": 150.0, "note": ""}];
const S_MANUTENZIONI = [{"id": "MP001", "cameraZona": "Ascensore Principale", "tipoManutenzione": "Verifica Normativa", "frequenza": "Mensile", "ultimaEsecuzione": "2026-06-19", "prossimaScadenza": "2026-07-10", "tecnico": "Ascensori Sicuri Srl", "note": ""}, {"id": "MP002", "cameraZona": "Struttura (tutti i piani)", "tipoManutenzione": "Verifica Normativa", "frequenza": "Trimestrale", "ultimaEsecuzione": "2026-04-20", "prossimaScadenza": "2026-07-20", "tecnico": "AntincendioPiu Srl", "note": ""}, {"id": "MP003", "cameraZona": "Cucina", "tipoManutenzione": "Verifica Normativa", "frequenza": "Semestrale", "ultimaEsecuzione": "2026-02-05", "prossimaScadenza": "2026-08-05", "tecnico": "AntincendioPiu Srl", "note": ""}];
const S_COSTI = [{"id": "C001", "idIntervento": "SI001", "tipo": "Fattura", "descrizione": "Rifacimento Bagno", "fornitore": "Giuseppe Verdi", "numeroDocumento": "FT-2026-041", "data": "2026-04-04", "importo": 45.0, "statoPagamento": "Pagato", "note": ""}, {"id": "C002", "idIntervento": "SI002", "tipo": "Fattura", "descrizione": "Riparazione abbattitore cucina", "fornitore": "ClimaTech Impianti", "numeroDocumento": "FT-2026-088", "data": "2026-04-12", "importo": 320.0, "statoPagamento": "Pagato", "note": ""}];

/* ---------- Camere ---------- */
function CamereScreen({ camere, onOpen, onAdd, onMenu, onHome, filtroStatoIniziale }) {
  const [q, setQ] = useState('');
  const [filtroPiano, setFiltroPiano] = useState('');
  const [filtroNucleo, setFiltroNucleo] = useState('');
  const [filtroStato, setFiltroStato] = useState(filtroStatoIniziale || '');

  // Tutti i piani disponibili
  const piani = useMemo(() => [...new Set(camere.map(c => c.piano).filter(Boolean))].sort(), [camere]);

  // I nuclei disponibili cambiano in base al piano selezionato
  const nucleiDisponibili = useMemo(() => {
    const base = filtroPiano ? camere.filter(c => c.piano === filtroPiano) : camere;
    return [...new Set(base.map(c => c.nucleo).filter(Boolean))].sort();
  }, [camere, filtroPiano]);

  // Mappa nucleo -> piano (ogni nucleo sta su un solo piano)
  const nucleoPiano = useMemo(() => {
    const map = {};
    camere.forEach(c => { if (c.nucleo && c.piano) map[c.nucleo] = c.piano; });
    return map;
  }, [camere]);

  function selezionaPiano(piano) {
    if (filtroPiano === piano) {
      // Deseleziona piano: rimuovi anche il nucleo se non è compatibile
      setFiltroPiano('');
    } else {
      setFiltroPiano(piano);
      // Se il nucleo corrente non appartiene al nuovo piano, deselezionalo
      if (filtroNucleo && nucleoPiano[filtroNucleo] !== piano) {
        setFiltroNucleo('');
      }
    }
  }

  function selezionaNucleo(nucleo) {
    if (filtroNucleo === nucleo) {
      setFiltroNucleo('');
    } else {
      setFiltroNucleo(nucleo);
      // Auto-seleziona il piano corrispondente
      const piano = nucleoPiano[nucleo];
      if (piano) setFiltroPiano(piano);
    }
  }

  const filtered = camere.filter(c => {
    if (filtroPiano && c.piano !== filtroPiano) return false;
    if (filtroNucleo && c.nucleo !== filtroNucleo) return false;
    if (filtroStato && c.stato !== filtroStato) return false;
    return `${c.codice} ${c.piano} ${c.nucleo}`.toLowerCase().includes(q.toLowerCase());
  });

  const pillBtn = (label, active, onPress) => (
    <button
      key={label}
      onClick={onPress}
      style={{ border: 'none', borderRadius: 999, padding: '5px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0,
        background: active ? STR_COLORS.primary : STR_COLORS.bg,
        color: active ? '#fff' : STR_COLORS.muted,
        boxShadow: active ? `0 2px 8px ${STR_COLORS.primary}44` : 'none',
      }}
    >{label}</button>
  );

  return (
    <>
      <TopBar theme={STR_COLORS} title="Camere" subtitle={`${filtered.length} di ${camere.length}`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />

      {/* Filtro Piano */}
      <div style={{ padding: '8px 14px 0', borderBottom: `1px solid ${STR_COLORS.line}` }}>
        <div style={{ fontSize: 10, fontWeight: 700, color: STR_COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 6 }}>Piano</div>
        <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 10 }}>
          {piani.map(p => pillBtn(p, filtroPiano === p, () => selezionaPiano(p)))}
        </div>
      </div>

      {/* Filtro Nucleo — aggiornato dinamicamente in base al piano */}
      <div style={{ padding: '8px 14px 0', borderBottom: `1px solid ${STR_COLORS.line}` }}>
        <div style={{ fontSize: 10, fontWeight: 700, color: STR_COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 6 }}>Nucleo</div>
        <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 10 }}>
          {nucleiDisponibili.map(n => pillBtn(n, filtroNucleo === n, () => selezionaNucleo(n)))}
        </div>
      </div>

      <div style={{ padding: 14 }}>
        <input placeholder="Cerca camera, piano, nucleo…" style={{ ...strInputStyle, marginBottom: 10 }} value={q} onChange={(e) => setQ(e.target.value)} />

        {/* Filtro Stato */}
        <div style={{ display: 'flex', gap: 7, marginBottom: 12, flexWrap: 'wrap' }}>
          {STR_STATI_CAMERA.map(s => (
            <button key={s} onClick={() => setFiltroStato(filtroStato === s ? '' : s)}
              style={{ border: 'none', borderRadius: 999, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                background: filtroStato === s ? STR_COLORS.primary : STR_STATO_CAMERA_STYLE[s].bg,
                color: filtroStato === s ? '#fff' : STR_STATO_CAMERA_STYLE[s].fg }}>
              {s}
            </button>
          ))}
        </div>

        {filtered.length === 0 && <Empty theme={STR_COLORS} icon={BedDouble} text="Nessuna camera trovata." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {filtered.map(c => (
            <Card theme={STR_COLORS} key={c.codice} onClick={() => onOpen(c)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Camera {c.codice}</div>
                  <div style={{ fontSize: 12, color: STR_COLORS.muted, marginBottom: 6 }}>{c.piano} · {c.tipo}</div>
                  <STR_NucleoTag nucleo={c.nucleo} />
                </div>
                <Pill style={STR_STATO_CAMERA_STYLE[c.stato] || {}}>{c.stato}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Camera" />
    </>
  );
}

function CameraDetail({ camera, interventi, onBack, onEdit, onOpenIntervento }) {
  const { isAdmin } = usePermessi();
  const puoScrivere = isAdmin;
  const own = interventi.filter(i => i.cameraZona === camera.codice).sort((a, b) => (b.dataSegnalazione || '').localeCompare(a.dataSegnalazione || ''));
  return (
    <>
      <TopBar theme={STR_COLORS} title={`Camera ${camera.codice}`} subtitle={camera.piano} onBack={onBack} />
      <div style={{ padding: 14 }}>
        <Card theme={STR_COLORS} style={{ marginBottom: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
            <STR_NucleoTag nucleo={camera.nucleo} />
            <Pill style={STR_STATO_CAMERA_STYLE[camera.stato] || {}}>{camera.stato}</Pill>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <InfoRow theme={STR_COLORS} icon={Building2} label="Piano" value={camera.piano} />
            <InfoRow theme={STR_COLORS} icon={BedDouble} label="Tipo" value={camera.tipo} />
          </div>
          {camera.note && <div style={{ marginTop: 10, fontSize: 13, color: STR_COLORS.muted }}>{camera.note}</div>}
          {puoScrivere && (
            <button onClick={() => onEdit(camera)} style={{ width: '100%', background: STR_COLORS.bg, border: `1px solid ${STR_COLORS.line}`, color: STR_COLORS.ink, borderRadius: 10, padding: '10px', fontWeight: 700, fontSize: 13.5, marginTop: 12 }}>
              Modifica camera
            </button>
          )}
        </Card>

        <SectionLabel theme={STR_COLORS}>Interventi in questa camera · {own.length}</SectionLabel>
        {own.length === 0 && <Empty theme={STR_COLORS} icon={ClipboardList} text="Nessun intervento registrato per questa camera." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {own.map(i => (
            <Card theme={STR_COLORS} key={i.id} onClick={() => onOpenIntervento(i)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 4 }}>{i.descrizione || '—'}</div>
                  <div style={{ fontSize: 11.5, color: STR_COLORS.muted }}>{fmtDate(i.dataSegnalazione)}{i.tecnico ? ' · ' + i.tecnico : ''}</div>
                </div>
                <Pill style={STR_STATO_INTERVENTO_STYLE[i.stato] || {}}>{i.stato}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
    </>
  );
}

function CameraForm({ initial, piani, nuclei, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { codice: '', piano: piani[0] || 'Piano Terra', nucleo: nuclei[0] || '', tipo: 'Singola', stato: 'Attiva', note: '' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica camera' : 'Nuova camera'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <STR_Field label="Codice camera *">
          <input style={strInputStyle} value={f.codice} onChange={set('codice')} />
        </STR_Field>
        {initial && (
          <div style={{ fontSize: 11.5, color: STR_COLORS.muted, marginTop: -8, marginBottom: 14 }}>
            Cambiando il codice, gli interventi e le scadenze gia' collegati a questa camera verranno aggiornati automaticamente.
          </div>
        )}
        <STR_Field label="Piano"><input list="str-piani" style={strInputStyle} value={f.piano} onChange={set('piano')} />
          <datalist id="str-piani">{piani.map(p => <option key={p} value={p} />)}</datalist>
        </STR_Field>
        <STR_Field label="Nucleo"><input list="str-nuclei" style={strInputStyle} value={f.nucleo} onChange={set('nucleo')} />
          <datalist id="str-nuclei">{nuclei.map(n => <option key={n} value={n} />)}</datalist>
        </STR_Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <STR_Field label="Tipo"><select style={strInputStyle} value={f.tipo} onChange={set('tipo')}>{STR_TIPI_CAMERA.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
          <STR_Field label="Stato"><select style={selectStyle(STR_STATO_CAMERA_STYLE[f.stato] || {})} value={f.stato} onChange={set('stato')}>{STR_STATI_CAMERA.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
        </div>
        <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button onClick={() => onSave(f)} style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}>
            Salva camera
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina camera
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          message={`Eliminare la camera ${initial?.codice}? Verranno eliminati anche gli interventi e le scadenze collegati.`}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Interventi ---------- */
function InterventiScreen({ interventi, onOpen, onAdd, onMenu, onHome, filtroStatoIniziale }) {
  const [filtro, setFiltro] = useState(filtroStatoIniziale || null);
  const [filtroTipologia, setFiltroTipologia] = useState('');
  const sorted = [...interventi]
    .filter(i => i.tipologia !== 'carrozzina')
    .filter(i => !filtro || i.stato === filtro)
    .filter(i => !filtroTipologia || i.tipologia === filtroTipologia)
    .sort((a, b) => (b.dataSegnalazione || '').localeCompare(a.dataSegnalazione || ''));
  return (
    <>
      <TopBar theme={STR_COLORS} title="Interventi" subtitle={`${interventi.length} segnalazioni`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'flex', gap: 7, marginBottom: 10, flexWrap: 'wrap' }}>
          {STR_STATI_INTERVENTO.map(s => (
            <button key={s} onClick={() => setFiltro(filtro === s ? null : s)}
              style={{ border: 'none', borderRadius: 999, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                background: filtro === s ? STR_COLORS.primary : STR_STATO_INTERVENTO_STYLE[s].bg,
                color: filtro === s ? '#fff' : STR_STATO_INTERVENTO_STYLE[s].fg }}>
              {s}
            </button>
          ))}
        </div>
        <select
          value={filtroTipologia}
          onChange={(e) => setFiltroTipologia(e.target.value)}
          style={{ ...strInputStyle, marginBottom: 12, width: 'auto', minWidth: 180, fontSize: 13, padding: '7px 10px' }}
        >
          <option value="">Tutte le tipologie</option>
          {STR_TIPOLOGIE_INTERVENTO.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        {sorted.length === 0 && <Empty theme={STR_COLORS} icon={ClipboardList} text="Nessun intervento ancora. Aggiungine uno." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {sorted.map(i => (
            <Card theme={STR_COLORS} key={i.id} onClick={() => onOpen(i)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 4 }}>{i.descrizione || '—'}</div>
                  <div style={{ fontSize: 12, color: STR_COLORS.ink, marginBottom: 6 }}>{i.cameraZona} · {fmtDate(i.dataSegnalazione)}{i.tipologia ? ' · ' + i.tipologia : ''}</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <Pill style={STR_PRIORITA_STYLE[i.priorita] || {}}>{i.priorita}</Pill>
                    {i.foto?.length > 0 && (
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, fontSize: 11, color: STR_COLORS.muted }}>
                        <ImageIcon size={12} /> {i.foto.length}
                      </span>
                    )}
                  </div>
                </div>
                <div style={{ textAlign: 'right', flexShrink: 0, marginLeft: 8 }}>
                  <Pill style={STR_STATO_INTERVENTO_STYLE[i.stato] || {}}>{i.stato}</Pill>
                  {i.costo > 0 && <div style={{ fontSize: 12, fontWeight: 700, marginTop: 6 }}>{fmtEuro(i.costo)}</div>}
                </div>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Intervento" visible />
    </>
  );
}

const FOTO_BUCKET = 'segnalazioni';

async function caricaFoto(interventoId, file) {
  const estensione = (file.name.split('.').pop() || 'jpg').toLowerCase().replace(/[^a-z0-9]/g, '') || 'jpg';
  const path = `${interventoId}/${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${estensione}`;
  const { error } = await supabase.storage.from(FOTO_BUCKET).upload(path, file, { upsert: false });
  if (error) throw error;
  return path;
}
async function eliminaFoto(paths) {
  if (!paths || paths.length === 0) return;
  try { await supabase.storage.from(FOTO_BUCKET).remove(paths); } catch (e) { /* non bloccante */ }
}
async function urlFirmateFoto(paths) {
  if (!paths || paths.length === 0) return {};
  const { data, error } = await supabase.storage.from(FOTO_BUCKET).createSignedUrls(paths, 3600);
  if (error || !data) return {};
  const map = {};
  data.forEach((d) => { if (d.signedUrl && d.path) map[d.path] = d.signedUrl; });
  return map;
}

function FotoPicker({ fotoEsistenti, onRemoveEsistente, fotoNuove, onAddNuove, onRemoveNuova, disabled }) {
  const [urls, setUrls] = useState({});
  const [zoom, setZoom] = useState(null);
  const fileInputRef = useRef(null);
  const chiave = fotoEsistenti.join(',');

  useEffect(() => {
    let mounted = true;
    urlFirmateFoto(fotoEsistenti).then((m) => { if (mounted) setUrls(m); });
    return () => { mounted = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chiave]);

  const anteprimeNuove = useMemo(() => fotoNuove.map((f) => URL.createObjectURL(f)), [fotoNuove]);
  useEffect(() => {
    return () => { anteprimeNuove.forEach(url => URL.revokeObjectURL(url)); };
  }, [anteprimeNuove]);

  const thumbStyle = { position: 'relative', width: 72, height: 72, borderRadius: 10, overflow: 'hidden', background: STR_COLORS.bg, border: `1px solid ${STR_COLORS.line}`, flexShrink: 0, cursor: 'pointer' };
  const xStyle = { position: 'absolute', top: 2, right: 2, background: 'rgba(0,0,0,0.55)', border: 'none', borderRadius: 999, width: 20, height: 20, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center' };

  return (
    <div style={{ marginBottom: 14 }}>
      <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: STR_COLORS.muted, marginBottom: 8, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Fotografie</span>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
        {fotoEsistenti.map((p) => (
          <div key={p} style={thumbStyle} onClick={() => urls[p] && setZoom(urls[p])}>
            {urls[p] && <img src={urls[p]} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />}
            {!disabled && <button type="button" onClick={(e) => { e.stopPropagation(); onRemoveEsistente(p); }} style={xStyle}><XIcon size={12} /></button>}
          </div>
        ))}
        {fotoNuove.map((file, i) => (
          <div key={i} style={thumbStyle} onClick={() => setZoom(anteprimeNuove[i])}>
            <img src={anteprimeNuove[i]} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            {!disabled && <button type="button" onClick={(e) => { e.stopPropagation(); onRemoveNuova(i); }} style={xStyle}><XIcon size={12} /></button>}
          </div>
        ))}
        {!disabled && (
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            style={{ width: 72, height: 72, borderRadius: 10, border: `1.5px dashed ${STR_COLORS.line}`, background: 'none', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: STR_COLORS.muted, gap: 3, flexShrink: 0 }}
          >
            <Plus size={18} />
            <span style={{ fontSize: 10 }}>Aggiungi</span>
          </button>
        )}
        <input
          ref={fileInputRef} type="file" accept="image/*" multiple
          style={{ display: 'none' }}
          onChange={(e) => { onAddNuove(Array.from(e.target.files || [])); e.target.value = ''; }}
        />
      </div>
      {zoom && (
        <div
          onClick={() => setZoom(null)}
          style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.9)', zIndex: 100, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}
        >
          <button
            onClick={() => setZoom(null)}
            style={{ position: 'absolute', top: 16, right: 16, background: 'rgba(255,255,255,0.15)', border: 'none', borderRadius: 999, width: 38, height: 38, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
          >
            <XIcon size={18} />
          </button>
          <img src={zoom} alt="" style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 6 }} onClick={(e) => e.stopPropagation()} />
        </div>
      )}
    </div>
  );
}

function InterventoForm({ initial, luoghi, camere, reparti, tecnici, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare, nomeVisualizzato } = usePermessi();
  const puoScrivere = !initial || isAdmin || initial.tecnico === nomeVisualizzato;

  // Inizializza i selettori a cascata dal valore esistente
  const [f, setF] = useState(initial || {
    dataSegnalazione: todayISO(), cameraZona: '', descrizione: '', priorita: 'Media',
    tecnico: tecnici[0] || '', stato: 'Aperto', dataChiusura: '', costo: '', note: '',
    segnalatoDa: nomeVisualizzato,
  });

  // Stato cascata
  const [pianoSel, setPianoSel] = useState(initial?.pianoSel || '');
  const [tipoSel, setTipoSel] = useState(initial?.tipoSel || '');   // 'Camere' | categoria reparto
  const [nucleoSel, setNucleoSel] = useState(initial?.nucleoSel || '');
  const [subSel, setSubSel] = useState(initial?.cameraZona || '');

  const [fotoEsistenti, setFotoEsistenti] = useState(initial?.foto || []);
  const [fotoDaRimuovere, setFotoDaRimuovere] = useState([]);
  const [fotoNuove, setFotoNuove] = useState([]);
  const [salvataggio, setSalvataggio] = useState(false);
  const [erroreFoto, setErroreFoto] = useState('');
  const [confermaElimina, setConfermaElimina] = useState(false);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  // ---- Dati per la cascata ----
  const pianiDisponibili = useMemo(() => {
    const fromCamere = camere.map(c => c.piano).filter(Boolean);
    const fromReparti = reparti.map(r => r.piano).filter(Boolean);
    return [...new Set([...fromCamere, ...fromReparti])].sort();
  }, [camere, reparti]);

  const tipiPerPiano = useMemo(() => {
    if (!pianoSel) return [];
    const tipi = [];
    const haCamere = camere.some(c => c.piano === pianoSel);
    if (haCamere) tipi.push('Camere');
    const categorie = [...new Set(reparti.filter(r => r.piano === pianoSel && r.categoria).map(r => r.categoria))].sort();
    return [...tipi, ...categorie];
  }, [camere, reparti, pianoSel]);

  const nucleiPerPiano = useMemo(() => {
    if (!pianoSel || tipoSel !== 'Camere') return [];
    return [...new Set(camere.filter(c => c.piano === pianoSel && c.nucleo).map(c => c.nucleo))].sort();
  }, [camere, pianoSel, tipoSel]);

  const camerePerNucleo = useMemo(() => {
    if (!pianoSel || tipoSel !== 'Camere' || !nucleoSel) return [];
    return camere.filter(c => c.piano === pianoSel && c.nucleo === nucleoSel).sort((a, b) => a.codice.localeCompare(b.codice));
  }, [camere, pianoSel, tipoSel, nucleoSel]);

  const repartiPerTipo = useMemo(() => {
    if (!pianoSel || tipoSel === 'Camere' || !tipoSel) return [];
    return reparti.filter(r => r.piano === pianoSel && r.categoria === tipoSel).sort((a, b) => a.nome.localeCompare(b.nome));
  }, [reparti, pianoSel, tipoSel]);

  // Aggiorna cameraZona (il campo salvato) quando la selezione finale cambia
  useEffect(() => {
    if (!subSel) return;
    setF(prev => ({ ...prev, cameraZona: subSel }));
  }, [subSel]);

  function handlePiano(piano) {
    setPianoSel(piano); setTipoSel(''); setNucleoSel(''); setSubSel('');
    setF(prev => ({ ...prev, cameraZona: '' }));
  }
  function handleTipo(tipo) {
    setTipoSel(tipo); setNucleoSel(''); setSubSel('');
    setF(prev => ({ ...prev, cameraZona: '' }));
  }
  function handleNucleo(nucleo) {
    setNucleoSel(nucleo); setSubSel('');
    setF(prev => ({ ...prev, cameraZona: '' }));
  }

  const selectPill = (label, value, options, onChange, placeholder = '— seleziona —') => (
    <STR_Field label={label}>
      <select
        style={strInputStyle}
        value={value}
        onChange={e => onChange(e.target.value)}
        disabled={options.length === 0}
      >
        <option value="">{options.length === 0 ? '—' : placeholder}</option>
        {options.map(o => (
          typeof o === 'string'
            ? <option key={o} value={o}>{o}</option>
            : <option key={o.codice || o.id} value={o.label || o.nome}>{o.label || o.nome}</option>
        ))}
      </select>
    </STR_Field>
  );

  async function handleSave() {
    if (salvataggio) return;
    setErroreFoto('');
    setSalvataggio(true);
    const id = f.id || uid();
    try {
      if (fotoDaRimuovere.length) await eliminaFoto(fotoDaRimuovere);
      const nuovePaths = [];
      for (const file of fotoNuove) nuovePaths.push(await caricaFoto(id, file));
      const fotoFinali = [...fotoEsistenti, ...nuovePaths];
      const recordPulito = { ...f, id, costo: f.costo ? Number(f.costo) : 0, foto: fotoFinali };
      delete recordPulito.pianoSel;
      delete recordPulito.tipoSel;
      delete recordPulito.nucleoSel;
      onSave(recordPulito);
    } catch (e) {
      setErroreFoto('Caricamento foto non riuscito: ' + (e.message || 'riprova.'));
      setSalvataggio(false);
    }
  }

  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica intervento' : 'Nuovo intervento'} onBack={onCancel} />
      <div style={{ padding: 16 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 7, marginBottom: 14, fontSize: 12.5, color: STR_COLORS.muted }}>
          <User size={13} /> Segnalato da <b style={{ color: STR_COLORS.ink }}>{f.segnalatoDa || nomeVisualizzato}</b>
        </div>
        <div style={{ pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>

          {/* ---- Selettore a cascata ---- */}
          <div style={{ background: STR_COLORS.bg, borderRadius: 12, padding: '12px 14px', marginBottom: 14, border: `1px solid ${STR_COLORS.line}` }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: STR_COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 10 }}>Posizione intervento *</div>

            {/* 1. Piano */}
            {selectPill('Piano', pianoSel, pianiDisponibili, handlePiano, '— seleziona piano —')}

            {/* 2. Tipo area */}
            {pianoSel && selectPill('Tipo area', tipoSel, tipiPerPiano, handleTipo, '— seleziona tipo —')}

            {/* 3a. Se Camere: Nucleo → Camera */}
            {tipoSel === 'Camere' && (
              <>
                {selectPill('Nucleo', nucleoSel, nucleiPerPiano, handleNucleo, '— seleziona nucleo —')}
                {nucleoSel && selectPill(
                  'Camera',
                  subSel,
                  camerePerNucleo.map(c => ({ codice: c.codice, label: `Camera ${c.codice}` })),
                  setSubSel,
                  '— seleziona camera —'
                )}
              </>
            )}

            {/* 3b. Se reparto/zona: scelta diretta */}
            {tipoSel && tipoSel !== 'Camere' && (
              selectPill(
                tipoSel,
                subSel,
                repartiPerTipo.map(r => ({ codice: r.codice, label: r.nome })),
                setSubSel,
                `— seleziona ${tipoSel.toLowerCase()} —`
              )
            )}

            {/* Riepilogo selezione */}
            {f.cameraZona && (
              <div style={{ marginTop: 6, fontSize: 12.5, color: STR_COLORS.primary, fontWeight: 600 }}>
                ✓ {f.cameraZona}{pianoSel ? ` · ${pianoSel}` : ''}{tipoSel && tipoSel !== 'Camere' ? ` · ${tipoSel}` : ''}{nucleoSel ? ` · ${nucleoSel}` : ''}
              </div>
            )}

            {/* Fallback per interventi già esistenti con cameraZona libera */}
            {!pianoSel && initial?.cameraZona && (
              <div style={{ fontSize: 12.5, color: STR_COLORS.muted, marginTop: 4 }}>
                Posizione attuale: <b>{initial.cameraZona}</b>
              </div>
            )}
          </div>

          <STR_Field label="Data segnalazione"><input type="date" style={strInputStyle} value={f.dataSegnalazione} onChange={set('dataSegnalazione')} /></STR_Field>
          <STR_Field label="Descrizione problema"><input style={strInputStyle} value={f.descrizione} onChange={set('descrizione')} placeholder="Cosa è successo" /></STR_Field>
          <FotoPicker
            fotoEsistenti={fotoEsistenti}
            onRemoveEsistente={(p) => { setFotoEsistenti(prev => prev.filter(x => x !== p)); setFotoDaRimuovere(prev => [...prev, p]); }}
            fotoNuove={fotoNuove}
            onAddNuove={(files) => setFotoNuove(prev => [...prev, ...files])}
            onRemoveNuova={(i) => setFotoNuove(prev => prev.filter((_, idx) => idx !== i))}
            disabled={!puoScrivere}
          />
          <STR_Field label="Tipologia"><select style={strInputStyle} value={f.tipologia || ''} onChange={set('tipologia')}><option value="">—</option>{STR_TIPOLOGIE_INTERVENTO.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <STR_Field label="Priorità"><select style={selectStyle(STR_PRIORITA_STYLE[f.priorita] || {})} value={f.priorita} onChange={set('priorita')}>{STR_PRIORITA_LIST.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
            <STR_Field label="Stato"><select style={selectStyle(STR_STATO_INTERVENTO_STYLE[f.stato] || {})} value={f.stato} onChange={set('stato')}>{STR_STATI_INTERVENTO.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
          </div>
          <STR_Field label="Tecnico / Ditta assegnato">
            <input list="str-tecnici" style={strInputStyle} value={f.tecnico} onChange={set('tecnico')} />
            <datalist id="str-tecnici">{tecnici.map(t => <option key={t} value={t} />)}</datalist>
          </STR_Field>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <STR_Field label="Data chiusura"><input type="date" style={strInputStyle} value={f.dataChiusura || ''} onChange={set('dataChiusura')} /></STR_Field>
            <STR_Field label="Costo (€)"><input type="number" step="0.01" style={strInputStyle} value={f.costo} onChange={set('costo')} placeholder="0.00" /></STR_Field>
          </div>
          <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
        </div>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          {erroreFoto && <div style={{ background: STR_COLORS.redBg, color: STR_COLORS.red, padding: '10px 12px', borderRadius: 10, fontSize: 12.5, marginBottom: 12 }}>{erroreFoto}</div>}
          <button
            onClick={handleSave}
            disabled={salvataggio}
            style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4, opacity: salvataggio ? 0.6 : 1 }}
          >
            {salvataggio ? 'Salvataggio…' : 'Salva intervento'}
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina intervento
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Scadenze (manutenzioni programmate) ---------- */
function ScadenzeStrScreen({ manutenzioni, onOpen, onAdd, onMenu, onHome }) {
  const rows = useMemo(() => {
    return manutenzioni.map(m => ({ ...m, days: daysUntil(m.prossimaScadenza) })).sort((a, b) => (a.days ?? 9999) - (b.days ?? 9999));
  }, [manutenzioni]);
  return (
    <>
      <TopBar theme={STR_COLORS} title="Scadenze" subtitle={`${manutenzioni.length} manutenzioni programmate`} onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        {rows.length === 0 && <Empty theme={STR_COLORS} icon={CalendarClock} text="Nessuna manutenzione programmata." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {rows.map(m => {
            const stato = strAlertStatus(m.days);
            return (
              <Card theme={STR_COLORS} key={m.id} onClick={() => onOpen(m)}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 3 }}>{m.cameraZona}</div>
                    <div style={{ fontSize: 12.5, color: STR_COLORS.muted, marginBottom: 6 }}>{m.tipoManutenzione} · {m.frequenza}</div>
                    <div style={{ fontSize: 11.5, color: STR_COLORS.muted }}>{m.tecnico}</div>
                  </div>
                  <div style={{ textAlign: 'right', flexShrink: 0 }}>
                    <Pill style={STR_ALERT_STYLE[stato]}>{m.days < 0 ? `${Math.abs(m.days)} gg fa` : stato === 'OK' ? `tra ${m.days} gg` : `tra ${m.days} gg`}</Pill>
                    <div style={{ fontSize: 12, color: STR_COLORS.muted, marginTop: 6 }}>{fmtDate(m.prossimaScadenza)}</div>
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Scadenza" />
    </>
  );
}

function ManutenzioneForm({ initial, luoghi, tecnici, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || {
    cameraZona: luoghi[0] || '', tipoManutenzione: STR_TIPI_MANUTENZIONE[0], frequenza: STR_FREQUENZE[1],
    ultimaEsecuzione: todayISO(), prossimaScadenza: todayISO(), tecnico: tecnici[0] || '', note: '',
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica scadenza' : 'Nuova scadenza'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <STR_Field label="Camera / Zona *">
          <input list="str-luoghi2" style={strInputStyle} value={f.cameraZona} onChange={set('cameraZona')} />
          <datalist id="str-luoghi2">{luoghi.map(l => <option key={l} value={l} />)}</datalist>
        </STR_Field>
        <STR_Field label="Tipo manutenzione"><select style={strInputStyle} value={f.tipoManutenzione} onChange={set('tipoManutenzione')}>{STR_TIPI_MANUTENZIONE.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
        <STR_Field label="Frequenza"><select style={strInputStyle} value={f.frequenza} onChange={set('frequenza')}>{STR_FREQUENZE.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <STR_Field label="Ultima esecuzione"><input type="date" style={strInputStyle} value={f.ultimaEsecuzione} onChange={set('ultimaEsecuzione')} /></STR_Field>
          <STR_Field label="Prossima scadenza"><input type="date" style={strInputStyle} value={f.prossimaScadenza} onChange={set('prossimaScadenza')} /></STR_Field>
        </div>
        <STR_Field label="Tecnico assegnato">
          <input list="str-tecnici2" style={strInputStyle} value={f.tecnico} onChange={set('tecnico')} />
          <datalist id="str-tecnici2">{tecnici.map(t => <option key={t} value={t} />)}</datalist>
        </STR_Field>
        <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button onClick={() => onSave({ ...f, id: f.id || uid() })} style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}>
            Salva scadenza
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina scadenza
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Costi ---------- */
function CostiStrScreen({ costi, onOpen, onAdd, onMenu, onHome, onBack }) {
  const [filtro, setFiltro] = useState(null);
  const sorted = [...costi].filter(c => !filtro || c.statoPagamento === filtro).sort((a, b) => (b.data || '').localeCompare(a.data || ''));
  const totale = costi.reduce((s, c) => s + (Number(c.importo) || 0), 0);
  return (
    <>
      <TopBar theme={STR_COLORS} title="Costi" subtitle={`${costi.length} voci · ${fmtEuro(totale)}`} onBack={onBack || onHome} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'flex', gap: 7, marginBottom: 12, flexWrap: 'wrap' }}>
          {STR_STATI_PAGAMENTO.map(s => (
            <button key={s} onClick={() => setFiltro(filtro === s ? null : s)}
              style={{ border: 'none', borderRadius: 999, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer',
                background: filtro === s ? STR_COLORS.primary : STR_STATO_PAGAMENTO_STYLE[s].bg,
                color: filtro === s ? '#fff' : STR_STATO_PAGAMENTO_STYLE[s].fg }}>
              {s}
            </button>
          ))}
        </div>
        {sorted.length === 0 && <Empty theme={STR_COLORS} icon={Wallet} text="Nessun costo registrato." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {sorted.map(c => (
            <Card theme={STR_COLORS} key={c.id} onClick={() => onOpen(c)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{c.descrizione || c.tipo}</div>
                  <div style={{ fontSize: 12, color: STR_COLORS.muted }}>{c.fornitore}{c.fornitore ? ' · ' : ''}{fmtDate(c.data)}</div>
                </div>
                <div style={{ textAlign: 'right', flexShrink: 0, marginLeft: 8 }}>
                  <div style={{ fontWeight: 700, fontSize: 14 }}>{fmtEuro(c.importo)}</div>
                  <div style={{ marginTop: 6 }}><Pill style={STR_STATO_PAGAMENTO_STYLE[c.statoPagamento] || {}}>{c.statoPagamento}</Pill></div>
                </div>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Costo" />
    </>
  );
}

function CostoForm({ initial, interventiIds, tecnici, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || {
    idIntervento: '', tipo: 'Preventivo', descrizione: '', fornitore: tecnici[0] || '',
    numeroDocumento: '', data: todayISO(), importo: '', statoPagamento: 'Da pagare', note: '',
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica costo' : 'Nuovo costo'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <STR_Field label="Descrizione"><input style={strInputStyle} value={f.descrizione} onChange={set('descrizione')} /></STR_Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <STR_Field label="Tipo"><select style={strInputStyle} value={f.tipo} onChange={set('tipo')}>{STR_TIPI_COSTO.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
          <STR_Field label="Importo (€)"><input type="number" step="0.01" style={strInputStyle} value={f.importo} onChange={set('importo')} placeholder="0.00" /></STR_Field>
        </div>
        <STR_Field label="Fornitore / Ditta">
          <input list="str-fornitori" style={strInputStyle} value={f.fornitore} onChange={set('fornitore')} />
          <datalist id="str-fornitori">{tecnici.map(t => <option key={t} value={t} />)}</datalist>
        </STR_Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <STR_Field label="Numero documento"><input style={strInputStyle} value={f.numeroDocumento} onChange={set('numeroDocumento')} /></STR_Field>
          <STR_Field label="Data"><input type="date" style={strInputStyle} value={f.data} onChange={set('data')} /></STR_Field>
        </div>
        <STR_Field label="ID intervento collegato">
          <input list="str-interventi-ids" style={strInputStyle} value={f.idIntervento} onChange={set('idIntervento')} placeholder="facoltativo" />
          <datalist id="str-interventi-ids">{interventiIds.map(t => <option key={t} value={t} />)}</datalist>
        </STR_Field>
        <STR_Field label="Stato pagamento"><select style={selectStyle(STR_STATO_PAGAMENTO_STYLE[f.statoPagamento] || {})} value={f.statoPagamento} onChange={set('statoPagamento')}>{STR_STATI_PAGAMENTO.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
        <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button onClick={() => onSave({ ...f, id: f.id || uid(), importo: f.importo ? Number(f.importo) : 0 })} style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}>
            Salva costo
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina costo
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Reparti (anagrafica, raggiungibile dal Riepilogo) ---------- */
function RepartiStrScreen({ reparti, onOpen, onAdd, onBack, onRinominaEnd }) {
  const { isAdmin } = usePermessi();
  const [filtroPiano, setFiltroPiano] = useState('');
  const [filtroCategoria, setFiltroCategoria] = useState('');
  const [rinominaCategoria, setRinominaCategoria] = useState(null);
  const [busy, setBusy] = useState(false);

  const piani = useMemo(() => [...new Set(reparti.map(r => r.piano).filter(Boolean))].sort(), [reparti]);
  const categorie = useMemo(() => [...new Set([...STR_CATEGORIE_REPARTO, ...reparti.map(r => r.categoria).filter(Boolean)])].sort(), [reparti]);

  const filtered = reparti.filter(r => {
    if (filtroPiano && r.piano !== filtroPiano) return false;
    if (filtroCategoria && r.categoria !== filtroCategoria) return false;
    return true;
  });

  async function handleRinominaCategoria(vecchia, nuova) {
    setBusy(true);
    await supabase.from('reparti').update({ categoria: nuova }).eq('categoria', vecchia);
    if (filtroCategoria === vecchia) setFiltroCategoria(nuova);
    setRinominaCategoria(null); setBusy(false);
    if (onRinominaEnd) onRinominaEnd();
  }

  const pillBtn = (label, active, onClick) => (
    <button
      key={label}
      onClick={() => onClick(active ? '' : label)}
      style={{ border: 'none', borderRadius: 999, padding: '5px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0,
        background: active ? STR_COLORS.primary : STR_COLORS.bg,
        color: active ? '#fff' : STR_COLORS.muted,
        boxShadow: active ? `0 2px 8px ${STR_COLORS.primary}44` : 'none',
      }}
    >{label}</button>
  );

  return (
    <>
      <TopBar theme={STR_COLORS} title="Reparti e Zone" subtitle={`${filtered.length} di ${reparti.length}`} onBack={onBack} />

      {/* Filtro Piano */}
      {piani.length > 0 && (
        <div style={{ padding: '8px 14px 0', borderBottom: `1px solid ${STR_COLORS.line}` }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: STR_COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 6 }}>Piano</div>
          <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 10 }}>
            {piani.map(p => pillBtn(p, filtroPiano === p, setFiltroPiano))}
          </div>
        </div>
      )}

      {/* Filtro Tipologia con rename per admin */}
      <div style={{ padding: '8px 14px 0', borderBottom: `1px solid ${STR_COLORS.line}` }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: STR_COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.06em' }}>Tipologia</div>
          {isAdmin && <span style={{ fontSize: 10.5, color: STR_COLORS.muted }}>✎ tocca per rinominare</span>}
        </div>
        <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 10, flexWrap: 'nowrap' }}>
          {categorie.map(c => (
            <div key={c} style={{ display: 'flex', alignItems: 'center', gap: 3, flexShrink: 0 }}>
              {pillBtn(c, filtroCategoria === c, setFiltroCategoria)}
              {isAdmin && (
                <button
                  onClick={e => { e.stopPropagation(); setRinominaCategoria(c); }}
                  style={{ background: 'none', border: 'none', padding: '2px 4px', color: STR_COLORS.muted, fontSize: 13, cursor: 'pointer' }}
                  title={`Rinomina "${c}"`}
                >✎</button>
              )}
            </div>
          ))}
        </div>
      </div>

      <div style={{ padding: 14 }}>
        {filtered.length === 0 && <Empty theme={STR_COLORS} icon={Building2} text="Nessun reparto per questo filtro." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {filtered.map(r => (
            <Card theme={STR_COLORS} key={r.codice} onClick={() => onOpen(r)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 3 }}>{r.nome}</div>
                  <div style={{ fontSize: 12, color: STR_COLORS.muted }}>{r.responsabile}</div>
                  {r.piano && <div style={{ fontSize: 11, color: STR_COLORS.muted, marginTop: 2 }}>{r.piano}</div>}
                </div>
                <Pill style={{ bg: STR_COLORS.bg, fg: STR_COLORS.primary }}>{r.categoria}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Reparto" />

      {rinominaCategoria && (
        <RenameDialog
          theme={STR_COLORS}
          label="tipologia"
          currentName={rinominaCategoria}
          onSave={nuova => handleRinominaCategoria(rinominaCategoria, nuova)}
          onCancel={() => setRinominaCategoria(null)}
        />
      )}
      {busy && <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.3)', zIndex:100, display:'flex', alignItems:'center', justifyContent:'center', color:'#fff', fontSize:14 }}>Aggiornamento…</div>}
    </>
  );
}

function RepartoForm({ initial, reparti, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { codice: uid(), nome: '', categoria: 'Servizi', piano: '', responsabile: '', note: '' });
  const [nuovaCategoria, setNuovaCategoria] = useState(false);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  // Tutte le categorie: fisse + quelle già usate nei reparti esistenti
  const categorieUsate = useMemo(() => {
    const tutte = new Set([...STR_CATEGORIE_REPARTO, ...(reparti || []).map(r => r.categoria).filter(Boolean)]);
    return [...tutte].sort();
  }, [reparti]);

  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica reparto' : 'Nuovo reparto'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <STR_Field label="Nome reparto / zona *"><input style={strInputStyle} value={f.nome} onChange={set('nome')} /></STR_Field>

        <STR_Field label="Categoria">
          {!nuovaCategoria ? (
            <>
              <select
                style={strInputStyle}
                value={f.categoria}
                onChange={e => {
                  if (e.target.value === '__nuova__') {
                    setNuovaCategoria(true);
                    setF({ ...f, categoria: '' });
                  } else {
                    setF({ ...f, categoria: e.target.value });
                  }
                }}
              >
                {categorieUsate.map(c => <option key={c} value={c}>{c}</option>)}
                {isAdmin && <option value="__nuova__">➕ Aggiungi nuova tipologia…</option>}
              </select>
            </>
          ) : (
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input
                style={{ ...strInputStyle, flex: 1 }}
                value={f.categoria}
                onChange={set('categoria')}
                placeholder="Scrivi la nuova tipologia…"
                autoFocus
              />
              <button
                type="button"
                onClick={() => {
                  if (!f.categoria.trim()) setF({ ...f, categoria: categorieUsate[0] || 'Servizi' });
                  setNuovaCategoria(false);
                }}
                style={{ background: STR_COLORS.bg, border: `1.5px solid ${STR_COLORS.line}`, borderRadius: 8, padding: '10px 12px', fontSize: 12.5, color: STR_COLORS.muted, whiteSpace: 'nowrap' }}
              >
                Annulla
              </button>
            </div>
          )}
          {nuovaCategoria && (
            <div style={{ fontSize: 11.5, color: STR_COLORS.muted, marginTop: 4 }}>
              Questa nuova tipologia sarà disponibile per tutti i reparti da ora in poi.
            </div>
          )}
        </STR_Field>

        <STR_Field label="Piano">
          <select style={strInputStyle} value={f.piano || ''} onChange={set('piano')}>
            <option value="">— non specificato</option>
            {STR_PIANI_REPARTO.map(p => <option key={p}>{p}</option>)}
          </select>
        </STR_Field>
        <STR_Field label="Responsabile"><input style={strInputStyle} value={f.responsabile} onChange={set('responsabile')} /></STR_Field>
        <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button onClick={() => onSave(f)} style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}>
            Salva reparto
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina reparto
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Tecnici e ditte (anagrafica, raggiungibile dal Riepilogo) ---------- */
function TecniciStrScreen({ tecnici, onOpen, onAdd, onBack, titolo }) {
  const titoloMostrato = titolo || 'Tecnici e Ditte';
  return (
    <>
      <TopBar theme={STR_COLORS} title={titoloMostrato} subtitle={`${tecnici.length} in anagrafica`} onBack={onBack} />
      <div style={{ padding: 14 }}>
        {tecnici.length === 0 && <Empty theme={STR_COLORS} icon={Users} text="Nessun tecnico registrato." />}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          {tecnici.map(t => (
            <Card theme={STR_COLORS} key={t.id} onClick={() => onOpen(t)}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 14.5, marginBottom: 3 }}>{t.nome}</div>
                  <div style={{ fontSize: 12, color: STR_COLORS.muted }}>{t.specializzazione}{t.telefono ? ' · ' + t.telefono : ''}</div>
                </div>
                <Pill style={t.tipo === 'Interno' ? { bg: '#DCEEE3', fg: '#1F6B45' } : t.tipo === 'Fornitore' ? { bg: '#FDF3D0', fg: '#8A6000' } : { bg: STR_COLORS.bg, fg: STR_COLORS.primary }}>{t.tipo}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <STR_FAB onClick={onAdd} label="Tecnico" />
    </>
  );
}

function TecnicoForm({ initial, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { id: uid(), nome: '', tipo: 'Esterno', specializzazione: '', indirizzo: '', telefono: '', email: '', note: '' });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <>
      <TopBar theme={STR_COLORS} title={initial ? 'Modifica tecnico' : 'Nuovo tecnico'} onBack={onCancel} />
      <div style={{ padding: 16, pointerEvents: puoScrivere ? 'auto' : 'none', opacity: puoScrivere ? 1 : 0.65 }}>
        <STR_Field label="Nome / Ragione sociale *"><input style={strInputStyle} value={f.nome} onChange={set('nome')} /></STR_Field>
        <STR_Field label="Tipo"><select style={strInputStyle} value={f.tipo} onChange={set('tipo')}>{STR_TIPI_TECNICO.map(t => <option key={t}>{t}</option>)}</select></STR_Field>
        <STR_Field label="Specializzazione"><input style={strInputStyle} value={f.specializzazione} onChange={set('specializzazione')} /></STR_Field>
        <STR_Field label="Indirizzo"><input style={strInputStyle} value={f.indirizzo || ''} onChange={set('indirizzo')} placeholder="Via, civico, città" /></STR_Field>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <STR_Field label="Telefono"><input style={strInputStyle} value={f.telefono} onChange={set('telefono')} /></STR_Field>
          <STR_Field label="Email"><input style={strInputStyle} value={f.email} onChange={set('email')} /></STR_Field>
        </div>
        <STR_Field label="Note"><input style={strInputStyle} value={f.note} onChange={set('note')} /></STR_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding: '0 16px' }}>
          <button onClick={() => onSave(f)} style={{ width: '100%', background: STR_COLORS.primary, color: '#fff', border: 'none', borderRadius: 12, padding: '14px', fontWeight: 700, fontSize: 15, marginTop: 4 }}>
            Salva tecnico
          </button>
          {initial && puoEliminare && (
            <button onClick={() => setConfermaElimina(true)} style={{ width: '100%', background: 'none', border: 'none', color: STR_COLORS.danger, fontWeight: 600, fontSize: 13.5, padding: '14px 0 4px' }}>
              Elimina tecnico
            </button>
          )}
        </div>
      )}
      {confermaElimina && (
        <ConfirmDelete
          theme={STR_COLORS}
          onConfirm={() => { setConfermaElimina(false); onDelete(initial); }}
          onCancel={() => setConfermaElimina(false)}
        />
      )}
    </>
  );
}

/* ---------- Riepilogo (dashboard) ---------- */
function RiepilogoStrScreen({ camere, reparti, tecnici, interventi, manutenzioni, costi, onMenu, onHome, onOpenTecnici, onOpenCosti, onOpenFornitori, onOpenDitte, onOpenCamereConStato, onOpenInterventiConStato, onOpenScadenze }) {
  const fuoriServizio = camere.filter(c => c.stato === 'Fuori Servizio').length;
  const inManutenzioneCamere = camere.filter(c => c.stato === 'In Manutenzione').length;
  const aperti = interventi.filter(i => i.tipologia !== 'carrozzina' && i.stato === 'Aperto').length;
  const inCorso = interventi.filter(i => i.stato === 'In corso').length;
  const withAlert = manutenzioni.map(m => strAlertStatus(daysUntil(m.prossimaScadenza)));
  const scadute = withAlert.filter(s => s === 'SCADUTO').length;
  const inScadenza = withAlert.filter(s => s === 'IN SCADENZA').length;
  const costoTotale = costi.reduce((s, c) => s + (Number(c.importo) || 0), 0);

  const perTipoCosto = useMemo(() => {
    const map = {};
    STR_TIPI_COSTO.forEach(t => { map[t] = 0; });
    costi.forEach(c => { map[c.tipo] = (map[c.tipo] || 0) + (Number(c.importo) || 0); });
    return Object.entries(map);
  }, [costi]);
  const maxCosto = Math.max(1, ...perTipoCosto.map(([, v]) => v));

  function ClickableStatCard({ label, value, accent, onClick, disabled }) {
    return (
      <div onClick={disabled ? undefined : onClick} style={{ cursor: disabled ? 'default' : 'pointer' }}>
        <StatCard theme={STR_COLORS} label={disabled ? label : `${label} →`} value={value} accent={accent} />
      </div>
    );
  }

  return (
    <>
      <TopBar theme={STR_COLORS} title="Riepilogo" subtitle="Struttura · camere e interventi" onBack={onHome} backIcon={Home} right={<MenuButton onClick={onMenu} />} />
      <div style={{ padding: 14 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 6 }}>
          <ClickableStatCard label="Camere totali" value={camere.length} accent={STR_COLORS.primary} onClick={() => onOpenCamereConStato(null)} disabled={!camere.length} />
          <ClickableStatCard label="Fuori servizio" value={fuoriServizio} accent={fuoriServizio ? STR_COLORS.danger : STR_COLORS.ok} onClick={() => onOpenCamereConStato('Fuori Servizio')} disabled={!fuoriServizio} />
          <ClickableStatCard label="In manutenzione" value={inManutenzioneCamere} accent={inManutenzioneCamere ? STR_COLORS.amber : STR_COLORS.ok} onClick={() => onOpenCamereConStato('In Manutenzione')} disabled={!inManutenzioneCamere} />
          <ClickableStatCard label="Interventi aperti" value={aperti + inCorso} accent={(aperti + inCorso) ? STR_COLORS.danger : STR_COLORS.ok} onClick={() => onOpenInterventiConStato('Aperto')} disabled={!(aperti + inCorso)} />
          <ClickableStatCard label="Scadenze scadute" value={scadute} accent={scadute ? STR_COLORS.danger : STR_COLORS.ok} onClick={onOpenScadenze} disabled={!scadute} />
          <ClickableStatCard label="In scadenza (7gg)" value={inScadenza} accent={inScadenza ? STR_COLORS.amber : STR_COLORS.ok} onClick={onOpenScadenze} disabled={!inScadenza} />
        </div>
        <div style={{ marginBottom: 6 }}>
          <StatCard theme={STR_COLORS} label="Costo totale registrato" value={fmtEuro(costoTotale)} accent={STR_COLORS.primary} />
        </div>

        <SectionLabel theme={STR_COLORS}>Costi per tipo</SectionLabel>
        <Card theme={STR_COLORS}>
          {perTipoCosto.map(([tipo, v], i) => (
            <div key={tipo} style={{ marginTop: i ? 10 : 0 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12.5, marginBottom: 4 }}>
                <span>{tipo}</span><span style={{ color: STR_COLORS.muted }}>{fmtEuro(v)}</span>
              </div>
              <div style={{ height: 7, background: '#EEEAE0', borderRadius: 999 }}>
                <div style={{ height: '100%', width: `${(v / maxCosto) * 100}%`, background: STR_COLORS.amber, borderRadius: 999 }} />
              </div>
            </div>
          ))}
        </Card>

        <SectionLabel theme={STR_COLORS}>Anagrafiche e Costi</SectionLabel>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
          <Card theme={STR_COLORS} onClick={onOpenTecnici}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Users size={17} color={STR_COLORS.primary} /><span style={{ fontWeight: 700, fontSize: 14 }}>Tecnici interni</span></div>
              <span style={{ fontSize: 12.5, color: STR_COLORS.muted }}>{tecnici.filter(t => t.tipo === 'Interno').length} →</span>
            </div>
          </Card>
          <Card theme={STR_COLORS} onClick={onOpenFornitori}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Building2 size={17} color={STR_COLORS.primary} /><span style={{ fontWeight: 700, fontSize: 14 }}>Fornitori</span></div>
              <span style={{ fontSize: 12.5, color: STR_COLORS.muted }}>{tecnici.filter(t => t.tipo === 'Fornitore').length} →</span>
            </div>
          </Card>
          <Card theme={STR_COLORS} onClick={onOpenDitte}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Wrench size={17} color={STR_COLORS.primary} /><span style={{ fontWeight: 700, fontSize: 14 }}>Ditte esterne</span></div>
              <span style={{ fontSize: 12.5, color: STR_COLORS.muted }}>{tecnici.filter(t => t.tipo === 'Esterno').length} →</span>
            </div>
          </Card>
          <Card theme={STR_COLORS} onClick={onOpenCosti}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Wallet size={17} color={STR_COLORS.primary} /><span style={{ fontWeight: 700, fontSize: 14 }}>Costi</span></div>
              <span style={{ fontSize: 12.5, color: STR_COLORS.muted }}>{costi.length} voci · {fmtEuro(costoTotale)} →</span>
            </div>
          </Card>
        </div>
      </div>
    </>
  );
}

/* ---------- Root ---------- */
function StrutturaModule({ onHome }) {
  const camereT = useSupaTable('camere', 'codice', S_CAMERE);
  const repartiT = useSupaTable('reparti', 'codice', S_REPARTI);
  const tecniciT = useSupaTable('tecnici', 'id', S_TECNICI);
  const interventiT = useSupaTable('interventi', 'id', S_INTERVENTI, ['dataSegnalazione', 'dataChiusura']);
  const manutenzioniT = useSupaTable('manutenzioni', 'id', S_MANUTENZIONI, ['ultimaEsecuzione', 'prossimaScadenza']);
  const costiT = useSupaTable('costi', 'id', S_COSTI, ['data']);
  const camere = camereT.rows, reparti = repartiT.rows, tecnici = tecniciT.rows;
  const interventi = interventiT.rows, manutenzioni = manutenzioniT.rows, costi = costiT.rows;
  const ready = camereT.ready && repartiT.ready && tecniciT.ready && interventiT.ready && manutenzioniT.ready && costiT.ready;
  const dataError = camereT.error || repartiT.error || tecniciT.error || interventiT.error || manutenzioniT.error || costiT.error;
  const { nomeVisualizzato } = usePermessi();

  const [tab, setTab] = useState('camere');
  const [subScreen, setSubScreen] = useState(null); // null | 'reparti' | 'tecnici' | ...
  const [view, setView] = useState(LIST_VIEW);
  const [toast, setToast] = useState('');
  const [showMenu, setShowMenu] = useState(false);
  const [filtroRiepilogo, setFiltroRiepilogo] = useState(null); // { tipo: 'camere_stato'|'interventi_stato', valore }
  useBackable(showMenu, setShowMenu);

  useBackable(tab, setTab);
  useBackable(subScreen, setSubScreen);
  useBackable(view, setView);

  useEffect(() => { setView(LIST_VIEW); setFiltroRiepilogo(null); }, [tab, subScreen]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 2400); };

  const saveCamera = async (c, originalCodice) => {
    const codiceNuovo = (c.codice || '').trim();
    if (!codiceNuovo) { flash('Il codice camera non puo\' essere vuoto'); return; }
    const matchCodice = originalCodice || c.codice;
    const duplicato = camere.some(x => x.codice === codiceNuovo && x.codice !== matchCodice);
    if (duplicato) { flash(`Esiste gia' una camera con codice ${codiceNuovo}`); return; }
    const cNorm = { ...c, codice: codiceNuovo };

    if (originalCodice && originalCodice !== codiceNuovo) {
      // Rinomina vera e propria (la chiave primaria cambia): aggiorno la riga
      // esistente e sposto i riferimenti negli interventi/scadenze collegati,
      // cosi' non restano orfani.
      const { data, error } = await supabase.from('camere').update(cNorm).eq('codice', originalCodice).select().single();
      if (error) { flash(traduciErroreDati(error.message)); return; }
      camereT.setRows((prev) => prev.map((x) => (x.codice === originalCodice ? data : x)));

      const { error: e2 } = await supabase.from('interventi').update({ cameraZona: codiceNuovo }).eq('cameraZona', originalCodice);
      if (!e2) interventiT.setRows((prev) => prev.map((x) => (x.cameraZona === originalCodice ? { ...x, cameraZona: codiceNuovo } : x)));

      const { error: e3 } = await supabase.from('manutenzioni').update({ cameraZona: codiceNuovo }).eq('cameraZona', originalCodice);
      if (!e3) manutenzioniT.setRows((prev) => prev.map((x) => (x.cameraZona === originalCodice ? { ...x, cameraZona: codiceNuovo } : x)));

      flash(`Camera rinominata in ${codiceNuovo}`);
      goBackMulti(1, setView, { name: 'detail', id: codiceNuovo });
    } else {
      const { error } = await camereT.save(cNorm);
      if (error) { flash(error.message); return; }
      flash('Camera salvata');
      goBack();
    }
  };
  const deleteCamera = async (c) => {
    const { error } = await camereT.remove(c);
    if (error) { flash(error.message); return; }
    flash('Camera eliminata');
    // La modifica si apre sempre dal dettaglio: eliminando, quel dettaglio
    // non avrebbe piu' senso (la camera non esiste piu'), quindi si salta
    // dritti alla lista invece di passare da una scheda ormai orfana.
    goBackMulti(2, setView, LIST_VIEW);
  };
  async function salvaOTorna(azione, entita, msgOk) {
    const { error } = await azione();
    if (error) { flash(error.message); return; }
    flash(msgOk);
    goBack();
  }
  const saveReparto = (r) => salvaOTorna(() => repartiT.save(r), 'reparto', 'Reparto salvato');
  const deleteReparto = (r) => salvaOTorna(() => repartiT.remove(r), 'reparto', 'Reparto eliminato');
  const saveTecnico = (t) => salvaOTorna(() => tecniciT.save(t), 'tecnico', 'Tecnico salvato');
  const deleteTecnico = (t) => salvaOTorna(() => tecniciT.remove(t), 'tecnico', 'Tecnico eliminato');
  const saveIntervento = async (i) => {
    const { error } = await interventiT.save(i);
    if (error) { flash(error.message); return; }
    flash('Intervento salvato');
    if (!i.id || !interventi.find(x => x.id === i.id)) {
      // è nuovo: invia notifica
      inviaNotifica({
        tipo: 'intervento_struttura',
        titolo: `🏥 Struttura: ${i.cameraZona || 'Nuovo intervento'}`,
        corpo: i.descrizione || 'Nuovo intervento aperto',
        linkId: i.id,
        inviata_da: i.segnalatoDa || nomeVisualizzato,
      });
    }
    goBack();
  };
  const deleteIntervento = (i) => salvaOTorna(async () => { await eliminaFoto(i.foto); return interventiT.remove(i); }, 'intervento', 'Intervento eliminato');
  const saveManutenzione = (m) => salvaOTorna(() => manutenzioniT.save(m), 'manutenzione', 'Scadenza salvata');
  const deleteManutenzione = (m) => salvaOTorna(() => manutenzioniT.remove(m), 'manutenzione', 'Scadenza eliminata');
  const saveCosto = (c) => salvaOTorna(() => costiT.save(c), 'costo', 'Costo salvato');
  const deleteCosto = (c) => salvaOTorna(() => costiT.remove(c), 'costo', 'Costo eliminato');
  // Intervento aperto dal dettaglio di una camera
  const saveInterventoDaCamera = async (i) => {
    const { error } = await interventiT.save(i);
    if (error) { flash(error.message); return; }
    flash('Intervento salvato');
    if (!i.id || !interventi.find(x => x.id === i.id)) {
      inviaNotifica({
        tipo: 'intervento_struttura',
        titolo: `🏥 Struttura: ${i.cameraZona || 'Camera'}`,
        corpo: i.descrizione || 'Nuovo intervento aperto',
        linkId: i.id,
        inviata_da: i.segnalatoDa || nomeVisualizzato,
      });
    }
    goBack();
  };
  const deleteInterventoDaCamera = (i) => salvaOTorna(async () => { await eliminaFoto(i.foto); return interventiT.remove(i); }, 'intervento', 'Intervento eliminato');

  const luoghi = useMemo(() => [...camere.map(c => c.codice), ...reparti.map(r => r.nome), 'Struttura (tutti i piani)'], [camere, reparti]);
  const tecniciNomi = useMemo(() => tecnici.map(t => t.nome), [tecnici]);
  const piani = useMemo(() => [...new Set(camere.map(c => c.piano).filter(Boolean))], [camere]);
  const nuclei = useMemo(() => [...new Set(camere.map(c => c.nucleo).filter(Boolean))], [camere]);
  const interventiIds = useMemo(() => interventi.map(i => i.id), [interventi]);

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();

    const wsCamere = XLSX.utils.json_to_sheet(camere.map(c => ({ 'Codice Camera': c.codice, 'Piano': c.piano, 'Nucleo': c.nucleo, 'Tipo': c.tipo, 'Stato': c.stato, 'Note': c.note })));
    wsCamere['!cols'] = [14, 16, 14, 12, 16, 26].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsCamere, 'Camere');

    const wsReparti = XLSX.utils.json_to_sheet(reparti.map(r => ({ 'Codice': r.codice, 'Nome': r.nome, 'Categoria': r.categoria, 'Responsabile': r.responsabile, 'Note': r.note })));
    wsReparti['!cols'] = [10, 30, 16, 26, 26].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsReparti, 'Reparti_Zone');

    const wsTecnici = XLSX.utils.json_to_sheet(tecnici.map(t => ({ 'ID': t.id, 'Nome': t.nome, 'Tipo': t.tipo, 'Specializzazione': t.specializzazione, 'Telefono': t.telefono, 'Email': t.email, 'Note': t.note })));
    wsTecnici['!cols'] = [8, 26, 10, 26, 15, 26, 20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsTecnici, 'Tecnici_Ditte');

    const dt = (iso) => (iso ? new Date(iso + 'T00:00:00') : '');
    const wsInterventi = XLSX.utils.json_to_sheet(interventi.map(i => ({
      'ID': i.id, 'Data Segnalazione': dt(i.dataSegnalazione), 'Camera/Zona': i.cameraZona, 'Descrizione': i.descrizione,
      'Priorita': i.priorita, 'Tecnico': i.tecnico, 'Stato': i.stato, 'Data Chiusura': dt(i.dataChiusura), 'Costo (EUR)': i.costo || 0, 'Note': i.note,
    })), { cellDates: true });
    wsInterventi['!cols'] = [8, 16, 14, 32, 10, 20, 12, 16, 12, 24].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsInterventi, 'Storico_Interventi');

    const wsManut = XLSX.utils.json_to_sheet(manutenzioni.map(m => ({
      'ID': m.id, 'Camera/Zona': m.cameraZona, 'Tipo': m.tipoManutenzione, 'Frequenza': m.frequenza,
      'Ultima Esecuzione': dt(m.ultimaEsecuzione), 'Prossima Scadenza': dt(m.prossimaScadenza),
      'Giorni alla Scadenza': daysUntil(m.prossimaScadenza), 'Stato Allerta': strAlertStatus(daysUntil(m.prossimaScadenza)), 'Tecnico': m.tecnico, 'Note': m.note,
    })), { cellDates: true });
    wsManut['!cols'] = [8, 22, 18, 14, 16, 16, 16, 14, 20, 24].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsManut, 'Manutenzioni_Programmate');

    const wsCosti = XLSX.utils.json_to_sheet(costi.map(c => ({
      'ID': c.id, 'ID Intervento': c.idIntervento, 'Tipo': c.tipo, 'Descrizione': c.descrizione, 'Fornitore': c.fornitore,
      'Numero Documento': c.numeroDocumento, 'Data': dt(c.data), 'Importo (EUR)': c.importo || 0, 'Stato Pagamento': c.statoPagamento, 'Note': c.note,
    })), { cellDates: true });
    wsCosti['!cols'] = [8, 14, 14, 28, 20, 16, 14, 12, 16, 20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsCosti, 'Costi');

    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `registro_manutenzione_rsa_${stamp}.xlsx`);
    goBack();
    flash('File Excel scaricato');
  };

  if (!ready) {
    return (
      <div style={{ minHeight: '100vh', background: STR_COLORS.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', color: STR_COLORS.muted, flexDirection: 'column', gap: 10, padding: 24, textAlign: 'center' }}>
        <span>Caricamento…</span>
        {dataError && <span style={{ color: STR_COLORS.red, fontSize: 13, maxWidth: 280 }}>{dataError}</span>}
      </div>
    );
  }

  const onMenu = () => setShowMenu(true);
  let content;

  if (subScreen === 'tecnici') {
    if (view.name === 'add') content = <TecnicoForm onSave={saveTecnico} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <TecnicoForm initial={view.t} onSave={saveTecnico} onCancel={() => goBack()} onDelete={deleteTecnico} />;
    else content = <TecniciStrScreen tecnici={tecnici.filter(t => t.tipo === 'Interno')} titolo="Tecnici interni" onOpen={(t) => setView({ name: 'edit', t })} onAdd={() => setView({ name: 'add' })} onBack={() => setSubScreen(null)} />;
  } else if (subScreen === 'fornitori') {
    if (view.name === 'add') content = <TecnicoForm onSave={saveTecnico} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <TecnicoForm initial={view.t} onSave={saveTecnico} onCancel={() => goBack()} onDelete={deleteTecnico} />;
    else content = <TecniciStrScreen tecnici={tecnici.filter(t => t.tipo === 'Fornitore')} titolo="Fornitori" onOpen={(t) => setView({ name: 'edit', t })} onAdd={() => setView({ name: 'add' })} onBack={() => setSubScreen(null)} />;
  } else if (subScreen === 'ditte') {
    if (view.name === 'add') content = <TecnicoForm onSave={saveTecnico} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <TecnicoForm initial={view.t} onSave={saveTecnico} onCancel={() => goBack()} onDelete={deleteTecnico} />;
    else content = <TecniciStrScreen tecnici={tecnici.filter(t => t.tipo === 'Esterno')} titolo="Ditte esterne" onOpen={(t) => setView({ name: 'edit', t })} onAdd={() => setView({ name: 'add' })} onBack={() => setSubScreen(null)} />;
  } else if (tab === 'camere') {
    if (view.name === 'detail') content = <CameraDetail camera={camere.find(c => c.codice === view.id)} interventi={interventi} onBack={() => goBack()} onEdit={(c) => setView({ name: 'edit', c })} onOpenIntervento={(i) => setView({ name: 'intervento', i, cameraId: view.id })} />;
    else if (view.name === 'add') content = <CameraForm piani={piani} nuclei={nuclei} onSave={(c) => saveCamera(c)} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <CameraForm initial={view.c} piani={piani} nuclei={nuclei} onSave={(c) => saveCamera(c, view.c.codice)} onCancel={() => goBack()} onDelete={deleteCamera} />;
    else if (view.name === 'intervento') content = <InterventoForm initial={view.i} luoghi={luoghi} camere={camere} reparti={reparti} tecnici={tecniciNomi} onSave={saveInterventoDaCamera} onCancel={() => goBack()} onDelete={deleteInterventoDaCamera} />;
    else content = <CamereScreen camere={camere} onOpen={(c) => setView({ name: 'detail', id: c.codice })} onAdd={() => setView({ name: 'add' })} onMenu={onMenu} onHome={onHome} filtroStatoIniziale={filtroRiepilogo?.tipo === 'camere_stato' ? filtroRiepilogo.valore : ''} />;
  } else if (tab === 'interventi') {
    if (view.name === 'add') content = <InterventoForm luoghi={luoghi} camere={camere} reparti={reparti} tecnici={tecniciNomi} onSave={saveIntervento} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <InterventoForm initial={view.i} luoghi={luoghi} camere={camere} reparti={reparti} tecnici={tecniciNomi} onSave={saveIntervento} onCancel={() => goBack()} onDelete={deleteIntervento} />;
    else content = <InterventiScreen interventi={interventi} onOpen={(i) => setView({ name: 'edit', i })} onAdd={() => setView({ name: 'add' })} onMenu={onMenu} onHome={onHome} filtroStatoIniziale={filtroRiepilogo?.tipo === 'interventi_stato' ? filtroRiepilogo.valore : null} />;
  } else if (tab === 'scadenze') {
    if (view.name === 'add') content = <ManutenzioneForm luoghi={luoghi} tecnici={tecniciNomi} onSave={saveManutenzione} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <ManutenzioneForm initial={view.m} luoghi={luoghi} tecnici={tecniciNomi} onSave={saveManutenzione} onCancel={() => goBack()} onDelete={deleteManutenzione} />;
    else content = <ScadenzeStrScreen manutenzioni={manutenzioni} onOpen={(m) => setView({ name: 'edit', m })} onAdd={() => setView({ name: 'add' })} onMenu={onMenu} onHome={onHome} />;
  } else if (tab === 'reparti') {
    if (view.name === 'add') content = <RepartoForm reparti={reparti} onSave={saveReparto} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <RepartoForm reparti={reparti} initial={view.r} onSave={saveReparto} onCancel={() => goBack()} onDelete={deleteReparto} />;
    else content = <RepartiStrScreen reparti={reparti} onOpen={(r) => setView({ name: 'edit', r })} onAdd={() => setView({ name: 'add' })} onBack={() => goBack()} onRinominaEnd={() => repartiT.reload()} />;
  } else if (subScreen === 'costi') {
    if (view.name === 'add') content = <CostoForm interventiIds={interventiIds} tecnici={tecniciNomi} onSave={saveCosto} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <CostoForm initial={view.c} interventiIds={interventiIds} tecnici={tecniciNomi} onSave={saveCosto} onCancel={() => goBack()} onDelete={deleteCosto} />;
    else content = <CostiStrScreen costi={costi} onOpen={(c) => setView({ name: 'edit', c })} onAdd={() => setView({ name: 'add' })} onMenu={onMenu} onHome={onHome} onBack={() => setSubScreen(null)} />;
  } else {
    content = <RiepilogoStrScreen camere={camere} reparti={reparti} tecnici={tecnici} interventi={interventi} manutenzioni={manutenzioni} costi={costi} onMenu={onMenu} onHome={onHome}
      onOpenTecnici={() => setSubScreen('tecnici')} onOpenCosti={() => setSubScreen('costi')}
      onOpenFornitori={() => setSubScreen('fornitori')} onOpenDitte={() => setSubScreen('ditte')}
      onOpenCamereConStato={(stato) => { setFiltroRiepilogo(stato ? { tipo: 'camere_stato', valore: stato } : null); setTab('camere'); }}
      onOpenInterventiConStato={(stato) => { setFiltroRiepilogo({ tipo: 'interventi_stato', valore: stato }); setTab('interventi'); }}
      onOpenScadenze={() => setTab('scadenze')}
    />;
  }

  const showBottomNav = view.name === 'list' && !['costi','tecnici','fornitori','ditte'].includes(subScreen);

  return (
    <div style={{ minHeight: '100vh', background: STR_COLORS.bg, fontFamily: 'Inter, sans-serif', color: STR_COLORS.ink, maxWidth: 480, margin: '0 auto', position: 'relative' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        input:focus, select:focus, textarea:focus { border-color: ${STR_COLORS.primary} !important; }
      `}</style>
      <div style={{ paddingBottom: showBottomNav ? 78 : 20 }}>{content}</div>
      {showBottomNav && <BottomNav theme={STR_COLORS} tab={tab} setTab={setTab} items={STR_NAV_ITEMS} badges={{ interventi: interventi.filter(i => i.tipologia !== 'carrozzina' && i.stato === 'Aperto').length }} />}
      {showMenu && <MenuSheet theme={STR_COLORS} onClose={() => goBack()} onExport={exportToExcel} exportSub="Scarica camere, interventi, scadenze e costi in .xlsx" />}
      {toast && (
        <div style={{ position: 'fixed', bottom: showBottomNav ? 92 : 20, left: '50%', transform: 'translateX(-50%)', background: STR_COLORS.primaryDeep, color: '#fff', padding: '10px 18px', borderRadius: 999, fontSize: 13, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 7, zIndex: 30, maxWidth: 400 }}>
          <Check size={15} /> {toast}
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   HUB — schermata di ingresso con i moduli disponibili
   ========================================================================= */

// ============================================================
//  MODULO PROCEDURE
// ============================================================

const PROC_COLORS = {
  primary: '#5E3A8A', primaryDeep: '#3D2466',
  bg: '#F2EEF8', surface: '#FFFFFF', line: '#DDD8EE',
  ink: '#1E1530', muted: '#7A6A9A', danger: '#C0392B', ok: '#27AE60',
};

const PROC_TIPOLOGIE = ['Idraulico','Elettrico','Muratura/Edile','Climatizzazione','Antincendio','Informatica','Sicurezza','Generale'];
const PROC_FREQUENZE = ['Giornaliera','Settimanale','Mensile','Trimestrale','Semestrale','Annuale'];
const PROC_ARMADI_TIPI = ['Elettrico','Idraulico','Medicale','Attrezzature','DPI','Documenti','Generale'];
const PROC_UNITA = ['pz','kg','lt','m','conf','scatola','rotolo','pacco'];
const FREQ_GIORNI = { Giornaliera:1, Settimanale:7, Mensile:30, Trimestrale:90, Semestrale:180, Annuale:365 };

const PROC_NAV_ITEMS = [
  ['procedure', BookOpen, 'Procedure'],
  ['ricorrenti', RefreshCw, 'Ricorrenti'],
  ['armadi', Package, 'Armadi'],
];

const PROC_BUCKET = 'procedure-files';

function PROC_Field({ label, children }) {
  return (
    <label style={{ display: 'block', marginBottom: 14 }}>
      <span style={{ display: 'block', fontSize: 12, fontWeight: 600, color: PROC_COLORS.muted, marginBottom: 5, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{label}</span>
      {children}
    </label>
  );
}
const procInputStyle = { width: '100%', boxSizing: 'border-box', padding: '11px 12px', borderRadius: 10, border: `1.5px solid ${PROC_COLORS.line}`, fontSize: 15, background: '#FAFAFE', fontFamily: 'Inter, sans-serif', outline: 'none', color: PROC_COLORS.ink };

function PROC_FAB({ onClick, label = 'Aggiungi', visible }) {
  const { isAdmin } = usePermessi();
  if (!isAdmin && !visible) return null;
  return <button onClick={onClick} style={{ position:'fixed', right:18, bottom:84, background:PROC_COLORS.primary, color:'#fff', border:'none', borderRadius:999, height:52, padding:'0 20px 0 16px', display:'flex', alignItems:'center', gap:7, fontWeight:700, fontSize:14.5, boxShadow:`0 6px 16px ${PROC_COLORS.primary}66`, zIndex:20 }}><Plus size={20} strokeWidth={2.6} />{label}</button>;
}

async function caricaProcFile(id, file) {
  const ext = (file.name.split('.').pop() || 'bin').toLowerCase().replace(/[^a-z0-9]/g, '');
  const path = `${id}/${Date.now()}-${Math.random().toString(36).slice(2,8)}.${ext}`;
  const { error } = await supabase.storage.from(PROC_BUCKET).upload(path, file, { upsert: false });
  if (error) throw error;
  return { nome: file.name, path, tipo: file.type };
}
async function eliminaProcFiles(items) {
  if (!items?.length) return;
  try { await supabase.storage.from(PROC_BUCKET).remove(items.map(i => i.path || i)); } catch {}
}
async function urlFirmateProcImmagini(immagini) {
  if (!immagini?.length) return {};
  const { data } = await supabase.storage.from(PROC_BUCKET).createSignedUrls(immagini.map(i => i.path || i), 3600);
  const map = {}; (data || []).forEach(d => { if (d.signedUrl && d.path) map[d.path] = d.signedUrl; }); return map;
}
async function urlFirmateFileProc(files) {
  if (!files?.length) return {};
  const { data } = await supabase.storage.from(PROC_BUCKET).createSignedUrls(files.map(f => f.path), 3600);
  const map = {}; (data || []).forEach(d => { if (d.signedUrl && d.path) map[d.path] = d.signedUrl; }); return map;
}
function calcProssima(ultima, frequenza) {
  if (!ultima) return '';
  const d = new Date(ultima + 'T00:00:00');
  d.setDate(d.getDate() + (FREQ_GIORNI[frequenza] || 30));
  return d.toISOString().slice(0, 10);
}

/* ---- Procedure ---- */
function ProcedureScreen({ procedure, onOpen, onAdd, onHome }) {
  const [filtro, setFiltro] = useState('');
  const filtered = filtro ? procedure.filter(p => p.tipologia === filtro) : procedure;
  return (
    <>
      <TopBar theme={PROC_COLORS} title="Procedure" subtitle={`${procedure.length} procedure`} onBack={onHome} backIcon={Home} />
      <div style={{ padding:'8px 14px 0', borderBottom:`1px solid ${PROC_COLORS.line}` }}>
        <div style={{ display:'flex', gap:6, overflowX:'auto', paddingBottom:10 }}>
          {PROC_TIPOLOGIE.map(t => (
            <button key={t} onClick={() => setFiltro(filtro === t ? '' : t)}
              style={{ border:'none', borderRadius:999, padding:'5px 12px', fontSize:12, fontWeight:700, cursor:'pointer', whiteSpace:'nowrap', flexShrink:0,
                background: filtro === t ? PROC_COLORS.primary : PROC_COLORS.bg,
                color: filtro === t ? '#fff' : PROC_COLORS.muted }}>
              {t}
            </button>
          ))}
        </div>
      </div>
      <div style={{ padding:14 }}>
        {filtered.length === 0 && <Empty theme={PROC_COLORS} icon={BookOpen} text="Nessuna procedura. Aggiungine una." />}
        <div style={{ display:'flex', flexDirection:'column', gap:9 }}>
          {filtered.map(p => (
            <Card theme={PROC_COLORS} key={p.id} onClick={() => onOpen(p)}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', gap:8 }}>
                <div style={{ minWidth:0 }}>
                  <div style={{ fontWeight:700, fontSize:14.5, marginBottom:4 }}>{p.titolo}</div>
                  {p.descrizione && <div style={{ fontSize:12.5, color:PROC_COLORS.muted, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{p.descrizione}</div>}
                  {(p.immagini?.length > 0 || p.videoLinks?.length > 0 || p.files?.length > 0) && (
                    <div style={{ display:'flex', gap:10, marginTop:6, fontSize:11, color:PROC_COLORS.muted }}>
                      {p.immagini?.length > 0 && <span><ImageIcon size={11} style={{ display:'inline', marginRight:3 }} />{p.immagini.length}</span>}
                      {p.videoLinks?.length > 0 && <span><ExternalLink size={11} style={{ display:'inline', marginRight:3 }} />{p.videoLinks.length} video</span>}
                      {p.files?.length > 0 && <span><FileText size={11} style={{ display:'inline', marginRight:3 }} />{p.files.length}</span>}
                    </div>
                  )}
                </div>
                <Pill style={{ bg:PROC_COLORS.bg, fg:PROC_COLORS.primary }}>{p.tipologia}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <PROC_FAB onClick={onAdd} />
    </>
  );
}

function ProceduraDetail({ procedura, onBack, onEdit, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [urls, setUrls] = useState({});
  const [fileUrls, setFileUrls] = useState({});
  const [zoom, setZoom] = useState(null);
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [erroreUrls, setErroreUrls] = useState('');

  useEffect(() => {
    let m = true;
    if (procedura.immagini?.length) {
      urlFirmateProcImmagini(procedura.immagini)
        .then(u => { if (m) setUrls(u); })
        .catch(() => { if (m) setErroreUrls('Impossibile caricare le immagini. Verificare che il bucket "procedure-files" esista su Supabase.'); });
    }
    if (procedura.files?.length) {
      urlFirmateFileProc(procedura.files)
        .then(u => { if (m) setFileUrls(u); })
        .catch(() => {});
    }
    return () => { m = false; };
  }, [procedura.id, procedura.immagini?.length]);

  return (
    <>
      <TopBar theme={PROC_COLORS} title={procedura.titolo} subtitle={procedura.tipologia} onBack={onBack} />
      <div style={{ padding:14 }}>
        {procedura.descrizione && <Card theme={PROC_COLORS} style={{ marginBottom:12 }}><p style={{ margin:0, fontSize:14, lineHeight:1.6 }}>{procedura.descrizione}</p></Card>}
        {procedura.steps && (<><SectionLabel theme={PROC_COLORS}>Procedura passo-passo</SectionLabel><Card theme={PROC_COLORS} style={{ marginBottom:12 }}><pre style={{ margin:0, fontSize:13.5, lineHeight:1.7, fontFamily:'Inter, sans-serif', whiteSpace:'pre-wrap', wordBreak:'break-word' }}>{procedura.steps}</pre></Card></>)}

        {procedura.immagini?.length > 0 && (
          <>
            <SectionLabel theme={PROC_COLORS}>Immagini ({procedura.immagini.length})</SectionLabel>
            {erroreUrls && <div style={{ fontSize:12, color:PROC_COLORS.danger, marginBottom:8 }}>{erroreUrls}</div>}
            <div style={{ display:'flex', gap:8, flexWrap:'wrap', marginBottom:16 }}>
              {procedura.immagini.map((img, i) => {
                const key = img?.path || img;
                const url = urls[key];
                return (
                  <div key={i} onClick={() => url && setZoom(url)}
                    style={{ width:90, height:90, borderRadius:10, overflow:'hidden', background:PROC_COLORS.bg, border:`1px solid ${PROC_COLORS.line}`, cursor:url?'pointer':'default', flexShrink:0, display:'flex', alignItems:'center', justifyContent:'center' }}>
                    {url
                      ? <img src={url} alt="" style={{ width:'100%', height:'100%', objectFit:'cover' }} />
                      : <ImageIcon size={22} color={PROC_COLORS.muted} />
                    }
                  </div>
                );
              })}
            </div>
          </>
        )}

        {procedura.videoLinks?.length > 0 && (<><SectionLabel theme={PROC_COLORS}>Video</SectionLabel><Card theme={PROC_COLORS} style={{ marginBottom:12 }}>{procedura.videoLinks.map((url, i) => (<div key={i} style={{ padding:i?'10px 0 0':'0', borderTop:i?`1px solid ${PROC_COLORS.line}`:'none', display:'flex', alignItems:'center', gap:10 }}><ExternalLink size={15} color={PROC_COLORS.primary} style={{ flexShrink:0 }} /><a href={url} target="_blank" rel="noopener noreferrer" style={{ color:PROC_COLORS.primary, fontSize:13.5, fontWeight:600, wordBreak:'break-all' }}>{url.replace(/^https?:\/\//, '').slice(0,50)}{url.length>55?'…':''}</a></div>))}</Card></>)}
        {procedura.files?.length > 0 && (<><SectionLabel theme={PROC_COLORS}>File allegati</SectionLabel><Card theme={PROC_COLORS} style={{ marginBottom:12 }}>{procedura.files.map((file, i) => { const url = fileUrls[file?.path]; return (<div key={i} style={{ padding:i?'10px 0 0':'0', borderTop:i?`1px solid ${PROC_COLORS.line}`:'none', display:'flex', alignItems:'center', gap:10 }}><FileText size={15} color={PROC_COLORS.primary} style={{ flexShrink:0 }} />{url?<a href={url} target="_blank" rel="noopener noreferrer" style={{ color:PROC_COLORS.primary, fontSize:13.5, fontWeight:600 }}>{file.nome}</a>:<span style={{ fontSize:13.5, color:PROC_COLORS.muted }}>{file.nome}</span>}</div>); })}</Card></>)}
        {procedura.note && (<><SectionLabel theme={PROC_COLORS}>Note</SectionLabel><Card theme={PROC_COLORS} style={{ marginBottom:12 }}><p style={{ margin:0, fontSize:13.5, color:PROC_COLORS.muted }}>{procedura.note}</p></Card></>)}

        {/* Pulsanti modifica/elimina espliciti */}
        {puoScrivere && (
          <div style={{ marginTop:8 }}>
            <button onClick={onEdit} style={{ width:'100%', background:PROC_COLORS.primary, color:'#fff', border:'none', borderRadius:12, padding:'13px', fontWeight:700, fontSize:15, marginBottom:10 }}>
              Modifica procedura
            </button>
            {puoEliminare && (
              <button onClick={() => setConfermaElimina(true)} style={{ width:'100%', background:'none', border:'none', color:PROC_COLORS.danger, fontWeight:600, fontSize:13.5, padding:'4px 0 16px' }}>
                Elimina procedura
              </button>
            )}
          </div>
        )}
      </div>

      {zoom && (<div onClick={() => setZoom(null)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.9)', zIndex:100, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }}><button onClick={() => setZoom(null)} style={{ position:'absolute', top:16, right:16, background:'rgba(255,255,255,0.15)', border:'none', borderRadius:999, width:38, height:38, color:'#fff', display:'flex', alignItems:'center', justifyContent:'center' }}><XIcon size={18} /></button><img src={zoom} alt="" style={{ maxWidth:'100%', maxHeight:'100%', objectFit:'contain', borderRadius:6 }} onClick={e => e.stopPropagation()} /></div>)}
      {confermaElimina && <ConfirmDelete theme={PROC_COLORS} message={`Eliminare "${procedura.titolo}"?`} onConfirm={() => { setConfermaElimina(false); onDelete(procedura); }} onCancel={() => setConfermaElimina(false)} />}
    </>
  );
}

function ProceduraForm({ initial, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { id:uid(), titolo:'', tipologia:PROC_TIPOLOGIE[0], descrizione:'', steps:'', videoLinks:[], files:[], immagini:[], note:'' });
  const [nuoveImmagini, setNuoveImmagini] = useState([]);
  const [nuoviFiles, setNuoviFiles] = useState([]);
  const [immaginiDaRimuovere, setImmaginiDaRimuovere] = useState([]);
  const [filesDaRimuovere, setFilesDaRimuovere] = useState([]);
  const [links, setLinks] = useState(initial?.videoLinks || []);
  const [salvataggio, setSalvataggio] = useState(false);
  const [errore, setErrore] = useState('');
  const imgRef = useRef(null); const fileRef = useRef(null);
  const anteprimeImg = useMemo(() => nuoveImmagini.map(f => URL.createObjectURL(f)), [nuoveImmagini]);
  useEffect(() => { return () => { anteprimeImg.forEach(url => URL.revokeObjectURL(url)); }; }, [anteprimeImg]);
  useEffect(() => { return () => { anteprimeImg.forEach(url => URL.revokeObjectURL(url)); }; }, [anteprimeImg]);
  useEffect(() => { return () => { anteprimeImg.forEach(url => URL.revokeObjectURL(url)); }; }, [anteprimeImg]);
  const set = (k) => (e) => setF(prev => ({ ...prev, [k]: e.target.value }));

  async function handleSave() {
    if (!f.titolo.trim()) { setErrore('Il titolo è obbligatorio.'); return; }
    setSalvataggio(true); setErrore('');
    try {
      if (immaginiDaRimuovere.length) await eliminaProcFiles(immaginiDaRimuovere);
      if (filesDaRimuovere.length) await eliminaProcFiles(filesDaRimuovere);
      const nuoviImg = [];
      for (const img of nuoveImmagini) nuoviImg.push(await caricaProcFile(f.id, img));
      const nuoviF = [];
      for (const file of nuoviFiles) nuoviF.push(await caricaProcFile(f.id, file));
      onSave({ ...f, immagini:[...(f.immagini||[]).filter(x => !immaginiDaRimuovere.includes(x)), ...nuoviImg], files:[...(f.files||[]).filter(x => !filesDaRimuovere.find(d => d.path===x.path)), ...nuoviF], videoLinks:links.filter(l => l.trim()) });
    } catch(e) { setErrore('Errore: ' + e.message); setSalvataggio(false); }
  }

  const xBtn = (onClick) => (<button type="button" onClick={onClick} style={{ background:'none', border:'none', color:PROC_COLORS.danger, padding:4, display:'flex', alignItems:'center', justifyContent:'center' }}><XIcon size={15} /></button>);
  const thumbStyle = { position:'relative', width:72, height:72, borderRadius:10, overflow:'hidden', background:PROC_COLORS.bg, border:`1px solid ${PROC_COLORS.line}`, flexShrink:0 };
  const xOverlay = (onClick) => (<button type="button" onClick={onClick} style={{ position:'absolute', top:2, right:2, background:'rgba(0,0,0,0.55)', border:'none', borderRadius:999, width:20, height:20, color:'#fff', display:'flex', alignItems:'center', justifyContent:'center' }}><XIcon size={12} /></button>);

  return (
    <>
      <TopBar theme={PROC_COLORS} title={initial ? 'Modifica procedura' : 'Nuova procedura'} onBack={onCancel} />
      <div style={{ padding:16, pointerEvents:puoScrivere?'auto':'none', opacity:puoScrivere?1:0.65 }}>
        <PROC_Field label="Titolo *"><input style={procInputStyle} value={f.titolo} onChange={set('titolo')} placeholder="Es. Spurgo radiatori" /></PROC_Field>
        <PROC_Field label="Tipologia"><select style={procInputStyle} value={f.tipologia} onChange={set('tipologia')}>{PROC_TIPOLOGIE.map(t => <option key={t}>{t}</option>)}</select></PROC_Field>
        <PROC_Field label="Descrizione breve"><input style={procInputStyle} value={f.descrizione} onChange={set('descrizione')} placeholder="Scopo e contesto" /></PROC_Field>
        <PROC_Field label="Procedura passo-passo"><textarea style={{ ...procInputStyle, minHeight:120, resize:'vertical' }} value={f.steps} onChange={set('steps')} placeholder={'1. Primo passo\n2. Secondo passo\n3. ...'} /></PROC_Field>

        <SectionLabel theme={PROC_COLORS}>Immagini</SectionLabel>
        <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginBottom:14 }}>
          {(f.immagini||[]).filter(img => !immaginiDaRimuovere.includes(img)).map((img, i) => (
            <div key={i} style={thumbStyle}><div style={{ width:'100%', height:'100%', display:'flex', alignItems:'center', justifyContent:'center' }}><ImageIcon size={20} color={PROC_COLORS.muted} /></div>{xOverlay(() => setImmaginiDaRimuovere(p => [...p, img]))}</div>
          ))}
          {nuoveImmagini.map((file, i) => (
            <div key={`n${i}`} style={thumbStyle}><img src={anteprimeImg[i]} alt="" style={{ width:'100%', height:'100%', objectFit:'cover' }} />{xOverlay(() => setNuoveImmagini(p => p.filter((_,idx)=>idx!==i)))}</div>
          ))}
          <button type="button" onClick={() => imgRef.current?.click()} style={{ width:72, height:72, borderRadius:10, border:`1.5px dashed ${PROC_COLORS.line}`, background:'none', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', color:PROC_COLORS.muted, gap:3, flexShrink:0 }}><Plus size={18}/><span style={{ fontSize:10 }}>Immagine</span></button>
          <input ref={imgRef} type="file" accept="image/*" multiple style={{ display:'none' }} onChange={e => { setNuoveImmagini(p=>[...p,...Array.from(e.target.files||[])]); e.target.value=''; }} />
        </div>

        <SectionLabel theme={PROC_COLORS}>Link video</SectionLabel>
        <div style={{ marginBottom:14 }}>
          {links.map((l, i) => (
            <div key={i} style={{ display:'flex', gap:8, marginBottom:8, alignItems:'center' }}>
              <input style={{ ...procInputStyle, flex:1, fontSize:13 }} value={l} onChange={e => setLinks(p => p.map((x,idx)=>idx===i?e.target.value:x))} placeholder="https://youtube.com/..." />
              {xBtn(() => setLinks(p => p.filter((_,idx)=>idx!==i)))}
            </div>
          ))}
          <button type="button" onClick={() => setLinks(p=>[...p,''])} style={{ fontSize:13, color:PROC_COLORS.primary, background:'none', border:'none', fontWeight:600, padding:0, display:'flex', alignItems:'center', gap:5 }}><Plus size={14}/> Aggiungi link</button>
        </div>

        <SectionLabel theme={PROC_COLORS}>File allegati</SectionLabel>
        <div style={{ marginBottom:14 }}>
          {(f.files||[]).filter(file => !filesDaRimuovere.find(x=>x.path===file.path)).map((file, i) => (
            <div key={i} style={{ display:'flex', alignItems:'center', gap:8, padding:'8px 0', borderBottom:`1px solid ${PROC_COLORS.line}` }}>
              <FileText size={14} color={PROC_COLORS.primary} />
              <span style={{ flex:1, fontSize:13 }}>{file.nome}</span>
              {xBtn(() => setFilesDaRimuovere(p=>[...p,file]))}
            </div>
          ))}
          {nuoviFiles.map((file, i) => (
            <div key={`n${i}`} style={{ display:'flex', alignItems:'center', gap:8, padding:'8px 0', borderBottom:`1px solid ${PROC_COLORS.line}` }}>
              <FileText size={14} color={PROC_COLORS.primary} />
              <span style={{ flex:1, fontSize:13 }}>{file.name}</span>
              {xBtn(() => setNuoviFiles(p=>p.filter((_,idx)=>idx!==i)))}
            </div>
          ))}
          <button type="button" onClick={() => fileRef.current?.click()} style={{ fontSize:13, color:PROC_COLORS.primary, background:'none', border:'none', fontWeight:600, padding:'8px 0', display:'flex', alignItems:'center', gap:5 }}><Plus size={14}/> Allega file</button>
          <input ref={fileRef} type="file" multiple style={{ display:'none' }} onChange={e => { setNuoviFiles(p=>[...p,...Array.from(e.target.files||[])]); e.target.value=''; }} />
        </div>
        <PROC_Field label="Note"><input style={procInputStyle} value={f.note||''} onChange={set('note')} /></PROC_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding:'0 16px' }}>
          {errore && <div style={{ background:'#F7DCD9', color:'#A3352A', padding:'10px 12px', borderRadius:10, fontSize:12.5, marginBottom:12 }}>{errore}</div>}
          <button onClick={handleSave} disabled={salvataggio} style={{ width:'100%', background:PROC_COLORS.primary, color:'#fff', border:'none', borderRadius:12, padding:'14px', fontWeight:700, fontSize:15, marginBottom:8, opacity:salvataggio?0.6:1 }}>{salvataggio?'Salvataggio…':'Salva procedura'}</button>
          {initial && puoEliminare && <button onClick={() => setConfermaElimina(true)} style={{ width:'100%', background:'none', border:'none', color:PROC_COLORS.danger, fontWeight:600, fontSize:13.5, padding:'8px 0 16px' }}>Elimina procedura</button>}
        </div>
      )}
      {confermaElimina && <ConfirmDelete theme={PROC_COLORS} message={`Eliminare "${initial?.titolo}"?`} onConfirm={() => { setConfermaElimina(false); onDelete(initial); }} onCancel={() => setConfermaElimina(false)} />}
    </>
  );
}

/* ---- Manutenzioni Ricorrenti ---- */
function RicorrentiScreen({ items, onOpen, onAdd, onHome }) {
  const lista = [...items].map(r => ({ ...r, days: r.prossimaScadenza ? daysUntil(r.prossimaScadenza) : null })).sort((a,b) => (a.days??9999)-(b.days??9999));
  return (
    <>
      <TopBar theme={PROC_COLORS} title="Manutenzioni ricorrenti" subtitle={`${items.length} attività`} onBack={onHome} backIcon={Home} />
      <div style={{ padding:14 }}>
        {lista.length === 0 && <Empty theme={PROC_COLORS} icon={RefreshCw} text="Nessuna manutenzione ricorrente." />}
        <div style={{ display:'flex', flexDirection:'column', gap:9 }}>
          {lista.map(r => {
            const u = r.days != null ? urgencyOf(r.days, { urgentDays:7, soonDays:30 }) : 'ok';
            return (
              <Card theme={PROC_COLORS} key={r.id} onClick={() => onOpen(r)}>
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', gap:8 }}>
                  <div style={{ minWidth:0 }}>
                    <div style={{ fontWeight:700, fontSize:14.5, marginBottom:3 }}>{r.titolo}</div>
                    <div style={{ fontSize:12, color:PROC_COLORS.muted, marginBottom:4 }}>{r.frequenza}{r.categoria ? ' · ' + r.categoria : ''}</div>
                    {r.ultimaEsecuzione && <div style={{ fontSize:12, color:PROC_COLORS.muted }}>Ultima: {fmtDate(r.ultimaEsecuzione)}</div>}
                  </div>
                  <div style={{ textAlign:'right', flexShrink:0 }}>
                    {r.days != null ? <><Pill style={URGENCY_STYLE[u]}>{r.days < 0 ? 'Scaduta' : `${r.days} gg`}</Pill><div style={{ fontSize:11, color:PROC_COLORS.muted, marginTop:5 }}>{fmtDate(r.prossimaScadenza)}</div></> : <Pill style={{ bg:PROC_COLORS.bg, fg:PROC_COLORS.muted }}>N/D</Pill>}
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      </div>
      <PROC_FAB onClick={onAdd} />
    </>
  );
}

function RicorrenteForm({ initial, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(() => { if (initial) { const { days, ...clean } = initial; return clean; } return { id:uid(), titolo:'', categoria:'', frequenza:'Mensile', descrizione:'', ultimaEsecuzione:'', prossimaScadenza:'', responsabile:'', note:'' }; });
  const set = (k) => (e) => {
    const v = e.target.value;
    setF(prev => {
      const n = { ...prev, [k]: v };
      if (k === 'ultimaEsecuzione' && v) n.prossimaScadenza = calcProssima(v, n.frequenza);
      if (k === 'frequenza' && prev.ultimaEsecuzione) n.prossimaScadenza = calcProssima(prev.ultimaEsecuzione, v);
      return n;
    });
  };
  function segnaOggi() {
    const oggi = todayISO();
    setF(prev => ({ ...prev, ultimaEsecuzione:oggi, prossimaScadenza:calcProssima(oggi, prev.frequenza) }));
  }
  return (
    <>
      <TopBar theme={PROC_COLORS} title={initial ? 'Modifica ricorrente' : 'Nuova ricorrente'} onBack={onCancel} />
      <div style={{ padding:16, pointerEvents:puoScrivere?'auto':'none', opacity:puoScrivere?1:0.65 }}>
        <PROC_Field label="Titolo *"><input style={procInputStyle} value={f.titolo} onChange={set('titolo')} placeholder="Es. Cambio bombole Sapio" /></PROC_Field>
        <PROC_Field label="Categoria"><input style={procInputStyle} value={f.categoria||''} onChange={set('categoria')} placeholder="Es. Gas medicale, Sicurezza…" /></PROC_Field>
        <PROC_Field label="Frequenza"><select style={procInputStyle} value={f.frequenza} onChange={set('frequenza')}>{PROC_FREQUENZE.map(fr => <option key={fr}>{fr}</option>)}</select></PROC_Field>
        <PROC_Field label="Descrizione"><textarea style={{ ...procInputStyle, minHeight:70, resize:'vertical' }} value={f.descrizione||''} onChange={set('descrizione')} /></PROC_Field>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
          <PROC_Field label="Ultima esecuzione"><input type="date" style={procInputStyle} value={f.ultimaEsecuzione||''} onChange={set('ultimaEsecuzione')} /></PROC_Field>
          <PROC_Field label="Prossima scadenza"><input type="date" style={procInputStyle} value={f.prossimaScadenza||''} onChange={set('prossimaScadenza')} /></PROC_Field>
        </div>
        <button type="button" onClick={segnaOggi} style={{ width:'100%', background:PROC_COLORS.bg, border:`1.5px solid ${PROC_COLORS.line}`, borderRadius:10, padding:'11px', fontWeight:700, fontSize:13.5, color:PROC_COLORS.primary, marginBottom:14, display:'flex', alignItems:'center', justifyContent:'center', gap:7 }}>
          <Check size={15} /> Segna come eseguita oggi
        </button>
        <PROC_Field label="Responsabile"><input style={procInputStyle} value={f.responsabile||''} onChange={set('responsabile')} /></PROC_Field>
        <PROC_Field label="Note"><input style={procInputStyle} value={f.note||''} onChange={set('note')} /></PROC_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding:'0 16px' }}>
          <button onClick={() => onSave(f)} style={{ width:'100%', background:PROC_COLORS.primary, color:'#fff', border:'none', borderRadius:12, padding:'14px', fontWeight:700, fontSize:15, marginBottom:8 }}>Salva</button>
          {initial && puoEliminare && <button onClick={() => setConfermaElimina(true)} style={{ width:'100%', background:'none', border:'none', color:PROC_COLORS.danger, fontWeight:600, fontSize:13.5, padding:'8px 0 16px' }}>Elimina</button>}
        </div>
      )}
      {confermaElimina && <ConfirmDelete theme={PROC_COLORS} message={`Eliminare "${initial?.titolo}"?`} onConfirm={() => { setConfermaElimina(false); onDelete(initial); }} onCancel={() => setConfermaElimina(false)} />}
    </>
  );
}

/* ---- Armadi ---- */
function ArmadiScreen({ armadi, onOpen, onAdd, onHome }) {
  const [q, setQ] = useState('');
  const filtered = armadi.filter(a => `${a.numero} ${a.posizione} ${a.tipologia}`.toLowerCase().includes(q.toLowerCase()));
  return (
    <>
      <TopBar theme={PROC_COLORS} title="Contenuto armadi" subtitle={`${armadi.length} armadi`} onBack={onHome} backIcon={Home} />
      <div style={{ padding:14 }}>
        <input placeholder="Cerca numero, posizione, tipologia…" style={{ ...procInputStyle, marginBottom:12 }} value={q} onChange={e => setQ(e.target.value)} />
        {filtered.length === 0 && <Empty theme={PROC_COLORS} icon={Package} text="Nessun armadio registrato." />}
        <div style={{ display:'flex', flexDirection:'column', gap:9 }}>
          {filtered.map(a => (
            <Card theme={PROC_COLORS} key={a.id} onClick={() => onOpen(a)}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', gap:8 }}>
                <div>
                  <div style={{ fontWeight:700, fontSize:14.5, marginBottom:3 }}>Armadio {a.numero}</div>
                  <div style={{ fontSize:12, color:PROC_COLORS.muted }}>{a.posizione}</div>
                  {a.contenuto?.length > 0 && <div style={{ fontSize:11.5, color:PROC_COLORS.muted, marginTop:4 }}>{a.contenuto.length} voci</div>}
                </div>
                <Pill style={{ bg:PROC_COLORS.bg, fg:PROC_COLORS.primary }}>{a.tipologia||'Generale'}</Pill>
              </div>
            </Card>
          ))}
        </div>
      </div>
      <PROC_FAB onClick={onAdd} />
    </>
  );
}

function ArmadioDetail({ armadio, onBack, onEdit }) {
  const { isAdmin } = usePermessi();
  const puoScrivere = isAdmin;
  return (
    <>
      <TopBar theme={PROC_COLORS} title={`Armadio ${armadio.numero}`} subtitle={armadio.tipologia||armadio.posizione} onBack={onBack}
        right={puoScrivere ? <button onClick={onEdit} style={{ background:'none', border:'none', color:PROC_COLORS.primary, fontWeight:700, fontSize:13.5, padding:'4px 8px' }}>Modifica</button> : null} />
      <div style={{ padding:14 }}>
        <Card theme={PROC_COLORS} style={{ marginBottom:12 }}>
          <InfoRow theme={PROC_COLORS} icon={MapPin} label="Posizione" value={armadio.posizione||'—'} />
          <InfoRow theme={PROC_COLORS} icon={Hash} label="Tipologia" value={armadio.tipologia||'—'} />
        </Card>
        {armadio.contenuto?.length > 0 && (
          <>
            <SectionLabel theme={PROC_COLORS}>Contenuto</SectionLabel>
            <Card theme={PROC_COLORS} style={{ marginBottom:12 }}>
              <div style={{ display:'grid', gridTemplateColumns:'1fr auto auto', gap:'6px 12px', fontSize:12, fontWeight:700, color:PROC_COLORS.muted, paddingBottom:8, borderBottom:`1px solid ${PROC_COLORS.line}`, marginBottom:4 }}>
                <span>Elemento</span><span style={{ textAlign:'right' }}>Qtà</span><span>U.M.</span>
              </div>
              {armadio.contenuto.map((r, i) => (
                <div key={i} style={{ display:'grid', gridTemplateColumns:'1fr auto auto', gap:'4px 12px', fontSize:13.5, padding:'6px 0', borderTop:i?`1px solid ${PROC_COLORS.line}`:'none' }}>
                  <span>{r.elemento}</span>
                  <span style={{ textAlign:'right', fontWeight:700 }}>{r.quantita}</span>
                  <span style={{ color:PROC_COLORS.muted }}>{r.unita}</span>
                </div>
              ))}
            </Card>
          </>
        )}
        {armadio.note && (<><SectionLabel theme={PROC_COLORS}>Note</SectionLabel><Card theme={PROC_COLORS}><p style={{ margin:0, fontSize:13.5, color:PROC_COLORS.muted }}>{armadio.note}</p></Card></>)}
      </div>
    </>
  );
}

function ArmadioForm({ initial, onSave, onCancel, onDelete }) {
  const { isAdmin, puoEliminare } = usePermessi();
  const puoScrivere = isAdmin;
  const [confermaElimina, setConfermaElimina] = useState(false);
  const [f, setF] = useState(initial || { id:uid(), numero:'', posizione:'', tipologia:'Generale', note:'' });
  const [righe, setRighe] = useState(initial?.contenuto || []);
  const set = (k) => (e) => setF(prev => ({ ...prev, [k]:e.target.value }));
  const setR = (i, campo, val) => setRighe(p => p.map((r,idx) => idx===i ? { ...r, [campo]:val } : r));
  return (
    <>
      <TopBar theme={PROC_COLORS} title={initial ? 'Modifica armadio' : 'Nuovo armadio'} onBack={onCancel} />
      <div style={{ padding:16, pointerEvents:puoScrivere?'auto':'none', opacity:puoScrivere?1:0.65 }}>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
          <PROC_Field label="Numero *"><input style={procInputStyle} value={f.numero} onChange={set('numero')} placeholder="Es. A01" /></PROC_Field>
          <PROC_Field label="Tipologia"><select style={procInputStyle} value={f.tipologia} onChange={set('tipologia')}>{PROC_ARMADI_TIPI.map(t => <option key={t}>{t}</option>)}</select></PROC_Field>
        </div>
        <PROC_Field label="Posizione"><input style={procInputStyle} value={f.posizione||''} onChange={set('posizione')} placeholder="Es. Piano Terra - Corridoio Nord" /></PROC_Field>
        <SectionLabel theme={PROC_COLORS}>Contenuto</SectionLabel>
        {righe.length > 0 && (
          <div style={{ marginBottom:8 }}>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 70px 80px 36px', gap:6, marginBottom:6 }}>
              {['Elemento','Qtà','Unità',''].map((h,i) => <span key={i} style={{ fontSize:11, fontWeight:600, color:PROC_COLORS.muted, textTransform:'uppercase' }}>{h}</span>)}
            </div>
            {righe.map((r, i) => (
              <div key={i} style={{ display:'grid', gridTemplateColumns:'1fr 70px 80px 36px', gap:6, marginBottom:6 }}>
                <input style={{ ...procInputStyle, padding:'8px 10px', fontSize:13 }} value={r.elemento} onChange={e => setR(i,'elemento',e.target.value)} placeholder="Es. Fusibili 16A" />
                <input type="number" style={{ ...procInputStyle, padding:'8px 10px', fontSize:13 }} value={r.quantita} onChange={e => setR(i,'quantita',e.target.value)} />
                <select style={{ ...procInputStyle, padding:'8px 6px', fontSize:13 }} value={r.unita||'pz'} onChange={e => setR(i,'unita',e.target.value)}>{PROC_UNITA.map(u => <option key={u}>{u}</option>)}</select>
                <button type="button" onClick={() => setRighe(p => p.filter((_,idx) => idx!==i))} style={{ background:'none', border:'none', color:PROC_COLORS.danger, padding:4, display:'flex', alignItems:'center', justifyContent:'center' }}><XIcon size={15} /></button>
              </div>
            ))}
          </div>
        )}
        <button type="button" onClick={() => setRighe(p => [...p, { elemento:'', quantita:'', unita:'pz' }])} style={{ fontSize:13, color:PROC_COLORS.primary, background:'none', border:'none', fontWeight:600, padding:'4px 0', display:'flex', alignItems:'center', gap:5, marginBottom:14 }}><Plus size={14}/> Aggiungi voce</button>
        <PROC_Field label="Note"><input style={procInputStyle} value={f.note||''} onChange={set('note')} /></PROC_Field>
      </div>
      {puoScrivere && (
        <div style={{ padding:'0 16px' }}>
          <button onClick={() => onSave({ ...f, contenuto:righe })} style={{ width:'100%', background:PROC_COLORS.primary, color:'#fff', border:'none', borderRadius:12, padding:'14px', fontWeight:700, fontSize:15, marginBottom:8 }}>Salva armadio</button>
          {initial && puoEliminare && <button onClick={() => setConfermaElimina(true)} style={{ width:'100%', background:'none', border:'none', color:PROC_COLORS.danger, fontWeight:600, fontSize:13.5, padding:'8px 0 16px' }}>Elimina armadio</button>}
        </div>
      )}
      {confermaElimina && <ConfirmDelete theme={PROC_COLORS} message={`Eliminare armadio ${initial?.numero}?`} onConfirm={() => { setConfermaElimina(false); onDelete(initial); }} onCancel={() => setConfermaElimina(false)} />}
    </>
  );
}

/* ---- Modulo Root ---- */
function ProcedureModule({ onHome }) {
  const procT  = useSupaTable('procedure_manuali', 'id', []);
  const ricT   = useSupaTable('manutenzioni_ricorrenti', 'id', [], ['ultimaEsecuzione','prossimaScadenza']);
  const armT   = useSupaTable('armadi', 'id', []);
  const ready  = procT.ready && ricT.ready && armT.ready;
  const [tab, setTab]   = useState('procedure');
  const [view, setView] = useState({ name:'list' });
  const [toast, setToast] = useState('');

  useBackable(tab, setTab);
  useBackable(view, setView);
  useEffect(() => { setView({ name:'list' }); }, [tab]);

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(''), 2400); };
  const tableOf = { procedure_manuali:procT, manutenzioni_ricorrenti:ricT, armadi:armT };

  async function salva(tbl, record, msg) {
    const { error } = await tableOf[tbl].save(record);
    if (error) { flash('Errore: ' + error.message); return; }
    flash(msg); goBack();
  }
  async function elimina(tbl, record, msg) {
    const { error } = await tableOf[tbl].remove(record);
    if (error) { flash('Errore: ' + error.message); return; }
    flash(msg); goBack();
  }

  if (!ready) return <div style={{ minHeight:'100vh', background:PROC_COLORS.bg, display:'flex', alignItems:'center', justifyContent:'center', color:PROC_COLORS.muted }}>Caricamento…</div>;

  let content;
  if (tab === 'procedure') {
    const proc = procT.rows;
    if (view.name === 'detail') { const p = proc.find(x => x.id === view.id); content = p ? <ProceduraDetail procedura={p} onBack={() => goBack()} onEdit={() => setView({ name:'edit', p })} onDelete={r => elimina('procedure_manuali', r, 'Eliminata')} /> : null; }
    else if (view.name === 'add') content = <ProceduraForm onSave={r => salva('procedure_manuali', r, 'Procedura salvata')} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <ProceduraForm initial={view.p} onSave={r => salva('procedure_manuali', r, 'Aggiornata')} onCancel={() => goBack()} onDelete={r => elimina('procedure_manuali', r, 'Eliminata')} />;
    else content = <ProcedureScreen procedure={proc} onOpen={p => setView({ name:'detail', id:p.id })} onAdd={() => setView({ name:'add' })} onHome={onHome} />;
  } else if (tab === 'ricorrenti') {
    if (view.name === 'add') content = <RicorrenteForm onSave={r => salva('manutenzioni_ricorrenti', r, 'Salvata')} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <RicorrenteForm initial={view.r} onSave={r => salva('manutenzioni_ricorrenti', r, 'Aggiornata')} onCancel={() => goBack()} onDelete={r => elimina('manutenzioni_ricorrenti', r, 'Eliminata')} />;
    else content = <RicorrentiScreen items={ricT.rows} onOpen={r => setView({ name:'edit', r })} onAdd={() => setView({ name:'add' })} onHome={onHome} />;
  } else {
    const arm = armT.rows;
    if (view.name === 'detail') { const a = arm.find(x => x.id === view.id); content = a ? <ArmadioDetail armadio={a} onBack={() => goBack()} onEdit={() => setView({ name:'edit', a })} /> : null; }
    else if (view.name === 'add') content = <ArmadioForm onSave={r => salva('armadi', r, 'Armadio salvato')} onCancel={() => goBack()} />;
    else if (view.name === 'edit') content = <ArmadioForm initial={view.a} onSave={r => salva('armadi', r, 'Aggiornato')} onCancel={() => goBack()} onDelete={r => elimina('armadi', r, 'Eliminato')} />;
    else content = <ArmadiScreen armadi={arm} onOpen={a => setView({ name:'detail', id:a.id })} onAdd={() => setView({ name:'add' })} onHome={onHome} />;
  }

  return (
    <div style={{ minHeight:'100vh', background:PROC_COLORS.bg, fontFamily:'Inter, sans-serif', color:PROC_COLORS.ink, maxWidth:480, margin:'0 auto', position:'relative' }}>
      <style>{GLOBAL_FONTS}</style>
      <div style={{ paddingBottom:78 }}>{content}</div>
      {view.name === 'list' && <BottomNav theme={PROC_COLORS} tab={tab} setTab={setTab} items={PROC_NAV_ITEMS} />}
      {toast && <div style={{ position:'fixed', bottom:view.name==='list'?92:20, left:'50%', transform:'translateX(-50%)', background:PROC_COLORS.primaryDeep, color:'#fff', padding:'10px 18px', borderRadius:999, fontSize:13, fontWeight:600, display:'flex', alignItems:'center', gap:7, zIndex:30 }}><Check size={15} /> {toast}</div>}
    </div>
  );
}

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
  {
    key: 'struttura',
    name: 'Struttura',
    desc: 'Camere, interventi, scadenze e costi',
    icon: BedDouble,
    color: '#6E4A2E',
    colorSoft: '#EDE3D6',
    stat: (d) => `${d.camere.length} camere`,
  },
  {
    key: 'procedure',
    name: 'Procedure',
    desc: 'Manuali operativi, ricorrenti e armadi',
    icon: BookOpen,
    color: '#5E3A8A',
    colorSoft: '#E8E0F5',
    stat: () => 'Procedure · Ricorrenti · Armadi',
  },
];

function HubScreen({ onOpen, counts, nomeVisualizzato, role, onSignOut, onOpenUsers, onOpenProfilo, onOpenNotifiche }) {
  return (
    <div style={{ minHeight: '100vh', background: HUB_COLORS.bg, fontFamily: 'Inter, sans-serif', color: HUB_COLORS.ink, maxWidth: 480, margin: '0 auto' }}>
      <style>{GLOBAL_FONTS}</style>
      <div style={{ background: '#1C2321', padding: '34px 20px 30px', color: '#fff' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
          <div style={{ width: 46, height: 46, borderRadius: 13, background: 'rgba(255,255,255,0.12)', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 14 }}>
            <WrenchHub size={24} />
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            <NotificationBell />
            {role === 'admin' && (
              <button onClick={onOpenUsers} title="Gestione utenti" style={{ background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: 10, width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff' }}>
                <UserCog size={17} />
              </button>
            )}
            {role === 'admin' && (
              <button onClick={onOpenNotifiche} title="Gestione notifiche" style={{ background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: 10, width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 16 }}>
                🔕
              </button>
            )}
            <button onClick={onSignOut} title="Esci" style={{ background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: 10, width: 36, height: 36, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff' }}>
              <LogOut size={16} />
            </button>
          </div>
        </div>
        <div style={{ fontFamily: "'Archivo', sans-serif", fontWeight: 800, fontSize: 26 }}>Manutenzione</div>
        <div style={{ fontSize: 13.5, opacity: 0.7, marginTop: 4 }}>Scegli un modulo per iniziare</div>
        <button
          onClick={onOpenProfilo}
          style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginTop: 12, background: 'rgba(255,255,255,0.08)', padding: '4px 10px 4px 8px', borderRadius: 999, border: 'none', color: '#fff', cursor: 'pointer' }}
        >
          <ShieldCheck size={13} />
          <span style={{ fontSize: 11.5, opacity: 0.85 }}>{nomeVisualizzato} · {RUOLO_LABEL[role] || role}</span>
        </button>
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

/* ============================================================
   NOTIFICHE PUSH — Service Worker + Web Push
   ============================================================ */

const VAPID_PUBLIC_KEY = import.meta.env.VITE_VAPID_PUBLIC_KEY || '';

function urlBase64ToUint8Array(base64String) {
  try {
    const padding = '='.repeat((4 - base64String.length % 4) % 4);
    const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
    const rawData = window.atob(base64);
    return Uint8Array.from([...rawData].map((char) => char.charCodeAt(0)));
  } catch (e) {
    if (process.env.NODE_ENV === 'development') console.warn('Invalid VAPID key:', e);
    return null;
  }
}

async function registerPush(userId) {
  if (!('serviceWorker' in navigator) || !('PushManager' in window) || !VAPID_PUBLIC_KEY) return;
  try {
    // Se userId non passato, ottieni dalla sessione attiva
    if (!userId) {
      const { data: sessionData } = await supabase.auth.getSession();
      userId = sessionData?.session?.user?.id;
    }
    if (!userId) return;

    const vapidKey = urlBase64ToUint8Array(VAPID_PUBLIC_KEY);
    if (!vapidKey) return;
    const registration = await navigator.serviceWorker.register((import.meta.env.BASE_URL || '/') + 'sw.js');
    await navigator.serviceWorker.ready;

    const existing = await registration.pushManager.getSubscription();
    const subscription = existing || await registration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: vapidKey
    });

    const json = subscription.toJSON();
    if (!json.endpoint || !json.keys?.p256dh || !json.keys?.auth) {
      throw new Error('Subscription push non valida');
    }

    const { error } = await supabase
      .from('push_subscriptions')
      .upsert({
        user_id: userId,
        endpoint: json.endpoint,
        p256dh: json.keys.p256dh,
        auth: json.keys.auth,
      }, { onConflict: 'endpoint' });

    if (error) {
      if (process.env.NODE_ENV === 'development') console.error('Errore salvataggio subscription push:', error);
    }
  } catch (e) {
    if (process.env.NODE_ENV === 'development') console.error('Push registration failed:', e);
  }
}

function playNotificationSound() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.type = 'sine';
    osc.frequency.setValueAtTime(880, ctx.currentTime);
    osc.frequency.setValueAtTime(1100, ctx.currentTime + 0.1);
    gain.gain.setValueAtTime(0.3, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.4);
    osc.start();
    osc.stop(ctx.currentTime + 0.4);
    if (navigator.vibrate) navigator.vibrate([200, 100, 200]);
  } catch (e) {
    if (process.env.NODE_ENV === 'development') console.warn('playNotificationSound failed:', e);
  }
}

export default function ManutenzioneApp() {
  if (!supabaseConfigured) {
    return (
      <div style={{ minHeight: '100vh', background: '#1C2321', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: 'sans-serif', textAlign: 'center' }}>
        <div style={{ maxWidth: 380 }}>
          <div style={{ fontWeight: 800, fontSize: 20, marginBottom: 10 }}>Configurazione mancante</div>
          <div style={{ fontSize: 14, opacity: 0.8, lineHeight: 1.5 }}>
            Mancano le variabili <code>VITE_SUPABASE_URL</code> e <code>VITE_SUPABASE_ANON_KEY</code>.<br /><br />
            Su GitHub: <b>Settings → Secrets and variables → Actions</b>, verifica che i due secret esistano con questi nomi esatti, poi rilancia il deploy (<b>Actions → il workflow più recente → Re-run all jobs</b>).
          </div>
        </div>
      </div>
    );
  }

  const [screen, setScreen] = useState('hub'); // 'hub' | 'mezzi' | 'carrozzine' | 'struttura' | 'utenti' | 'profilo' | 'notifiche'
  const [counts, setCounts] = useState({ vehicles: [], carrozzine: [], camere: [] });
  const { session, profile, refreshProfile, passwordRecovery, clearPasswordRecovery, authLoading, signOut } = useAuth();
  const notificheCtx = useNotifications(session?.user?.id || null, profile?.role || 'lettore');

  useEffect(() => {
    if (session?.user?.id) registerPush(session.user.id);
  }, [session?.user?.id]);

  useHardwareBack();
  useBackable(screen, setScreen);

  // Carica conteggi reali da Supabase per l'Hub
  useEffect(() => {
    if (!session) return;
    let mounted = true;
    (async () => {
      const [vRes, cRes, caRes] = await Promise.all([
        supabase.from('vehicles').select('*', { count: 'exact', head: true }),
        supabase.from('carrozzine').select('*', { count: 'exact', head: true }),
        supabase.from('camere').select('*', { count: 'exact', head: true }),
      ]);
      if (!mounted) return;
      const vc = vRes.count ?? 0;
      const cc = cRes.count ?? 0;
      const ca = caRes.count ?? 0;
      setCounts({
        vehicles: Array(vc || 0).fill(null),
        carrozzine: Array(cc || 0).fill(null),
        camere: Array(ca || 0).fill(null),
      });
    })();
    return () => { mounted = false; };
  }, [session]);

  if (authLoading) {
    return (
      <div style={{ minHeight: '100vh', background: '#1C2321', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13.5, opacity: 0.7 }}>
        Caricamento…
      </div>
    );
  }
  if (!session) return <AuthScreen />;
  if (passwordRecovery) return <NewPasswordScreen onDone={clearPasswordRecovery} />;

  const role = profile?.role || 'lettore';
  const nome = profile?.nome || '';
  const cognome = profile?.cognome || '';

  return (
    <NotificheContext.Provider value={notificheCtx}>
    <RoleContext.Provider value={{ role, email: session.user.email, nome, cognome }}>
      {screen === 'mezzi' && <MezziModule onHome={() => setScreen('hub')} />}
      {screen === 'carrozzine' && <CarrozzineModule onHome={() => setScreen('hub')} />}
      {screen === 'struttura' && <StrutturaModule onHome={() => setScreen('hub')} />}
      {screen === 'procedure' && <ProcedureModule onHome={() => setScreen('hub')} />}
      {screen === 'utenti' && <UtentiScreen onHome={() => setScreen('hub')} myUserId={session.user.id} />}
      {screen === 'gestione-notifiche' && <GestioneNotificheScreen onHome={() => setScreen('hub')} myUserId={session.user.id} />}
      {screen === 'profilo' && (
        <ProfiloScreen onHome={() => setScreen('hub')} profile={{ ...profile, email: session.user.email }} onSaved={refreshProfile} />
      )}
      {screen === 'hub' && (
        <HubScreen
          onOpen={setScreen}
          counts={counts}
          nomeVisualizzato={nomeVisualizzato({ nome, cognome, email: session.user.email })}
          role={role}
          onSignOut={signOut}
          onOpenUsers={() => setScreen('utenti')}
          onOpenProfilo={() => setScreen('profilo')}
          onOpenNotifiche={() => setScreen('gestione-notifiche')}
        />
      )}
    </RoleContext.Provider>
    </NotificheContext.Provider>
  );
}