/* =========================================================================
   GESTÃO DE SALAS DE REUNIÃO — App React (sem build step, via Babel/CDN)
   ========================================================================= */
const { useState, useEffect, useMemo, useCallback, useRef } = React;
const DB = window.DB;
const CFG = window.APP_CONFIG;

// Defina os domínios de e-mail autorizados no config.js (ALLOWED_DOMAINS)
// ou deixe vazio para aceitar qualquer domínio.
const ALLOWED_DOMAINS = (window.APP_CONFIG && window.APP_CONFIG.ALLOWED_DOMAINS) || [];
const validDomain = (email) => ALLOWED_DOMAINS.includes((email.split("@")[1] || "").toLowerCase());

const toggleSt  = { flex: 1, border: "none", background: "transparent", padding: "8px 12px", borderRadius: "999px", fontSize: 13.5, fontWeight: 600, cursor: "pointer", color: "var(--ink-2)", transition: "all .15s" };
const toggleAct = { background: "#fff", color: "var(--brand)", boxShadow: "0 1px 4px rgba(0,0,0,.12)" };

/* ----------------------------- Helpers de data ------------------------- */
const pad = (n) => String(n).padStart(2, "0");
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const fmtDate = (iso) => new Date(String(iso).length <= 10 ? iso + "T12:00:00" : iso).toLocaleDateString("pt-BR", { day: "2-digit", month: "short", year: "numeric" });
const fmtTime = (iso) => new Date(iso).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
const fmtDateLong = (d) => d.toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "long", year: "numeric" });

const toISO = (dateStr, timeStr) => new Date(`${dateStr}T${timeStr}:00`).toISOString();
const localDateStr = (iso) => { const d = new Date(iso); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; };
const OPEN_HOUR = 7, CLOSE_HOUR = 20;

/* ---- Perfil do usuário (localStorage) ---- */
const PROFILE_KEY = "pluma_user_profile";
const getProfile  = () => { try { return JSON.parse(localStorage.getItem(PROFILE_KEY) || "null"); } catch { return null; } };
const saveProfile = (p) => localStorage.setItem(PROFILE_KEY, JSON.stringify(p));

const EMAIL_HIST_KEY = "pluma_email_history";
const getEmailHistory  = () => { try { return JSON.parse(localStorage.getItem(EMAIL_HIST_KEY) || "[]"); } catch { return []; } };
const saveEmailHistory = (emails) => {
  const prev = getEmailHistory();
  const merged = [...new Set([...emails, ...prev])].slice(0, 100); /* mantém até 100 e-mails únicos */
  localStorage.setItem(EMAIL_HIST_KEY, JSON.stringify(merged));
};

/* ---- Helpers de calendário ---- */
const fmtICSDate = (iso) => iso.replace(/[-:]/g, "").replace(/\.\d{3}/, "");

const generateICS = ({ title, start, end, location, description, organizer }) => {
  const lines = [
    "BEGIN:VCALENDAR", "VERSION:2.0",
    "PRODID:-//Salas de Reuniao//PT", "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
    "BEGIN:VEVENT",
    `DTSTART:${fmtICSDate(start)}`, `DTEND:${fmtICSDate(end)}`,
    `SUMMARY:${title}`,
    `UID:${Date.now()}@salas-reuniao`,
  ];
  if (location)    lines.push(`LOCATION:${location}`);
  if (description) lines.push(`DESCRIPTION:${description.replace(/\n/g, "\\n")}`);
  if (organizer)   lines.push(`ORGANIZER;CN=${organizer}:MAILTO:${organizer}`);
  lines.push("END:VEVENT", "END:VCALENDAR");
  return lines.join("\r\n");
};

const downloadICS = (ics, filename = "reuniao.ics") => {
  const blob = new Blob([ics], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
};

const teamsCalendarUrl = ({ title, start, end, location, organizer }) => {
  const p = new URLSearchParams({ subject: title, startTime: start, endTime: end, content: `Local: ${location || ""}` });
  if (organizer) p.append("attendees", organizer);
  return `https://teams.microsoft.com/l/meeting/new?${p.toString()}`;
};

const outlookWebUrl = ({ title, start, end, location, description }) =>
  `https://outlook.office.com/calendar/deeplink/compose?path=/calendar/action/compose&rru=addevent` +
  `&subject=${encodeURIComponent(title)}&startdt=${encodeURIComponent(start)}&enddt=${encodeURIComponent(end)}` +
  `&body=${encodeURIComponent(description || "")}&location=${encodeURIComponent(location || "")}`;

// Atualiza o horário atual a cada minuto para que os badges Livre/Ocupada reflitam o tempo real
function useNow() {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 60000);
    return () => clearInterval(id);
  }, []);
  return now;
}

const ICONS = {
  home: '<path d="M3.5 10.5 12 4l8.5 6.5"/><path d="M5.5 9.2V20h13V9.2"/><path d="M10 20v-5h4v5"/>',
  check: '<circle cx="12" cy="12" r="8.5"/><path d="m8.5 12 2.4 2.4L16 9"/>',
  ok: '<circle cx="12" cy="12" r="8.5"/><path d="m8.5 12 2.4 2.4L16 9"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  list: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><path d="M7.5 9h9M7.5 12.2h9M7.5 15.4h5.5"/>',
  settings: '<path d="M4 7.5h9M17 7.5h3M4 16.5h3M11 16.5h9"/><circle cx="15" cy="7.5" r="2"/><circle cx="8" cy="16.5" r="2"/>',
  building: '<rect x="5" y="3.5" width="14" height="17" rx="1.5"/><path d="M9 7.5h2M13 7.5h2M9 11h2M13 11h2M9 14.5h2M13 14.5h2M10.5 20.5v-3h3v3"/>',
  layout: '<rect x="4" y="4.5" width="16" height="15" rx="2.5"/><path d="M4 9.5h16M10 9.5v10"/>',
  calendar: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 9.5h17M8 3.5v3M16 3.5v3"/>',
  calendarPlus: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 9.5h17M8 3.5v3M16 3.5v3M12 13v4M10 15h4"/>',
  logo: '<g transform="rotate(-42 11 12)"><path fill="currentColor" stroke="none" d="M17 5A9 9 0 1 0 17 19L15 15A5 5 0 1 1 15 9Z"/><circle fill="currentColor" stroke="none" cx="14" cy="10" r="2.4"/></g>',
  users: '<circle cx="9" cy="8" r="3"/><path d="M3.8 18.5c0-2.8 2.3-4.6 5.2-4.6s5.2 1.8 5.2 4.6M16 5.5a3 3 0 0 1 0 5M20.2 18.5c0-2-1.2-3.6-3.2-4.2"/>',
  pin: '<path d="M12 20.5s6.5-5 6.5-10.3A6.5 6.5 0 0 0 5.5 10.2C5.5 15.5 12 20.5 12 20.5Z"/><circle cx="12" cy="10" r="2.3"/>',
  clock: '<circle cx="12" cy="12" r="8"/><path d="M12 7.8V12l2.8 1.8"/>',
  edit: '<path d="M4 20h4.2L19 9.2 14.8 5 4 15.8V20Z"/><path d="M13.5 6.3 17.7 10.5"/>',
  trash: '<path d="M5 7h14M9.5 7V5.2h5V7M7.5 7l.8 12.5h7.4L16.5 7"/>',
  warn: '<path d="M12 4 2.8 20h18.4L12 4Z"/><path d="M12 10v4.5M12 17.3v.4"/>',
  info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5M12 8v.4"/>',
  inbox: '<path d="M3.5 13.5 6 5.5h12l2.5 8M3.5 13.5V19h17v-5.5M3.5 13.5h5l1.5 2.5h4l1.5-2.5h5"/>',
  menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  x: '<path d="m6 6 12 12M18 6 6 18"/>',
  share: '<circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="m8.6 13.5 6.8 3.9M15.4 6.6 8.6 10.5"/>',
  mail: '<rect x="3" y="5.5" width="18" height="13" rx="2.5"/><path d="m3 6.5 9 6.5 9-6.5"/>',
  phone: '<path d="M6.6 10.8a15 15 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.25 11.4 11.4 0 0 0 3.6.6 1 1 0 0 1 1 1V17a1 1 0 0 1-1 1A16 16 0 0 1 5 6a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1c0 1.26.2 2.47.6 3.6a1 1 0 0 1-.24 1L6.6 10.8Z"/>',
  whatsapp: '<path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5Z"/>',
  teams: '<path d="M17 10h3a1 1 0 0 1 1 1v5a2 2 0 0 1-2 2h-2M3 10h11v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-7Z"/><circle cx="9" cy="5" r="2.5"/><circle cx="18" cy="5" r="2"/>',
  outlook: '<rect x="2" y="4" width="20" height="16" rx="2.5"/><path d="M2 9h20M8 4v16"/><circle cx="15" cy="14" r="2.5"/><path d="M17.5 14h2M15 11.5v-1"/>',
  download: '<path d="M12 3v12M8 11l4 4 4-4"/><path d="M3 18h18"/>',
  send: '<path d="M22 2L11 13"/><path d="M22 2L15 22l-4-9-9-4 20-7z"/>',
  compose: '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>',
  user: '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
  arrowRight: '<path d="M5 12h14M12 5l7 7-7 7"/>',
  arrowLeft: '<path d="M19 12H5M12 19l-7-7 7-7"/>'
};
function Icon({ name, size = 18 }) {
  return <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" dangerouslySetInnerHTML={{ __html: ICONS[name] || "" }} />;
}

/* ----------------------------- UI primitives --------------------------- */
function Modal({ title, onClose, children, footer }) {
  return (
    <div className="modal-bg" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" role="dialog" aria-modal="true">
        <div className="m-head">
          <h3>{title}</h3>
          <button className="x-btn" onClick={onClose} aria-label="Fechar"><Icon name="x" size={18} /></button>
        </div>
        <div className="m-body">{children}</div>
        {footer && <div className="m-foot">{footer}</div>}
      </div>
    </div>
  );
}

function Field({ label, hint, children }) {
  return (
    <div className="field">
      {label && <label>{label}</label>}
      {children}
      {hint && <span className="hint">{hint}</span>}
    </div>
  );
}

function Stat({ ico, color, num, lbl }) {
  return (
    <div className="card"><div className="card-body stat">
      <div className={`ico ${color}`}><Icon name={ico} size={22} /></div>
      <div><div className="num">{num}</div><div className="lbl">{lbl}</div></div>
    </div></div>
  );
}

function RoomBadge({ busy }) {
  return <span className={`badge ${busy ? "busy" : "free"}`}><span className="bdot"></span>{busy ? "Ocupada" : "Livre"}</span>;
}

/* ============================== LOGIN ================================== */
function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState("login");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);

  const submit = async (e) => {
    e.preventDefault();
    setMsg(null);
    if (!validDomain(email))
      return setMsg({ type: "err", text: "E-mail não pertence a um domínio autorizado." });
    if (mode === "register" && !name.trim())
      return setMsg({ type: "err", text: "Informe seu nome completo." });
    if (password.length < 6)
      return setMsg({ type: "err", text: "A senha deve ter pelo menos 6 caracteres." });
    setBusy(true);
    try {
      if (mode === "register") {
        const data = await DB.signUp(email.trim(), password, name.trim());
        if (data.session) { onLogin(data.session.user); }
        else setMsg({ type: "ok", text: "Cadastro realizado! Verifique seu e-mail para confirmar antes de entrar." });
      } else {
        const data = await DB.signIn(email.trim(), password);
        if (!data.user) return setMsg({ type: "err", text: "Confirme seu e-mail antes de entrar." });
        onLogin(data.user);
      }
    } catch (err) {
      const m = err.message || "";
      if (m.includes("Invalid login credentials")) setMsg({ type: "err", text: "E-mail ou senha incorretos." });
      else if (m.includes("already registered"))    setMsg({ type: "err", text: "E-mail já cadastrado. Faça login." });
      else if (m.includes("Email not confirmed"))   setMsg({ type: "err", text: "Confirme seu e-mail antes de entrar." });
      else setMsg({ type: "err", text: m || "Erro ao autenticar." });
    } finally { setBusy(false); }
  };

  return (
    <div style={{ minHeight: "100vh", background: "var(--brand-deep)", display: "grid", placeItems: "center", padding: "24px", position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", inset: 0, background: "url('bg.png') center/cover no-repeat", opacity: .05, pointerEvents: "none" }} />
      <div style={{ width: "100%", maxWidth: 420, position: "relative" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ width: 72, height: 72, borderRadius: 18, overflow: "hidden", margin: "0 auto 14px", boxShadow: "0 8px 28px rgba(0,0,0,.45)", background: "#fff", display: "grid", placeItems: "center" }}>
            <img src="logo.png" alt="Logo" style={{ width: "100%", height: "100%", objectFit: "contain", padding: "6px" }} />
          </div>
          <h1 style={{ color: "#fff", fontSize: 21, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-.01em" }}>Salas de Reunião</h1>
          <p style={{ color: "#7f978f", fontSize: 12.5, margin: 0 }}>{CFG.COMPANY_NAME}</p>
        </div>

        <div style={{ background: "#fff", borderRadius: 20, padding: "26px 26px 22px", boxShadow: "0 28px 56px rgba(0,0,0,.35)" }}>
          <div style={{ display: "flex", background: "var(--line-soft)", borderRadius: "999px", padding: 3, marginBottom: 20 }}>
            <button style={{ ...toggleSt, ...(mode === "login" ? toggleAct : {}) }} onClick={() => { setMode("login"); setMsg(null); }}>Entrar</button>
            <button style={{ ...toggleSt, ...(mode === "register" ? toggleAct : {}) }} onClick={() => { setMode("register"); setMsg(null); }}>Criar conta</button>
          </div>

          <form onSubmit={submit}>
            {mode === "register" && (
              <Field label="Nome completo">
                <input className="input" placeholder="Seu nome" value={name} onChange={(e) => setName(e.target.value)} autoFocus />
              </Field>
            )}
            <Field label="E-mail corporativo">
              <input className="input" type="email" placeholder="voce@empresa.com.br" value={email} onChange={(e) => setEmail(e.target.value)} autoFocus={mode === "login"} />
            </Field>
            <Field label="Senha">
              <input className="input" type="password" placeholder="••••••••" value={password} onChange={(e) => setPassword(e.target.value)} />
            </Field>

            {msg && <div className={`alert ${msg.type}`} style={{ marginBottom: 14 }}><Icon name={msg.type === "ok" ? "ok" : "warn"} size={17} />{msg.text}</div>}

            <button className="btn primary block" style={{ marginTop: 6 }} disabled={busy}>
              {busy ? "Aguarde..." : mode === "login" ? "Entrar" : "Criar conta"}
            </button>
          </form>

          <p style={{ marginTop: 16, fontSize: 11, color: "var(--muted)", textAlign: "center", lineHeight: 1.7 }}>
            Acesso restrito a colaboradores da empresa
          </p>
        </div>
      </div>
    </div>
  );
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ floors, rooms, bookings, go }) {
  const now = useNow();
  const todayISO = localDateStr(now.toISOString());
  const todays = bookings.filter((b) => b.status !== "cancelled" && localDateStr(b.start_time) === todayISO && new Date(b.end_time) > now);
  const busyRoomIds = new Set(
    bookings.filter((b) => b.status !== "cancelled" && new Date(b.start_time) <= now && new Date(b.end_time) > now).map((b) => b.room_id)
  );
  const freeNow = rooms.filter((r) => r.active && !busyRoomIds.has(r.id)).length;
  const roomName  = (id) => rooms.find((r) => r.id === id)?.name  || "—";
  const floorName = (id) => { const r = rooms.find((r) => r.id === id); return floors.find((f) => f.id === r?.floor_id)?.name || ""; };
  const roomColor = (id) => rooms.find((r) => r.id === id)?.color || "var(--brand)";

  /* reuniões futuras + hoje, ordenadas por data/hora, agrupadas por data */
  const upcoming = bookings
    .filter((b) => b.status !== "cancelled" && new Date(b.end_time) > now)
    .sort((a, b) => new Date(a.start_time) - new Date(b.start_time));

  const byDate = {};
  upcoming.forEach((b) => {
    const d = localDateStr(b.start_time);
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(b);
  });
  const dateKeys = Object.keys(byDate).sort();

  const dayLabel = (iso) => {
    if (iso === todayISO) return "Hoje";
    const tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);
    if (iso === localDateStr(tomorrow.toISOString())) return "Amanhã";
    return new Date(iso + "T12:00:00").toLocaleDateString("pt-BR", { weekday: "long", day: "numeric", month: "long" });
  };

  return (
    <div>
      <div className="page-head">
        <h2>Visão geral</h2>
        <p style={{ textTransform: "capitalize" }}>{fmtDateLong(now)}</p>
      </div>

      <div className="grid cols-4 mb">
        <Stat ico="building" color="blue"   num={floors.length}                      lbl="Andares" />
        <Stat ico="layout"   color="violet" num={rooms.filter((r) => r.active).length} lbl="Salas ativas" />
        <Stat ico="ok"       color="green"  num={freeNow}                            lbl="Livres agora" />
        <Stat ico="calendar" color="amber"  num={todays.length}                      lbl="Reuniões hoje" />
      </div>

      <div className="grid cols-2">
        {/* ── Reuniões agendadas agrupadas por data ── */}
        <div className="card">
          <div className="card-head">
            <h3>Reuniões agendadas</h3>
            <button className="btn ghost sm right" onClick={() => go("schedule")}><Icon name="plus" size={14} /> Agendar</button>
          </div>
          <div className="card-body" style={{ paddingTop: 4, maxHeight: 480, overflowY: "auto" }}>
            {dateKeys.length === 0 ? (
              <div className="empty"><div className="big"><Icon name="calendar" size={22} /></div>Nenhuma reunião agendada.</div>
            ) : (
              dateKeys.map((dk) => (
                <div key={dk} style={{ marginBottom: 18 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em",
                    color: dk === todayISO ? "var(--brand)" : "var(--muted)", marginBottom: 8, paddingBottom: 6,
                    borderBottom: "1px solid var(--line)" }}>
                    {dayLabel(dk)}
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {byDate[dk].map((b) => (
                      <div key={b.id} style={{ display: "flex", alignItems: "center", gap: 10,
                        padding: "10px 12px", borderRadius: 10, background: "var(--surface-2)",
                        border: "1px solid var(--line)" }}>
                        <span style={{ width: 8, height: 8, borderRadius: "50%", background: roomColor(b.room_id), flexShrink: 0 }} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: 13.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{b.title}</div>
                          <div className="muted" style={{ fontSize: 12 }}>{roomName(b.room_id)}{floorName(b.room_id) ? ` · ${floorName(b.room_id)}` : ""}</div>
                        </div>
                        <div style={{ textAlign: "right", flexShrink: 0 }}>
                          <div style={{ fontWeight: 600, fontSize: 12.5 }}>{fmtTime(b.start_time)}–{fmtTime(b.end_time)}</div>
                          <div className="muted" style={{ fontSize: 11.5 }}>{b.organizer_name}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>

        <div className="card">
          <div className="card-head"><h3>Status das salas agora</h3></div>
          <div className="card-body" style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {rooms.filter((r) => r.active).map((r) => {
              const busy = busyRoomIds.has(r.id);
              const fl = floors.find((f) => f.id === r.floor_id);
              return (
                <div key={r.id} className="flex between" style={{ borderBottom: "1px solid var(--line)", paddingBottom: 10 }}>
                  <div className="flex gap">
                    <span style={{ width: 10, height: 10, borderRadius: 999, background: r.color }} />
                    <div><b>{r.name}</b><div className="muted" style={{ fontSize: 12 }}>{fl?.name} · {r.capacity} lugares</div></div>
                  </div>
                  <RoomBadge busy={busy} />
                </div>
              );
            })}
            <button className="btn primary mt block" onClick={() => go("schedule")}><Icon name="plus" size={16} /> Novo agendamento</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== DISPONIBILIDADE ======================== */
function Availability({ floors, rooms, bookings, onBook }) {
  const [floorId, setFloorId] = useState("all");
  const [date, setDate] = useState(todayStr());
  const [detailTarget, setDetailTarget] = useState(null);
  const now = useNow();

  const dayBookings = useCallback((roomId) =>
    bookings.filter((b) => b.status !== "cancelled" && b.room_id === roomId && localDateStr(b.start_time) === date)
      .sort((a, b) => new Date(a.start_time) - new Date(b.start_time)), [bookings, date]);

  const visible = rooms.filter((r) => r.active && (floorId === "all" || r.floor_id === floorId));

  return (
    <div>
      <div className="page-head"><h2>Disponibilidade das salas</h2><p>Veja o que está livre e agende em poucos cliques.</p></div>

      <div className="card mb"><div className="card-body flex gap wrap">
        <Field label="Andar"><div className="chips">
          <button className={`chip ${floorId === "all" ? "active" : ""}`} onClick={() => setFloorId("all")}>Todos</button>
          {floors.filter((f) => rooms.some((r) => r.floor_id === f.id && r.active)).map((f) => (
            <button key={f.id} className={`chip ${floorId === f.id ? "active" : ""}`} onClick={() => setFloorId(f.id)}>{f.name}</button>
          ))}
        </div></Field>
        <div style={{ marginLeft: "auto" }}>
          <Field label="Data"><input className="input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        </div>
      </div></div>

      {visible.length === 0 ? (
        <div className="card"><div className="empty"><div className="big"><Icon name="layout" size={22} /></div>Nenhuma sala neste filtro.</div></div>
      ) : (
        <div className="grid cols-3">
          {visible.map((r) => {
            const evs = dayBookings(r.id);
            const fl = floors.find((f) => f.id === r.floor_id);
            const busyNow = date === todayStr() && evs.some((b) => new Date(b.start_time) <= now && new Date(b.end_time) > now);
            return (
              <div key={r.id} className="room">
                <span className="bar" style={{ background: r.color }} />
                <div className="top">
                  <div><h4>{r.name}</h4><div className="muted" style={{ fontSize: 12 }}>{fl?.name}</div></div>
                  {date === todayStr() && <RoomBadge busy={busyNow} />}
                </div>
                <div className="meta"><span><Icon name="users" size={15} />{r.capacity} lugares</span>{r.location && <span><Icon name="pin" size={15} />{r.location}</span>}</div>
                {r.equipment?.length > 0 && <div className="equip">{r.equipment.map((e) => <span key={e} className="tag">{e}</span>)}</div>}
                <div style={{ borderTop: "1px solid var(--line)", paddingTop: 10 }}>
                  <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Reservas em {fmtDate(date)}:</div>
                  {evs.length === 0 ? (
                    <div style={{ fontSize: 13, color: "var(--ok)", fontWeight: 500 }}>Livre o dia todo</div>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                      {evs.map((b) => (
                        <div key={b.id} style={{ fontSize: 12.5, cursor: "pointer", borderRadius: 6, padding: "3px 6px", margin: "0 -6px" }}
                          className="flex between"
                          onClick={() => setDetailTarget({ b, room: r, floor: fl })}
                          title="Clique para ver detalhes">
                          <span className="t-time" style={{ color: "var(--ink-2)" }}>{fmtTime(b.start_time)}–{fmtTime(b.end_time)}</span>
                          <span className="muted">{b.title}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <button className="btn primary block" onClick={() => onBook(r, date)}>Agendar nesta sala</button>
              </div>
            );
          })}
        </div>
      )}

      {detailTarget && (
        <BookingDetailModal
          booking={detailTarget.b}
          room={detailTarget.room}
          floor={detailTarget.floor}
          onClose={() => setDetailTarget(null)}
          onCancel={() => setDetailTarget(null)}
        />
      )}
    </div>
  );
}

/* ============================== AGENDAR (wizard 2 etapas) ============= */
function StepIndicator({ step, onBack }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 0, marginBottom: 24 }}>
      {[{ n: 1, label: "Detalhes da reunião" }, { n: 2, label: "Convite por e-mail" }].map((s, i) => (
        <React.Fragment key={s.n}>
          {i > 0 && <div style={{ flex: 1, height: 2, background: step >= s.n ? "var(--brand)" : "var(--line)", margin: "0 8px", transition: "background .3s" }} />}
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{
              width: 30, height: 30, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
              fontWeight: 800, fontSize: 13, transition: "all .3s",
              background: step >= s.n ? "var(--brand)" : "var(--line)",
              color: step >= s.n ? "#fff" : "var(--muted)",
            }}>{step > s.n ? <Icon name="ok" size={14} /> : s.n}</div>
            <span style={{ fontSize: 12.5, fontWeight: step === s.n ? 700 : 500, color: step === s.n ? "var(--ink)" : "var(--muted)", whiteSpace: "nowrap" }}>{s.label}</span>
          </div>
        </React.Fragment>
      ))}
    </div>
  );
}

/* Deriva nome legível a partir do e-mail corporativo.
   joao.silva@empresa.com.br → "Joao Silva" */
function nameFromEmail(email) {
  if (!email) return "";
  const local = email.split("@")[0] || "";
  return local
    .split(/[._-]/)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

function ScheduleForm({ floors, rooms, bookings, preset, onSaved, user, onOpenProfile }) {
  const userEmail  = user?.email || "";
  const userName   = nameFromEmail(userEmail);
  const firstFloor = floors[0]?.id || "";
  const [step, setStep] = useState(1);

  /* ── Etapa 1: dados da reunião ── */
  const [floorId, setFloorId] = useState(preset?.room ? preset.room.floor_id : firstFloor);
  const roomsOfFloor = rooms.filter((r) => r.active && r.floor_id === floorId);
  const [roomId, setRoomId]   = useState(preset?.room?.id || roomsOfFloor[0]?.id || "");
  const [date, setDate]       = useState(preset?.date || todayStr());
  const [start, setStart]     = useState("09:00");
  const [end, setEnd]         = useState("10:00");
  const [form, setForm]       = useState({ title: "", organizer_name: userName, organizer_email: userEmail, department: "", attendees: 1, notes: "", phone: "" });
  const [participants, setParticipants] = useState([{ name: "", email: "", phone: "" }]);
  const [phoneError, setPhoneError] = useState(false);

  /* ── Etapa 2: e-mail de convite ── */
  const [emailTo, setEmailTo]         = useState([]);
  const [emailCc, setEmailCc]         = useState([]);
  const [emailSubject, setEmailSubject] = useState("");
  const [emailMessage, setEmailMessage] = useState("");
  const [profile, setProfile]         = useState(getProfile);
  const hasSig = !!(profile && (
    profile.sig_html?.trim() ||
    profile.sig_image?.trim() ||
    profile.name?.trim() ||
    profile.department?.trim()
  ));
  const [sigWarningDismissed, setSigWarningDismissed] = useState(false);

  const [busy, setBusy]       = useState(false);
  const [msg, setMsg]         = useState(null);
  const [lastBooking, setLastBooking] = useState(null);

  const addParticipant = () => setParticipants((p) => [...p, { name: "", email: "", phone: "" }]);
  const removeParticipant = (i) => setParticipants((p) => p.filter((_, idx) => idx !== i));
  const setParticipant = (i, field, val) => setParticipants((p) => p.map((x, idx) => idx === i ? { ...x, [field]: val } : x));

  useEffect(() => {
    const list = rooms.filter((r) => r.active && r.floor_id === floorId);
    setRoomId((prev) => (list.find((r) => r.id === prev) ? prev : (list[0]?.id || "")));
  }, [floorId, rooms]);

  const set = (k) => (e) => { setForm({ ...form, [k]: e.target.value }); if (k === "phone") setPhoneError(false); };

  const now2 = new Date();
  const sameDay = bookings
    .filter((b) =>
      b.room_id === roomId &&
      localDateStr(b.start_time) === date &&
      b.status !== "cancelled" &&
      new Date(b.end_time) > now2          /* ignora reuniões já encerradas */
    )
    .sort((a, b) => new Date(a.start_time) - new Date(b.start_time));

  /* valida etapa 1 */
  const validate1 = () => {
    if (!roomId)                   return setMsg({ type: "err", text: "Selecione uma sala." })                   || false;
    if (!form.title.trim())        return setMsg({ type: "err", text: "Informe o assunto da reunião." })        || false;
    if (!form.organizer_name.trim()) return setMsg({ type: "err", text: "Informe o nome do responsável." })   || false;
    if (!form.phone.trim())        { setPhoneError(true); return setMsg({ type: "err", text: "Preencha o telefone para continuar." }) || false; }
    if (end <= start)              return setMsg({ type: "err", text: "O horário final deve ser maior que o inicial." }) || false;
    if (parseInt(start) < OPEN_HOUR || parseInt(end) > CLOSE_HOUR)
      return setMsg({ type: "err", text: `Horário permitido: ${OPEN_HOUR}h às ${CLOSE_HOUR}h.` }) || false;
    const room = rooms.find((r) => r.id === roomId);
    if (room && Number(form.attendees) > room.capacity)
      return setMsg({ type: "err", text: `Esta sala comporta no máximo ${room.capacity} pessoas.` }) || false;
    return true;
  };

  /* avança para etapa 2 e pré-preenche o e-mail */
  const goStep2 = () => {
    setMsg(null);
    if (!validate1()) return;
    const room  = rooms.find((r) => r.id === roomId);
    const floor = floors.find((f) => f.id === floorId);
    const to = [];
    participants.forEach((p) => { if (p.email.trim()) to.push(p.email.trim()); });
    setEmailTo(to);
    setEmailSubject(form.title.trim() || "Reunião de trabalho");
    const d = new Date(`${date}T${start}`);
    const dateStr = d.toLocaleDateString("pt-BR", { weekday: "long", day: "numeric", month: "long" });
    setEmailMessage(`Boa tarde,\n\nSegue o convite para a reunião "${form.title.trim()}".\n\nData: ${dateStr}\nHorário: ${start} – ${end}\nLocal: ${room?.name || ""}${floor ? ` · ${floor.name}` : ""}\n\nContamos com sua presença!`);
    setProfile(getProfile());
    setSigWarningDismissed(false);
    setStep(2);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  /* confirma agendamento (chamado da etapa 2) */
  const confirm = async (skipEmail = false) => {
    setBusy(true);
    setMsg(null);
    try {
      const startISO = toISO(date, start);
      const endISO   = toISO(date, end);
      const room  = rooms.find((r) => r.id === roomId);
      const floor = floors.find((f) => f.id === floorId);

      await DB.createBooking({
        room_id: roomId, title: form.title.trim(),
        organizer_name: form.organizer_name.trim(),
        organizer_email: form.organizer_email.trim() || null,
        phone: form.phone.trim() || null,
        department: form.department.trim() || null,
        attendees: Number(form.attendees) || 1,
        start_time: startISO, end_time: endISO,
        notes: form.notes.trim() || null,
        participants: participants.filter((p) => p.email.trim() || p.phone.trim()),
      });

      setLastBooking({
        title: form.title.trim(), start: startISO, end: endISO,
        location: `${room?.name || ""}${floor ? ` · ${floor.name}` : ""}`,
        organizer: form.organizer_email.trim(),
        description: `Responsável: ${form.organizer_name.trim()}${form.phone.trim() ? ` | Tel: ${form.phone.trim()}` : ""}${form.notes.trim() ? `\n${form.notes.trim()}` : ""}`,
      });

      /* e-mail desabilitado temporariamente — aguardando verificação de domínio no Resend */
      const emailOk = true;
      const emailErr = "";

      /* reset */
      setForm({ title: "", organizer_name: userName, organizer_email: userEmail, department: "", attendees: 1, notes: "", phone: "" });
      setParticipants([{ name: "", email: "", phone: "" }]);
      /* salva destinatários no histórico para sugestões futuras */
      const sentEmails = [...emailTo, ...emailCc].filter(Boolean);
      if (sentEmails.length > 0) saveEmailHistory(sentEmails);

      setEmailTo([]); setEmailCc([]); setEmailSubject(""); setEmailMessage("");
      setStep(1);

      if (skipEmail) {
        setMsg({ type: "ok", text: "Reunião agendada com sucesso!" });
      } else if (emailOk) {
        const total = emailTo.length + emailCc.length;
        setMsg({ type: "ok", text: `Reunião agendada! Convite enviado para ${total > 0 ? `${total} destinatário${total > 1 ? "s" : ""}` : "os participantes"}.` });
      } else {
        setMsg({ type: "warn", text: `Reunião agendada, mas o e-mail não foi enviado: ${emailErr}` });
      }
      onSaved && onSaved();
    } catch (err) {
      setMsg({ type: "err", text: err.message || "Erro ao agendar." });
    } finally {
      setBusy(false);
    }
  };

  /* ─────────── RENDER ─────────── */
  return (
    <div>
      <div className="page-head"><h2>Agendar reunião</h2><p>Reserve uma sala em 2 etapas rápidas.</p></div>

      <StepIndicator step={step} />

      {/* ═══════════════ ETAPA 1 ═══════════════ */}
      {step === 1 && (
      <div className="grid schedule-grid">
        <div className="card">
          <div className="card-head">
            <h3>Salas em uso · {fmtDate(date)}</h3>
          </div>
          <div className="card-body" style={{ maxHeight: 520, overflowY: "auto", paddingTop: 4 }}>
            {(() => {
              const nowTs = new Date();
              const isToday = date === todayStr();
              const activeRooms = rooms.filter((r) => r.active);
              /* bookings do dia para TODAS as salas (exceto cancelados e já encerrados hoje) */
              const dayAll = bookings.filter((b) =>
                b.status !== "cancelled" &&
                localDateStr(b.start_time) === date &&
                (!isToday || new Date(b.end_time) > nowTs)
              );
              /* salas que têm pelo menos um booking no dia */
              const busyRoomIds = new Set(dayAll.map((b) => b.room_id));
              const busyNowIds  = new Set(
                bookings.filter((b) => b.status !== "cancelled" && new Date(b.start_time) <= nowTs && new Date(b.end_time) > nowTs).map((b) => b.room_id)
              );

              if (busyRoomIds.size === 0) return (
                <div className="alert ok" style={{ marginBottom: 0 }}>
                  <Icon name="ok" size={17} /> Todas as salas estão livres neste dia.
                </div>
              );

              return activeRooms
                .filter((r) => busyRoomIds.has(r.id))
                .map((r) => {
                  const fl = floors.find((f) => f.id === r.floor_id);
                  const rBookings = dayAll.filter((b) => b.room_id === r.id)
                    .sort((a, b) => new Date(a.start_time) - new Date(b.start_time));
                  const busyNow = isToday && busyNowIds.has(r.id);
                  return (
                    <div key={r.id} style={{ marginBottom: 14 }}>
                      {/* cabeçalho da sala */}
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6,
                        paddingBottom: 6, borderBottom: "1px solid var(--line)" }}>
                        <span style={{ width: 10, height: 10, borderRadius: "50%", background: r.color || "var(--brand)", flexShrink: 0 }} />
                        <span style={{ fontWeight: 700, fontSize: 13 }}>{r.name}</span>
                        <span className="muted" style={{ fontSize: 11.5 }}>{fl?.name} · {r.capacity} lug.</span>
                        {busyNow && (
                          <span style={{ marginLeft: "auto", fontSize: 11, fontWeight: 700, color: "#dc2626",
                            background: "#fee2e2", borderRadius: 99, padding: "2px 8px" }}>● Em uso agora</span>
                        )}
                      </div>
                      {/* bookings da sala */}
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {rBookings.map((b) => {
                          const inProgress = isToday && new Date(b.start_time) <= nowTs && new Date(b.end_time) > nowTs;
                          return (
                            <div key={b.id} style={{ display: "flex", alignItems: "flex-start", gap: 10,
                              background: inProgress ? "#fff5f5" : "var(--surface-2)",
                              border: `1px solid ${inProgress ? "#fca5a5" : "var(--line)"}`,
                              borderRadius: 9, padding: "9px 12px" }}>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 700, fontSize: 13, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                  {b.title}
                                </div>
                                <div className="muted" style={{ fontSize: 11.5, marginTop: 2 }}>
                                  {b.organizer_name}{b.organizer_email ? ` · ${b.organizer_email}` : ""}
                                </div>
                                {b.phone && <div className="muted" style={{ fontSize: 11.5 }}>📞 {b.phone}</div>}
                              </div>
                              <div style={{ textAlign: "right", flexShrink: 0 }}>
                                <div style={{ fontWeight: 700, fontSize: 12.5, color: inProgress ? "#dc2626" : "var(--ink)" }}>
                                  {fmtTime(b.start_time)}–{fmtTime(b.end_time)}
                                </div>
                                {b.attendees > 1 && <div className="muted" style={{ fontSize: 11 }}>{b.attendees} pessoas</div>}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                });
            })()}
            <div className="alert info mt" style={{ fontSize: 12 }}>
              <Icon name="info" size={15} /> Conflitos de horário na mesma sala são bloqueados automaticamente.
            </div>
          </div>
        </div>

        <form className="card" onSubmit={(e) => e.preventDefault()}>
          <div className="card-head"><h3>Detalhes da reserva</h3></div>
          <div className="card-body">

            <div className="row">
              <Field label="Andar"><select className="select" value={floorId} onChange={(e) => setFloorId(e.target.value)}>
                {floors.filter((f) => rooms.some((r) => r.floor_id === f.id && r.active)).map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
              </select></Field>
              <Field label="Sala"><select className="select" value={roomId} onChange={(e) => setRoomId(e.target.value)}>
                {roomsOfFloor.length === 0 && <option value="">— sem salas —</option>}
                {roomsOfFloor.map((r) => <option key={r.id} value={r.id}>{r.name} ({r.capacity} lug.)</option>)}
              </select></Field>
            </div>

            <div className="row">
              <Field label="Data"><input className="input" type="date" value={date} min={todayStr()} onChange={(e) => setDate(e.target.value)} /></Field>
              <div className="row" style={{ gridTemplateColumns: "1fr 1fr" }}>
                <Field label="Início"><input className="input" type="time" value={start} step="900" onChange={(e) => setStart(e.target.value)} /></Field>
                <Field label="Término"><input className="input" type="time" value={end} step="900" onChange={(e) => setEnd(e.target.value)} /></Field>
              </div>
            </div>

            <Field label="Assunto da reunião"><input className="input" placeholder="Ex.: Alinhamento de projeto" value={form.title} onChange={set("title")} /></Field>

            <div className="row">
              <Field label="Responsável"><input className="input" placeholder="Seu nome" value={form.organizer_name} onChange={set("organizer_name")} /></Field>
              <Field label="E-mail corporativo"><input className="input" type="email" value={form.organizer_email} readOnly style={{ background: "var(--surface-2)", cursor: "default", color: "var(--ink-2)" }} /></Field>
            </div>

            <Field label="Telefone para contato *">
              <input
                className="input"
                type="tel"
                placeholder="(00) 00000-0000"
                value={form.phone}
                onChange={set("phone")}
                style={phoneError ? { borderColor: "var(--err)", background: "#fff5f5" } : {}}
              />
              {phoneError && <span style={{ fontSize: 12, color: "var(--err)", marginTop: 4, display: "block" }}>Preencha para continuar.</span>}
            </Field>

            <div className="row">
              <Field label="Departamento (opcional)"><input className="input" placeholder="Ex.: Comercial" value={form.department} onChange={set("department")} /></Field>
              <Field label="Nº de participantes"><input className="input" type="number" min="1" value={form.attendees} onChange={set("attendees")} /></Field>
            </div>

            <Field label="Observações (opcional)"><textarea className="textarea" placeholder="Detalhes, recursos extras..." value={form.notes} onChange={set("notes")} /></Field>

            {msg && <div className={`alert ${msg.type} mb`}><Icon name={msg.type === "ok" ? "ok" : "warn"} size={17} />{msg.text}</div>}

            <button type="button" className="btn primary block" style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              onClick={goStep2}>
              Próximo <Icon name="arrowRight" size={16} />
            </button>
          </div>
        </form>
      </div>
      )} {/* fecha step === 1 */}

      {/* ═══════════════ ETAPA 2 ═══════════════ */}
      {step === 2 && (
      <div className="grid compose-grid">

        {/* ── Coluna esquerda: e-mail ── */}
        <div>
          {/* Aviso de assinatura não configurada */}
          {!hasSig && !sigWarningDismissed && (
            <div style={{ background: "var(--warn-soft)", border: "1px solid #d4a814", borderRadius: 12, padding: "14px 16px", marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 13.5, color: "var(--warn)", marginBottom: 4 }}>
                <Icon name="warn" size={16} /> Assinatura de e-mail não configurada
              </div>
              <div style={{ fontSize: 12.5, color: "var(--ink-2)", marginBottom: 12 }}>
                Seus e-mails de convite serão enviados sem assinatura. Deseja configurar agora?
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button className="btn primary sm" onClick={() => {
                  onOpenProfile?.((savedProfile) => {
                    setProfile(savedProfile || getProfile());
                    setSigWarningDismissed(true);
                  });
                }}>
                  Configurar no perfil
                </button>
                <button className="btn ghost sm" onClick={() => setSigWarningDismissed(true)}>
                  Continuar sem assinatura
                </button>
              </div>
            </div>
          )}

          <div className="card" style={{ marginBottom: 0, borderBottomLeftRadius: 0, borderBottomRightRadius: 0, borderBottom: "none" }}>
            <div className="card-body" style={{ padding: "4px 0" }}>
              <ChipInput label="Para" chips={emailTo} onChange={setEmailTo} />
              <div style={{ height: 1, background: "var(--line)", margin: "0 16px" }} />
              <ChipInput label="Cc" chips={emailCc} onChange={setEmailCc} placeholder="Adicionar em cópia (opcional)" />
              <div style={{ height: 1, background: "var(--line)", margin: "0 16px" }} />
              <div className="ce-field">
                <div className="ce-label">Assunto</div>
                <input className="ce-subject" value={emailSubject} onChange={(e) => setEmailSubject(e.target.value)} />
              </div>
            </div>
          </div>

          <div className="card" style={{ borderTopLeftRadius: 0, borderTopRightRadius: 0 }}>
            <div className="card-body" style={{ padding: 0 }}>
              <textarea className="ce-body" value={emailMessage} onChange={(e) => setEmailMessage(e.target.value)} rows={8} />
              <div style={{ padding: "0 20px 20px" }}>
                {hasSig && <SignaturePreview sig={profile} />}
              </div>
            </div>
          </div>

          {msg && <div className={`alert ${msg.type}`} style={{ marginTop: 12 }}><Icon name={msg.type === "ok" ? "ok" : "warn"} size={17} />{msg.text}</div>}

          <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
            <button className="btn ghost" style={{ display: "flex", alignItems: "center", gap: 7 }}
              onClick={() => { setStep(1); setMsg(null); }}>
              <Icon name="arrowLeft" size={15} /> Voltar
            </button>
            <button className="btn primary" style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}
              disabled={busy} onClick={() => confirm(emailTo.length === 0)}>
              <Icon name={emailTo.length > 0 ? "send" : "ok"} size={15} />
              {busy ? "Aguarde…" : emailTo.length > 0 ? "Confirmar e enviar convite" : "Confirmar agendamento"}
            </button>
          </div>
          {emailTo.length > 0 && (
            <button className="btn ghost sm block" style={{ marginTop: 8, color: "var(--muted)" }}
              disabled={busy} onClick={() => confirm(true)}>
              Confirmar sem enviar e-mail
            </button>
          )}
        </div>

        {/* ── Coluna direita: prévia ── */}
        <div className="card">
          <div className="card-head"><h3>Prévia do convite</h3></div>
          <div className="card-body" style={{ background: "var(--bg)", borderRadius: "0 0 var(--radius) var(--radius)", padding: 16 }}>
            <div style={{ background: "#fff", borderRadius: 12, overflow: "hidden", border: "1px solid var(--line)", fontSize: 13 }}>
              <div style={{ background: "var(--brand-deep)", padding: "14px 18px", display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", background: "var(--accent)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 12, color: "#0d2420" }}>
                  {profile?.name ? profile.name[0].toUpperCase() : user?.email?.[0]?.toUpperCase() || "P"}
                </div>
                <div>
                  <div style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>{profile?.name || form.organizer_name || "Remetente"}</div>
                  <div style={{ color: "rgba(255,255,255,.55)", fontSize: 10.5 }}>{profile?.teams || form.organizer_email || ""}</div>
                </div>
              </div>
              <div style={{ padding: "10px 16px", borderBottom: "1px solid var(--line)", fontSize: 12 }}>
                {emailTo.length > 0  && <div style={{ marginBottom: 3 }}><b>Para:</b> {emailTo.join(", ")}</div>}
                {emailCc.length > 0  && <div style={{ marginBottom: 3 }}><b>Cc:</b> {emailCc.join(", ")}</div>}
                {emailSubject        && <div><b>Assunto:</b> {emailSubject}</div>}
              </div>
              <div style={{ padding: "16px 18px" }}>
                {emailMessage
                  ? <div style={{ color: "var(--ink)", lineHeight: 1.75, whiteSpace: "pre-wrap", fontSize: 13 }}>{emailMessage}</div>
                  : <div className="muted" style={{ fontStyle: "italic", fontSize: 12 }}>A mensagem aparecerá aqui…</div>}
                {hasSig && <SignaturePreview sig={profile} />}
              </div>
            </div>
          </div>
        </div>

      </div>
      )} {/* fecha step === 2 */}

      {/* calendário pós-agendamento */}
      {lastBooking && msg?.type === "ok" && (
        <div style={{ marginTop: 10, display: "flex", flexWrap: "wrap", gap: 8 }}>
          <div className="muted" style={{ width: "100%", fontSize: 12, marginBottom: 2 }}>Adicionar ao calendário:</div>
          <button className="btn ghost sm" style={{ display: "flex", alignItems: "center", gap: 6 }} onClick={() => {
            const iframe = document.createElement("iframe");
            iframe.style.display = "none";
            iframe.src = "ms-outlook://";
            document.body.appendChild(iframe);
            setTimeout(() => { document.body.removeChild(iframe); downloadICS(generateICS(lastBooking)); }, 1500);
          }}>
            <Icon name="outlook" size={14} /> Outlook
          </button>
          <a className="btn ghost sm" style={{ display: "flex", alignItems: "center", gap: 6, textDecoration: "none" }} href={outlookWebUrl(lastBooking)} target="_blank" rel="noopener">
            <Icon name="outlook" size={14} /> Outlook Web
          </a>
          <a className="btn ghost sm" style={{ display: "flex", alignItems: "center", gap: 6, textDecoration: "none" }} href={teamsCalendarUrl(lastBooking)} target="_blank" rel="noopener">
            <Icon name="teams" size={14} /> Microsoft Teams
          </a>
        </div>
      )}
    </div>
  );
}

/* ============================== MODAL DETALHE DA REUNIÃO ============== */
function CalendarButtons({ booking, room, floor }) {
  const calData = {
    title: booking.title,
    start: booking.start_time,
    end: booking.end_time,
    location: room ? `${room.name}${floor ? ` · ${floor.name}` : ""}` : "",
    organizer: booking.organizer_email || "",
    description: [
      booking.organizer_name ? `Responsável: ${booking.organizer_name}` : "",
      booking.phone ? `Tel: ${booking.phone}` : "",
      booking.notes || "",
    ].filter(Boolean).join("\n"),
  };

  const openOutlook = () => {
    const ics      = generateICS(calData);
    const filename = `${booking.title.replace(/\s+/g, "-")}.ics`;

    /* 1º tenta abrir o Outlook Clássico via protocolo ms-outlook: */
    let opened = false;
    const onFocus = () => { opened = true; };
    window.addEventListener("focus", onFocus, { once: true });

    const iframe = document.createElement("iframe");
    iframe.style.display = "none";
    iframe.src = "ms-outlook://";          /* protocolo do Outlook clássico */
    document.body.appendChild(iframe);

    /* Após 1.5 s, se a janela voltou ao foco (app não abriu), baixa o .ics
       que o Windows abre com o app de calendário padrão (Outlook novo / outro) */
    setTimeout(() => {
      document.body.removeChild(iframe);
      window.removeEventListener("focus", onFocus);
      if (!opened) downloadICS(ics, filename);   /* clássico abriu, não precisa */
      else         downloadICS(ics, filename);   /* fallback: .ics para Outlook novo */
    }, 1500);
  };

  const openTeams = () => {
    const webUrl = teamsCalendarUrl(calData);
    /* tenta abrir no app instalado via protocolo msteams:// */
    const desktopUrl = webUrl.replace("https://teams.microsoft.com", "msteams://teams.microsoft.com");
    const iframe = document.createElement("iframe");
    iframe.style.display = "none";
    iframe.src = desktopUrl;
    document.body.appendChild(iframe);
    /* fallback para web após 1.5s caso o app não abra */
    setTimeout(() => { document.body.removeChild(iframe); window.open(webUrl, "_blank"); }, 1500);
  };

  const btnStyle = {
    display: "flex", alignItems: "center", gap: 7,
    padding: "9px 13px", borderRadius: 8, border: "1.5px solid var(--line)",
    background: "var(--surface)", cursor: "pointer", fontSize: 12.5,
    fontWeight: 600, color: "var(--ink-2)", textDecoration: "none",
    transition: "border-color .15s, background .15s",
  };

  return (
    <div style={{ borderTop: "1px solid var(--line)", paddingTop: 14, marginTop: 4 }}>
      <div className="muted" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".05em", marginBottom: 10, textTransform: "uppercase" }}>
        Adicionar ao calendário
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        <button style={btnStyle} onClick={openOutlook}
          title="Abre o Outlook Clássico; se não instalado, baixa o .ics">
          <Icon name="outlook" size={15} /> Outlook
        </button>
        <a style={btnStyle} href={outlookWebUrl(calData)} target="_blank" rel="noopener"
          title="Abre o Outlook na web">
          <Icon name="mail" size={15} /> Outlook Web
        </a>
        <button style={btnStyle} onClick={openTeams}
          title="Abre o Microsoft Teams">
          <Icon name="teams" size={15} /> Teams
        </button>
      </div>
    </div>
  );
}

function BookingDetailModal({ booking, room, floor, onClose, onCancel }) {
  const now = new Date();
  const past = new Date(booking.end_time) < now;
  const cancelled = booking.status === "cancelled";
  const parts = Array.isArray(booking.participants) ? booking.participants : [];

  return (
    <Modal title="Detalhes da reunião" onClose={onClose}
      footer={
        <div style={{ display: "flex", justifyContent: "space-between", width: "100%" }}>
          <button className="btn ghost" onClick={onClose}>Fechar</button>
          {!past && !cancelled && (
            <button className="btn danger" onClick={() => { onClose(); onCancel(booking); }}>Cancelar reunião</button>
          )}
        </div>
      }>

      {/* Cabeçalho */}
      <div style={{ background: "var(--surface-2)", borderRadius: 10, padding: "14px 16px", marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 }}>
          <b style={{ fontSize: 15 }}>{booking.title}</b>
          {cancelled
            ? <span className="badge busy"><span className="bdot"></span>Cancelado</span>
            : past
              ? <span className="badge" style={{ background: "var(--line)", color: "var(--muted)" }}>Encerrado</span>
              : <span className="badge free"><span className="bdot"></span>Confirmado</span>}
        </div>
        <div className="muted" style={{ fontSize: 12.5, marginTop: 6, display: "flex", flexWrap: "wrap", gap: "4px 12px" }}>
          <span><Icon name="calendar" size={13} /> {fmtDate(booking.start_time)}</span>
          <span><Icon name="clock" size={13} /> {fmtTime(booking.start_time)}–{fmtTime(booking.end_time)}</span>
          {room && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><span style={{ width: 8, height: 8, borderRadius: 999, background: room.color }} />{room.name}{floor ? ` · ${floor.name}` : ""}</span>}
        </div>
      </div>

      {/* Responsável */}
      <div style={{ marginBottom: 14 }}>
        <div className="muted" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".05em", marginBottom: 6, textTransform: "uppercase" }}>Responsável</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          <span style={{ fontWeight: 600 }}>{booking.organizer_name}</span>
          {booking.organizer_email && <span className="muted" style={{ fontSize: 13 }}><Icon name="mail" size={13} /> {booking.organizer_email}</span>}
          {booking.phone && <span className="muted" style={{ fontSize: 13 }}><Icon name="phone" size={13} /> {booking.phone}</span>}
          {booking.department && <span className="muted" style={{ fontSize: 12 }}>{booking.department} · {booking.attendees} pessoas</span>}
        </div>
      </div>

      {/* Participantes */}
      {parts.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div className="muted" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".05em", marginBottom: 8, textTransform: "uppercase" }}>
            Participantes ({parts.length})
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {parts.map((p, i) => (
              <div key={i} style={{ background: "var(--surface-2)", borderRadius: 8, padding: "10px 12px" }}>
                {p.name && <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 4 }}>{p.name}</div>}
                <div style={{ display: "flex", flexWrap: "wrap", gap: "3px 14px" }}>
                  {p.email && <span className="muted" style={{ fontSize: 12.5 }}><Icon name="mail" size={12} /> {p.email}</span>}
                  {p.phone && <span className="muted" style={{ fontSize: 12.5 }}><Icon name="phone" size={12} /> {p.phone}</span>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Notas */}
      {booking.notes && (
        <div style={{ marginBottom: 14 }}>
          <div className="muted" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".05em", marginBottom: 6, textTransform: "uppercase" }}>Observações</div>
          <p style={{ fontSize: 13, margin: 0, color: "var(--ink-2)", lineHeight: 1.6 }}>{booking.notes}</p>
        </div>
      )}

      {/* Motivo do cancelamento */}
      {cancelled && booking.cancel_reason && (
        <div style={{ background: "#fff0f0", borderRadius: 8, padding: "10px 12px", marginBottom: 14 }}>
          <div className="muted" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".05em", marginBottom: 4, textTransform: "uppercase", color: "var(--err)" }}>Motivo do cancelamento</div>
          <p style={{ fontSize: 13, margin: 0, color: "var(--ink-2)" }}>{booking.cancel_reason}</p>
        </div>
      )}

      {/* Botões de calendário — sempre visível para reuniões não canceladas */}
      {!cancelled && <CalendarButtons booking={booking} room={room} floor={floor} />}
    </Modal>
  );
}

/* ============================== MODAL CANCELAMENTO ==================== */
function CancelModal({ booking, onClose, onConfirm }) {
  const [reason, setReason] = useState("");
  const [busy, setBusy] = useState(false);

  const handleConfirm = async () => {
    if (!reason.trim()) return;
    setBusy(true);
    try { await onConfirm(reason.trim()); }
    finally { setBusy(false); }
  };

  return (
    <Modal title="Cancelar reunião" onClose={onClose}
      footer={
        <>
          <button className="btn ghost" onClick={onClose}>Voltar</button>
          <button className="btn danger" disabled={busy || !reason.trim()} onClick={handleConfirm}>
            {busy ? "Cancelando…" : "Confirmar cancelamento"}
          </button>
        </>
      }>
      <div style={{ background: "var(--surface-2)", borderRadius: 10, padding: "12px 14px", marginBottom: 16 }}>
        <b style={{ fontSize: 14 }}>{booking.title}</b>
        <div className="muted" style={{ fontSize: 12.5, marginTop: 3 }}>
          {fmtDate(booking.start_time)} · {fmtTime(booking.start_time)}–{fmtTime(booking.end_time)}
        </div>
      </div>
      <Field label="Motivo do cancelamento">
        <textarea
          className="input"
          placeholder="Descreva o motivo do cancelamento…"
          rows={4}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          style={{ resize: "vertical", fontFamily: "inherit", fontSize: 13.5 }}
          autoFocus
        />
      </Field>
    </Modal>
  );
}

/* ============================== AGENDAMENTOS ========================== */
function BookingsList({ rooms, floors, bookings, onChanged }) {
  const [filter, setFilter] = useState("today");
  const [cancelTarget, setCancelTarget] = useState(null);
  const [detailTarget, setDetailTarget] = useState(null);
  const now = new Date();
  const roomOf = (id) => rooms.find((r) => r.id === id);
  const floorOf = (roomId) => { const r = roomOf(roomId); return floors.find((f) => f.id === r?.floor_id); };

  let rows = [...bookings];
  if (filter === "upcoming") rows = rows.filter((b) => new Date(b.end_time) >= now && b.status !== "cancelled");
  if (filter === "past") rows = rows.filter((b) => new Date(b.end_time) < now);
  if (filter === "today") rows = rows.filter((b) => localDateStr(b.start_time) === todayStr());
  /* Todos e Hoje: ativas/futuras primeiro (crescente), encerradas depois (decrescente).
     Passados: mais recente primeiro. Próximos: próximo primeiro. */
  if (filter === "past") {
    rows.sort((a, b) => new Date(b.start_time) - new Date(a.start_time));
  } else if (filter === "upcoming") {
    rows.sort((a, b) => new Date(a.start_time) - new Date(b.start_time));
  } else {
    rows.sort((a, b) => {
      const aEnded = new Date(a.end_time) < now;
      const bEnded = new Date(b.end_time) < now;
      if (aEnded !== bEnded) return aEnded ? 1 : -1; // encerradas vão para o fim
      return aEnded
        ? new Date(b.start_time) - new Date(a.start_time) // encerradas: mais recente primeiro
        : new Date(a.start_time) - new Date(b.start_time); // ativas/futuras: mais próxima primeiro
    });
  }

  const handleCancel = async (reason) => {
    await DB.cancelBooking(cancelTarget.id, reason);
    setCancelTarget(null);
    onChanged && onChanged();
  };

  return (
    <div>
      <div className="page-head"><h2>Agendamentos</h2><p>Consulte e gerencie as reservas de salas.</p></div>
      <div className="chips mb">
        {[["upcoming", "Próximos"], ["today", "Hoje"], ["past", "Passados"], ["all", "Todos"]].map(([k, l]) => (
          <button key={k} className={`chip ${filter === k ? "active" : ""}`} onClick={() => setFilter(k)}>{l}</button>
        ))}
      </div>
      <div className="card"><div className="card-body" style={{ padding: 0 }}>
        {rows.length === 0 ? (
          <div className="empty"><div className="big"><Icon name="inbox" size={22} /></div>Nenhum agendamento neste filtro.</div>
        ) : (
          <div className="table-wrap"><table className="table">
            <thead><tr><th>Reunião</th><th>Sala</th><th>Data</th><th>Horário</th><th>Responsável</th><th></th></tr></thead>
            <tbody>
              {rows.map((b) => {
                const past = new Date(b.end_time) < now;
                const cancelled = b.status === "cancelled";
                return (
                  <tr key={b.id} style={{ ...(past || cancelled ? { opacity: .6 } : {}), cursor: "pointer" }} onClick={() => setDetailTarget(b)}>
                    <td><b>{b.title}</b>{b.department && <div className="muted" style={{ fontSize: 12 }}>{b.department} · {b.attendees} pessoas</div>}</td>
                    <td><div className="flex gap"><span style={{ width: 8, height: 8, borderRadius: 999, background: roomOf(b.room_id)?.color }} />{roomOf(b.room_id)?.name || "—"}</div><div className="muted" style={{ fontSize: 12 }}>{floorOf(b.room_id)?.name}</div></td>
                    <td>{fmtDate(b.start_time)}</td>
                    <td><span className="t-time">{fmtTime(b.start_time)}–{fmtTime(b.end_time)}</span></td>
                    <td>{b.organizer_name}{b.organizer_email && <div className="muted" style={{ fontSize: 12 }}>{b.organizer_email}</div>}</td>
                    <td style={{ textAlign: "right" }}>
                      {cancelled ? (
                        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 3 }}>
                          <span className="badge busy"><span className="bdot"></span>Cancelado</span>
                          {b.cancel_reason && <span className="muted" style={{ fontSize: 11, maxWidth: 160, textAlign: "right" }}>{b.cancel_reason}</span>}
                        </div>
                      ) : !past && (
                        <button className="btn danger sm" onClick={(e) => { e.stopPropagation(); setCancelTarget(b); }}>Cancelar</button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        )}
      </div></div>

      {cancelTarget && (
        <CancelModal booking={cancelTarget} onClose={() => setCancelTarget(null)} onConfirm={handleCancel} />
      )}
      {detailTarget && (
        <BookingDetailModal
          booking={detailTarget}
          room={roomOf(detailTarget.room_id)}
          floor={floorOf(detailTarget.room_id)}
          onClose={() => setDetailTarget(null)}
          onCancel={(b) => setCancelTarget(b)}
        />
      )}
    </div>
  );
}

/* ============================== COMPARTILHAR ========================== */
/* ============================== MODAL PERFIL ========================== */
function ProfileModal({ onClose, user, onSaved }) {
  const [tab, setTab]         = useState("info");
  const [sigMode, setSigMode] = useState("builder");
  const [draft, setDraft]     = useState(() => getProfile() || {});
  const [uploading, setUploading] = useState(false);
  const [uploadErr, setUploadErr] = useState("");
  const fileRef = useRef(null);

  /* pré-preenche nome do usuário logado se vazio */
  useEffect(() => {
    if (user && !draft.name) {
      const name = user.user_metadata?.name || "";
      if (name) setDraft((d) => ({ ...d, name }));
    }
    if (draft.sig_mode) setSigMode(draft.sig_mode);
  }, []);

  const set = (k) => (e) => setDraft((d) => ({ ...d, [k]: e.target.value }));

  const handleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadErr("");
    setUploading(true);
    try {
      const url = await window.DB.uploadSignature(file, user?.id);
      setDraft((d) => ({ ...d, sig_image: url, sig_mode: "image" }));
      setSigMode("image");
    } catch (err) {
      setUploadErr(err.message || "Erro ao fazer upload.");
    } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const save = () => {
    const final = { ...draft, sig_mode: sigMode };
    saveProfile(final);
    onSaved?.(final);
    onClose();
  };

  const tabBtn = (id, label) => (
    <button className={`chip${tab === id ? " active" : ""}`} onClick={() => setTab(id)}>{label}</button>
  );

  const sigTabBtn = (id, label) => (
    <button
      className="btn ghost sm"
      style={{ fontWeight: sigMode === id ? 700 : 400, borderColor: sigMode === id ? "var(--brand)" : "var(--line)", color: sigMode === id ? "var(--brand)" : "var(--ink-2)" }}
      onClick={() => setSigMode(id)}>{label}
    </button>
  );

  const hasSig = (sigMode === "builder" && (draft.name || draft.department)) ||
                 (sigMode === "html"    && draft.sig_html?.trim()) ||
                 (sigMode === "image"   && draft.sig_image?.trim());

  return (
    <Modal title="Meu Perfil" onClose={onClose}
      footer={<><button className="btn ghost" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={save}>Salvar perfil</button></>}>

      {/* Tabs */}
      <div className="chips mb" style={{ marginBottom: 20 }}>
        {tabBtn("info", "Dados pessoais")}
        {tabBtn("signature", "Assinatura de e-mail")}
      </div>

      {/* ---- DADOS PESSOAIS ---- */}
      {tab === "info" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <Field label="Nome completo"><input className="input" value={draft.name || ""} onChange={set("name")} placeholder="Seu nome completo" /></Field>
          <Field label="Cargo / Departamento"><input className="input" value={draft.department || ""} onChange={set("department")} placeholder="Ex: Departamento" /></Field>
          <div className="row">
            <Field label="Telefone"><input className="input" value={draft.phone || ""} onChange={set("phone")} placeholder="(00) 9 0000-0000" /></Field>
            <Field label="E-mail Teams"><input className="input" value={draft.teams || ""} onChange={set("teams")} placeholder="nome@empresa.com.br" /></Field>
          </div>
          <Field label="Empresa"><input className="input" value={draft.company || ""} onChange={set("company")} placeholder="Nome da Empresa" /></Field>
          <Field label="Endereço (opcional)"><input className="input" value={draft.address || ""} onChange={set("address")} placeholder="Rua Exemplo, 100, 1° andar · Cidade/UF" /></Field>
        </div>
      )}

      {/* ---- ASSINATURA ---- */}
      {tab === "signature" && (
        <div>
          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
            {sigTabBtn("builder", "Construir")}
            {sigTabBtn("html",    "Colar HTML do Outlook")}
            {sigTabBtn("image",   "Imagem")}
          </div>

          {sigMode === "builder" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div className="alert info" style={{ fontSize: 12.5, marginBottom: 4 }}>
                <Icon name="warn" size={15} /> Os campos de Dados pessoais são usados automaticamente nesta assinatura.
              </div>
              <Field label="Nome (da aba Dados pessoais)"><input className="input" value={draft.name || ""} onChange={set("name")} placeholder="Seu nome" /></Field>
              <Field label="Cargo"><input className="input" value={draft.department || ""} onChange={set("department")} placeholder="Seu cargo" /></Field>
              <div className="row">
                <Field label="Telefone"><input className="input" value={draft.phone || ""} onChange={set("phone")} /></Field>
                <Field label="E-mail Teams"><input className="input" value={draft.teams || ""} onChange={set("teams")} /></Field>
              </div>
              <Field label="Empresa"><input className="input" value={draft.company || ""} onChange={set("company")} /></Field>
              <Field label="Endereço"><input className="input" value={draft.address || ""} onChange={set("address")} /></Field>
            </div>
          )}

          {sigMode === "html" && (
            <div>
              <div className="alert info" style={{ fontSize: 12.5, marginBottom: 10 }}>
                <Icon name="warn" size={15} /> No Outlook: <b>Arquivo → Opções → Email → Assinaturas</b> → selecione sua assinatura → copie o HTML e cole aqui.
              </div>
              <Field label="Cole o HTML da sua assinatura">
                <textarea className="input" rows={8} style={{ fontFamily: "var(--mono)", fontSize: 12, resize: "vertical" }}
                  placeholder="<table>...</table>"
                  value={draft.sig_html || ""}
                  onChange={set("sig_html")} />
              </Field>
              {draft.sig_html?.trim() && (
                <div style={{ marginTop: 12 }}>
                  <div className="muted" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 8 }}>Prévia</div>
                  <div style={{ border: "1px solid var(--line)", borderRadius: 8, padding: 14, background: "#fff", fontSize: 13 }}
                    dangerouslySetInnerHTML={{ __html: draft.sig_html }} />
                </div>
              )}
            </div>
          )}

          {sigMode === "image" && (
            <div>
              {/* Upload direto */}
              <input ref={fileRef} type="file" accept="image/png,image/jpeg,image/gif,image/webp"
                style={{ display: "none" }} onChange={handleUpload} />

              <div style={{ border: "2px dashed var(--line)", borderRadius: 12, padding: "28px 20px",
                textAlign: "center", marginBottom: 14, background: "var(--surface-2)" }}>
                <div style={{ fontSize: 13, color: "var(--ink-2)", marginBottom: 12 }}>
                  {draft.sig_image
                    ? "Imagem carregada — clique para substituir"
                    : "Selecione a imagem da sua assinatura (PNG ou JPG)"}
                </div>
                <button className="btn primary sm" disabled={uploading}
                  onClick={() => fileRef.current?.click()}>
                  <Icon name={uploading ? "warn" : "download"} size={15} />
                  {uploading ? "Enviando…" : draft.sig_image ? "Trocar imagem" : "Selecionar imagem"}
                </button>
                {uploadErr && (
                  <div style={{ marginTop: 10, fontSize: 12, color: "var(--err)" }}>{uploadErr}</div>
                )}
              </div>

              {draft.sig_image?.trim() && (
                <div>
                  <div className="muted" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 8 }}>Prévia</div>
                  <div style={{ border: "1px solid var(--line)", borderRadius: 8, padding: 14, background: "#fff" }}>
                    <img src={draft.sig_image} alt="Assinatura" style={{ maxWidth: "100%" }} />
                  </div>
                  <button className="btn ghost sm" style={{ marginTop: 8, color: "var(--err)" }}
                    onClick={() => setDraft((d) => ({ ...d, sig_image: "" }))}>
                    Remover imagem
                  </button>
                </div>
              )}
            </div>
          )}

          {/* Preview assinatura construída */}
          {sigMode === "builder" && hasSig && (
            <div style={{ marginTop: 16 }}>
              <div className="muted" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 8 }}>Prévia</div>
              <SignaturePreview sig={draft} />
            </div>
          )}
        </div>
      )}
    </Modal>
  );
}

/* ============================== CHIP INPUT ============================ */
function ChipInput({ label, chips, onChange, placeholder = "nome@empresa.com" }) {
  const [val, setVal]           = useState("");
  const [showSug, setShowSug]   = useState(false);
  const [activeSug, setActiveSug] = useState(-1);
  const inputRef  = useRef(null);
  const history   = useMemo(() => getEmailHistory(), []);

  const suggestions = useMemo(() => {
    if (!val.trim()) return [];
    const q = val.toLowerCase();
    return history
      .filter((e) => !chips.includes(e) && e.toLowerCase().includes(q))
      .slice(0, 6);
  }, [val, chips, history]);

  const tryAdd = (raw) => {
    const emails = raw.split(/[,;\s]+/).map((e) => e.trim()).filter((e) => e.includes("@"));
    if (!emails.length) return;
    const next = [...chips];
    emails.forEach((e) => { if (!next.includes(e)) next.push(e); });
    onChange(next);
    setVal("");
    setShowSug(false);
    setActiveSug(-1);
  };

  const pickSug = (email) => {
    const next = [...chips];
    if (!next.includes(email)) next.push(email);
    onChange(next);
    setVal("");
    setShowSug(false);
    setActiveSug(-1);
    inputRef.current?.focus();
  };

  const onKeyDown = (e) => {
    if (showSug && suggestions.length > 0) {
      if (e.key === "ArrowDown") { e.preventDefault(); setActiveSug((i) => Math.min(i + 1, suggestions.length - 1)); return; }
      if (e.key === "ArrowUp")   { e.preventDefault(); setActiveSug((i) => Math.max(i - 1, -1)); return; }
      if (e.key === "Enter" && activeSug >= 0) { e.preventDefault(); pickSug(suggestions[activeSug]); return; }
      if (e.key === "Escape") { setShowSug(false); setActiveSug(-1); return; }
    }
    if (["Enter", ",", ";", "Tab"].includes(e.key)) { e.preventDefault(); tryAdd(val); }
    if (e.key === "Backspace" && !val && chips.length) onChange(chips.slice(0, -1));
  };

  const onPaste = (e) => { e.preventDefault(); tryAdd(e.clipboardData.getData("text")); };

  return (
    <div className="ce-field" style={{ position: "relative" }}>
      <div className="ce-label">{label}</div>
      <div className="ce-chip-box" onClick={() => inputRef.current?.focus()}>
        {chips.map((c, i) => (
          <span key={i} className="ce-chip">
            {c}
            <button className="ce-chip-x" onClick={(e) => { e.stopPropagation(); onChange(chips.filter((_, idx) => idx !== i)); }}>×</button>
          </span>
        ))}
        <input
          ref={inputRef}
          className="ce-chip-input"
          value={val}
          onChange={(e) => { setVal(e.target.value); setShowSug(true); setActiveSug(-1); }}
          onKeyDown={onKeyDown}
          onBlur={() => setTimeout(() => setShowSug(false), 150)}
          onFocus={() => val && setShowSug(true)}
          onPaste={onPaste}
          placeholder={chips.length === 0 ? placeholder : ""}
        />
      </div>

      {/* Dropdown de sugestões */}
      {showSug && suggestions.length > 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 99,
          background: "#fff", border: "1px solid var(--line)", borderRadius: 10,
          boxShadow: "0 6px 20px rgba(0,0,0,.1)", overflow: "hidden", marginTop: 4 }}>
          {suggestions.map((s, i) => (
            <div key={s}
              onMouseDown={(e) => { e.preventDefault(); pickSug(s); }}
              style={{ padding: "9px 14px", fontSize: 13, cursor: "pointer",
                background: i === activeSug ? "var(--brand-soft, #e8f5ee)" : "#fff",
                color: "var(--ink)", borderBottom: i < suggestions.length - 1 ? "1px solid var(--line)" : "none" }}>
              {s}
            </div>
          ))}
        </div>
      )}

      <div className="ce-hint">Pressione Enter, Tab ou vírgula para adicionar · Cole vários e-mails de uma vez</div>
    </div>
  );
}

/* ============================== SIGNATURE BUILDER ==================== */
const SIG_KEY = "pluma_email_sig";

function SignatureBuilder({ sig, onChange }) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(sig || {});

  const save = () => {
    localStorage.setItem(SIG_KEY, JSON.stringify(draft));
    onChange(draft);
    setOpen(false);
  };

  const set = (k) => (e) => setDraft((d) => ({ ...d, [k]: e.target.value }));

  return (
    <>
      <button className="btn ghost sm" style={{ display: "flex", alignItems: "center", gap: 6 }}
        onClick={() => { setDraft(sig || {}); setOpen(true); }}>
        <Icon name="settings" size={13} /> Configurar assinatura
      </button>

      {open && (
        <Modal title="Minha assinatura" onClose={() => setOpen(false)}
          footer={
            <>
              <button className="btn ghost" onClick={() => setOpen(false)}>Cancelar</button>
              <button className="btn primary" onClick={save}>Salvar assinatura</button>
            </>
          }>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <Field label="Nome completo"><input className="input" value={draft.name || ""} onChange={set("name")} placeholder="João Silva" /></Field>
            <Field label="Cargo / Departamento"><input className="input" value={draft.department || ""} onChange={set("department")} placeholder="Departamento" /></Field>
            <div className="row">
              <Field label="Telefone"><input className="input" value={draft.phone || ""} onChange={set("phone")} placeholder="(00) 9 0000-0000" /></Field>
              <Field label="E-mail Teams"><input className="input" value={draft.teams || ""} onChange={set("teams")} placeholder="joao.silva@empresa.com.br" /></Field>
            </div>
            <Field label="Empresa"><input className="input" value={draft.company || ""} onChange={set("company")} placeholder="Nome da Empresa" /></Field>
            <Field label="Endereço (opcional)"><input className="input" value={draft.address || ""} onChange={set("address")} placeholder="Rua Minas Gerais, 1932, Edifício Unique, 14° andar · Cascavel/PR" /></Field>
          </div>

          {/* preview da assinatura */}
          {(draft.name || draft.department) && (
            <div style={{ marginTop: 18 }}>
              <div className="muted" style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 8 }}>Prévia</div>
              <SignaturePreview sig={draft} />
            </div>
          )}
        </Modal>
      )}
    </>
  );
}

function SignaturePreview({ sig }) {
  if (!sig) return null;

  /* HTML importado */
  if (sig.sig_mode === "html" && sig.sig_html?.trim()) {
    return (
      <div style={{ borderTop: "2px solid var(--line)", paddingTop: 16, marginTop: 4 }}
        dangerouslySetInnerHTML={{ __html: sig.sig_html }} />
    );
  }

  /* Imagem */
  if (sig.sig_mode === "image" && sig.sig_image?.trim()) {
    return (
      <div style={{ borderTop: "2px solid var(--line)", paddingTop: 16, marginTop: 4 }}>
        <img src={sig.sig_image} alt="Assinatura" style={{ maxWidth: "100%" }} />
      </div>
    );
  }

  /* Construída — builder padrão */
  if (!sig.name && !sig.department) return null;
  return (
    <div style={{ borderTop: "2px solid var(--line)", paddingTop: 16, marginTop: 4 }}>
      <div style={{ display: "flex", gap: 16, alignItems: "center" }}>
        <div style={{ background: "var(--brand-deep)", borderRadius: 10, padding: "12px 16px", flex: "0 0 auto", minWidth: 180 }}>
          {sig.name && <div style={{ color: "var(--accent)", fontWeight: 800, fontSize: 12.5, letterSpacing: ".04em", marginBottom: 2 }}>{sig.name.toUpperCase()}</div>}
          {sig.department && <div style={{ color: "#a8c5bb", fontSize: 11, marginBottom: 8 }}>{sig.department}</div>}
          {sig.phone && <div style={{ color: "rgba(255,255,255,.7)", fontSize: 11 }}>📞 {sig.phone}</div>}
          {sig.teams && <div style={{ color: "rgba(255,255,255,.7)", fontSize: 11 }}>Teams: {sig.teams}</div>}
        </div>
        <div style={{ borderLeft: "3px solid var(--accent)", paddingLeft: 16 }}>
          <div style={{ color: "var(--ink-2)", fontSize: 10, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase" }}>GRUPO</div>
          <div style={{ color: "var(--ink)", fontSize: 22, fontWeight: 900, letterSpacing: ".03em", lineHeight: 1.1 }}>PLUMA</div>
          {sig.company && <div style={{ color: "var(--muted)", fontSize: 10.5, marginTop: 3 }}>{sig.company}</div>}
        </div>
      </div>
      {sig.address && <div style={{ fontSize: 10.5, color: "var(--muted)", marginTop: 10 }}>{sig.address}</div>}
    </div>
  );
}

/* ============================== COMPOR E-MAIL ========================= */
function ComposeEmail({ user }) {
  const [to, setTo]           = useState([]);
  const [cc, setCc]           = useState([]);
  const [subject, setSubject] = useState("");
  const [message, setMessage] = useState("");
  const [busy, setBusy]       = useState(false);
  const [msg, setMsg]         = useState(null);
  const [sig, setSig]         = useState(() => {
    try { return JSON.parse(localStorage.getItem(SIG_KEY) || "null"); } catch { return null; }
  });

  /* pré-preenche nome da assinatura com o nome do usuário logado */
  useEffect(() => {
    if (user && !sig?.name) {
      const name = user.user_metadata?.name || "";
      if (name) {
        const initial = { name };
        localStorage.setItem(SIG_KEY, JSON.stringify(initial));
        setSig(initial);
      }
    }
  }, [user]);

  const handleSigChange = (s) => setSig(s);

  const send = async () => {
    setMsg(null);
    if (!to.length) return setMsg({ type: "err", text: "Adicione ao menos um destinatário em 'Para'." });
    if (!subject.trim()) return setMsg({ type: "err", text: "Preencha o assunto." });
    if (!message.trim()) return setMsg({ type: "err", text: "Escreva a mensagem." });

    setBusy(true);
    try {
      const res = await fetch("/api/send-email", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ to, cc, subject: subject.trim(), message: message.trim(), signature: sig, senderEmail: user?.email, senderName: sig?.name }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erro ao enviar.");
      setMsg({ type: "ok", text: `E-mail enviado para ${to.length + cc.length} destinatário(s).` });
      setTo([]); setCc([]); setSubject(""); setMessage("");
    } catch (e) {
      setMsg({ type: "err", text: e.message });
    } finally {
      setBusy(false);
    }
  };

  return (
    <div>
      <div className="page-head">
        <h2>Compor E-mail</h2>
        <p>Envie e-mails com a identidade visual da empresa.</p>
      </div>

      <div className="grid compose-grid">

        {/* ---- Formulário ---- */}
        <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
          <div className="card" style={{ marginBottom: 0, borderBottomLeftRadius: 0, borderBottomRightRadius: 0, borderBottom: "none" }}>
            <div className="card-body" style={{ padding: "4px 0", display: "flex", flexDirection: "column" }}>
              <ChipInput label="Para" chips={to} onChange={setTo} />
              <div style={{ height: 1, background: "var(--line)", margin: "0 16px" }} />
              <ChipInput label="Cc" chips={cc} onChange={setCc} placeholder="Adicionar em cópia (opcional)" />
              <div style={{ height: 1, background: "var(--line)", margin: "0 16px" }} />
              <div className="ce-field">
                <div className="ce-label">Assunto</div>
                <input className="ce-subject" value={subject}
                  onChange={(e) => setSubject(e.target.value)}
                  placeholder="Ex: Reunião com a Diretoria" />
              </div>
            </div>
          </div>

          <div className="card" style={{ borderTopLeftRadius: 0, borderTopRightRadius: 0 }}>
            <div className="card-body" style={{ padding: 0 }}>
              <textarea
                className="ce-body"
                value={message}
                onChange={(e) => setMessage(e.target.value)}
                placeholder="Boa tarde,&#10;&#10;Segue agendamento da reunião para discutir os temas…"
                rows={8}
              />

              {/* Assinatura */}
              <div style={{ padding: "0 20px 20px" }}>
                <SignaturePreview sig={sig} />
              </div>
            </div>
          </div>

          {msg && (
            <div className={`alert ${msg.type}`} style={{ marginTop: 12 }}>
              <Icon name={msg.type === "ok" ? "ok" : "warn"} size={17} />{msg.text}
            </div>
          )}

          <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
            <button className="btn primary" style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              disabled={busy} onClick={send}>
              <Icon name="send" size={16} />
              {busy ? "Enviando…" : "Enviar e-mail"}
            </button>
            <SignatureBuilder sig={sig} onChange={handleSigChange} />
          </div>
        </div>

        {/* ---- Prévia ---- */}
        <div className="card">
          <div className="card-head"><h3>Prévia do e-mail</h3></div>
          <div className="card-body" style={{ background: "var(--bg)", borderRadius: "0 0 var(--radius) var(--radius)", padding: 16 }}>
            <div style={{ background: "#fff", borderRadius: 12, overflow: "hidden", border: "1px solid var(--line)", fontSize: 13 }}>
              {/* Header simulado */}
              <div style={{ background: "var(--brand-deep)", padding: "14px 18px", display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 28, height: 28, borderRadius: 50, background: "var(--accent)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 12, color: "#0d2420" }}>
                  {sig?.name ? sig.name[0].toUpperCase() : "P"}
                </div>
                <div>
                  <div style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>{sig?.name || "Remetente"}</div>
                  <div style={{ color: "rgba(255,255,255,.55)", fontSize: 10.5 }}>{sig?.teams || user?.email || ""}</div>
                </div>
              </div>

              {/* Campos Para/Assunto */}
              <div style={{ padding: "10px 16px", borderBottom: "1px solid var(--line)", fontSize: 12 }}>
                {to.length > 0 && <div style={{ marginBottom: 4 }}><b>Para:</b> {to.join(", ")}</div>}
                {cc.length > 0 && <div style={{ marginBottom: 4 }}><b>Cc:</b> {cc.join(", ")}</div>}
                {subject && <div><b>Assunto:</b> {subject}</div>}
              </div>

              {/* Corpo */}
              <div style={{ padding: "16px 18px" }}>
                {message ? (
                  <div style={{ color: "var(--ink)", lineHeight: 1.75, whiteSpace: "pre-wrap", fontSize: 13 }}>{message}</div>
                ) : (
                  <div className="muted" style={{ fontStyle: "italic", fontSize: 12 }}>A mensagem aparecerá aqui…</div>
                )}
                <SignaturePreview sig={sig} />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== SHARE VIEW ============================ */
function ShareView() {
  const APP_URL = "https://salas-de-reuniao.vercel.app";
  const [copied, setCopied] = useState(false);

  const copy = () => {
    navigator.clipboard.writeText(APP_URL).then(() => { setCopied(true); setTimeout(() => setCopied(false), 2500); });
  };

  const waText  = encodeURIComponent(`Acesse o sistema de salas de reunião:\n${APP_URL}`);
  const mailSub = encodeURIComponent("Gestão de Salas de Reunião");
  const mailBody = encodeURIComponent(`Olá,\n\nSegue o link para o sistema de agendamento de salas de reunião:\n\n${APP_URL}\n\nAté mais!`);
  const teamsMsg = encodeURIComponent(`Acesse o sistema de salas de reunião: ${APP_URL}`);
  const qrUrl   = `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(APP_URL)}&size=220x220&margin=12&color=1f574b&bgcolor=ffffff`;

  const shareBtn = (href, label, icon, bg) => (
    <a href={href} target="_blank" rel="noopener noreferrer"
      style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, background: bg, color: "#fff",
        borderRadius: 10, padding: "12px 18px", fontWeight: 600, fontSize: 14, textDecoration: "none", cursor: "pointer" }}>
      <Icon name={icon} size={18} />{label}
    </a>
  );

  return (
    <div>
      <div className="page-head"><h2>Compartilhar</h2><p>Envie o link do sistema para outros colaboradores.</p></div>

      <div className="grid cols-2">
        <div className="card">
          <div className="card-head"><h3>QR Code</h3></div>
          <div className="card-body" style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
            <img src={qrUrl} alt="QR Code" width={220} height={220}
              style={{ borderRadius: 12, border: "1px solid var(--line)", display: "block" }} />
            <p className="muted" style={{ fontSize: 12.5, textAlign: "center", margin: 0, lineHeight: 1.6 }}>
              Aponte a câmera do celular para abrir o sistema direto no browser.
            </p>
            <a href={qrUrl} download="qrcode-salas.png" className="btn ghost sm" style={{ width: "100%", textAlign: "center" }}>
              Baixar QR Code
            </a>
          </div>
        </div>

        <div className="card">
          <div className="card-head"><h3>Enviar link</h3></div>
          <div className="card-body" style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div>
              <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Link do sistema</div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <code style={{ flex: 1, background: "var(--surface-2)", padding: "9px 12px", borderRadius: 8,
                  fontSize: 12, wordBreak: "break-all", border: "1px solid var(--line)" }}>{APP_URL}</code>
                <button className="btn ghost sm" onClick={copy} style={{ flexShrink: 0, minWidth: 80 }}>
                  {copied ? <><Icon name="ok" size={14} /> Copiado</> : "Copiar"}
                </button>
              </div>
            </div>

            <div style={{ borderTop: "1px solid var(--line)", paddingTop: 14, display: "flex", flexDirection: "column", gap: 10 }}>
              <div className="muted" style={{ fontSize: 12, marginBottom: 2 }}>Compartilhar via</div>
              {shareBtn(`https://wa.me/?text=${waText}`, "WhatsApp", "whatsapp", "#25d366")}
              <button
                onClick={() => {
                  const iframe = document.createElement("iframe");
                  iframe.style.display = "none";
                  iframe.src = "ms-outlook://";
                  document.body.appendChild(iframe);
                  setTimeout(() => {
                    document.body.removeChild(iframe);
                    window.open(`mailto:?subject=${mailSub}&body=${mailBody}`, "_blank");
                  }, 1500);
                }}
                style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, background: "#0078d4", color: "#fff",
                  borderRadius: 10, padding: "12px 18px", fontWeight: 600, fontSize: 14, textDecoration: "none", cursor: "pointer", border: "none", width: "100%" }}>
                <Icon name="outlook" size={18} /> Outlook / E-mail
              </button>
              {shareBtn(`https://teams.microsoft.com/l/chat/0/0?message=${teamsMsg}`, "Microsoft Teams", "teams", "#6264a7")}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== ADMIN ================================= */
function RoomNameCell({ room, onSave }) {
  const [editing, setEditing] = useState(false);
  const [name, setName] = useState(room.name);
  const inputRef = useRef(null);

  const start = () => { setName(room.name); setEditing(true); };
  const cancel = () => setEditing(false);
  const save = async () => {
    const trimmed = name.trim();
    if (!trimmed || trimmed === room.name) { cancel(); return; }
    await onSave(trimmed);
    setEditing(false);
  };
  const onKey = (e) => { if (e.key === "Enter") save(); if (e.key === "Escape") cancel(); };

  useEffect(() => { if (editing) inputRef.current?.focus(); }, [editing]);

  if (editing) {
    return (
      <div className="flex gap" style={{ alignItems: "center" }}>
        <span style={{ width: 10, height: 10, borderRadius: 999, background: room.color, flexShrink: 0 }} />
        <input
          ref={inputRef}
          className="input"
          style={{ padding: "4px 8px", fontSize: 13.5, height: 32, width: 160 }}
          value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={onKey}
          onBlur={save}
        />
      </div>
    );
  }

  return (
    <div>
      <div className="flex gap" style={{ alignItems: "center", cursor: "text" }} onClick={start} title="Clique para renomear">
        <span style={{ width: 10, height: 10, borderRadius: 999, background: room.color, flexShrink: 0 }} />
        <b>{room.name}</b>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="var(--muted)" strokeWidth="2" style={{ marginLeft: 2, opacity: .6 }}>
          <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
        </svg>
      </div>
      {room.location && <div className="muted" style={{ fontSize: 12, marginTop: 3 }}>{room.location}</div>}
    </div>
  );
}

function Admin({ floors, rooms, bookings, onChanged }) {
  const [floorModal, setFloorModal] = useState(null);
  const [roomModal, setRoomModal] = useState(null);

  const roomsCount = (fid) => rooms.filter((r) => r.floor_id === fid).length;
  const futureBookings = (roomIds) => bookings.filter((b) => roomIds.includes(b.room_id) && new Date(b.end_time) > new Date()).length;
  const delFloor = async (f) => {
    const rids = rooms.filter((r) => r.floor_id === f.id).map((r) => r.id);
    const n = futureBookings(rids);
    const msg = n > 0 ? `Excluir o andar "${f.name}"? ${n} reserva(s) futura(s) serão perdidas.` : `Excluir o andar "${f.name}" e todas as suas salas?`;
    if (confirm(msg)) { await DB.deleteFloor(f.id); onChanged(); }
  };
  const delRoom = async (r) => {
    const n = futureBookings([r.id]);
    const msg = n > 0 ? `Excluir a sala "${r.name}"? ${n} reserva(s) futura(s) serão perdidas.` : `Excluir a sala "${r.name}"?`;
    if (confirm(msg)) { await DB.deleteRoom(r.id); onChanged(); }
  };
  const renameRoom = async (r, newName) => {
    await DB.updateRoom(r.id, { name: newName });
    onChanged();
  };

  return (
    <div>
      <div className="page-head"><h2>Administração</h2><p>Cadastre andares e salas de reunião.</p></div>

      <div className="card mb">
        <div className="card-head"><h3>Andares</h3><button className="btn ghost sm right" onClick={() => setFloorModal({})}><Icon name="plus" size={15} /> Novo andar</button></div>
        <div className="card-body" style={{ padding: 0 }}>
          <table className="table admin-table">
            <thead><tr><th>Andar</th><th>Nº</th><th>Descrição</th><th>Salas</th><th></th></tr></thead>
            <tbody>
              {floors.map((f) => (
                <tr key={f.id}>
                  <td><b>{f.name}</b></td><td>{f.number ?? "—"}</td><td className="muted">{f.description || "—"}</td><td>{roomsCount(f.id)} sala(s)</td>
                  <td style={{ textAlign: "right" }}>
                    <button className="btn ghost sm" onClick={() => setFloorModal(f)}>Editar</button>{" "}
                    <button className="btn danger sm" onClick={() => delFloor(f)}>Excluir</button>
                  </td>
                </tr>
              ))}
              {floors.length === 0 && <tr><td colSpan="5"><div className="empty">Nenhum andar cadastrado.</div></td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card">
        <div className="card-head"><h3>Salas de reunião</h3>
          <button className="btn ghost sm right" disabled={floors.length === 0} onClick={() => setRoomModal({})}><Icon name="plus" size={15} /> Nova sala</button>
        </div>
        <div className="card-body" style={{ padding: 0 }}>
          <table className="table admin-table">
            <thead><tr><th>Sala</th><th>Andar</th><th>Capacidade</th><th>Equipamentos</th><th>Status</th><th></th></tr></thead>
            <tbody>
              {rooms.map((r) => (
                <tr key={r.id}>
                  <td><RoomNameCell room={r} onSave={(newName) => renameRoom(r, newName)} /></td>
                  <td>{floors.find((f) => f.id === r.floor_id)?.name || "—"}</td>
                  <td>{r.capacity} lug.</td>
                  <td>{(r.equipment || []).map((e) => <span key={e} className="tag" style={{ marginRight: 4 }}>{e}</span>) || "—"}</td>
                  <td><span className={`badge ${r.active ? "free" : "busy"}`}><span className="bdot"></span>{r.active ? "Ativa" : "Inativa"}</span></td>
                  <td style={{ textAlign: "right" }}>
                    <button className="btn ghost sm" onClick={() => setRoomModal(r)}>Editar</button>{" "}
                    <button className="btn danger sm" onClick={() => delRoom(r)}>Excluir</button>
                  </td>
                </tr>
              ))}
              {rooms.length === 0 && <tr><td colSpan="6"><div className="empty">Nenhuma sala cadastrada.</div></td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {floorModal && <FloorModal floor={floorModal} rooms={rooms} onClose={() => setFloorModal(null)} onSaved={() => { setFloorModal(null); onChanged(); }} />}
      {roomModal && <RoomModal room={roomModal} floors={floors} onClose={() => setRoomModal(null)} onSaved={() => { setRoomModal(null); onChanged(); }} />}
    </div>
  );
}

function FloorModal({ floor, rooms, onClose, onSaved }) {
  const editing = !!floor.id;
  const floorRooms = (rooms || []).filter((r) => r.floor_id === floor.id);

  const [f, setF] = useState({ name: floor.name || "", number: floor.number ?? "", description: floor.description || "" });
  // Nomes das salas existentes (só em modo edição)
  const [roomEdits, setRoomEdits] = useState(() =>
    Object.fromEntries(floorRooms.map((r) => [r.id, r.name]))
  );
  // Novas salas a criar
  const [roomCount, setRoomCount] = useState(0);
  const [roomNames, setRoomNames] = useState([]);
  const [busy, setBusy] = useState(false);

  const handleCountChange = (val) => {
    const n = Math.min(20, Math.max(0, Number(val) || 0));
    setRoomCount(n);
    setRoomNames((prev) => Array.from({ length: n }, (_, i) => prev[i] ?? ""));
  };

  const save = async () => {
    if (!f.name.trim()) return;
    setBusy(true);
    const payload = { name: f.name.trim(), number: f.number === "" ? null : Number(f.number), description: f.description.trim() || null };
    try {
      let floorId;
      if (editing) {
        await DB.updateFloor(floor.id, payload);
        floorId = floor.id;
        // Atualiza nomes das salas que mudaram
        await Promise.all(
          floorRooms
            .filter((r) => roomEdits[r.id]?.trim() && roomEdits[r.id].trim() !== r.name)
            .map((r) => DB.updateRoom(r.id, { name: roomEdits[r.id].trim() }))
        );
      } else {
        const created = await DB.createFloor(payload);
        floorId = created.id;
      }
      // Cria novas salas
      const toCreate = roomNames.filter((n) => n.trim());
      if (toCreate.length > 0) {
        await Promise.all(
          toCreate.map((name, i) =>
            DB.createRoom({ floor_id: floorId, name: name.trim(), capacity: 4, equipment: [], color: COLORS[i % COLORS.length], active: true, location: null })
          )
        );
      }
      onSaved();
    } finally { setBusy(false); }
  };

  return (
    <Modal title={editing ? "Editar andar" : "Novo andar"} onClose={onClose}
      footer={<><button className="btn ghost" onClick={onClose}>Cancelar</button><button className="btn primary" disabled={busy} onClick={save}>{busy ? "Salvando…" : "Salvar"}</button></>}>
      <Field label="Nome do andar"><input className="input" placeholder="Ex.: 3º Andar" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <Field label="Número / ordem" hint="Usado para ordenar a exibição."><input className="input" type="number" value={f.number} onChange={(e) => setF({ ...f, number: e.target.value })} /></Field>
      <Field label="Descrição (opcional)"><input className="input" value={f.description} onChange={(e) => setF({ ...f, description: e.target.value })} /></Field>

      {editing && floorRooms.length > 0 && (
        <div style={{ borderTop: "1px solid var(--line)", marginTop: 14, paddingTop: 14 }}>
          <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>Salas deste andar</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {floorRooms.map((r) => (
              <Field key={r.id} label={<span style={{ display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 8, height: 8, borderRadius: 999, background: r.color, display: "inline-block" }} />{r.name}</span>}>
                <input
                  className="input"
                  value={roomEdits[r.id] ?? r.name}
                  onChange={(e) => setRoomEdits((prev) => ({ ...prev, [r.id]: e.target.value }))}
                  placeholder="Nome da sala"
                />
              </Field>
            ))}
          </div>
        </div>
      )}

      <div style={{ borderTop: "1px solid var(--line)", marginTop: 14, paddingTop: 14 }}>
        <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>Adicionar novas salas</div>
        <Field label="Nº de salas a criar">
          <input className="input" type="number" min="0" max="20" value={roomCount} onChange={(e) => handleCountChange(e.target.value)} />
        </Field>
        {roomCount > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
            {roomNames.map((name, i) => (
              <Field key={i} label={`Nova sala ${i + 1}`}>
                <input className="input" placeholder="Ex.: Sala Atlântico" value={name}
                  onChange={(e) => setRoomNames((prev) => prev.map((v, idx) => idx === i ? e.target.value : v))} />
              </Field>
            ))}
          </div>
        )}
      </div>
    </Modal>
  );
}

const EQUIP_OPTS = ["TV", "Projetor", "Webcam", "Telefone", "Quadro", "Ar-condicionado"];
const COLORS = ["#1f574b", "#3c6e6a", "#2f7d5b", "#5a7d4e", "#b6862f", "#9a6a3a", "#274b6e", "#7a4a4a"];

function RoomModal({ room, floors, onClose, onSaved }) {
  const editing = !!room.id;
  const [r, setR] = useState({
    floor_id: room.floor_id || floors[0]?.id || "",
    name: room.name || "", capacity: room.capacity ?? 4, location: room.location || "",
    equipment: room.equipment || [], color: room.color || COLORS[0], active: room.active ?? true,
  });
  const [busy, setBusy] = useState(false);
  const toggleEq = (e) => setR({ ...r, equipment: r.equipment.includes(e) ? r.equipment.filter((x) => x !== e) : [...r.equipment, e] });
  const save = async () => {
    if (!r.name.trim() || !r.floor_id) return;
    setBusy(true);
    const payload = { floor_id: r.floor_id, name: r.name.trim(), capacity: Number(r.capacity) || 1, location: r.location.trim() || null, equipment: r.equipment, color: r.color, active: r.active };
    try {
      if (editing) await DB.updateRoom(room.id, payload); else await DB.createRoom(payload);
      onSaved();
    } finally { setBusy(false); }
  };
  return (
    <Modal title={editing ? "Editar sala" : "Nova sala"} onClose={onClose}
      footer={<><button className="btn ghost" onClick={onClose}>Cancelar</button><button className="btn primary" disabled={busy} onClick={save}>Salvar</button></>}>
      <div className="row">
        <Field label="Andar"><select className="select" value={r.floor_id} onChange={(e) => setR({ ...r, floor_id: e.target.value })}>
          {floors.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select></Field>
        <Field label="Capacidade"><input className="input" type="number" min="1" value={r.capacity} onChange={(e) => setR({ ...r, capacity: e.target.value })} /></Field>
      </div>
      <Field label="Nome da sala"><input className="input" placeholder="Ex.: Sala Atlântico" value={r.name} onChange={(e) => setR({ ...r, name: e.target.value })} /></Field>
      <Field label="Localização (opcional)"><input className="input" placeholder="Ex.: Ala leste" value={r.location} onChange={(e) => setR({ ...r, location: e.target.value })} /></Field>
      <Field label="Equipamentos"><div className="chips">
        {EQUIP_OPTS.map((e) => <button key={e} type="button" className={`chip ${r.equipment.includes(e) ? "active" : ""}`} onClick={() => toggleEq(e)}>{e}</button>)}
      </div></Field>
      <Field label="Cor de identificação"><div className="chips">
        {COLORS.map((c) => <button key={c} type="button" onClick={() => setR({ ...r, color: c })}
          style={{ width: 30, height: 30, borderRadius: 8, background: c, border: r.color === c ? "3px solid var(--ink)" : "2px solid #fff", boxShadow: "var(--shadow-sm)", cursor: "pointer" }} />)}
      </div></Field>
      <Field><label className="flex gap" style={{ cursor: "pointer" }}>
        <input type="checkbox" checked={r.active} onChange={(e) => setR({ ...r, active: e.target.checked })} /> Sala ativa (disponível para reservas)
      </label></Field>
    </Modal>
  );
}

/* ============================== APP SHELL ============================= */
const NAV = [
  { id: "dashboard",    label: "Início",          ic: "home" },
  { id: "availability", label: "Disponibilidade",  ic: "check" },
  { id: "schedule",     label: "Agendar",          ic: "calendarPlus" },
  { id: "bookings",     label: "Agendamentos",     ic: "list" },
  { id: "share",        label: "Compartilhar",     ic: "share" },
  { id: "admin",        label: "Administração",    ic: "settings" },
];

function App() {
  const [user, setUser] = useState(undefined); // undefined=carregando, null=não logado, obj=logado
  const [view, setView] = useState("dashboard");
  const [floors, setFloors] = useState([]);
  const [rooms, setRooms] = useState([]);
  const [bookings, setBookings] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState(null);
  const [preset, setPreset] = useState(null);
  const [profileOpen, setProfileOpen]       = useState(false);
  const [profileCallback, setProfileCallback] = useState(null);

  const openProfile = (onSaved) => {
    setProfileCallback(() => onSaved || null);
    setProfileOpen(true);
  };
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // Verifica sessão salva ao abrir
  useEffect(() => {
    let authEventFired = false;
    const unsub = DB.onAuthChange((session) => {
      authEventFired = true;
      if (!session) { setFloors([]); setRooms([]); setBookings([]); }
      setUser(session?.user ?? null);
    });
    // Fallback caso onAuthChange demore
    DB.getSession().then((session) => {
      if (!authEventFired) setUser(session?.user ?? null);
    });
    return unsub;
  }, []);

  const logout = async () => {
    await DB.signOut();
    setFloors([]); setRooms([]); setBookings([]);
    setUser(null);
  };

  const userName = user ? (user.user_metadata?.name || user.email?.split("@")[0] || "Usuário") : "";

  const load = useCallback(async () => {
    try {
      setErr(null);
      const [fl, rm, bk] = await Promise.all([DB.listFloors(), DB.listRooms(), DB.listBookings()]);
      setFloors(fl); setRooms(rm); setBookings(bk);
    } catch (e) {
      setErr(e.message || "Erro ao carregar dados.");
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { if (user) load(); }, [load, user]);

  // Recarrega dados a cada 2 minutos para sincronizar entre usuários
  useEffect(() => {
    if (!user) return;
    const id = setInterval(load, 2 * 60 * 1000);
    return () => clearInterval(id);
  }, [user, load]);

  // Recarrega quando o usuário volta à aba do browser
  useEffect(() => {
    const onVisible = () => { if (document.visibilityState === "visible" && user) load(); };
    document.addEventListener("visibilitychange", onVisible);
    return () => document.removeEventListener("visibilitychange", onVisible);
  }, [user, load]);

  const goBook = (room, date) => { setPreset({ room, date }); setView("schedule"); };
  const go = (v) => { setPreset(null); setView(v); setSidebarOpen(false); };

  const current = NAV.find((n) => n.id === view);

  if (user === undefined) return <div style={{ minHeight: "100vh", background: "var(--brand-deep)", display: "grid", placeItems: "center" }}><div className="spin" style={{ borderTopColor: "var(--accent)" }} /></div>;
  if (user === null) return <LoginScreen onLogin={setUser} />;

  return (
    <div className="app">
      <div className={`overlay-bg ${sidebarOpen ? "show" : ""}`} onClick={() => setSidebarOpen(false)} />
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div className="brand">
          <div className="logo"><img src="logo.png" alt="Logo" style={{width:"100%",height:"100%",objectFit:"contain",padding:"6px"}} /></div>
          <div><b>Salas de Reunião</b><span>{CFG.COMPANY_NAME}</span></div>
        </div>
        <nav className="nav">
          <div className="group-label">Menu</div>
          {NAV.map((n) => (
            <button key={n.id} className={view === n.id ? "active" : ""} onClick={() => go(n.id)}>
              <span className="ic"><Icon name={n.ic} size={19} /></span>{n.label}
            </button>
          ))}
        </nav>
        <div className="foot">Sistema de Gestão de Salas</div>
      </aside>

      <div className="main">
        <header className="topbar">
          <button className="menu-btn" onClick={() => setSidebarOpen(true)}><Icon name="menu" size={22} /></button>
          <div><h1>{current?.label}</h1><div className="sub">{CFG.COMPANY_NAME} · Salas de Reunião</div></div>
          <div className="spacer" />
          <div className="flex gap" style={{ alignItems: "center", gap: 10 }}>
            <button
              onClick={() => openProfile()}
              title="Meu perfil e assinatura"
              style={{ width: 34, height: 34, borderRadius: "50%", background: "var(--brand)", color: "#fff", border: "none", cursor: "pointer", fontWeight: 800, fontSize: 13.5, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              {userName?.[0]?.toUpperCase() || <Icon name="user" size={16} />}
            </button>
            <span style={{ fontSize: 12.5, color: "var(--ink-2)", fontWeight: 500, display: "inline-block", maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{userName}</span>
            <button className="btn ghost sm" onClick={logout} title="Sair"><Icon name="x" size={15} /> Sair</button>
          </div>
        </header>

        <div className="content">
          {loading ? <div className="spin" /> : err ? (
            <div className="alert err"><Icon name="warn" size={17} />{err}</div>
          ) : (
            <>
              {view === "dashboard" && <Dashboard floors={floors} rooms={rooms} bookings={bookings} go={go} />}
              {view === "availability" && <Availability floors={floors} rooms={rooms} bookings={bookings} onBook={goBook} />}
              {view === "schedule" && <ScheduleForm floors={floors} rooms={rooms} bookings={bookings} preset={preset} onSaved={load} user={user} onOpenProfile={openProfile} />}
              {view === "bookings" && <BookingsList floors={floors} rooms={rooms} bookings={bookings} onChanged={load} />}
              {view === "share"    && <ShareView />}
              {view === "admin"    && <Admin floors={floors} rooms={rooms} bookings={bookings} onChanged={load} />}
            </>
          )}
        </div>
      </div>

      {profileOpen && (
        <ProfileModal
          user={user}
          onClose={() => { setProfileOpen(false); setProfileCallback(null); }}
          onSaved={(p) => { profileCallback?.(p); setProfileOpen(false); setProfileCallback(null); }}
        />
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
