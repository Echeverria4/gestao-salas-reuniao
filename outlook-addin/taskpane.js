/* ==========================================================
   OUTLOOK ADD-IN — Reservar Sala de Reunião
   ========================================================== */

let supa = null;
let currentUser = null;
let outlookItem = null;

/* ---- promisifica os callbacks do Office.js ---- */
function getAsync(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(new Error(result.error.message));
      else resolve(result.value);
    });
  });
}

function setAsync(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(new Error(result.error.message));
      else resolve();
    });
  });
}

const pad = (n) => String(n).padStart(2, "0");
const fmtTime = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const fmtDate = (d) => d.toLocaleDateString("pt-BR", { weekday: "short", day: "numeric", month: "short" });
const setHTML = (html) => { document.getElementById("app").innerHTML = html; };

/* ================================================================
   PONTO DE ENTRADA
   ================================================================ */
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    setHTML('<div class="loading">Este add-in funciona apenas no Outlook.</div>');
    return;
  }

  outlookItem = Office.context.mailbox.item;

  const cfg = window.APP_CONFIG;
  if (!cfg || !cfg.SUPABASE_URL || cfg.SUPABASE_URL.includes("COLE_AQUI")) {
    setHTML('<div class="loading">Supabase não configurado.</div>');
    return;
  }

  supa = supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);

  try {
    const { data } = await supa.auth.getSession();
    if (data.session) {
      currentUser = data.session.user;
      await renderForm();
    } else {
      renderLogin();
    }
  } catch (e) {
    renderError("Erro ao conectar: " + e.message);
  }
});

/* ================================================================
   CABEÇALHO COMUM
   ================================================================ */
function headerHTML() {
  return `
    <div class="header">
      <div class="header-logo">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <rect x="3" y="4" width="18" height="18" rx="2.5"/>
          <path d="M16 2v4M8 2v4M3 10h18"/>
        </svg>
      </div>
      <div>
        <div class="header-title">Reservar Sala</div>
        <div class="header-sub">Sistema de Salas</div>
      </div>
    </div>`;
}

/* ================================================================
   LOGIN
   ================================================================ */
function renderLogin(errMsg) {
  setHTML(`
    ${headerHTML()}
    <div class="login-form">
      <h2>Entrar no sistema</h2>
      <p>Use o mesmo e-mail e senha do sistema de salas.</p>
      ${errMsg ? `<div class="alert alert-err">⚠️ ${errMsg}</div>` : ""}
      <div class="field">
        <label>E-mail corporativo</label>
        <input id="l-email" type="email" placeholder="seu@email.com" autocomplete="email">
      </div>
      <div class="field">
        <label>Senha</label>
        <input id="l-pass" type="password" placeholder="••••••••">
      </div>
      <button class="btn btn-primary" id="btn-login">Entrar</button>
    </div>`);

  document.getElementById("btn-login").addEventListener("click", doLogin);
  document.getElementById("l-pass").addEventListener("keydown", (e) => { if (e.key === "Enter") doLogin(); });
}

async function doLogin() {
  const email = document.getElementById("l-email").value.trim();
  const pass  = document.getElementById("l-pass").value;
  const btn   = document.getElementById("btn-login");

  if (!email || !pass) { renderLogin("Preencha e-mail e senha."); return; }
  btn.disabled = true;
  btn.textContent = "Entrando…";

  try {
    const { data, error } = await supa.auth.signInWithPassword({ email, password: pass });
    if (error) throw error;
    currentUser = data.user;
    await renderForm();
  } catch (e) {
    renderLogin(e.message || "E-mail ou senha incorretos.");
  }
}

/* ================================================================
   FORMULÁRIO DE RESERVA
   ================================================================ */
let state = {
  floors: [], rooms: [], busyRooms: new Set(),
  selectedRoomId: null,
  subject: "", start: null, end: null,
};

async function renderForm() {
  setHTML(`${headerHTML()}<div class="loading"><div class="spinner"></div>Carregando salas…</div>`);

  try {
    /* Lê os dados do Outlook em paralelo */
    const [subject, start, end] = await Promise.all([
      getAsync((cb) => outlookItem.subject.getAsync(cb)).catch(() => ""),
      getAsync((cb) => outlookItem.start.getAsync(cb)),
      getAsync((cb) => outlookItem.end.getAsync(cb)),
    ]);

    /* Carrega salas e verifica conflitos */
    const [floorsRes, roomsRes, bookingsRes] = await Promise.all([
      supa.from("floors").select("*").order("number", { ascending: true }),
      supa.from("rooms").select("*").eq("active", true).order("name", { ascending: true }),
      supa.from("bookings").select("room_id")
        .eq("status", "confirmed")
        .lt("start_time", end.toISOString())
        .gt("end_time", start.toISOString()),
    ]);

    if (floorsRes.error) throw floorsRes.error;
    if (roomsRes.error)  throw roomsRes.error;

    const busyRooms = new Set((bookingsRes.data || []).map((b) => b.room_id));

    state = {
      floors: floorsRes.data,
      rooms: roomsRes.data,
      busyRooms,
      selectedRoomId: null,
      subject, start, end,
    };

    drawForm();
  } catch (e) {
    renderError("Erro ao carregar: " + e.message);
  }
}

function drawForm() {
  const { floors, rooms, busyRooms, selectedRoomId, subject, start, end } = state;
  const floorOf = (fid) => floors.find((f) => f.id === fid);

  /* Agrupa salas por andar */
  const byFloor = {};
  rooms.forEach((r) => {
    if (!byFloor[r.floor_id]) byFloor[r.floor_id] = [];
    byFloor[r.floor_id].push(r);
  });

  const roomsHTML = Object.entries(byFloor).map(([fid, rs]) => {
    const floor = floorOf(fid);
    const items = rs.map((r) => {
      const busy = busyRooms.has(r.id);
      const sel  = r.id === selectedRoomId;
      return `
        <div class="room-item ${sel ? "selected" : ""} ${busy ? "busy" : ""}"
          data-id="${r.id}" role="button">
          <span class="room-dot" style="background:${r.color || "#888"}"></span>
          <span class="room-name">${r.name}</span>
          <span class="room-cap">${r.capacity} lug.</span>
          ${busy ? '<span class="room-busy-badge">Ocupada</span>' : ""}
        </div>`;
    }).join("");

    return `
      <div style="margin-bottom:10px">
        <div class="card-label" style="margin-bottom:6px">${floor ? floor.name : "Andar"}</div>
        <div class="room-list">${items}</div>
      </div>`;
  }).join("");

  const noRooms = rooms.length === 0
    ? '<div class="alert alert-info">Nenhuma sala cadastrada.</div>'
    : "";

  setHTML(`
    ${headerHTML()}

    <!-- Info da reunião do Outlook -->
    <div class="card">
      <div class="card-label">Reunião no Outlook</div>
      <div class="meeting-row">
        <span class="meeting-title">${subject || "(sem título)"}</span>
        <span class="meeting-time">${fmtTime(start)}–${fmtTime(end)}</span>
      </div>
      <div style="font-size:11.5px;color:var(--muted)">${fmtDate(start)}</div>
    </div>

    <!-- Seleção de sala -->
    <div class="card">
      <div class="card-label">Selecione uma sala disponível</div>
      ${noRooms}
      ${roomsHTML}
    </div>

    <div id="msg-area"></div>

    <button class="btn btn-primary" id="btn-book" disabled>Reservar sala</button>
    <button class="btn btn-ghost" id="btn-logout">Sair da conta</button>
  `);

  /* Eventos */
  document.querySelectorAll(".room-item:not(.busy)").forEach((el) => {
    el.addEventListener("click", () => {
      state.selectedRoomId = el.dataset.id;
      document.querySelectorAll(".room-item").forEach((x) => x.classList.remove("selected"));
      el.classList.add("selected");
      document.getElementById("btn-book").disabled = false;
    });
  });

  document.getElementById("btn-book").addEventListener("click", doBooking);
  document.getElementById("btn-logout").addEventListener("click", async () => {
    await supa.auth.signOut();
    currentUser = null;
    renderLogin();
  });
}

/* ================================================================
   SALVAR RESERVA
   ================================================================ */
async function doBooking() {
  const { selectedRoomId, rooms, floors, subject, start, end } = state;
  if (!selectedRoomId) return;

  const btn = document.getElementById("btn-book");
  btn.disabled = true;
  btn.textContent = "Reservando…";
  document.getElementById("msg-area").innerHTML = "";

  try {
    const room  = rooms.find((r) => r.id === selectedRoomId);
    const floor = floors.find((f) => f.id === room.floor_id);
    const userName = currentUser.user_metadata?.name || currentUser.email;

    /* Verifica conflito novamente (corrida de clique) */
    const { data: conflict } = await supa.from("bookings").select("id")
      .eq("room_id", selectedRoomId).eq("status", "confirmed")
      .lt("start_time", end.toISOString())
      .gt("end_time", start.toISOString());

    if (conflict && conflict.length > 0) {
      showMsgError("Esta sala ficou ocupada agora. Escolha outra.");
      btn.disabled = false;
      btn.textContent = "Reservar sala";
      return;
    }

    /* Cria o agendamento */
    const { error } = await supa.from("bookings").insert({
      room_id: selectedRoomId,
      title: subject || "Reunião",
      organizer_name: userName,
      organizer_email: currentUser.email,
      start_time: start.toISOString(),
      end_time: end.toISOString(),
      attendees: 1,
      status: "confirmed",
    });
    if (error) throw error;

    /* Atualiza o campo Local da reunião no Outlook */
    const locationStr = `${room.name}${floor ? ` · ${floor.name}` : ""}`;
    await setAsync((cb) => outlookItem.location.setAsync(locationStr, cb)).catch(() => {});

    renderSuccess(room, floor);
  } catch (e) {
    showMsgError(e.message || "Erro ao reservar.");
    btn.disabled = false;
    btn.textContent = "Reservar sala";
  }
}

/* ================================================================
   SUCESSO
   ================================================================ */
function renderSuccess(room, floor) {
  setHTML(`
    ${headerHTML()}
    <div style="text-align:center;padding:16px 0">
      <div class="success-icon">
        <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#1a6b3a" stroke-width="2.5">
          <path d="M20 6L9 17l-5-5"/>
        </svg>
      </div>
      <div style="font-weight:700;font-size:15px;margin-bottom:6px">Sala reservada!</div>
      <div style="display:flex;align-items:center;justify-content:center;gap:8px;margin-bottom:4px">
        <span style="width:10px;height:10px;border-radius:50%;background:${room.color || "#1a6b3a"};display:inline-block"></span>
        <span style="font-weight:600;font-size:13.5px">${room.name}</span>
      </div>
      ${floor ? `<div style="font-size:12px;color:var(--muted)">${floor.name}</div>` : ""}
      <div style="font-size:12px;color:var(--muted);margin-top:6px">
        ${fmtDate(state.start)} · ${fmtTime(state.start)}–${fmtTime(state.end)}
      </div>
    </div>

    <div class="alert alert-ok" style="margin-top:8px">
      O campo <strong>Local</strong> da reunião foi atualizado automaticamente no Outlook.
    </div>

    <button class="btn btn-ghost" onclick="renderForm()">Escolher outra sala</button>
  `);
}

/* ================================================================
   HELPERS
   ================================================================ */
function showMsgError(text) {
  const el = document.getElementById("msg-area");
  if (el) el.innerHTML = `<div class="alert alert-err" style="margin-bottom:10px">⚠️ ${text}</div>`;
}

function renderError(msg) {
  setHTML(`${headerHTML()}<div class="alert alert-err">⚠️ ${msg}</div>`);
}
