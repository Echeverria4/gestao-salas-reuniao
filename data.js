/* =========================================================================
   CAMADA DE DADOS
   -------------------------------------------------------------------------
   Expõe window.DB com métodos para floors / rooms / bookings.
   - Se o Supabase estiver configurado (config.js), usa o banco real.
   - Caso contrário, cai no MODO DEMO: dados em memória (não persistem).
   ========================================================================= */
(function () {
  const cfg = window.APP_CONFIG;
  const configured = window.IS_SUPABASE_CONFIGURED;

  let supa = null;
  if (configured && window.supabase) {
    supa = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);
  }

  // ---------------------------------------------------------------------
  // MODO DEMO — base de dados falsa em memória
  // ---------------------------------------------------------------------
  const uid = () => "demo-" + Math.abs(Date.now() ^ (performance.now() * 1000 | 0)).toString(36) + (demoSeq++);
  let demoSeq = 0;

  const demo = {
    floors: [
      { id: "f0", name: "Térreo", number: 0, description: "Recepção e salas de apoio" },
      { id: "f1", name: "1º Andar", number: 1, description: "Área comercial" },
      { id: "f3", name: "3º Andar", number: 3, description: "Diretoria e reuniões executivas" },
    ],
    rooms: [
      { id: "r1", floor_id: "f0", name: "Sala Recepção", capacity: 4, location: "Próx. à entrada", equipment: ["TV"], color: "#0ea5e9", active: true },
      { id: "r2", floor_id: "f0", name: "Sala Apoio", capacity: 6, location: "Corredor B", equipment: ["TV", "Quadro"], color: "#14b8a6", active: true },
      { id: "r3", floor_id: "f1", name: "Sala Atlântico", capacity: 8, location: "Ala leste", equipment: ["TV", "Webcam"], color: "#2563eb", active: true },
      { id: "r4", floor_id: "f1", name: "Sala Pacífico", capacity: 12, location: "Ala oeste", equipment: ["Projetor", "Webcam"], color: "#7c3aed", active: true },
      { id: "r5", floor_id: "f3", name: "Sala Diretoria", capacity: 16, location: "Sala principal", equipment: ["Projetor", "TV", "Webcam", "Telefone"], color: "#dc2626", active: true },
    ],
    bookings: [],
  };

  // alguns agendamentos de exemplo para hoje
  (function seedBookings() {
    const today = new Date();
    const at = (h, m = 0) => {
      const d = new Date(today);
      d.setHours(h, m, 0, 0);
      return d.toISOString();
    };
    demo.bookings = [
      { id: "b1", room_id: "r3", title: "Reunião de Vendas", organizer_name: "Ana Souza", organizer_email: "ana@empresa.com", department: "Comercial", attendees: 5, start_time: at(9), end_time: at(10), status: "confirmed", notes: "" },
      { id: "b2", room_id: "r4", title: "Planejamento Q3", organizer_name: "Carlos Lima", organizer_email: "carlos@empresa.com", department: "Diretoria", attendees: 8, start_time: at(14), end_time: at(15, 30), status: "confirmed", notes: "" },
      { id: "b3", room_id: "r5", title: "Comitê Executivo", organizer_name: "Marina Reis", organizer_email: "marina@empresa.com", department: "Diretoria", attendees: 12, start_time: at(11), end_time: at(12), status: "confirmed", notes: "" },
    ];
  })();

  const clone = (x) => JSON.parse(JSON.stringify(x));
  const overlaps = (aS, aE, bS, bE) => new Date(aS) < new Date(bE) && new Date(bS) < new Date(aE);

  // ---------------------------------------------------------------------
  // API pública
  // ---------------------------------------------------------------------
  const DB = {
    isDemo: !supa,

    // ---------- FLOORS ----------
    async listFloors() {
      if (!supa) return clone(demo.floors).sort((a, b) => (a.number ?? 0) - (b.number ?? 0));
      const { data, error } = await supa.from("floors").select("*").order("number", { ascending: true });
      if (error) throw error;
      return data;
    },
    async createFloor(f) {
      if (!supa) { const row = { id: uid(), ...f }; demo.floors.push(row); return row; }
      const { data, error } = await supa.from("floors").insert(f).select().single();
      if (error) throw error;
      return data;
    },
    async updateFloor(id, patch) {
      if (!supa) { const r = demo.floors.find((x) => x.id === id); Object.assign(r, patch); return clone(r); }
      const { data, error } = await supa.from("floors").update(patch).eq("id", id).select().single();
      if (error) throw error;
      return data;
    },
    async deleteFloor(id) {
      if (!supa) {
        demo.floors = demo.floors.filter((x) => x.id !== id);
        const roomIds = demo.rooms.filter((r) => r.floor_id === id).map((r) => r.id);
        demo.rooms = demo.rooms.filter((r) => r.floor_id !== id);
        demo.bookings = demo.bookings.filter((b) => !roomIds.includes(b.room_id));
        return;
      }
      const { error } = await supa.from("floors").delete().eq("id", id);
      if (error) throw error;
    },

    // ---------- ROOMS ----------
    async listRooms() {
      if (!supa) return clone(demo.rooms);
      const { data, error } = await supa.from("rooms").select("*").order("name", { ascending: true });
      if (error) throw error;
      return data;
    },
    async createRoom(r) {
      if (!supa) { const row = { id: uid(), active: true, equipment: [], ...r }; demo.rooms.push(row); return row; }
      const { data, error } = await supa.from("rooms").insert(r).select().single();
      if (error) throw error;
      return data;
    },
    async updateRoom(id, patch) {
      if (!supa) { const r = demo.rooms.find((x) => x.id === id); Object.assign(r, patch); return clone(r); }
      const { data, error } = await supa.from("rooms").update(patch).eq("id", id).select().single();
      if (error) throw error;
      return data;
    },
    async deleteRoom(id) {
      if (!supa) {
        demo.rooms = demo.rooms.filter((x) => x.id !== id);
        demo.bookings = demo.bookings.filter((b) => b.room_id !== id);
        return;
      }
      const { error } = await supa.from("rooms").delete().eq("id", id);
      if (error) throw error;
    },

    // ---------- BOOKINGS ----------
    async listBookings({ from, to } = {}) {
      if (!supa) {
        let rows = clone(demo.bookings).filter((b) => b.status === "confirmed");
        if (from) rows = rows.filter((b) => new Date(b.end_time) >= new Date(from));
        if (to) rows = rows.filter((b) => new Date(b.start_time) <= new Date(to));
        return rows.sort((a, b) => new Date(a.start_time) - new Date(b.start_time));
      }
      let q = supa.from("bookings").select("*").eq("status", "confirmed").order("start_time", { ascending: true });
      if (from) q = q.gte("end_time", from);
      if (to) q = q.lte("start_time", to);
      const { data, error } = await q;
      if (error) throw error;
      return data;
    },

    // Verifica conflito antes de salvar (no banco há também a constraint).
    async hasConflict(roomId, startISO, endISO, ignoreId = null) {
      if (!supa) {
        return demo.bookings.some(
          (b) => b.status === "confirmed" && b.room_id === roomId && b.id !== ignoreId &&
            overlaps(startISO, endISO, b.start_time, b.end_time)
        );
      }
      let q = supa.from("bookings").select("id").eq("room_id", roomId).eq("status", "confirmed")
        .lt("start_time", endISO).gt("end_time", startISO);
      if (ignoreId) q = q.neq("id", ignoreId);
      const { data, error } = await q;
      if (error) throw error;
      return data.length > 0;
    },

    async createBooking(b) {
      if (await this.hasConflict(b.room_id, b.start_time, b.end_time)) {
        throw new Error("Já existe um agendamento confirmado nesse horário para esta sala.");
      }
      if (!supa) { const row = { id: uid(), status: "confirmed", ...b }; demo.bookings.push(row); return row; }
      const { data, error } = await supa.from("bookings").insert({ ...b, status: "confirmed" }).select().single();
      if (error) {
        if (String(error.message || "").includes("bookings_no_overlap"))
          throw new Error("Já existe um agendamento confirmado nesse horário para esta sala.");
        throw error;
      }
      return data;
    },

    async cancelBooking(id, reason) {
      const patch = { status: "cancelled", cancel_reason: reason || null };
      if (!supa) {
        const b = demo.bookings.find((x) => x.id === id);
        if (b) Object.assign(b, patch);
        return;
      }
      const { error } = await supa.from("bookings").update(patch).eq("id", id);
      if (error) throw error;
    },

    // ---------- AUTH ----------
    async getSession() {
      if (!supa) return { user: { id: "demo", email: "demo@empresa.com", user_metadata: { name: "Demonstração" } } };
      const { data } = await supa.auth.getSession();
      return data.session;
    },

    onAuthChange(callback) {
      if (!supa) return () => {};
      const { data: { subscription } } = supa.auth.onAuthStateChange((_event, session) => {
        callback(session);
      });
      return () => subscription.unsubscribe();
    },

    async signUp(email, password, name) {
      if (!supa) throw new Error("Auth requer Supabase configurado.");
      const { data, error } = await supa.auth.signUp({
        email,
        password,
        options: { data: { name } },
      });
      if (error) throw error;
      return data;
    },

    async signIn(email, password) {
      if (!supa) throw new Error("Auth requer Supabase configurado.");
      const { data, error } = await supa.auth.signInWithPassword({ email, password });
      if (error) throw error;
      return data;
    },

    async signOut() {
      if (!supa) return;
      const { error } = await supa.auth.signOut();
      if (error) throw error;
    },

    /* Faz upload da imagem de assinatura para o Supabase Storage
       e retorna a URL pública. Bucket: "signatures" (público, somente leitura). */
    async uploadSignature(file, userId) {
      if (!supa) throw new Error("Supabase não configurado.");
      const ext  = (file.name.split(".").pop() || "png").toLowerCase();
      const path = `${userId || "anon"}/${Date.now()}.${ext}`;
      const { error } = await supa.storage
        .from("signatures")
        .upload(path, file, { upsert: true, contentType: file.type });
      if (error) throw error;
      const { data } = supa.storage.from("signatures").getPublicUrl(path);
      return data.publicUrl;
    },
  };

  window.DB = DB;
})();
