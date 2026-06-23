/* ==========================================================
   Vercel Serverless Function — dispara e-mail de convite
   após agendamento de sala.

   Variáveis de ambiente necessárias (Vercel Dashboard):
     RESEND_API_KEY  — chave da API do Resend (resend.com)
     EMAIL_FROM      — fallback: "Salas de Reunião <salas@plumaagro.com.br>"
   ========================================================== */

const VERIFIED_DOMAINS = new Set([
  "maisfrango.com.br", "plumaagro.com.br", "bellofoods.com.br",
  "belloalimentos.com.br", "plusvalagro.com.br", "levoalimentos.com.br",
]);
const DEFAULT_DOMAIN = "plumaagro.com.br";

/* Se o domínio do organizador está verificado, o e-mail sai do próprio endereço
   dele — aparece corretamente no Outlook/Google Calendar como organizador real.
   Caso contrário, usa o endereço genérico de fallback. */
function buildFrom(organizerEmail, organizerName) {
  const domain = (organizerEmail || "").split("@")[1] || "";
  if (VERIFIED_DOMAINS.has(domain)) {
    const name = organizerName || organizerEmail;
    return `${name} <${organizerEmail}>`;
  }
  return `Salas de Reunião <salas@${DEFAULT_DOMAIN}>`;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "https://salas-de-reuniao.vercel.app");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const RESEND_KEY = process.env.RESEND_API_KEY;
  if (!RESEND_KEY) return res.status(500).json({ error: "Email service not configured (missing RESEND_API_KEY)" });

  const { booking, room, floor, participants = [] } = req.body || {};
  if (!booking) return res.status(400).json({ error: "Missing booking data" });

  /* ---- destinatários ---- */
  const recipients = [];
  if (booking.organizer_email) recipients.push(booking.organizer_email);
  participants.forEach((p) => { if (p?.email?.trim()) recipients.push(p.email.trim()); });

  if (recipients.length === 0) return res.status(200).json({ ok: true, sent: 0 });

  /* ---- formatação de data / hora ---- */
  const toLocal = (iso) => {
    const d = new Date(iso);
    const pad = (n) => String(n).padStart(2, "0");
    return {
      date: d.toLocaleDateString("pt-BR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }),
      time: `${pad(d.getHours())}:${pad(d.getMinutes())}`,
    };
  };
  const start = toLocal(booking.start_time);
  const end   = toLocal(booking.end_time);
  const location = room ? `${room.name}${floor ? ` · ${floor.name}` : ""}` : "A definir";
  const APP_URL = "https://salas-de-reuniao.vercel.app";

  /* ---- bloco de participantes ---- */
  const participantsBlock = participants.length > 0 ? `
    <div style="margin-top:20px">
      <div style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6b7280;margin-bottom:10px">
        Participantes (${participants.length})
      </div>
      ${participants.map((p) => `
        <div style="display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:1px solid #f3f4f2">
          <div style="width:30px;height:30px;border-radius:50%;background:#e8f5ee;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0">
            👤
          </div>
          <div>
            ${p.name ? `<div style="font-weight:600;font-size:13px;color:#1c1f1a">${p.name}</div>` : ""}
            ${p.email ? `<div style="font-size:12px;color:#6b7280;margin-top:2px">📧 ${p.email}</div>` : ""}
            ${p.phone ? `<div style="font-size:12px;color:#6b7280;margin-top:1px">📞 ${p.phone}</div>` : ""}
          </div>
        </div>`).join("")}
    </div>` : "";

  /* ---- template HTML do e-mail ---- */
  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${booking.title}</title>
</head>
<body style="margin:0;padding:0;background:#f3f4f2;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
  <div style="max-width:580px;margin:32px auto;padding:0 16px 40px">

    <!-- Brand -->
    <div style="text-align:center;margin-bottom:24px">
      <div style="display:inline-block;background:#1a6b3a;color:#fff;border-radius:10px;padding:9px 20px;font-weight:700;font-size:13.5px;letter-spacing:.01em">
        📅 Salas de Reunião · Grupo Pluma
      </div>
    </div>

    <!-- Card principal -->
    <div style="background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e5e7eb;box-shadow:0 2px 8px rgba(0,0,0,.06)">

      <!-- Header verde -->
      <div style="background:linear-gradient(135deg,#1a6b3a 0%,#22863f 100%);padding:28px 30px">
        <div style="color:rgba(255,255,255,.72);font-size:11.5px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;margin-bottom:8px">
          Nova reunião agendada
        </div>
        <div style="color:#fff;font-size:22px;font-weight:800;line-height:1.25;margin-bottom:4px">
          ${booking.title}
        </div>
        ${booking.department ? `<div style="color:rgba(255,255,255,.7);font-size:13px;margin-top:6px">${booking.department}</div>` : ""}
      </div>

      <!-- Info da reunião -->
      <div style="margin:22px 30px 0;background:#e8f5ee;border-radius:12px;padding:18px 20px">
        <table style="width:100%;border-collapse:collapse">
          <tr>
            <td style="padding:5px 0;font-size:13.5px;color:#145c30;font-weight:600">
              📍 &nbsp;${location}
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;font-size:13.5px;color:#145c30">
              📅 &nbsp;${start.date}
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;font-size:13.5px;color:#145c30">
              🕐 &nbsp;${start.time} – ${end.time}
            </td>
          </tr>
          ${booking.attendees ? `<tr>
            <td style="padding:5px 0;font-size:13px;color:#1a6b3a">
              👥 &nbsp;${booking.attendees} participante${booking.attendees > 1 ? "s" : ""}
            </td>
          </tr>` : ""}
        </table>
      </div>

      <!-- Conteúdo -->
      <div style="padding:20px 30px 28px">

        <!-- Responsável -->
        <div style="margin-bottom:20px">
          <div style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6b7280;margin-bottom:10px">
            Responsável pelo agendamento
          </div>
          <div style="display:flex;align-items:center;gap:12px">
            <div style="width:38px;height:38px;border-radius:50%;background:#1a6b3a;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px;flex-shrink:0">
              ${(booking.organizer_name || "?")[0].toUpperCase()}
            </div>
            <div>
              <div style="font-weight:700;font-size:14px;color:#1c1f1a">${booking.organizer_name || ""}</div>
              ${booking.organizer_email ? `<div style="font-size:12.5px;color:#6b7280;margin-top:2px">📧 ${booking.organizer_email}</div>` : ""}
              ${booking.phone ? `<div style="font-size:12.5px;color:#6b7280;margin-top:1px">📞 ${booking.phone}</div>` : ""}
            </div>
          </div>
        </div>

        <!-- Participantes -->
        ${participantsBlock}

        <!-- Observações -->
        ${booking.notes ? `
        <div style="margin-top:20px;padding:14px 16px;background:#f9fafb;border-radius:10px;border-left:4px solid #1a6b3a">
          <div style="font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#6b7280;margin-bottom:6px">Observações</div>
          <div style="font-size:13px;color:#374151;line-height:1.65">${booking.notes.replace(/\n/g, "<br>")}</div>
        </div>` : ""}

        <!-- Divider -->
        <div style="height:1px;background:#f0f0f0;margin:24px 0"></div>

        <!-- CTA -->
        <div style="text-align:center">
          <a href="${APP_URL}"
            style="display:inline-block;background:#1a6b3a;color:#fff;text-decoration:none;padding:13px 32px;border-radius:9px;font-weight:700;font-size:14px;letter-spacing:.01em">
            Abrir sistema de salas →
          </a>
          <div style="margin-top:10px;font-size:11.5px;color:#9ca3af">
            ${APP_URL}
          </div>
        </div>
      </div>
    </div>

    <!-- Footer -->
    <div style="text-align:center;margin-top:24px;font-size:11.5px;color:#9ca3af;line-height:1.7">
      Você recebeu este e-mail porque foi incluído como participante desta reunião.<br>
      Grupo Pluma Agroavícola · Sistema de Gestão de Salas de Reunião
    </div>
  </div>
</body>
</html>`;

  /* ---- gera o arquivo .ics (convite de calendário) ---- */
  const fmtICS = (iso) => {
    const d = new Date(iso);
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getUTCFullYear()}${p(d.getUTCMonth()+1)}${p(d.getUTCDate())}T${p(d.getUTCHours())}${p(d.getUTCMinutes())}00Z`;
  };
  const uid = `booking-${Date.now()}@salas-de-reuniao.vercel.app`;
  const attendeeLines = recipients
    .map((e) => `ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:${e}`)
    .join("\r\n");
  const icsContent = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Salas de Reunião//Grupo Pluma//PT",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${fmtICS(new Date().toISOString())}`,
    `DTSTART:${fmtICS(booking.start_time)}`,
    `DTEND:${fmtICS(booking.end_time)}`,
    `SUMMARY:${booking.title}`,
    `LOCATION:${location}`,
    `DESCRIPTION:Responsável: ${booking.organizer_name || ""}${booking.phone ? ` | Tel: ${booking.phone}` : ""}${booking.notes ? `\\n\\n${booking.notes}` : ""}`,
    booking.organizer_email ? `ORGANIZER;CN="${booking.organizer_name || booking.organizer_email}":mailto:${booking.organizer_email}` : "",
    attendeeLines,
    "STATUS:CONFIRMED",
    "SEQUENCE:0",
    "BEGIN:VALARM",
    "TRIGGER:-PT15M",
    "ACTION:DISPLAY",
    "DESCRIPTION:Lembrete de reunião",
    "END:VALARM",
    "END:VEVENT",
    "END:VCALENDAR",
  ].filter(Boolean).join("\r\n");

  const icsBase64 = Buffer.from(icsContent).toString("base64");

  /* ---- envia via Resend ---- */
  try {
    const from = process.env.EMAIL_FROM || buildFrom(booking.organizer_email, booking.organizer_name);

    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${RESEND_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from,
        to: recipients,
        reply_to: booking.organizer_email || undefined,
        subject: `[Reunião] ${booking.title} — ${location} · ${start.time}–${end.time}`,
        html,
        attachments: [{
          filename: "convite-reuniao.ics",
          content: icsBase64,
        }],
      }),
    });

    const result = await response.json();

    if (!response.ok) {
      console.error("Resend API error:", result);
      return res.status(500).json({ error: result.message || "Failed to send email" });
    }

    return res.status(200).json({ ok: true, sent: recipients.length, resend_id: result.id });
  } catch (err) {
    console.error("send-booking-email error:", err);
    return res.status(500).json({ error: err.message });
  }
}
