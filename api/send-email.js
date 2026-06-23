/* ==========================================================
   Vercel Serverless Function — compositor de e-mail geral
   Variáveis necessárias:
     RESEND_API_KEY   chave Resend
     EMAIL_FROM       fallback: "Salas de Reunião <salas@suaempresa.com.br>"
                      (usado quando o domínio do remetente não está verificado)

   Domínios verificados no Resend → e-mail sai do próprio domínio do remetente.
   Domínios não verificados       → cai no EMAIL_FROM de fallback.
   ========================================================== */

/* Domínios verificados no Resend — configure via variável de ambiente RESEND_VERIFIED_DOMAINS
   (lista separada por vírgula) ou edite diretamente aqui após verificar no painel do Resend. */
const _verifiedEnv = process.env.RESEND_VERIFIED_DOMAINS || "";
const VERIFIED_DOMAINS = new Set(_verifiedEnv.split(",").map(d => d.trim()).filter(Boolean));

const DEFAULT_DOMAIN = process.env.RESEND_DEFAULT_DOMAIN || "suaempresa.com.br";

function buildFrom(senderEmail, senderName) {
  const domain = (senderEmail || "").split("@")[1] || "";
  if (VERIFIED_DOMAINS.has(domain) && senderEmail) {
    return `${senderName || senderEmail} <${senderEmail}>`;
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

  const { to = [], cc = [], subject, message, signature, senderEmail, senderName } = req.body || {};

  if (!to.length)       return res.status(400).json({ error: "Informe ao menos um destinatário." });
  if (!subject?.trim()) return res.status(400).json({ error: "Informe o assunto." });
  if (!message?.trim()) return res.status(400).json({ error: "Informe a mensagem." });

  const APP_URL = "https://salas-de-reuniao.vercel.app";

  /* ---- assinatura HTML ---- */
  let sigHtml = "";
  if (signature) {
    const mode = signature.sig_mode;
    if (mode === "html" && signature.sig_html?.trim()) {
      sigHtml = `<div style="margin-top:28px;border-top:2px solid #e4e9e7;padding-top:20px">${signature.sig_html}</div>`;
    } else if (mode === "image" && signature.sig_image?.trim()) {
      sigHtml = `<div style="margin-top:28px;border-top:2px solid #e4e9e7;padding-top:20px"><img src="${signature.sig_image}" alt="Assinatura" style="max-width:600px"></div>`;
    } else if (signature.name || signature.department) {
      /* builder padrão */
      sigHtml = `
        <table cellpadding="0" cellspacing="0" border="0"
          style="margin-top:28px;border-top:2px solid #e4e9e7;padding-top:20px;font-family:-apple-system,'Segoe UI',sans-serif;width:100%;max-width:520px">
          <tr>
            <td style="vertical-align:top;padding-right:20px">
              <div style="background:#0d2420;border-radius:10px;padding:14px 18px;min-width:200px">
                ${signature.name ? `<div style="color:#c8960c;font-weight:800;font-size:13.5px;letter-spacing:.02em;margin-bottom:2px">${signature.name.toUpperCase()}</div>` : ""}
                ${signature.department ? `<div style="color:#a8c5bb;font-size:11.5px;margin-bottom:8px">${signature.department}</div>` : ""}
                ${signature.phone ? `<div style="color:rgba(255,255,255,.75);font-size:11px;margin-bottom:2px">📞 ${signature.phone}</div>` : ""}
                ${signature.teams ? `<div style="color:rgba(255,255,255,.75);font-size:11px">Teams: ${signature.teams}</div>` : ""}
              </div>
            </td>
            <td style="vertical-align:middle;border-left:3px solid #c8960c;padding-left:18px">
              <div style="color:#0d2420;font-weight:900;font-size:11px;letter-spacing:.12em;text-transform:uppercase">GRUPO</div>
              <div style="color:#0d2420;font-weight:900;font-size:24px;letter-spacing:.04em;line-height:1">PLUMA</div>
              ${signature.company ? `<div style="color:#6c7d78;font-size:10.5px;margin-top:3px">${signature.company}</div>` : ""}
            </td>
          </tr>
          ${signature.address ? `<tr><td colspan="2" style="padding-top:10px"><div style="font-size:10.5px;color:#6c7d78">${signature.address}</div></td></tr>` : ""}
        </table>`;
    }
  }

  /* ---- corpo do e-mail ---- */
  const bodyHtml = message.trim().replace(/\n/g, "<br>");

  const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eef1ef;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif">
  <div style="max-width:600px;margin:32px auto;padding:0 16px 40px">

    <div style="text-align:center;margin-bottom:20px">
      <div style="display:inline-block;background:#1f574b;color:#fff;border-radius:8px;padding:7px 18px;font-size:12px;font-weight:700">
        📧 Salas de Reunião
      </div>
    </div>

    <div style="background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e4e9e7;box-shadow:0 2px 8px rgba(13,40,35,.06)">

      <div style="padding:28px 32px;border-bottom:1px solid #e4e9e7">
        <div style="font-size:15px;color:#16201d;line-height:1.75">${bodyHtml}</div>
        ${sigHtml}
      </div>

      <div style="padding:14px 32px;background:#f5f8f6">
        <a href="${APP_URL}" style="font-size:11.5px;color:#1f574b;text-decoration:none">
          🔗 Sistema de Salas de Reunião · ${APP_URL}
        </a>
      </div>
    </div>

  </div>
</body>
</html>`;

  try {
    /* from dinâmico: usa domínio do remetente se verificado, senão fallback */
    const from = process.env.EMAIL_FROM || buildFrom(senderEmail, senderName);
    const replyTo = senderEmail || signature?.teams || undefined;

    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${RESEND_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from,
        to,
        ...(cc.length ? { cc } : {}),
        ...(replyTo ? { reply_to: replyTo } : {}),
        subject: subject.trim(),
        html,
      }),
    });

    const result = await response.json();
    if (!response.ok) {
      console.error("Resend error:", result);
      return res.status(500).json({ error: result.message || "Falha ao enviar e-mail." });
    }

    return res.status(200).json({ ok: true, sent: to.length + cc.length, id: result.id });
  } catch (err) {
    console.error("send-email error:", err);
    return res.status(500).json({ error: err.message });
  }
}
