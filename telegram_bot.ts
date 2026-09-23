// ═══════════════════════════════════════════════════════════
// Телеграм-бот для чеков за Яндекс и Индрайвер
//
// Что делает: принимает PDF-чек из Kaspi, вытаскивает из него сумму,
// дату, время и службу, спрашивает № заказа и кладёт всё в таблицу
// receipts. Инструмент потом сводит эти строки с отчётом 0080.
//
// Развернуть:
//   supabase functions deploy telegram --no-verify-jwt
//   supabase secrets set TELEGRAM_TOKEN=... TELEGRAM_SECRET=... ALLOWED_CHAT_ID=...
// ═══════════════════════════════════════════════════════════

import { extractText, getDocumentProxy } from "npm:unpdf@0.12.1";

const TG_TOKEN = Deno.env.get("TELEGRAM_TOKEN") ?? "";
const TG_SECRET = Deno.env.get("TELEGRAM_SECRET") ?? "";
const ALLOWED = (Deno.env.get("ALLOWED_CHAT_ID") ?? "")
  .split(",").map((s) => s.trim()).filter(Boolean);

const SB_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SB_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";

const api = (m: string) => `https://api.telegram.org/bot${TG_TOKEN}/${m}`;

async function send(chat: number, text: string, keyboard?: unknown) {
  await fetch(api("sendMessage"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chat, text, parse_mode: "HTML",
      ...(keyboard ? { reply_markup: keyboard } : {}),
    }),
  });
}

// ── Работа с базой напрямую по REST, сервисным ключом ──
async function db(path: string, init: RequestInit = {}) {
  const r = await fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`,
      "Content-Type": "application/json", ...(init.headers ?? {}),
    },
  });
  if (!r.ok) throw new Error(await r.text());
  return r.status === 204 ? null : await r.json();
}

// ── Разбор чека ────────────────────────────────────────────
// Kaspi печатает суммы как «2 551,00 ₸», дату как «01.09.26 14:17».
// Берём сумму рядом со словом «Сумма», иначе самую крупную из найденных.
function parseReceipt(text: string) {
  const t = text.replace(/ /g, " ").replace(/\s+/g, " ");

  const money = (s: string) =>
    parseFloat(s.replace(/[\s ]/g, "").replace(",", "."));
  const RE_SUM = /(\d[\d\s ]*(?:[.,]\d{2})?)\s*(?:₸|тг|KZT|тенге)/gi;

  let amount = 0;
  const near = t.match(/Сумма[^0-9]{0,12}(\d[\d\s ]*(?:[.,]\d{2})?)/i);
  if (near) amount = money(near[1]);
  if (!amount) {
    for (const m of t.matchAll(RE_SUM)) {
      const v = money(m[1]);
      if (v > amount) amount = v;
    }
  }

  let paid_date = "", paid_time = "";
  const d = t.match(/(\d{2})[.\-/](\d{2})[.\-/](\d{2,4})(?:[^\d]{0,8}(\d{2}):(\d{2}))?/);
  if (d) {
    const year = d[3].length === 2 ? "20" + d[3] : d[3];
    paid_date = `${year}-${d[2]}-${d[1]}`;
    paid_time = d[4] ? `${d[4]}:${d[5]}` : "";
  }

  let service = "";
  if (/индрайвер|indriv|indrive/i.test(t)) service = "Индрайвер";
  else if (/яндекс|yandex/i.test(t)) service = "Яндекс";

  return { amount, paid_date, paid_time, service };
}

async function pdfText(bytes: Uint8Array) {
  try {
    const pdf = await getDocumentProxy(bytes);
    const { text } = await extractText(pdf, { mergePages: true });
    return String(text || "");
  } catch {
    return "";
  }
}

// ── Файлы Telegram ─────────────────────────────────────────
async function tgFile(fileId: string) {
  const info = await (await fetch(api(`getFile?file_id=${fileId}`))).json();
  const path = info?.result?.file_path;
  if (!path) return null;
  const r = await fetch(`https://api.telegram.org/file/bot${TG_TOKEN}/${path}`);
  return { bytes: new Uint8Array(await r.arrayBuffer()), name: path.split("/").pop() ?? "file" };
}

async function upload(name: string, bytes: Uint8Array, type: string) {
  const key = `${new Date().toISOString().slice(0, 10)}/${Date.now()}_${name}`;
  const r = await fetch(`${SB_URL}/storage/v1/object/receipts/${key}`, {
    method: "POST",
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, "Content-Type": type },
    body: bytes,
  });
  return r.ok ? key : "";
}

const fmt = (n: number) =>
  new Intl.NumberFormat("ru-RU").format(Math.round(n)) + " ₸";

const SVC_KB = {
  inline_keyboard: [[
    { text: "Яндекс", callback_data: "svc:Яндекс" },
    { text: "Индрайвер", callback_data: "svc:Индрайвер" },
  ]],
};

// Последний незакрытый чек этого чата — к нему относится следующий ответ
async function pending(chat: number) {
  const rows = await db(
    `receipts?select=*&tg_chat_id=eq.${chat}&status=eq.pending&order=created_at.desc&limit=1`,
  );
  return rows?.[0] ?? null;
}

async function askNext(chat: number, row: Record<string, unknown>) {
  if (!row.service) {
    await send(chat, "Какая служба?", SVC_KB);
    return;
  }
  if (!row.amount) {
    await send(chat, "Не разобрала сумму. Напишите её числом, например <b>2551</b>");
    return;
  }
  if (!row.order_num) {
    await send(
      chat,
      `<b>${row.service}</b> · ${row.paid_date ? String(row.paid_date).split("-").reverse().slice(0, 2).join(".") : "дата ?"}` +
        `${row.paid_time ? " " + row.paid_time : ""} · <b>${fmt(Number(row.amount))}</b>\n\n` +
        "Отправьте <b>№ заказа</b> из iiko. Если его нет — отправьте <b>-</b>",
    );
    return;
  }
  await db(`receipts?id=eq.${row.id}`, {
    method: "PATCH", headers: { Prefer: "return=minimal" },
    body: JSON.stringify({ status: "ready" }),
  });
  await send(chat, `✅ Записала: ${row.service} · ${fmt(Number(row.amount))} · заказ №${row.order_num}`);
}

// ═══════════════════════════════════════════════════════════
Deno.serve(async (req) => {
  if (req.headers.get("x-telegram-bot-api-secret-token") !== TG_SECRET) {
    return new Response("forbidden", { status: 403 });
  }
  let update: any;
  try { update = await req.json(); } catch { return new Response("ok"); }

  try {
    // Нажатие на кнопку выбора службы
    const cq = update.callback_query;
    if (cq) {
      const chat = cq.message.chat.id;
      await fetch(api("answerCallbackQuery"), {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ callback_query_id: cq.id }),
      });
      if (ALLOWED.length && !ALLOWED.includes(String(chat))) return new Response("ok");
      const row = await pending(chat);
      if (row && String(cq.data).startsWith("svc:")) {
        const service = String(cq.data).slice(4);
        await db(`receipts?id=eq.${row.id}`, {
          method: "PATCH", headers: { Prefer: "return=minimal" },
          body: JSON.stringify({ service }),
        });
        await askNext(chat, { ...row, service });
      }
      return new Response("ok");
    }

    const msg = update.message ?? update.channel_post;
    if (!msg) return new Response("ok");
    const chat = msg.chat.id;

    if (ALLOWED.length && !ALLOWED.includes(String(chat))) {
      await send(chat, `Этот чат не в списке разрешённых. Ваш ID: <code>${chat}</code>`);
      return new Response("ok");
    }

    if (msg.text === "/start" || msg.text === "/help") {
      await send(
        chat,
        "Пришлите <b>PDF-чек</b> из Kaspi за Яндекс или Индрайвер.\n" +
          "Я вытащу сумму, дату и службу, спрошу № заказа — и запишу.\n\n" +
          "Можно и скриншот: тогда сумму спрошу отдельно.\n" +
          "<b>/last</b> — последние 5 записей.",
      );
      return new Response("ok");
    }

    if (msg.text === "/last") {
      const rows = await db("receipts?select=service,paid_date,amount,order_num,status&order=created_at.desc&limit=5");
      await send(
        chat,
        rows?.length
          ? rows.map((r: any) =>
            `${r.paid_date ?? "—"} · ${r.service ?? "—"} · ${fmt(Number(r.amount || 0))} · №${r.order_num ?? "—"}${r.status === "pending" ? " ⏳" : ""}`
          ).join("\n")
          : "Пока пусто",
      );
      return new Response("ok");
    }

    // ── Файл: PDF или картинка ──
    const doc = msg.document;
    const photo = msg.photo?.[msg.photo.length - 1];
    if (doc || photo) {
      const f = await tgFile(doc ? doc.file_id : photo.file_id);
      if (!f) { await send(chat, "Не смогла скачать файл, попробуйте ещё раз"); return new Response("ok"); }

      const isPdf = !!doc && /pdf$/i.test(doc.file_name ?? doc.mime_type ?? "");
      const raw = isPdf ? await pdfText(f.bytes) : "";
      const p = raw ? parseReceipt(raw) : { amount: 0, paid_date: "", paid_time: "", service: "" };
      const key = await upload(f.name, f.bytes, doc?.mime_type ?? "image/jpeg");

      const ins = await db("receipts", {
        method: "POST", headers: { Prefer: "return=representation" },
        body: JSON.stringify({
          service: p.service || null,
          paid_date: p.paid_date || null,
          paid_time: p.paid_time || null,
          amount: p.amount || null,
          file_path: key || null,
          raw_text: raw ? raw.slice(0, 4000) : null,
          tg_chat_id: chat,
          tg_user: [msg.from?.first_name, msg.from?.username].filter(Boolean).join(" @"),
        }),
      });
      const row = ins?.[0];
      if (!isPdf) await send(chat, "Это не PDF — сумму и службу спрошу вручную.");
      await askNext(chat, row);
      return new Response("ok");
    }

    // ── Текст: сумма или № заказа для последнего чека ──
    const text = (msg.text ?? "").trim();
    if (!text) return new Response("ok");
    const row = await pending(chat);
    if (!row) { await send(chat, "Сначала пришлите чек."); return new Response("ok"); }

    if (!row.amount) {
      const v = parseFloat(text.replace(/[^\d.,]/g, "").replace(",", "."));
      if (!v) { await send(chat, "Не похоже на сумму. Напишите числом, например 2551"); return new Response("ok"); }
      await db(`receipts?id=eq.${row.id}`, {
        method: "PATCH", headers: { Prefer: "return=minimal" },
        body: JSON.stringify({ amount: v }),
      });
      await askNext(chat, { ...row, amount: v });
      return new Response("ok");
    }

    const num = text === "-" ? "-" : (text.match(/\d{3,6}/)?.[0] ?? "");
    if (!num) { await send(chat, "Не нашла номер. Отправьте № заказа цифрами или <b>-</b>"); return new Response("ok"); }
    await db(`receipts?id=eq.${row.id}`, {
      method: "PATCH", headers: { Prefer: "return=minimal" },
      body: JSON.stringify({ order_num: num, status: "ready" }),
    });
    await send(
      chat,
      num === "-"
        ? `✅ Записала без номера: ${row.service} · ${fmt(Number(row.amount))}. Свяжете в отчёте вручную.`
        : `✅ Записала: ${row.service} · ${fmt(Number(row.amount))} · заказ №${num}`,
    );
  } catch (e) {
    console.error(e);
  }
  return new Response("ok");
});
