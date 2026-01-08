# app.py
from __future__ import annotations

import io
import json
import os
import re
from typing import List, Dict
from urllib.parse import quote

import requests
from flask import Flask, Response, render_template_string, request

app = Flask(__name__)

HTML = r"""
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>単語帳</title>
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,"Noto Sans JP",sans-serif;margin:24px;background:#f6f7fb;}
    .wrap{max-width:860px;margin:0 auto;}
    h1{font-size:22px;margin:0;}
    .card{background:#fff;border-radius:14px;box-shadow:0 6px 18px rgba(0,0,0,.06);padding:16px;margin-bottom:14px;}
    form{display:grid;grid-template-columns:1fr 1fr auto auto;gap:10px;align-items:end;}
    label{display:block;font-size:12px;color:#444;margin-bottom:6px;}
    input{width:100%;padding:10px 12px;border:1px solid #d7dbe7;border-radius:10px;font-size:14px;}
    button,.btn{appearance:none;border:0;border-radius:10px;padding:10px 14px;font-size:14px;cursor:pointer;text-decoration:none;display:inline-block}
    button{background:#2563eb;color:#fff;}
    .btn{background:#111827;color:#fff;}
    .row{display:flex;gap:10px;flex-wrap:wrap}
    table{width:100%;border-collapse:collapse;}
    th,td{padding:10px 8px;border-bottom:1px solid #edf0f6;font-size:14px;vertical-align:top}
    th{color:#374151;text-align:left;font-weight:600;}
    .del{background:#ef4444;}
    .muted{color:#6b7280;font-size:12px;margin-top:8px}
    .topbar{display:flex;gap:10px;flex-wrap:wrap;justify-content:space-between;align-items:center;margin-bottom:12px}
    .hint{color:#6b7280;font-size:12px;margin-top:10px}
    .smallbtn{background:#111827;}
    .smallbtn:disabled{opacity:.5;cursor:not-allowed;}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <h1>単語帳</h1>
      <div class="row">
        <button class="btn" id="btnCsv" type="button">CSV</button>
        <button class="btn" id="btnPdf" type="button">PDF</button>
        <button class="btn" id="btnClear" type="button" style="background:#6b7280">全消去</button>
      </div>
    </div>

    <div class="card">
      <form id="addForm">
        <div>
          <label>word</label>
          <input id="word" autocomplete="off" required maxlength="80" />
        </div>
        <div>
          <label>単語</label>
          <input id="meaning" autocomplete="off" required maxlength="200" />
        </div>
        <div>
          <button class="smallbtn" id="btnLookup" type="button">翻訳</button>
        </div>
        <div>
          <button type="submit">記録</button>
        </div>
      </form>
      <div class="hint" id="hint"></div>
      <div class="muted">データはこのブラウザ内に保存されます。</div>
    </div>

    <div class="card">
      <table>
        <thead>
          <tr>
            <th style="width:30%">word</th>
            <th>単語</th>
            <th style="width:90px">操作</th>
          </tr>
        </thead>
        <tbody id="tbody">
          <tr><td colspan="3" style="color:#6b7280;padding:16px 8px">まだありません。</td></tr>
        </tbody>
      </table>
    </div>
  </div>

<script>
(() => {
  const KEY = "wordbook_items_v1";

  const $ = (id) => document.getElementById(id);
  const tbody = $("tbody");
  const form = $("addForm");
  const wordEl = $("word");
  const meaningEl = $("meaning");
  const btnCsv = $("btnCsv");
  const btnPdf = $("btnPdf");
  const btnClear = $("btnClear");
  const btnLookup = $("btnLookup");
  const hint = $("hint");

  function loadItems() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return [];
      const v = JSON.parse(raw);
      if (!Array.isArray(v)) return [];
      return v.filter(x => x && typeof x.word === "string" && typeof x.meaning === "string");
    } catch {
      return [];
    }
  }

  function saveItems(items) {
    localStorage.setItem(KEY, JSON.stringify(items));
  }

  function escapeHtml(s) {
    return s.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#39;");
  }

  function setHint(s) {
    hint.textContent = s || "";
  }

  function render() {
    const items = loadItems();
    if (items.length === 0) {
      tbody.innerHTML = '<tr><td colspan="3" style="color:#6b7280;padding:16px 8px">まだありません。</td></tr>';
      return;
    }
    tbody.innerHTML = items.map((it, i) => `
      <tr>
        <td>${escapeHtml(it.word)}</td>
        <td>${escapeHtml(it.meaning)}</td>
        <td><button class="del" type="button" data-del="${i}">削除</button></td>
      </tr>
    `).join("");
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const word = wordEl.value.trim();
    const meaning = meaningEl.value.trim();
    if (!word || !meaning) return;

    const items = loadItems();
    items.push({word, meaning});
    saveItems(items);

    wordEl.value = "";
    meaningEl.value = "";
    setHint("");
    wordEl.focus();

    render();
  });

  tbody.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-del]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-del"));
    const items = loadItems();
    if (Number.isInteger(idx) && idx >= 0 && idx < items.length) {
      items.splice(idx, 1);
      saveItems(items);
      render();
    }
  });

  btnClear.addEventListener("click", () => {
    saveItems([]);
    render();
  });

  btnCsv.addEventListener("click", () => {
    const items = loadItems();
    const bom = "\uFEFF";
    const lines = ["word,meaning"].concat(
      items.map(it => {
        const w = String(it.word).replaceAll('"','""');
        const m = String(it.meaning).replaceAll('"','""');
        return `"${w}","${m}"`;
      })
    );
    const csv = bom + lines.join("\r\n");
    const blob = new Blob([csv], {type: "text/csv;charset=utf-8"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "wordbook.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnPdf.addEventListener("click", async () => {
    const items = loadItems();
    const res = await fetch("/export.pdf", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({items})
    });
    if (!res.ok) {
      alert("PDF出力に失敗しました");
      return;
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "wordbook.pdf";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnLookup.addEventListener("click", async () => {
    const word = wordEl.value.trim();
    if (!word) return;

    btnLookup.disabled = true;
    setHint("検索中…");

    try {
      const res = await fetch("/lookup", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({word})
      });
      if (!res.ok) {
        setHint("失敗しました");
        return;
      }
      const data = await res.json();
      if (data && data.meaning) {
        meaningEl.value = data.meaning;
        setHint("");
        meaningEl.focus();
      } else {
        setHint("見つかりませんでした");
      }
    } catch {
      setHint("失敗しました");
    } finally {
      btnLookup.disabled = false;
    }
  });

  render();
})();
</script>
</body>
</html>
"""

@app.get("/")
def index():
    return render_template_string(HTML)


def _draw_pdf_word_sheet(rows: List[Dict[str, str]]) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    jp_font = "HeiseiKakuGo-W5"
    pdfmetrics.registerFont(UnicodeCIDFont(jp_font))

    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    width, height = A4

    margin_x = 12 * mm
    margin_top = 12 * mm
    margin_bottom = 12 * mm

    table_top = height - margin_top
    table_bottom = margin_bottom

    total_rows = 25
    row_h = (table_top - table_bottom) / total_rows

    gap = 8 * mm
    half_w = (width - 2 * margin_x - gap) / 2

    no_w = 10 * mm
    text_w = (half_w - no_w) / 2
    word_w = text_w
    meaning_w = text_w

    pad = 2 * mm

    def clean(s: str) -> str:
        return (s or "").replace("\r", " ").replace("\n", " ").strip()

    def text_width(font_name: str, size: float, s: str) -> float:
        return pdfmetrics.stringWidth(s, font_name, size)

    def truncate_to_fit(font_name: str, size: float, s: str, max_w: float) -> str:
        s = clean(s)
        if text_width(font_name, size, s) <= max_w:
            return s
        ell = "…"
        if text_width(font_name, size, ell) > max_w:
            return ""
        t = s
        while t and text_width(font_name, size, t + ell) > max_w:
            t = t[:-1]
        return (t + ell) if t else ell

    def draw_fit_text(font_name: str, base_size: float, min_size: float, x: float, y: float, s: str, max_w: float):
        s = clean(s)
        if not s:
            return
        w0 = text_width(font_name, base_size, s)
        size = base_size
        if w0 > max_w and w0 > 0:
            size = base_size * (max_w / w0)
            if size < min_size:
                size = min_size
        s2 = truncate_to_fit(font_name, size, s, max_w)
        c.setFont(font_name, size)
        c.drawString(x, y, s2)

    def draw_page(page_rows: List[Dict[str, str]], start_no: int):
        left_x0 = margin_x
        left_x1 = left_x0 + half_w
        right_x0 = left_x1 + gap
        right_x1 = right_x0 + half_w

        c.setLineWidth(0.6)

        for x0, x1 in ((left_x0, left_x1), (right_x0, right_x1)):
            c.line(x0, table_bottom, x0, table_top)
            c.line(x1, table_bottom, x1, table_top)
            c.line(x0 + no_w, table_bottom, x0 + no_w, table_top)
            c.line(x0 + no_w + word_w, table_bottom, x0 + no_w + word_w, table_top)
            for r in range(total_rows + 1):
                y = table_top - r * row_h
                c.line(x0, y, x1, y)

        no_size = 10
        meaning_base = 10
        word_base = 11
        min_size = 6

        for i in range(min(50, len(page_rows))):
            n = start_no + i
            side = 0 if i < 25 else 1
            row = i if i < 25 else i - 25

            x0 = left_x0 if side == 0 else right_x0
            y_top = table_top - row * row_h
            y_text = y_top - 0.72 * row_h

            word = str(page_rows[i].get("word", ""))
            meaning = str(page_rows[i].get("meaning", ""))

            c.setFont(jp_font, no_size)
            c.drawString(x0 + pad, y_text, str(n))

            draw_fit_text("Helvetica", word_base, min_size, x0 + no_w + pad, y_text, word, word_w - 2 * pad)
            draw_fit_text(jp_font, meaning_base, min_size, x0 + no_w + word_w + pad, y_text, meaning, meaning_w - 2 * pad)

    cleaned: List[Dict[str, str]] = []
    for it in rows:
        if isinstance(it, dict):
            cleaned.append(
                {
                    "word": clean(str(it.get("word", "")))[:80],
                    "meaning": clean(str(it.get("meaning", "")))[:200],
                }
            )

    page_size = 50
    start_no = 1
    for p in range(0, len(cleaned) if cleaned else 1, page_size):
        chunk = cleaned[p : p + page_size] if cleaned else []
        draw_page(chunk, start_no)
        start_no += page_size
        if p + page_size < len(cleaned):
            c.showPage()

    c.save()
    return out.getvalue()


@app.post("/export.pdf")
def export_pdf():
    data = request.get_data(cache=False, as_text=True) or "{}"
    try:
        payload = json.loads(data)
    except Exception:
        payload = {}

    items = payload.get("items")
    if not isinstance(items, list):
        items = []

    pdf = _draw_pdf_word_sheet(items)
    return Response(
        pdf,
        mimetype="application/pdf",
        headers={"Content-Disposition": 'attachment; filename="wordbook.pdf"'},
    )


def _fetch_wiktionary_raw(title: str) -> str:
    t = title.strip()
    if not t:
        return ""
    url = "https://en.wiktionary.org/wiki/" + quote(t) + "?action=raw"
    r = requests.get(url, timeout=6, headers={"User-Agent": "wordbook-app/1.0"})
    if r.status_code != 200:
        return ""
    txt = r.text or ""
    m = re.match(r"(?is)^\s*#redirect\s*\[\[(.+?)\]\]", txt)
    if m:
        target = m.group(1).strip()
        if target and target.lower() != t.lower():
            url2 = "https://en.wiktionary.org/wiki/" + quote(target) + "?action=raw"
            r2 = requests.get(url2, timeout=6, headers={"User-Agent": "wordbook-app/1.0"})
            if r2.status_code == 200:
                return r2.text or ""
    return txt


def _extract_ja_translations(wikitext: str, limit: int = 6) -> List[str]:
    if not wikitext:
        return []
    # {{t|ja|...}}, {{t+|ja|...}} から日本語語形を抜く
    hits = re.findall(r"\{\{t\+?\|ja\|([^|}\n]+)", wikitext)
    out: List[str] = []
    seen = set()
    for h in hits:
        s = (h or "").strip()
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
        if len(out) >= limit:
            break
    return out


@app.post("/lookup")
def lookup():
    data = request.get_data(cache=False, as_text=True) or "{}"
    try:
        payload = json.loads(data)
    except Exception:
        payload = {}

    word = str(payload.get("word", "")).strip()
    if not word:
        return Response(json.dumps({"meaning": ""}, ensure_ascii=False), mimetype="application/json")

    raw = _fetch_wiktionary_raw(word)
    ja = _extract_ja_translations(raw)

    meaning = "、".join(ja) if ja else ""
    return Response(json.dumps({"meaning": meaning}, ensure_ascii=False), mimetype="application/json")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
