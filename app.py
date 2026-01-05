# app.py
from __future__ import annotations

import csv
import io
import os
import threading
from dataclasses import dataclass
from typing import List

from flask import Flask, Response, redirect, render_template_string, request, url_for

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

app = Flask(__name__)

LOCK = threading.Lock()


@dataclass
class CardItem:
    word: str
    meaning: str


ITEMS: List[CardItem] = []


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
    form{display:grid;grid-template-columns:1fr 1fr auto;gap:10px;align-items:end;}
    label{display:block;font-size:12px;color:#444;margin-bottom:6px;}
    input{width:100%;padding:10px 12px;border:1px solid #d7dbe7;border-radius:10px;font-size:14px;}
    button,.btn{appearance:none;border:0;border-radius:10px;padding:10px 14px;font-size:14px;cursor:pointer;text-decoration:none;display:inline-block}
    button{background:#2563eb;color:#fff;}
    .btn{background:#111827;color:#fff;}
    .btn2{background:#10b981;color:#fff;}
    .row{display:flex;gap:10px;flex-wrap:wrap}
    table{width:100%;border-collapse:collapse;}
    th,td{padding:10px 8px;border-bottom:1px solid #edf0f6;font-size:14px;vertical-align:top}
    th{color:#374151;text-align:left;font-weight:600;}
    .del{background:#ef4444;}
    .muted{color:#6b7280;font-size:12px;margin-top:8px}
    .topbar{display:flex;gap:10px;flex-wrap:wrap;justify-content:space-between;align-items:center;margin-bottom:12px}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <h1>単語帳</h1>
      <div class="row">
        <a class="btn2" href="{{ url_for('export_csv') }}">CSV</a>
        <a class="btn" href="{{ url_for('export_pdf') }}">PDF</a>
        <a class="btn" style="background:#6b7280" href="{{ url_for('clear_all') }}">全消去</a>
      </div>
    </div>

    <div class="card">
      <form method="post" action="{{ url_for('add_item') }}">
        <div>
          <label>英単語</label>
          <input name="word" autocomplete="off" required maxlength="80" />
        </div>
        <div>
          <label>日本語の意味</label>
          <input name="meaning" autocomplete="off" required maxlength="200" />
        </div>
        <div>
          <button type="submit">記録</button>
        </div>
      </form>
      <div class="muted">入力して「記録」を押すと下のリストに追加されます。</div>
    </div>

    <div class="card">
      <table>
        <thead>
          <tr>
            <th style="width:30%">英単語</th>
            <th>意味</th>
            <th style="width:90px">操作</th>
          </tr>
        </thead>
        <tbody>
          {% if items %}
            {% for i, it in items %}
              <tr>
                <td>{{ it.word }}</td>
                <td>{{ it.meaning }}</td>
                <td>
                  <form method="post" action="{{ url_for('delete_item', idx=i) }}" style="margin:0">
                    <button class="del" type="submit">削除</button>
                  </form>
                </td>
              </tr>
            {% endfor %}
          {% else %}
            <tr><td colspan="3" style="color:#6b7280;padding:16px 8px">まだありません。</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>
"""


@app.get("/")
def index():
    with LOCK:
        items = list(enumerate(ITEMS))
    return render_template_string(HTML, items=items)


@app.post("/add")
def add_item():
    word = (request.form.get("word") or "").strip()
    meaning = (request.form.get("meaning") or "").strip()
    if word and meaning:
        with LOCK:
            ITEMS.append(CardItem(word=word, meaning=meaning))
    return redirect(url_for("index"))


@app.post("/delete/<int:idx>")
def delete_item(idx: int):
    with LOCK:
        if 0 <= idx < len(ITEMS):
            ITEMS.pop(idx)
    return redirect(url_for("index"))


@app.get("/clear")
def clear_all():
    with LOCK:
        ITEMS.clear()
    return redirect(url_for("index"))


@app.get("/export.csv")
def export_csv():
    with LOCK:
        rows = [(it.word, it.meaning) for it in ITEMS]

    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["word", "meaning"])
    for r in rows:
        w.writerow(r)

    data = buf.getvalue().encode("utf-8-sig")
    return Response(
        data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="wordbook.csv"'},
    )


def _draw_pdf(items: List[CardItem]) -> bytes:
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))

    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle("単語帳")

    width, height = A4
    left = 16 * mm
    right = 16 * mm
    top = 18 * mm
    bottom = 18 * mm

    c.setFont("HeiseiKakuGo-W5", 14)
    c.drawString(left, height - top, "単語帳")

    y = height - top - 12 * mm
    col1_w = 55 * mm

    def new_page():
        nonlocal y
        c.showPage()
        c.setFont("HeiseiKakuGo-W5", 10)
        y = height - top

    def header():
        nonlocal y
        c.setFont("HeiseiKakuGo-W5", 11)
        c.drawString(left, y, "英単語")
        c.drawString(left + col1_w + 8 * mm, y, "意味")
        c.setFont("HeiseiKakuGo-W5", 10)
        y -= 8 * mm
        c.line(left, y + 4 * mm, width - right, y + 4 * mm)

    def wrap(text: str, max_chars: int) -> List[str]:
        lines: List[str] = []
        s = text
        while len(s) > max_chars:
            lines.append(s[:max_chars])
            s = s[max_chars:]
        lines.append(s)
        return lines

    header()

    for it in items:
        meaning_lines = wrap(it.meaning, 38)
        row_h = (max(1, len(meaning_lines)) * 5.5 + 4) * mm

        if y - row_h < bottom:
            new_page()
            header()

        c.drawString(left, y, it.word)
        yy = y
        for ln in meaning_lines:
            c.drawString(left + col1_w + 8 * mm, yy, ln)
            yy -= 5.5 * mm

        y -= row_h
        c.line(left, y + 2 * mm, width - right, y + 2 * mm)

    c.save()
    return out.getvalue()


@app.get("/export.pdf")
def export_pdf():
    with LOCK:
        items = list(ITEMS)
    data = _draw_pdf(items)
    return Response(
        data,
        mimetype="application/pdf",
        headers={"Content-Disposition": 'attachment; filename="wordbook.pdf"'},
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
