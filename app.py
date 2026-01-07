# app.py
from __future__ import annotations

import io
import json
import os
from typing import List, Dict

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
    form{display:grid;grid-template-columns:1fr 1fr auto;gap:10px;align-items:end;}
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
          <button type="submit">記録</button>
        </div>
      </form>
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
