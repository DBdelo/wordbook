# app.py
from __future__ import annotations

import csv
import io
import json
import os
import re
from typing import List, Dict, Tuple
from urllib.parse import quote

import requests
from flask import Flask, Response, make_response, render_template_string, request

app = Flask(__name__)

_APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _candidate_preset_base_dirs() -> List[str]:
    """よくある配置向けに、プリセットCSVのルート候補（既存のwordbook フォルダ）を列挙する。"""
    return [
        os.path.join(_APP_DIR, "既存のwordbook"),
        os.path.join(_APP_DIR, "..", "既存のwordbook"),
        os.path.join(_APP_DIR, "..", "wordbook", "既存のwordbook"),
        os.path.join(_APP_DIR, "wordbook", "既存のwordbook"),
    ]


def _abs_norm(path: str) -> str:
    return os.path.normpath(os.path.abspath(path))


# 後方互換・参照用（実際の走査は都度 _pick_wordbook_preset_base() を使う）
_DEFAULT_PRESET_DIR = _abs_norm(os.path.join(_APP_DIR, "..", "既存のwordbook"))
WORDBOOK_PRESET_DIR = os.environ.get("WORDBOOK_PRESET_DIR", _DEFAULT_PRESET_DIR)


def _iter_preset_csv_rel_paths(base: str) -> List[str]:
    if not base or not os.path.isdir(base):
        return []
    base_real = os.path.realpath(base)
    out: List[str] = []
    for root, _dirs, filenames in os.walk(base_real):
        for fn in filenames:
            if not fn.lower().endswith(".csv"):
                continue
            full = os.path.join(root, fn)
            rel = os.path.relpath(full, base_real).replace("\\", "/")
            out.append(rel)
    return sorted(out)


def _pick_wordbook_preset_base() -> str:
    """
    環境変数 WORDBOOK_PRESET_DIR があればそれのみ。
    なければ候補ディレクトリのうち、CSVが1件でもある最初のルートを採用。
    どれにもCSVがなければ、存在する最初の候補、なければ従来どおり app のひとつ上の 既存のwordbook。
    """
    env = (os.environ.get("WORDBOOK_PRESET_DIR") or "").strip()
    if env:
        return _abs_norm(env)
    seen: List[str] = []
    for c in _candidate_preset_base_dirs():
        p = _abs_norm(c)
        if p in seen:
            continue
        seen.append(p)
        if not os.path.isdir(p):
            continue
        if _iter_preset_csv_rel_paths(p):
            return p
    for p in seen:
        if os.path.isdir(p):
            return p
    return _abs_norm(os.path.join(_APP_DIR, "..", "既存のwordbook"))


def _safe_preset_csv_path(base: str, rel: str) -> str | None:
    if not base or not rel:
        return None
    base_real = os.path.realpath(base)
    if not os.path.isdir(base_real):
        return None
    rel_norm = rel.replace("\\", "/").strip("/")
    if not rel_norm or ".." in rel_norm.split("/"):
        return None
    parts = [p for p in rel_norm.split("/") if p and p != "."]
    candidate = os.path.realpath(os.path.join(base_real, *parts))
    if not candidate.startswith(base_real + os.sep):
        return None
    if not candidate.lower().endswith(".csv"):
        return None
    if not os.path.isfile(candidate):
        return None
    return candidate


def _is_preset_csv_header(a: str, b: str) -> bool:
    a_st = a.strip().lstrip("\ufeff")
    b_st = b.strip()
    if a_st == "英単語":
        return True
    al = a_st.lower()
    bl = b_st.lower()
    if al == "word" and "meaning" in bl:
        return True
    if al in ("english", "en") and ("japanese" in bl or "日本語" in b_st):
        return True
    return False


def _parse_preset_csv_text(text: str) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    f = io.StringIO(text)
    reader = csv.reader(f)
    for i, row in enumerate(reader):
        if len(row) < 2:
            continue
        a, b = row[0].strip().lstrip("\ufeff"), row[1].strip()
        if not a or not b:
            continue
        if i == 0 and _is_preset_csv_header(a, b):
            continue
        if len(a) > 80 or len(b) > 200:
            continue
        rows.append({"word": a, "meaning": b})
    return rows


HTML = r"""
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <meta http-equiv="Cache-Control" content="no-store, max-age=0" />
  <title>あなたの単語帳</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{
      font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,"Noto Sans JP",sans-serif;
      margin:0;min-height:100vh;
      transition:background .4s ease;
    }
    body.mode-quiz-en{background:linear-gradient(135deg,#dbeafe 0%,#bfdbfe 100%);}
    body.mode-quiz-ja{background:linear-gradient(135deg,#d1fae5 0%,#a7f3d0 100%);}
    body.mode-record{background:#f6f7fb;}

    /* ── Navbar ── */
    .navbar{
      display:flex;justify-content:space-between;align-items:center;
      padding:12px 20px;
      background:rgba(255,255,255,.88);
      backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
      border-bottom:1px solid rgba(0,0,0,.07);
      position:sticky;top:0;z-index:100;
    }
    .navbar h1{font-size:18px;font-weight:800;color:#1e293b;letter-spacing:-.3px;}
    .quiz-source-strip{
      width:100%;background:rgba(255,255,255,.95);
      border-bottom:1px solid rgba(0,0,0,.08);
      box-shadow:0 4px 14px rgba(0,0,0,.04);
      padding:12px 20px 14px;z-index:90;
    }
    .quiz-source-strip-inner{
      max-width:560px;margin:0 auto;
      display:flex;flex-direction:column;gap:8px;align-items:stretch;
    }
    .quiz-source-strip-inner .quiz-source-row{
      display:flex;flex-wrap:wrap;align-items:center;gap:10px 14px;
    }
    .quiz-source-strip-inner label{
      font-size:13px;font-weight:800;color:#1e293b;margin:0;flex-shrink:0;
    }
    .quiz-source-strip-inner select#quizBookSource{
      flex:1;min-width:min(100%,260px);max-width:100%;
      padding:11px 12px;border:2px solid #cbd5e1;border-radius:10px;
      font-size:15px;font-weight:600;background:#fff;color:#0f172a;
    }
    .quiz-book-hint{margin:0;font-size:12px;line-height:1.5;color:#475569;text-align:left;}
    .mode-tabs{display:flex;background:#e2e8f0;border-radius:10px;padding:3px;gap:2px;}
    .mode-tab{
      padding:8px 18px;border-radius:8px;font-size:13px;font-weight:600;
      cursor:pointer;border:none;background:transparent;color:#64748b;transition:all .2s;
    }
    .mode-tab.active{background:#fff;color:#1e293b;box-shadow:0 1px 4px rgba(0,0,0,.1);}

    /* ── Quiz Mode ── */
    #quizView{display:none;}
    #recordView{display:none;}

    .quiz-container{
      display:flex;flex-direction:column;align-items:center;
      min-height:calc(100vh - 56px);padding:24px 20px 40px;
    }
    .quiz-header{
      display:flex;gap:14px;align-items:center;flex-wrap:wrap;
      justify-content:center;margin-bottom:20px;width:100%;max-width:520px;
    }
    .quiz-lang-toggle{
      display:flex;align-items:center;gap:6px;
      background:rgba(255,255,255,.65);padding:6px;border-radius:12px;
    }
    .lang-btn{
      padding:8px 14px;border-radius:8px;font-size:13px;font-weight:600;
      cursor:pointer;border:2px solid transparent;
      background:transparent;color:#64748b;transition:all .2s;
    }
    .lang-btn.sel-en{background:#3b82f6;color:#fff;border-color:#3b82f6;}
    .lang-btn.sel-ja{background:#10b981;color:#fff;border-color:#10b981;}

    .quiz-stats{
      display:flex;gap:10px;font-size:14px;font-weight:600;
    }
    .stat-badge{
      display:flex;align-items:center;gap:6px;
      padding:6px 14px;border-radius:20px;
      background:rgba(255,255,255,.7);
    }
    .stat-badge.good{color:#059669;}
    .stat-badge.bad{color:#dc2626;}
    .stat-badge .cnt{font-size:20px;font-weight:800;}

    .quiz-progress-wrap{width:100%;max-width:520px;margin-bottom:20px;text-align:center;}
    .quiz-progress-text{font-size:13px;color:rgba(0,0,0,.45);font-weight:700;margin-bottom:6px;}
    .quiz-progress{height:6px;background:rgba(255,255,255,.5);border-radius:3px;overflow:hidden;}
    .quiz-progress-bar{height:100%;border-radius:3px;transition:width .35s ease;}
    body.mode-quiz-en .quiz-progress-bar{background:#3b82f6;}
    body.mode-quiz-ja .quiz-progress-bar{background:#10b981;}

    .quiz-card{
      background:#fff;border-radius:22px;
      box-shadow:0 12px 44px rgba(0,0,0,.08);
      padding:48px 36px;text-align:center;
      max-width:520px;width:100%;margin-bottom:24px;
      min-height:180px;display:flex;flex-direction:column;
      align-items:center;justify-content:center;
      animation:cardIn .3s ease both;
    }
    @keyframes cardIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}

    .quiz-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;opacity:.55;}
    .quiz-word{font-size:34px;font-weight:800;color:#1e293b;word-break:break-word;line-height:1.3;}
    .quiz-answer-divider{width:60px;height:2px;background:#e2e8f0;margin:18px 0;border-radius:1px;}
    .quiz-answer-label{font-size:12px;font-weight:700;letter-spacing:.5px;margin-bottom:6px;opacity:.5;}
    .quiz-answer{
      font-size:24px;font-weight:700;color:#475569;word-break:break-word;
      animation:fadeUp .3s ease both;line-height:1.4;
    }
    @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}

    .quiz-speak{
      margin-top:14px;background:#f1f5f9;border:none;border-radius:10px;
      padding:8px 18px;font-size:14px;cursor:pointer;color:#475569;transition:background .2s;
    }
    .quiz-speak:hover{background:#e2e8f0;}

    .quiz-buttons{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;}
    .qbtn{
      padding:15px 36px;border-radius:14px;font-size:16px;font-weight:700;
      border:none;cursor:pointer;transition:all .15s;min-width:140px;
    }
    .qbtn:active{transform:scale(.96);}
    .qbtn-good{background:#10b981;color:#fff;}
    .qbtn-good:hover{background:#059669;}
    .qbtn-bad{background:#ef4444;color:#fff;}
    .qbtn-bad:hover{background:#dc2626;}
    .qbtn-next{background:#3b82f6;color:#fff;}
    body.mode-quiz-ja .qbtn-next{background:#10b981;}
    .qbtn-next:hover{filter:brightness(.92);}
    .qbtn-restart{background:#6366f1;color:#fff;}
    .qbtn-restart:hover{background:#4f46e5;}
    .qbtn-retry{background:#f59e0b;color:#fff;}
    .qbtn-retry:hover{background:#d97706;}
    .qbtn-goto{background:#64748b;color:#fff;}
    .qbtn-goto:hover{background:#475569;}

    .quiz-empty{text-align:center;color:#64748b;font-size:16px;line-height:2;}
    .quiz-empty .big-icon{font-size:52px;margin-bottom:8px;}

    .quiz-complete{text-align:center;}
    .quiz-complete .big-icon{font-size:52px;margin-bottom:10px;}
    .quiz-complete h2{font-size:24px;margin-bottom:8px;color:#1e293b;}
    .result-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:18px 0;}
    .result-cell{padding:18px 12px;border-radius:14px;text-align:center;}
    .result-cell.good{background:#d1fae5;color:#065f46;}
    .result-cell.bad{background:#fee2e2;color:#991b1b;}
    .result-cell .big{font-size:36px;font-weight:800;}
    .result-cell .lbl{font-size:12px;margin-top:4px;font-weight:600;}
    .pct{font-size:22px;font-weight:800;color:#1e293b;margin:8px 0 18px;}
    .result-actions{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;}

    .shortcut-hint{
      margin-top:18px;font-size:12px;color:rgba(0,0,0,.3);
      text-align:center;user-select:none;
    }
    .shortcut-hint kbd{
      display:inline-block;padding:2px 7px;border-radius:4px;
      background:rgba(255,255,255,.6);font-size:11px;font-family:inherit;
      border:1px solid rgba(0,0,0,.1);
    }

    /* ── Record Mode (existing, refined) ── */
    .wrap{max-width:860px;margin:0 auto;padding:0 20px;}
    .card{background:#fff;border-radius:14px;box-shadow:0 6px 18px rgba(0,0,0,.06);padding:16px;margin-bottom:14px;}
    form#addForm{display:grid;grid-template-columns:1fr 1fr auto;gap:10px;align-items:end;}
    label{display:block;font-size:12px;color:#444;margin-bottom:6px;}
    input,select{width:100%;padding:10px 12px;border:1px solid #d7dbe7;border-radius:10px;font-size:14px;background:#fff;}
    .rbtn{
      appearance:none;border:0;border-radius:10px;padding:10px 14px;font-size:14px;
      cursor:pointer;display:inline-flex;align-items:center;gap:8px;
      background:#111827;color:#fff;text-decoration:none;
    }
    .rbtn-primary{background:#2563eb;color:#fff;}
    .row{display:flex;gap:10px;flex-wrap:wrap}
    table{width:100%;border-collapse:collapse;}
    th,td{padding:10px 8px;border-bottom:1px solid #edf0f6;font-size:14px;vertical-align:top}
    th{color:#374151;text-align:left;font-weight:600;}
    .del{
      appearance:none;border:0;border-radius:10px;padding:8px 12px;
      font-size:13px;cursor:pointer;background:#ef4444;color:#fff;
    }
    .muted{color:#6b7280;font-size:12px;margin-top:8px}
    .topbar{display:flex;gap:10px;flex-wrap:wrap;justify-content:space-between;align-items:center;margin-bottom:12px;padding-top:16px;}
    .hint{color:#6b7280;font-size:12px;margin-top:10px;min-height:1em}
    .speak{
      background:#111827;padding:8px 10px;border-radius:10px;
      color:#fff;font-size:13px;line-height:1;border:none;cursor:pointer;
    }
    .speak:disabled{opacity:.5;cursor:not-allowed;}
    .wordcell{display:flex;align-items:center;gap:10px}
    .wordtext{font-weight:600}
    .voicebox{margin-top:12px;display:grid;grid-template-columns:1fr;gap:8px;}
    .voicehint{color:#6b7280;font-size:12px}
    /* ── Autocomplete ── */
    .ac-wrap{position:relative;}
    .ac-list{
      position:absolute;left:0;right:0;top:100%;
      background:#fff;border:1px solid #d7dbe7;border-top:none;
      border-radius:0 0 10px 10px;
      max-height:220px;overflow-y:auto;
      z-index:50;box-shadow:0 8px 24px rgba(0,0,0,.1);
      display:none;
    }
    .ac-list.open{display:block;}
    .ac-item{
      padding:10px 14px;font-size:14px;color:#1e293b;cursor:pointer;
      border-bottom:1px solid #f1f5f9;
      transition:background .1s;
    }
    .ac-item:last-child{border-bottom:none;}
    .ac-item:hover,.ac-item.active{background:#eff6ff;color:#2563eb;}
    .ac-item mark{background:none;color:#2563eb;font-weight:700;}

    /* ── Modal ── */
    .modal-overlay{
      position:fixed;inset:0;background:rgba(0,0,0,.45);
      display:flex;align-items:center;justify-content:center;
      z-index:200;padding:20px;animation:fadeIn .2s ease;
    }
    @keyframes fadeIn{from{opacity:0}to{opacity:1}}
    .modal-box{
      background:#fff;border-radius:18px;padding:28px 24px;
      max-width:400px;width:100%;box-shadow:0 16px 48px rgba(0,0,0,.15);
      animation:cardIn .25s ease both;
    }
    .modal-box h3{font-size:18px;font-weight:700;color:#1e293b;margin-bottom:12px;}
    .modal-box p{font-size:14px;color:#475569;line-height:1.7;margin-bottom:16px;}
    .modal-box .highlight{font-weight:700;color:#1e293b;}
    .modal-actions{display:flex;flex-direction:column;gap:10px;}
    .modal-btn{
      width:100%;padding:14px 16px;border-radius:12px;font-size:15px;font-weight:700;
      border:none;cursor:pointer;text-align:left;transition:filter .15s;
      display:flex;align-items:center;gap:12px;
    }
    .modal-btn:active{filter:brightness(.9);}
    .modal-btn .modal-icon{font-size:22px;flex-shrink:0;}
    .modal-btn .modal-text{display:flex;flex-direction:column;}
    .modal-btn .modal-title{font-size:15px;font-weight:700;}
    .modal-btn .modal-desc{font-size:12px;font-weight:400;opacity:.75;margin-top:2px;}
    .modal-btn-add{background:#059669;color:#fff;}
    .modal-btn-replace{background:#dc2626;color:#fff;}
    .modal-btn-cancel{background:#e2e8f0;color:#475569;justify-content:center;}

    /* ── Responsive ── */
    @media(max-width:600px){
      .navbar{padding:10px 14px;}
      .navbar h1{font-size:15px;}
      .mode-tab{padding:7px 12px;font-size:12px;}
      form#addForm{grid-template-columns:1fr;gap:8px;}
      .quiz-word{font-size:26px;}
      .quiz-answer{font-size:20px;}
      .quiz-card{padding:36px 22px;border-radius:18px;}
      .qbtn{min-width:120px;padding:13px 24px;font-size:15px;}
      .lang-btn{padding:7px 10px;font-size:12px;}
      .stat-badge{padding:5px 10px;font-size:13px;}
      .stat-badge .cnt{font-size:17px;}
      .shortcut-hint{display:none;}
    }
  </style>
</head>
<body class="mode-quiz-en">

  <!-- ── Navigation ── -->
  <nav class="navbar">
    <h1>あなたの単語帳</h1>
    <div class="mode-tabs">
      <button class="mode-tab active" id="tabQuiz" type="button">単語テスト</button>
      <button class="mode-tab" id="tabRecord" type="button">記録モード</button>
    </div>
  </nav>

  <!-- ── Quiz View ── -->
  <div id="quizView">
    <div class="quiz-source-strip" id="quizSourceStrip">
      <div class="quiz-source-strip-inner">
        <div class="quiz-source-row">
          <label for="quizBookSource">出題する単語帳</label>
          <select id="quizBookSource" title="マイ単語帳か、サーバー上のCSVを選びます"></select>
        </div>
        <div class="quiz-book-hint" id="quizBookHint"></div>
      </div>
    </div>
    <div class="quiz-container">
      <div class="quiz-header">
        <div class="quiz-lang-toggle">
          <button class="lang-btn sel-en" id="toggleEn" type="button">英語→日本語</button>
          <button class="lang-btn" id="toggleJa" type="button">日本語→英語</button>
        </div>
        <div class="quiz-stats">
          <div class="stat-badge good"><span>覚えた</span><span class="cnt" id="cntGood">0</span></div>
          <div class="stat-badge bad"><span>忘れた</span><span class="cnt" id="cntBad">0</span></div>
        </div>
      </div>

      <div class="quiz-progress-wrap">
        <div class="quiz-progress-text" id="progText"></div>
        <div class="quiz-progress"><div class="quiz-progress-bar" id="progBar" style="width:0%"></div></div>
      </div>

      <div id="quizArea"></div>

      <div class="shortcut-hint" id="shortcutHint">
        <kbd>→</kbd> 覚えた　<kbd>←</kbd> 忘れた　<kbd>Space</kbd> 次へ
      </div>
    </div>
  </div>

  <!-- ── Record View ── -->
  <div id="recordView">
    <div class="wrap">
      <div class="topbar">
        <h2 style="font-size:20px;margin:0;">単語の記録</h2>
        <div class="row">
          <button class="rbtn" id="btnImport" type="button" style="background:#059669">CSV取込</button>
          <button class="rbtn" id="btnCsv" type="button">CSV出力</button>
          <button class="rbtn" id="btnPdf" type="button">PDF</button>
          <button class="rbtn" id="btnClear" type="button" style="background:#6b7280">全消去</button>
          <input type="file" id="csvFile" accept=".csv,.txt" style="display:none" />
        </div>
      </div>
      <div class="card">
        <form id="addForm">
          <div class="ac-wrap"><label>word</label><input id="word" autocomplete="off" required maxlength="80" /><div class="ac-list" id="acList"></div></div>
          <div><label>単語</label><input id="meaning" autocomplete="off" required maxlength="200" /></div>
          <div><button type="submit" class="rbtn-primary" style="border:0;border-radius:10px;padding:10px 14px;font-size:14px;cursor:pointer;color:#fff;background:#2563eb;">記録</button></div>
        </form>
        <div class="voicebox">
          <div><label>読み上げ音声</label><select id="voiceSelect"></select></div>
          <div class="voicehint" id="voiceHint"></div>
        </div>
        <div class="hint" id="hint"></div>
        <div class="muted" id="storageNote">データはこのブラウザ内に保存されます。</div>
      </div>
      <div class="card">
        <table>
          <thead><tr><th style="width:34%">word</th><th>単語</th><th style="width:90px">操作</th></tr></thead>
          <tbody id="tbody"><tr><td colspan="3" style="color:#6b7280;padding:16px 8px">まだありません。</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

<div id="modalRoot"></div>
<script>
(() => {
  const KEY  = "wordbook_items_v1";
  const VKEY = "wordbook_voice_v1_en";
  const LKEY = "wordbook_quiz_lang_v1";
  const QBOOK_KEY = "wordbook_quiz_book_v1";
  const QBOOK_KEY_LEGACY = "wordbook_active_book_v1";

  const $ = id => document.getElementById(id);
  const quizBookSelect=$("quizBookSource"),quizBookHint=$("quizBookHint"),storageNote=$("storageNote");
  let presetItems=[];
  let presetCsvCount=null;
  function isPersonalQuizBook(){return !quizBookSelect||quizBookSelect.value==="personal";}
  function isPresetQuizBook(){return !!(quizBookSelect&&quizBookSelect.value.startsWith("preset:"));}

  /* ── Shared helpers ── */
  function loadItems(){
    try{const r=localStorage.getItem(KEY);if(!r)return[];const v=JSON.parse(r);
    if(!Array.isArray(v))return[];return v.filter(x=>x&&typeof x.word==="string"&&typeof x.meaning==="string");}catch{return[];}
  }
  function saveItems(a){localStorage.setItem(KEY,JSON.stringify(a));}
  function esc(s){return s.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#39;");}

  /* ── Speech ── */
  function canSpeak(){return !!(window.speechSynthesis&&window.SpeechSynthesisUtterance);}
  function voices(){try{return(window.speechSynthesis&&window.speechSynthesis.getVoices)?window.speechSynthesis.getVoices():[];}catch{return[];}}
  function isEn(v){return(v&&v.lang||"").toLowerCase().startsWith("en");}
  function vKey(v){return[v.name||"",v.lang||"",v.voiceURI||""].join("|");}
  function loadVC(){try{return localStorage.getItem(VKEY)||"";}catch{return"";}}
  function saveVC(k){try{localStorage.setItem(VKEY,k||"");}catch{}}

  const voiceSelect=$("voiceSelect"),voiceHint=$("voiceHint");

  function rebuildVoiceSelect(){
    voiceSelect.innerHTML="";
    if(!canSpeak()){voiceSelect.disabled=true;const o=document.createElement("option");o.textContent="非対応";voiceSelect.appendChild(o);return;}
    const vs=voices().filter(isEn);
    if(!vs.length){voiceSelect.disabled=true;const o=document.createElement("option");o.textContent="英語音声なし";voiceSelect.appendChild(o);return;}
    voiceSelect.disabled=false;const saved=loadVC();let sel=saved||vKey(vs[0]);
    vs.forEach(v=>{const o=document.createElement("option");o.value=vKey(v);o.textContent=v.name+" ("+v.lang+")";if(o.value===sel)o.selected=true;voiceSelect.appendChild(o);});
    saveVC(voiceSelect.value||"");
    const sv=vs.find(v=>vKey(v)===voiceSelect.value)||null;
    voiceHint.textContent=sv?("現在："+sv.name+"（"+sv.lang+"）"):"";
  }
  function getSelVoice(){const vs=voices().filter(isEn);const k=voiceSelect.value||loadVC();return vs.find(x=>vKey(x)===k)||(vs[0]||null);}
  function speakWord(w){
    if(!canSpeak())return;try{window.speechSynthesis.cancel();}catch{}
    const u=new SpeechSynthesisUtterance(w);const v=getSelVoice();
    if(v)u.voice=v;else u.lang="en-US";u.rate=1;u.pitch=1;window.speechSynthesis.speak(u);
  }
  voiceSelect.addEventListener("change",()=>{saveVC(voiceSelect.value||"");const v=getSelVoice();voiceHint.textContent=v?("現在："+v.name+"（"+v.lang+"）"):"";});
  if(canSpeak())window.speechSynthesis.onvoiceschanged=()=>{rebuildVoiceSelect();renderTable();};
  rebuildVoiceSelect();

  const quizArea=$("quizArea");

  /* ═══════════════════════════════════════
     MODE MANAGEMENT
     ═══════════════════════════════════════ */
  const tabQuiz=$("tabQuiz"),tabRecord=$("tabRecord");
  const quizView=$("quizView"),recordView=$("recordView");
  let curMode="quiz";

  async function loadPresetForQuiz(){
    if(isPersonalQuizBook()){presetItems=[];return true;}
    if(!isPresetQuizBook()){presetItems=[];return false;}
    const rel=quizBookSelect.value.slice(7);
    if(!rel){presetItems=[];return false;}
    try{
      const r=await fetch("/api/preset-csv/file?path="+encodeURIComponent(rel));
      const d=await r.json();
      if(r.ok&&d&&d.ok&&Array.isArray(d.items)){
        presetItems=d.items.filter(it=>it&&typeof it.word==="string"&&typeof it.meaning==="string");
        return true;
      }
      presetItems=[];
      return false;
    }catch{presetItems=[];return false;}
  }

  async function switchMode(m){
    curMode=m;
    tabQuiz.classList.toggle("active",m==="quiz");
    tabRecord.classList.toggle("active",m==="record");
    if(m==="quiz"){
      quizView.style.display="block";recordView.style.display="none";
      updateBg();
      if(isPresetQuizBook()){
        quizArea.innerHTML='<div class="quiz-card"><div class="quiz-empty" style="color:#64748b">単語帳を読み込み中…</div></div>';
        await loadPresetForQuiz();
      }
      resetQuiz();
    }else{
      quizView.style.display="none";recordView.style.display="block";
      document.body.className="mode-record";renderTable();
    }
  }
  tabQuiz.addEventListener("click",()=>{void switchMode("quiz");});
  tabRecord.addEventListener("click",()=>{void switchMode("record");});

  /* ═══════════════════════════════════════
     QUIZ MODE
     ═══════════════════════════════════════ */
  const cntGood=$("cntGood"),cntBad=$("cntBad");
  const progBar=$("progBar"),progText=$("progText");
  const toggleEn=$("toggleEn"),toggleJa=$("toggleJa");

  let qLang=localStorage.getItem(LKEY)||"en";
  let queue=[],qIdx=0,good=0,bad=0,forgot=[],answered=false;

  function updateBg(){document.body.className=qLang==="en"?"mode-quiz-en":"mode-quiz-ja";}
  function updateLangBtns(){
    toggleEn.className="lang-btn"+(qLang==="en"?" sel-en":"");
    toggleJa.className="lang-btn"+(qLang==="ja"?" sel-ja":"");
  }

  function setLang(l){qLang=l;localStorage.setItem(LKEY,l);updateLangBtns();updateBg();resetQuiz();}
  toggleEn.addEventListener("click",()=>setLang("en"));
  toggleJa.addEventListener("click",()=>setLang("ja"));
  updateLangBtns();

  function shuffle(a){const b=[...a];for(let i=b.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[b[i],b[j]]=[b[j],b[i]];}return b;}

  function activeQuizItems(){
    if(isPresetQuizBook())return presetItems;
    return loadItems();
  }

  function resetQuiz(){
    queue=shuffle(activeQuizItems().slice());qIdx=0;good=0;bad=0;forgot=[];answered=false;
    updateStats();renderQ();
  }
  function updateStats(){
    cntGood.textContent=good;cntBad.textContent=bad;
    const tot=queue.length,done=qIdx;
    progBar.style.width=tot>0?((done/tot)*100)+"%":"0%";
    progText.textContent=tot>0?(done+" / "+tot):"";
  }

  function renderQ(){
    const items=activeQuizItems();
    if(items.length===0){
      const isP=!isPresetQuizBook();
      const msg=isP
        ?'<p>あなたの単語帳に単語がありません。</p>'
        +'<p style="margin-top:8px;font-size:14px">画面上部の<strong>出題する単語帳</strong>で、サーバー上のCSVを選ぶとその内容でテストできます。</p>'
        +'<p style="margin-top:8px;font-size:14px">自分の単語で試す場合は「記録モード」で追加してください。</p>'
        +'<button class="qbtn qbtn-goto" type="button" id="qGoRecord" style="margin-top:18px">記録モードへ</button>'
        :'<p>このCSVに単語がありません。</p><p style="margin-top:8px">CSVの形式（英単語と意味の列）を確認するか、別のファイルを選んでください。</p>';
      quizArea.innerHTML='<div class="quiz-card"><div class="quiz-empty">'
        +'<div class="big-icon">📚</div>'+msg+'</div></div>';
      progText.textContent="";progBar.style.width="0%";
      return;
    }
    if(queue.length===0){resetQuiz();return;}

    if(qIdx>=queue.length){
      const tot=queue.length;
      const pct=tot>0?Math.round((good/tot)*100):0;
      const icon=pct>=80?"🎉":pct>=50?"👍":"💪";
      const hasForgot=forgot.length>0;
      quizArea.innerHTML='<div class="quiz-card"><div class="quiz-complete">'
        +'<div class="big-icon">'+icon+'</div>'
        +'<h2>テスト完了！</h2>'
        +'<div class="result-grid">'
        +'<div class="result-cell good"><div class="big">'+good+'</div><div class="lbl">覚えた</div></div>'
        +'<div class="result-cell bad"><div class="big">'+bad+'</div><div class="lbl">忘れた</div></div>'
        +'</div>'
        +'<div class="pct">正答率 '+pct+'%</div>'
        +'<div class="result-actions">'
        +'<button class="qbtn qbtn-restart" type="button" id="qRestart">もう一度</button>'
        +(hasForgot?'<button class="qbtn qbtn-retry" type="button" id="qRetry">忘れた単語だけ ('+forgot.length+')</button>':'')
        +'</div>'
        +'</div></div>';
      return;
    }

    const it=queue[qIdx];
    const question=qLang==="en"?it.word:it.meaning;
    const answer=qLang==="en"?it.meaning:it.word;
    const qLabel=qLang==="en"?"English":"日本語";
    const aLabel=qLang==="en"?"日本語":"English";
    const showSpeakQ=qLang==="en"&&canSpeak()&&!voiceSelect.disabled;
    const showSpeakA=qLang==="ja"&&canSpeak()&&!voiceSelect.disabled;

    if(!answered){
      quizArea.innerHTML='<div class="quiz-card">'
        +'<div class="quiz-label">'+qLabel+'</div>'
        +'<div class="quiz-word">'+esc(question)+'</div>'
        +(showSpeakQ?'<button class="quiz-speak" type="button" id="qSpeak">🔊 読み上げ</button>':'')
        +'</div>'
        +'<div class="quiz-buttons">'
        +'<button class="qbtn qbtn-bad" type="button" id="qBad">忘れた</button>'
        +'<button class="qbtn qbtn-good" type="button" id="qGood">覚えた</button>'
        +'</div>';
    }else{
      quizArea.innerHTML='<div class="quiz-card">'
        +'<div class="quiz-label">'+qLabel+'</div>'
        +'<div class="quiz-word">'+esc(question)+'</div>'
        +'<div class="quiz-answer-divider"></div>'
        +'<div class="quiz-answer-label">'+aLabel+'</div>'
        +'<div class="quiz-answer">'+esc(answer)+'</div>'
        +((showSpeakQ||showSpeakA)?'<button class="quiz-speak" type="button" id="qSpeak">🔊 読み上げ</button>':'')
        +'</div>'
        +'<div class="quiz-buttons">'
        +'<button class="qbtn qbtn-next" type="button" id="qNext">次へ →</button>'
        +'</div>';
    }
  }

  quizArea.addEventListener("click",e=>{
    const t=e.target;
    if(t.id==="qGood"){good++;answered=true;updateStats();renderQ();}
    if(t.id==="qBad"){bad++;forgot.push(queue[qIdx]);answered=true;updateStats();renderQ();}
    if(t.id==="qNext"){qIdx++;answered=false;updateStats();renderQ();}
    if(t.id==="qRestart"){resetQuiz();}
    if(t.id==="qRetry"){queue=shuffle(forgot);qIdx=0;good=0;bad=0;forgot=[];answered=false;updateStats();renderQ();}
    if(t.id==="qGoRecord"){void switchMode("record");}
    if(t.id==="qSpeak"){
      const it=queue[qIdx];
      if(it)speakWord(it.word);
    }
  });

  document.addEventListener("keydown",e=>{
    if(curMode!=="quiz")return;
    if(document.activeElement&&(document.activeElement.tagName==="INPUT"||document.activeElement.tagName==="SELECT"||document.activeElement.tagName==="TEXTAREA"))return;
    const items=activeQuizItems();if(!items.length)return;
    if(qIdx>=queue.length)return;

    if(!answered){
      if(e.key==="ArrowRight"){e.preventDefault();good++;answered=true;updateStats();renderQ();}
      if(e.key==="ArrowLeft"){e.preventDefault();bad++;forgot.push(queue[qIdx]);answered=true;updateStats();renderQ();}
    }else{
      if(e.key===" "||e.key==="Enter"||e.key==="ArrowRight"||e.key==="ArrowLeft"){e.preventDefault();qIdx++;answered=false;updateStats();renderQ();}
    }
  });

  /* ═══════════════════════════════════════
     RECORD MODE (existing logic preserved)
     ═══════════════════════════════════════ */
  const tbody=$("tbody"),form=$("addForm"),wordEl=$("word"),meaningEl=$("meaning");
  const btnCsv=$("btnCsv"),btnPdf=$("btnPdf"),btnClear=$("btnClear"),btnImport=$("btnImport"),hint=$("hint");
  const csvFile=$("csvFile");
  const submitBtn=form.querySelector('button[type="submit"]');
  let lastQ="",debT=null,inflight=null;

  function setHint(s){hint.textContent=s||"";}

  function updateQuizBookHint(){
    if(!quizBookHint)return;
    const noPreset=presetCsvCount===0;
    if(isPersonalQuizBook()){
      quizBookHint.textContent="「あなたの単語帳」に登録した単語から出題します。"
        +(noPreset?" サーバーの「既存のwordbook」内に.csvが見つかりません。サブフォルダ（例: EX準一級/）内も列挙します。フォルダ配置または環境変数 WORDBOOK_PRESET_DIR を確認してください。":"");
    }else{
      quizBookHint.textContent="選択したCSVから出題します。記録モードの一覧はあなたの単語帳のみです。";
    }
  }

  function saveQuizBookChoice(){try{localStorage.setItem(QBOOK_KEY,quizBookSelect.value||"personal");}catch{}}

  function refreshRecordEditor(){
    wordEl.disabled=false;meaningEl.disabled=false;
    if(submitBtn)submitBtn.disabled=false;
    btnImport.disabled=false;btnClear.disabled=false;
    acClose();
    if(storageNote)storageNote.style.display="";
  }

  function renderTable(){
    const items=loadItems();
    if(!items.length){tbody.innerHTML='<tr><td colspan="3" style="color:#6b7280;padding:16px 8px">まだありません。</td></tr>';return;}
    const ok=canSpeak()&&!voiceSelect.disabled;
    tbody.innerHTML=items.map((it,i)=>'<tr><td><div class="wordcell">'
      +'<button class="speak" type="button" data-speak="'+i+'"'+(ok?"":" disabled")+' title="'+(ok?"読み上げ":"英語音声なし")+'">🔊</button>'
      +'<span class="wordtext">'+esc(it.word)+'</span></div></td>'
      +'<td>'+esc(it.meaning)+'</td>'
      +'<td><button class="del" type="button" data-del="'+i+'">削除</button></td></tr>').join("");
  }

  function groupFilesForOptgroups(files){
    const m=new Map();
    files.forEach(f=>{
      const rel=f.rel;
      const slash=rel.lastIndexOf("/");
      const dir=slash>=0?rel.slice(0,slash):"";
      const name=slash>=0?rel.slice(slash+1):rel;
      const key=dir||"\0ROOT";
      if(!m.has(key))m.set(key,[]);
      m.get(key).push({rel,name});
    });
    const keys=[...m.keys()].sort((a,b)=>{
      if(a==="\0ROOT")return -1;if(b==="\0ROOT")return 1;
      return a.localeCompare(b,"ja");
    });
    return keys.map(k=>[k==="\0ROOT"?"":k,m.get(k)]);
  }

  async function onQuizBookChange(){
    saveQuizBookChoice();
    updateQuizBookHint();
    if(isPersonalQuizBook()){
      presetItems=[];
      resetQuiz();
      return;
    }
    if(curMode==="quiz"){
      quizArea.innerHTML='<div class="quiz-card"><div class="quiz-empty" style="color:#64748b">読み込み中…</div></div>';
    }
    await loadPresetForQuiz();
    resetQuiz();
  }

  function initQuizBookSelect(){
    quizBookSelect.innerHTML="";
    const o=document.createElement("option");
    o.value="personal";o.textContent="あなたの単語帳";
    quizBookSelect.appendChild(o);
    quizBookSelect.addEventListener("change",onQuizBookChange);
    updateQuizBookHint();
    refreshRecordEditor();
    renderTable();
    resetQuiz();
    fetch("/api/preset-csv/list").then(r=>{
      if(!r.ok)throw new Error("preset list "+r.status);
      return r.json();
    }).then(d=>{
      if(d&&d.ok&&Array.isArray(d.files)){
        presetCsvCount=d.files.length;
        if(d.files.length){
          groupFilesForOptgroups(d.files).forEach(([dir,list])=>{
            const og=document.createElement("optgroup");
            og.label=dir?dir:"（フォルダ直下）";
            list.sort((a,b)=>a.name.localeCompare(b.name,"ja",{numeric:true})).forEach(f=>{
              const opt=document.createElement("option");
              opt.value="preset:"+f.rel;
              opt.textContent=f.name.replace(/\.csv$/i,"");
              og.appendChild(opt);
            });
            quizBookSelect.appendChild(og);
          });
        }
      }else{presetCsvCount=null;}
      let saved="";
      try{
        saved=localStorage.getItem(QBOOK_KEY)||"";
        if(!saved){saved=localStorage.getItem(QBOOK_KEY_LEGACY)||"personal";}
      }catch{}
      if([...quizBookSelect.options].some(x=>x.value===saved))quizBookSelect.value=saved;
      else quizBookSelect.value="personal";
      onQuizBookChange();
    }).catch(()=>{
      presetCsvCount=null;
      try{localStorage.setItem(QBOOK_KEY,"personal");}catch{}
      quizBookSelect.value="personal";
      presetItems=[];
      updateQuizBookHint();
      refreshRecordEditor();
      renderTable();
      resetQuiz();
    });
  }

  async function doLookup(w){
    w=w.trim();if(!w){setHint("");return;}if(w===lastQ)return;lastQ=w;
    if(inflight)inflight.abort();inflight=new AbortController();setHint("検索中…");
    try{
      const r=await fetch("/lookup",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({word:w}),signal:inflight.signal});
      if(!r.ok){setHint("失敗しました");return;}const d=await r.json();
      if(d&&d.meaning){meaningEl.value=d.meaning;setHint("");}else setHint("見つかりませんでした");
    }catch(e){if(e&&e.name==="AbortError")return;setHint("失敗しました");}
  }

  /* ── Autocomplete (Datamuse API) ── */
  const acList=$("acList");
  let acTimer=null,acFlight=null,acIdx=-1,acItems=[];

  function acClose(){acList.classList.remove("open");acList.innerHTML="";acIdx=-1;acItems=[];}

  function acHighlight(word,query){
    const q=query.toLowerCase(),w=word;
    const i=w.toLowerCase().indexOf(q);
    if(i<0)return esc(w);
    return esc(w.slice(0,i))+"<mark>"+esc(w.slice(i,i+q.length))+"</mark>"+esc(w.slice(i+q.length));
  }

  function acRender(words,query){
    if(!words.length){acClose();return;}
    acItems=words;acIdx=-1;
    acList.innerHTML=words.map((w,i)=>'<div class="ac-item" data-ac="'+i+'">'+acHighlight(w,query)+'</div>').join("");
    acList.classList.add("open");
  }

  function acSelect(word){
    wordEl.value=word;acClose();lastQ="";doLookup(word);
  }

  async function acFetch(q){
    q=q.trim();if(q.length<2){acClose();return;}
    if(acFlight)acFlight.abort();acFlight=new AbortController();
    try{
      const r=await fetch("https://api.datamuse.com/sug?s="+encodeURIComponent(q)+"&max=8",{signal:acFlight.signal});
      if(!r.ok){acClose();return;}
      const data=await r.json();
      const words=data.map(d=>d.word).filter(w=>w&&/^[a-zA-Z\s'-]+$/.test(w)).slice(0,8);
      acRender(words,q);
    }catch(e){if(e&&e.name!=="AbortError")acClose();}
  }

  acList.addEventListener("click",e=>{
    const el=e.target.closest(".ac-item");if(!el)return;
    const i=Number(el.getAttribute("data-ac"));
    if(i>=0&&i<acItems.length)acSelect(acItems[i]);
  });

  wordEl.addEventListener("input",()=>{
    if(debT)clearTimeout(debT);
    debT=setTimeout(()=>doLookup(wordEl.value),450);
    if(acTimer)clearTimeout(acTimer);
    acTimer=setTimeout(()=>acFetch(wordEl.value),200);
  });

  wordEl.addEventListener("keydown",e=>{
    if(!acList.classList.contains("open"))return;
    if(e.key==="ArrowDown"){
      e.preventDefault();acIdx=Math.min(acIdx+1,acItems.length-1);acUpdateActive();
    }else if(e.key==="ArrowUp"){
      e.preventDefault();acIdx=Math.max(acIdx-1,-1);acUpdateActive();
    }else if((e.key==="Enter"||e.key==="Tab")&&acIdx>=0){
      e.preventDefault();acSelect(acItems[acIdx]);
    }else if(e.key==="Escape"){acClose();}
  });

  function acUpdateActive(){
    acList.querySelectorAll(".ac-item").forEach((el,i)=>{
      el.classList.toggle("active",i===acIdx);
      if(i===acIdx)el.scrollIntoView({block:"nearest"});
    });
  }

  document.addEventListener("click",e=>{if(!e.target.closest(".ac-wrap"))acClose();});

  form.addEventListener("submit",e=>{
    e.preventDefault();
    const w=wordEl.value.trim(),m=meaningEl.value.trim();if(!w||!m)return;
    const items=loadItems();items.push({word:w,meaning:m});saveItems(items);
    wordEl.value="";meaningEl.value="";lastQ="";setHint("");wordEl.focus();renderTable();
  });

  tbody.addEventListener("click",e=>{
    const sb=e.target.closest("button[data-speak]");
    if(sb){const i=Number(sb.getAttribute("data-speak"));const items=loadItems();if(Number.isInteger(i)&&i>=0&&i<items.length)speakWord(items[i].word);return;}
    const db=e.target.closest("button[data-del]");if(!db)return;
    const i=Number(db.getAttribute("data-del"));const items=loadItems();
    if(Number.isInteger(i)&&i>=0&&i<items.length){items.splice(i,1);saveItems(items);renderTable();}
  });

  btnClear.addEventListener("click",()=>{
    if(!confirm("全ての単語を削除しますか？"))return;saveItems([]);renderTable();
  });

  /* ── CSV Import ── */
  btnImport.addEventListener("click",()=>{csvFile.value="";csvFile.click();});

  const modalRoot=$("modalRoot");

  function showModal(html){modalRoot.innerHTML=html;}
  function closeModal(){modalRoot.innerHTML="";}

  function showToast(msg){
    const el=document.createElement("div");
    el.textContent=msg;
    Object.assign(el.style,{position:"fixed",bottom:"24px",left:"50%",transform:"translateX(-50%)",
      background:"#1e293b",color:"#fff",padding:"12px 24px",borderRadius:"12px",fontSize:"15px",
      fontWeight:"600",zIndex:"300",boxShadow:"0 4px 16px rgba(0,0,0,.2)",animation:"fadeUp .3s ease both"});
    document.body.appendChild(el);
    setTimeout(()=>{el.style.opacity="0";el.style.transition="opacity .3s";setTimeout(()=>el.remove(),300);},2200);
  }

  csvFile.addEventListener("change",()=>{
    const f=csvFile.files&&csvFile.files[0];if(!f)return;
    const reader=new FileReader();
    reader.onload=function(){
      let text=reader.result||"";
      if(text.charCodeAt(0)===0xFEFF)text=text.slice(1);
      const rows=parseCSV(text);
      if(!rows.length){showToast("読み込める単語がありませんでした");return;}
      const existing=loadItems();
      if(existing.length===0){
        saveItems(rows);
        renderTable();
        showToast(rows.length+"件をあなたの単語帳に追加しました");
        return;
      }
      const existSet=new Set(existing.map(x=>x.word.toLowerCase()));
      const fresh=rows.filter(r=>!existSet.has(r.word.toLowerCase()));
      const dupes=rows.length-fresh.length;
      let addDesc="重複をスキップして追加";
      if(dupes>0)addDesc=fresh.length+"件を追加（"+dupes+"件は登録済みのためスキップ）";
      else addDesc=fresh.length+"件を既存リストに追加";
      showModal(
        '<div class="modal-overlay" id="modalBg">'
        +'<div class="modal-box">'
        +'<h3>CSV取り込み</h3>'
        +'<p>CSVから <span class="highlight">'+rows.length+'件</span> の単語が見つかりました。<br>'
        +'現在 <span class="highlight">'+existing.length+'件</span> の単語が登録されています。</p>'
        +'<div class="modal-actions">'
        +(fresh.length>0
          ?'<button class="modal-btn modal-btn-add" id="modalAdd" type="button">'
           +'<span class="modal-icon">＋</span>'
           +'<span class="modal-text"><span class="modal-title">追加する</span>'
           +'<span class="modal-desc">'+esc(addDesc)+'</span></span></button>'
          :'')
        +'<button class="modal-btn modal-btn-replace" id="modalReplace" type="button">'
        +'<span class="modal-icon">↻</span>'
        +'<span class="modal-text"><span class="modal-title">上書きする</span>'
        +'<span class="modal-desc">既存の'+existing.length+'件を削除して'+rows.length+'件に置き換え</span></span></button>'
        +'<button class="modal-btn modal-btn-cancel" id="modalCancel" type="button">キャンセル</button>'
        +'</div>'
        +(fresh.length===0?'<p style="margin-top:12px;font-size:13px;color:#dc2626;">※ すべて登録済みの単語です。追加する新しい単語はありません。</p>':'')
        +'</div></div>'
      );
      $("modalCancel").onclick=closeModal;
      $("modalBg").addEventListener("click",e=>{if(e.target.id==="modalBg")closeModal();});
      if($("modalAdd"))$("modalAdd").onclick=function(){
        saveItems(existing.concat(fresh));
        renderTable();
        closeModal();
        showToast(fresh.length+"件をあなたの単語帳に追加しました");
      };
      $("modalReplace").onclick=function(){
        closeModal();
        showModal(
          '<div class="modal-overlay" id="modalBg2">'
          +'<div class="modal-box">'
          +'<h3>本当に上書きしますか？</h3>'
          +'<p>既存の <span class="highlight">'+existing.length+'件</span> はすべて削除され、'
          +'CSVの <span class="highlight">'+rows.length+'件</span> に置き換わります。<br>この操作は元に戻せません。</p>'
          +'<div class="modal-actions">'
          +'<button class="modal-btn modal-btn-replace" id="modalConfirm" type="button">'
          +'<span class="modal-icon">↻</span>'
          +'<span class="modal-text"><span class="modal-title">上書きする</span></span></button>'
          +'<button class="modal-btn modal-btn-cancel" id="modalCancel2" type="button">キャンセル</button>'
          +'</div></div></div>'
        );
        $("modalCancel2").onclick=closeModal;
        $("modalBg2").addEventListener("click",e=>{if(e.target.id==="modalBg2")closeModal();});
        $("modalConfirm").onclick=function(){
          saveItems(rows);
          renderTable();
          closeModal();
          showToast(rows.length+"件であなたの単語帳を上書きしました");
        };
      };
    };
    reader.readAsText(f,"utf-8");
  });

  function parseCSV(text){
    const lines=text.split(/\r?\n/);
    const result=[];
    let hasHeader=false;
    for(let i=0;i<lines.length;i++){
      const fields=splitCSVLine(lines[i]);
      if(fields.length<2)continue;
      const a=fields[0].trim(),b=fields[1].trim();
      if(!a||!b)continue;
      if(i===0&&/^word$/i.test(a)&&/^meaning$/i.test(b)){hasHeader=true;continue;}
      if(a.length>80||b.length>200)continue;
      result.push({word:a,meaning:b});
    }
    return result;
  }

  function splitCSVLine(line){
    const fields=[];let cur="",inQ=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(inQ){
        if(ch==='"'){
          if(i+1<line.length&&line[i+1]==='"'){cur+='"';i++;}
          else inQ=false;
        }else cur+=ch;
      }else{
        if(ch==='"')inQ=true;
        else if(ch===','){fields.push(cur);cur="";}
        else cur+=ch;
      }
    }
    fields.push(cur);
    return fields;
  }

  btnCsv.addEventListener("click",()=>{
    const items=loadItems();const bom="\uFEFF";
    const lines=["word,meaning"].concat(items.map(it=>{
      const w=String(it.word).replaceAll('"','""'),m=String(it.meaning).replaceAll('"','""');return '"'+w+'","'+m+'"';
    }));
    const csv=bom+lines.join("\r\n");const blob=new Blob([csv],{type:"text/csv;charset=utf-8"});
    const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;
    a.download="wordbook.csv";
    document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
  });

  btnPdf.addEventListener("click",async()=>{
    const items=loadItems();
    const r=await fetch("/export.pdf",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({items})});
    if(!r.ok){alert("PDF出力に失敗しました");return;}
    const blob=await r.blob();const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;
    a.download="wordbook.pdf";
    document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
  });

  /* ── Boot ── */
  initQuizBookSelect();
  void switchMode("quiz");
})();
</script>
</body>
</html>
"""

@app.get("/")
def index():
    resp = make_response(render_template_string(HTML))
    resp.headers["Cache-Control"] = "no-store, max-age=0"
    return resp


@app.get("/api/preset-csv/list")
def preset_csv_list():
    base = os.path.realpath(_pick_wordbook_preset_base())
    rels = _iter_preset_csv_rel_paths(base)
    files = [{"rel": r, "label": r.replace("/", " / ")} for r in rels]
    return Response(
        json.dumps({"ok": True, "files": files}, ensure_ascii=False),
        mimetype="application/json",
    )


@app.get("/api/preset-csv/file")
def preset_csv_file():
    rel = (request.args.get("path") or request.args.get("f") or "").strip()
    base = os.path.realpath(_pick_wordbook_preset_base())
    path = _safe_preset_csv_path(base, rel)
    if not path:
        return Response(
            json.dumps({"ok": False, "error": "not_found"}, ensure_ascii=False),
            status=404,
            mimetype="application/json",
        )
    try:
        with open(path, "rb") as fh:
            raw = fh.read()
    except OSError:
        return Response(
            json.dumps({"ok": False, "error": "read_error"}, ensure_ascii=False),
            status=500,
            mimetype="application/json",
        )
    text = raw.decode("utf-8-sig", errors="replace")
    items = _parse_preset_csv_text(text)
    return Response(
        json.dumps({"ok": True, "rel": rel, "items": items}, ensure_ascii=False),
        mimetype="application/json",
    )


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

    pad_x = 2 * mm
    pad_y = 1.5 * mm

    def clean(s: str) -> str:
        return (s or "").replace("\r", " ").replace("\n", " ").strip()

    def text_width(font_name: str, size: float, s: str) -> float:
        return pdfmetrics.stringWidth(s, font_name, size)

    def truncate_to_fit(font_name: str, size: float, s: str, max_w: float) -> str:
        s = clean(s)
        if not s:
            return ""
        if text_width(font_name, size, s) <= max_w:
            return s
        ell = "…"
        if text_width(font_name, size, ell) > max_w:
            return ""
        t = s
        while t and text_width(font_name, size, t + ell) > max_w:
            t = t[:-1]
        return (t + ell) if t else ell

    def draw_fit_text_single_line(font_name: str, base_size: float, min_size: float,
                                 x: float, y: float, s: str, max_w: float):
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

    def wrap_text(font_name: str, size: float, s: str, max_w: float) -> List[str]:
        s = clean(s)
        if not s:
            return []

        has_space = (" " in s)
        tokens = s.split(" ") if has_space else list(s)

        lines: List[str] = []
        cur = ""

        def push_line(line: str):
            if line != "":
                lines.append(line)

        for t in tokens:
            piece = (t if not has_space else (t if cur == "" else " " + t))
            trial = cur + piece
            if text_width(font_name, size, trial) <= max_w:
                cur = trial
                continue

            if cur == "":
                if has_space:
                    w = t
                    buf = ""
                    for ch in w:
                        if text_width(font_name, size, buf + ch) <= max_w:
                            buf += ch
                        else:
                            push_line(buf)
                            buf = ch
                    if buf:
                        push_line(buf)
                else:
                    if text_width(font_name, size, t) <= max_w:
                        push_line(t)
                cur = ""
            else:
                push_line(cur)
                cur = (t if not has_space else t)

        if cur:
            push_line(cur)

        return lines

    def truncate_lines_with_ellipsis(font_name: str, size: float,
                                    lines: List[str], max_lines: int, max_w: float) -> List[str]:
        if len(lines) <= max_lines:
            return lines
        kept = lines[:max_lines]
        last = kept[-1]
        ell = "…"
        if text_width(font_name, size, last + ell) <= max_w:
            kept[-1] = last + ell
            return kept
        t = last
        while t and text_width(font_name, size, t + ell) > max_w:
            t = t[:-1]
        kept[-1] = (t + ell) if t else ell
        return kept

    def draw_wrapped_fit_text(font_name: str, base_size: float, min_size: float,
                              x: float, y_top: float, cell_w: float, cell_h: float, s: str):
        s = clean(s)
        if not s:
            return

        max_w = max(1.0, cell_w - 2 * pad_x)
        max_h = max(1.0, cell_h - 2 * pad_y)

        size = base_size
        best_lines: List[str] = []
        best_size = size

        while True:
            lines = wrap_text(font_name, size, s, max_w)
            if not lines:
                return

            leading = size * 1.15
            max_lines = int(max_h // leading) if leading > 0 else 1
            if max_lines < 1:
                max_lines = 1

            fits = (len(lines) <= max_lines)

            if fits:
                best_lines = lines
                best_size = size
                break

            if size <= min_size:
                best_lines = truncate_lines_with_ellipsis(font_name, size, lines, max_lines, max_w)
                best_size = size
                break

            size -= 0.5
            if size < min_size:
                size = min_size

        c.setFont(font_name, best_size)
        leading = best_size * 1.15

        y = y_top - pad_y - best_size
        y_min = y_top - cell_h + pad_y
        for line in best_lines:
            if y < y_min:
                break
            c.drawString(x + pad_x, y, line)
            y -= leading

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
        word_base = 16
        meaning_base = 8
        min_size = 6

        for i in range(min(50, len(page_rows))):
            n = start_no + i
            side = 0 if i < 25 else 1
            row = i if i < 25 else i - 25

            x0 = left_x0 if side == 0 else right_x0
            y_top = table_top - row * row_h
            y_text_single = y_top - 0.72 * row_h

            word = str(page_rows[i].get("word", ""))
            meaning = str(page_rows[i].get("meaning", ""))

            c.setFont(jp_font, no_size)
            c.drawString(x0 + pad_x, y_text_single, str(n))

            draw_fit_text_single_line(
                "Helvetica",
                word_base,
                min_size,
                x0 + no_w + pad_x,
                y_text_single,
                word,
                word_w - 2 * pad_x,
            )

            cell_x = x0 + no_w + word_w
            draw_wrapped_fit_text(
                jp_font,
                meaning_base,
                min_size,
                cell_x,
                y_top,
                meaning_w,
                row_h,
                meaning,
            )

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
        chunk = cleaned[p:p + page_size] if cleaned else []
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


POS_MAP = {
    "Verb": "（動）",
    "Pronoun": "（代名）",
    "Adverb": "（副）",
    "Noun": "（名）",
    "Adjective": "（形）",
    "Conjunction": "（接）",
    "Auxiliary verb": "（助動）",
    "Preposition": "（前）",
    "Article": "（冠）",
    "Interjection": "（間投）",
}

POS_RE = re.compile(
    r"(?m)^={3,6}\s*(Noun|Verb|Adjective|Adverb|Pronoun|Conjunction|Auxiliary verb|Preposition|Article|Interjection)\s*={3,6}\s*$"
)
EN_HEAD_RE = re.compile(r"(?m)^==\s*English\s*==\s*$")
JA_T_RE = re.compile(r"\{\{t\+?\|ja\|([^|}\n]+)")


def _extract_ja_and_pos_nearby(wikitext: str, limit: int = 6) -> Tuple[str, List[str]]:
    if not wikitext:
        return ("", [])

    matches = list(JA_T_RE.finditer(wikitext))
    if not matches:
        return ("", [])

    first_pos = matches[0].start()

    out: List[str] = []
    seen = set()
    for m in matches:
        s = (m.group(1) or "").strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
        if len(out) >= limit:
            break

    english_start = 0
    for mh in EN_HEAD_RE.finditer(wikitext):
        if mh.start() < first_pos:
            english_start = mh.start()
        else:
            break

    region = wikitext[english_start:first_pos]

    last_pos = ""
    for mp in POS_RE.finditer(region):
        last_pos = mp.group(1)

    prefix = POS_MAP.get(last_pos, "") if last_pos else ""
    return (prefix, out)


def _lookup_case_insensitive_with_pos(word: str) -> str:
    w = word.strip()
    if not w:
        return ""

    variants = []
    for v in (w, w.lower(), w.capitalize(), w.title(), w.upper()):
        if v and v not in variants:
            variants.append(v)

    for v in variants:
        raw = _fetch_wiktionary_raw(v)
        prefix, ja_list = _extract_ja_and_pos_nearby(raw, limit=6)
        if ja_list:
            body = "、".join(ja_list)
            return (prefix + body) if prefix else body

    return ""


@app.post("/lookup")
def lookup():
    data = request.get_data(cache=False, as_text=True) or "{}"
    try:
        payload = json.loads(data)
    except Exception:
        payload = {}

    word = str(payload.get("word", "")).strip()
    meaning = _lookup_case_insensitive_with_pos(word)
    return Response(json.dumps({"meaning": meaning}, ensure_ascii=False), mimetype="application/json")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
