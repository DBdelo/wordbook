# app.py の _draw_pdf_word_sheet を丸ごとこれに置き換え

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

    def draw_fit_text_single_line(font_name: str, base_size: float, min_size: float, x: float, y: float, s: str, max_w: float):
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
                # 1トークンでも入らない場合は、文字単位で切る
                if has_space:
                    # 単語が長すぎる
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
                    # 1文字でも入らないなら捨てる
                    if text_width(font_name, size, t) <= max_w:
                        push_line(t)
                cur = ""
            else:
                push_line(cur)
                cur = (t if not has_space else t)

        if cur:
            push_line(cur)

        return lines

    def truncate_lines_with_ellipsis(font_name: str, size: float, lines: List[str], max_lines: int, max_w: float) -> List[str]:
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

        # 上から下へ
        y = y_top - pad_y - best_size
        for line in best_lines:
            if y < (y_top - cell_h + pad_y):
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
        word_base = 11
        meaning_base = 9   # 標準を下げる
        min_size = 6       # 最小 6

        for i in range(min(50, len(page_rows))):
            n = start_no + i
            side = 0 if i < 25 else 1
            row = i if i < 25 else i - 25

            x0 = left_x0 if side == 0 else right_x0

            y_top = table_top - row * row_h
            y_text_single = y_top - 0.72 * row_h  # No/word は従来通り1行

            word = str(page_rows[i].get("word", ""))
            meaning = str(page_rows[i].get("meaning", ""))

            # No
            c.setFont(jp_font, no_size)
            c.drawString(x0 + pad_x, y_text_single, str(n))

            # word（1行：縮小+省略）
            draw_fit_text_single_line(
                "Helvetica",
                word_base,
                min_size,
                x0 + no_w + pad_x,
                y_text_single,
                word,
                word_w - 2 * pad_x,
            )

            # meaning（折り返し+必要なら縮小、最小6）
            cell_x = x0 + no_w + word_w
            cell_w = meaning_w
            cell_h = row_h
            draw_wrapped_fit_text(
                jp_font,
                meaning_base,
                min_size,
                cell_x,
                y_top,
                cell_w,
                cell_h,
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
        chunk = cleaned[p : p + page_size] if cleaned else []
        draw_page(chunk, start_no)
        start_no += page_size
        if p + page_size < len(cleaned):
            c.showPage()

    c.save()
    return out.getvalue()
