from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# カラーパレット
NAVY   = RGBColor(0x1F, 0x35, 0x64)
BLUE   = RGBColor(0x2E, 0x75, 0xB6)
LIGHT  = RGBColor(0xD6, 0xE4, 0xF0)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GRAY   = RGBColor(0xF2, 0xF2, 0xF2)
DARK   = RGBColor(0x26, 0x26, 0x26)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

# ── ユーティリティ ──────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             size=18, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txb

def section_bar(slide, title, top=Inches(1.1)):
    add_rect(slide, Inches(0.4), top, Inches(12.5), Inches(0.5), fill=BLUE)
    add_text(slide, title, Inches(0.55), top+Inches(0.04), Inches(12), Inches(0.44),
             size=18, bold=True, color=WHITE)

def slide_footer(slide, num, total=7):
    add_text(slide, "コンフォール井の頭公園南 修繕委員会 第1回定例　2026.6.7",
             Inches(0.4), Inches(7.1), Inches(10), Inches(0.3),
             size=10, color=RGBColor(0x80,0x80,0x80))
    add_text(slide, f"{num} / {total}",
             Inches(12.5), Inches(7.1), Inches(0.7), Inches(0.3),
             size=10, color=RGBColor(0x80,0x80,0x80), align=PP_ALIGN.RIGHT)

def title_bar(slide):
    add_rect(slide, 0, 0, W, Inches(1.05), fill=NAVY)

def add_table(slide, headers, rows, l, t, w, h,
              hdr_size=14, data_size=13, col_widths=None):
    cols = len(headers)
    tbl = slide.shapes.add_table(len(rows)+1, cols, l, t, w, h).table
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = cw
    else:
        cw = w // cols
        for i in range(cols):
            tbl.columns[i].width = cw

    for ci, hd in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0] if p.runs else p.add_run()
        run.text = hd
        run.font.size = Pt(hdr_size)
        run.font.bold = True
        run.font.color.rgb = WHITE

    for ri, row in enumerate(rows):
        bg = GRAY if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row):
            cell = tbl.cell(ri+1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.text = str(val)
            run.font.size = Pt(data_size)
            run.font.color.rgb = DARK
    return tbl

# ════════════════════════════════════════════════
# SLIDE 1 ｜ 表紙
# ════════════════════════════════════════════════
s1 = prs.slides.add_slide(BLANK)
add_rect(s1, 0, 0, W, H, fill=NAVY)
add_rect(s1, 0, Inches(2.8), W, Inches(1.9), fill=BLUE)

add_text(s1, "コンフォール井の頭公園南", 0, Inches(1.5), W, Inches(0.8),
         size=26, bold=False, color=LIGHT, align=PP_ALIGN.CENTER)
add_text(s1, "修繕委員会  第1回定例（キックオフ）", 0, Inches(2.9), W, Inches(0.9),
         size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(s1, "2026年6月7日（日）　14:00〜15:00", 0, Inches(3.88), W, Inches(0.6),
         size=20, color=LIGHT, align=PP_ALIGN.CENTER)
add_text(s1, "三鷹市下連雀八丁目地区公会堂", 0, Inches(4.5), W, Inches(0.5),
         size=16, color=LIGHT, align=PP_ALIGN.CENTER)
add_text(s1, "出席：小橋　／　石川　／　十文字　／　阿部　／　武藤　／　新関",
         0, Inches(5.7), W, Inches(0.5),
         size=15, color=RGBColor(0xCC,0xDD,0xEE), align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════
# SLIDE 2 ｜ 委員会の目的・アジェンダ
# ════════════════════════════════════════════════
s2 = prs.slides.add_slide(BLANK)
title_bar(s2)
add_text(s2, "委員会の目的・本日のアジェンダ", Inches(0.4), Inches(0.15),
         Inches(12), Inches(0.75), size=24, bold=True, color=WHITE)
slide_footer(s2, 2)

# 左：目的
add_rect(s2, Inches(0.4), Inches(1.15), Inches(6.2), Inches(0.46), fill=BLUE)
add_text(s2, "委員会の目的", Inches(0.55), Inches(1.18), Inches(6.0), Inches(0.4),
         size=16, bold=True, color=WHITE)

purpose = [
    "● 建物・設備の状態を把握し、長期修繕計画の策定と\n　 必要に応じた修繕積立金の見直しを推進",
    "● 住民の意見を反映した計画立案と情報共有を支援",
    "● 決定は理事会・総会が行い、委員会は\n　 調査・提案・情報発信を担う",
]
y = Inches(1.75)
for p in purpose:
    add_text(s2, p, Inches(0.55), y, Inches(5.9), Inches(0.82), size=15, color=DARK)
    y += Inches(0.85)

add_rect(s2, Inches(0.4), Inches(4.45), Inches(6.2), Inches(0.04), fill=BLUE)
add_text(s2, "位置づけ：理事会と連携する住民委員会　決定機関は理事会・総会",
         Inches(0.55), Inches(4.55), Inches(6.0), Inches(0.5),
         size=14, color=NAVY)

# 右：アジェンダ
add_rect(s2, Inches(7.0), Inches(1.15), Inches(5.9), Inches(0.46), fill=BLUE)
add_text(s2, "本日のアジェンダ", Inches(7.15), Inches(1.18), Inches(5.7), Inches(0.4),
         size=16, bold=True, color=WHITE)

agenda = [
    ("1", "メンバー紹介・役割分担"),
    ("2", "活動スコープ・スケジュールの確認"),
    ("3", "コンサル業者選定の進め方"),
    ("4", "省エネ改修・緊急修繕の総会提案（新テーマ）"),
    ("5", "運営ルール・決定事項・次回予定"),
]
y = Inches(1.75)
for num, item in agenda:
    add_rect(s2, Inches(7.0), y, Inches(0.5), Inches(0.52), fill=NAVY)
    add_text(s2, num, Inches(7.0), y+Inches(0.05), Inches(0.5), Inches(0.44),
             size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s2, Inches(7.53), y, Inches(5.35), Inches(0.52),
             fill=GRAY if int(num)%2==0 else WHITE, line=LIGHT)
    add_text(s2, item, Inches(7.68), y+Inches(0.07), Inches(5.1), Inches(0.44),
             size=15, color=DARK)
    y += Inches(0.6)

# ════════════════════════════════════════════════
# SLIDE 3 ｜ メンバー紹介・役割分担
# ════════════════════════════════════════════════
s3 = prs.slides.add_slide(BLANK)
title_bar(s3)
add_text(s3, "メンバー紹介・役割分担", Inches(0.4), Inches(0.15),
         Inches(12), Inches(0.75), size=24, bold=True, color=WHITE)
slide_footer(s3, 3)

add_rect(s3, Inches(0.4), Inches(1.15), Inches(12.5), Inches(0.5), fill=BLUE)
add_text(s3, "出席メンバー", Inches(0.55), Inches(1.18), Inches(3.0), Inches(0.42),
         size=16, bold=True, color=WHITE)
add_text(s3, "小橋　／　石川　／　十文字　／　阿部　／　武藤　／　新関",
         Inches(3.8), Inches(1.18), Inches(9.0), Inches(0.42),
         size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

section_bar(s3, "【議題】役割分担の決定", top=Inches(1.85))

add_table(s3,
    ["役割", "内容"],
    [
        ["委員長",   "委員会の取りまとめ・対外窓口・理事会への報告"],
        ["副委員長", "委員長補佐・議事録管理（AIツール活用）・情報共有の整理"],
    ],
    Inches(0.6), Inches(2.45), Inches(12.1), Inches(1.3),
    hdr_size=15, data_size=15,
    col_widths=[Inches(2.2), Inches(9.9)]
)

add_rect(s3, Inches(0.4), Inches(3.95), Inches(12.5), Inches(0.46), fill=LIGHT)
add_text(s3, "追加担当について（本日議論）", Inches(0.55), Inches(3.98),
         Inches(12.0), Inches(0.4), size=15, bold=True, color=NAVY)

disc = [
    "● 積立金評価担当（修繕積立金の適正性検証）を設けるか",
    "● その他、必要な担当・役割があるかを議論する",
]
y = Inches(4.55)
for d in disc:
    add_text(s3, d, Inches(0.7), y, Inches(12.0), Inches(0.55), size=15, color=DARK)
    y += Inches(0.58)

# ════════════════════════════════════════════════
# SLIDE 4 ｜ 活動スコープ・改定スケジュール
# ════════════════════════════════════════════════
s4 = prs.slides.add_slide(BLANK)
title_bar(s4)
add_text(s4, "活動スコープ・改定スケジュール", Inches(0.4), Inches(0.15),
         Inches(12), Inches(0.75), size=24, bold=True, color=WHITE)
slide_footer(s4, 4)

section_bar(s4, "活動スコープ（3本柱）", top=Inches(1.1))

scope = [
    ("①", "長期修繕計画の策定（30年）",   "コンサル選定 → 現地調査 → 計画書作成 → 総会議決"),
    ("②", "短期修繕計画の立案（直近5年）", "優先修繕箇所の整理 → コンサル提案 → 総会議決"),
    ("③", "省エネ改修・緊急修繕の提案【NEW】",
     "断熱窓を起点に → 積立金との整合性を確認しながら来年度総会への提案を並行推進"),
]
y = Inches(1.72)
for num, title, detail in scope:
    add_rect(s4, Inches(0.4), y, Inches(0.55), Inches(0.6), fill=NAVY)
    add_text(s4, num, Inches(0.4), y+Inches(0.06), Inches(0.55), Inches(0.5),
             size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s4, Inches(0.98), y, Inches(11.92), Inches(0.6),
             fill=GRAY if num != "③" else RGBColor(0xFF,0xF0,0xCC), line=LIGHT)
    add_text(s4, f"{title}　　{detail}",
             Inches(1.12), y+Inches(0.07), Inches(11.6), Inches(0.5), size=13, color=DARK)
    y += Inches(0.68)

note = "※キックオフが5月予定から6月7日に遅延。以下は1ヶ月後ろ倒しの改定版。"
add_text(s4, note, Inches(0.4), Inches(3.82), Inches(12.5), Inches(0.36),
         size=12, color=RGBColor(0x80,0x60,0x00))

# 左：①長期
add_rect(s4, Inches(0.4), Inches(4.22), Inches(6.4), Inches(0.4), fill=BLUE)
add_text(s4, "① 長期修繕計画", Inches(0.55), Inches(4.25), Inches(6.2), Inches(0.34),
         size=14, bold=True, color=WHITE)
long_rows = [
    ["6月〜",  "コンサル3社へプレゼン依頼・日程調整"],
    ["7月",    "コンサル3社プレゼン実施"],
    ["8月",    "業者選定・理事会決議"],
    ["9月",    "実施計画案の提示"],
    ["10月",   "実施計画案の決議・現地調査"],
    ["11月",   "調査報告・中長期計画 納品・理事会報告"],
    ["12月",   "住民向け全戸配布 / 積立金変更案の検討"],
    ["来年総会","修繕計画・積立金変更の議決"],
]
add_table(s4, ["時期", "内容（委員会中心）"],
          long_rows, Inches(0.4), Inches(4.66), Inches(6.4), Inches(2.55),
          hdr_size=13, data_size=12,
          col_widths=[Inches(1.2), Inches(5.2)])

# 右：②短期
add_rect(s4, Inches(7.1), Inches(4.22), Inches(5.8), Inches(0.4), fill=BLUE)
add_text(s4, "② 短期修繕計画（直近5年）", Inches(7.25), Inches(4.25),
         Inches(5.6), Inches(0.34), size=14, bold=True, color=WHITE)
short_rows = [
    ["8月",    "要求仕様提示（インフラ緊急性・省エネ優先）"],
    ["9月",    "計画案提示（省エネ改修を中心に）"],
    ["11月",   "現地調査後の優先修繕案提示"],
    ["12月",   "計画案の検討 / 総会提案資料作成"],
    ["来年総会","議決"],
]
add_table(s4, ["時期", "内容（委員会中心）"],
          short_rows, Inches(7.1), Inches(4.66), Inches(5.8), Inches(2.1),
          hdr_size=13, data_size=12,
          col_widths=[Inches(1.2), Inches(4.6)])

# ════════════════════════════════════════════════
# SLIDE 5 ｜ コンサル業者選定の進め方
# ════════════════════════════════════════════════
s5 = prs.slides.add_slide(BLANK)
title_bar(s5)
add_text(s5, "コンサル業者選定の進め方", Inches(0.4), Inches(0.15),
         Inches(12), Inches(0.75), size=24, bold=True, color=WHITE)
slide_footer(s5, 5)

section_bar(s5, "候補3社：TDS　／　オフィスレコン　／　センターオフィス", top=Inches(1.1))

# 左：プレゼン方針
add_rect(s5, Inches(0.4), Inches(1.75), Inches(5.9), Inches(0.4), fill=NAVY)
add_text(s5, "プレゼン実施方針", Inches(0.55), Inches(1.78), Inches(5.7), Inches(0.34),
         size=15, bold=True, color=WHITE)
pres_items = [
    "● 7月中に3社それぞれ30分のプレゼンを実施",
    "● 場所：会場を1日確保して3社連続\n　 または管理事務室（Web参加可）",
    "● 本日：連絡・日程調整・会場確保の担当を決定",
]
y = Inches(2.25)
for item in pres_items:
    add_text(s5, item, Inches(0.6), y, Inches(5.6), Inches(0.65), size=14, color=DARK)
    y += Inches(0.65)

add_rect(s5, Inches(0.4), Inches(4.25), Inches(5.9), Inches(0.4), fill=NAVY)
add_text(s5, "依頼内容", Inches(0.55), Inches(4.28), Inches(5.7), Inches(0.34),
         size=15, bold=True, color=WHITE)
dep_items = [
    "● 劣化診断調査（第3回大規模修繕に向けた建物調査）",
    "● 30年中長期修繕計画の作成・修繕積立金の評価",
    "● 短期修繕計画（5年）の立案（追加依頼）",
    "● 省エネ改修提案（追加依頼）",
]
y = Inches(4.75)
for item in dep_items:
    add_text(s5, item, Inches(0.6), y, Inches(5.6), Inches(0.5), size=14, color=DARK)
    y += Inches(0.5)

# 右：選定基準
add_rect(s5, Inches(6.7), Inches(1.75), Inches(6.2), Inches(0.4), fill=NAVY)
add_text(s5, "選定基準（3点）", Inches(6.85), Inches(1.78), Inches(6.0), Inches(0.34),
         size=15, bold=True, color=WHITE)
add_table(s5,
    ["#", "基準"],
    [
        ["①", "要求仕様を満たしているか"],
        ["②", "金額が妥当か（想定100〜200万円）"],
        ["③", "短期計画・省エネ改修のコンサル提案が可能か"],
    ],
    Inches(6.7), Inches(2.23), Inches(6.2), Inches(1.5),
    hdr_size=14, data_size=14,
    col_widths=[Inches(0.5), Inches(5.7)]
)

add_rect(s5, Inches(6.7), Inches(3.9), Inches(6.2), Inches(0.44),
         fill=RGBColor(0xFF,0xF0,0xCC))
add_text(s5, "【議題】見積事業者選定プロセスの追加について",
         Inches(6.85), Inches(3.93), Inches(6.0), Inches(0.38),
         size=14, bold=True, color=RGBColor(0x80,0x40,0x00))
disc_items = [
    "● 管理会社任せにせず、委員会として選定基準・\n　 比較評価プロセスを持つことで透明性が高まる",
    "● 一方、現時点では計画策定が優先。\n　 コンサル選定後に整備する選択肢もある",
    "● 論点：今の段階で検討するか、後回しにするか",
]
y = Inches(4.45)
for item in disc_items:
    add_text(s5, item, Inches(6.85), y, Inches(5.9), Inches(0.75), size=14, color=DARK)
    y += Inches(0.75)

# ════════════════════════════════════════════════
# SLIDE 6 ｜ 省エネ改修・緊急修繕の来年総会提案
# ════════════════════════════════════════════════
s6 = prs.slides.add_slide(BLANK)
title_bar(s6)
add_text(s6, "新テーマ：省エネ改修・緊急修繕の来年総会提案",
         Inches(0.4), Inches(0.15), Inches(12), Inches(0.75),
         size=24, bold=True, color=WHITE)
slide_footer(s6, 6)

add_rect(s6, Inches(0.4), Inches(1.12), Inches(12.5), Inches(0.48), fill=LIGHT)
add_text(s6,
    "長期計画策定と並行し、省エネ改修・緊急修繕は早期に独立した検討ラインで動く。各提案は修繕積立金への影響を確認しながら進める。",
    Inches(0.55), Inches(1.14), Inches(12.2), Inches(0.42), size=14, color=NAVY)

add_table(s6,
    ["テーマ", "内容", "目標"],
    [
        ["省エネ改修",
         "断熱窓改修を優先検討（専有部対象）\n補助金・助成金（都・市・国）の活用可能性を調査\nコンサルにも提案を依頼",
         "2027年総会への提案"],
        ["緊急修繕",
         "インフラ系（給排水管・電気設備等）の現状把握を優先\nコンサル調査と並行して早期実態把握\n緊急度が高いものは総会を待たず理事会決議で対応",
         "随時・理事会判断"],
    ],
    Inches(0.4), Inches(1.72), Inches(12.5), Inches(2.9),
    hdr_size=15, data_size=14,
    col_widths=[Inches(1.8), Inches(8.5), Inches(2.2)]
)

add_rect(s6, Inches(0.4), Inches(4.82), Inches(12.5), Inches(0.44), fill=NAVY)
add_text(s6, "修繕積立金との整合性確認（積立金評価担当が設置された場合）",
         Inches(0.55), Inches(4.85), Inches(12.2), Inches(0.38),
         size=15, bold=True, color=WHITE)
integ = [
    "● 省エネ改修：費用・補助金差引後の実質負担と積立金残高を照合",
    "● 緊急修繕：修繕費用の積立金への影響を試算・確認し、緊急度判断の材料とする",
]
y = Inches(5.4)
for item in integ:
    add_text(s6, item, Inches(0.6), y, Inches(12.2), Inches(0.52), size=15, color=DARK)
    y += Inches(0.55)

# ════════════════════════════════════════════════
# SLIDE 7 ｜ 運営ルール・決定事項・次回予定
# ════════════════════════════════════════════════
s7 = prs.slides.add_slide(BLANK)
title_bar(s7)
add_text(s7, "運営ルール・本日の決定事項・次回予定",
         Inches(0.4), Inches(0.15), Inches(12), Inches(0.75),
         size=24, bold=True, color=WHITE)
slide_footer(s7, 7)

# 左：運営ルール
add_rect(s7, Inches(0.4), Inches(1.12), Inches(5.9), Inches(0.44), fill=BLUE)
add_text(s7, "運営ルール", Inches(0.55), Inches(1.15), Inches(5.7), Inches(0.38),
         size=16, bold=True, color=WHITE)
add_table(s7,
    ["項目", "内容"],
    [
        ["打合せ頻度", "月1回程度"],
        ["コミュニケーション", "Google Meet・チャット＋メール\n必要に応じ書面・対面"],
        ["議事録", "副委員長（AIツール活用）"],
    ],
    Inches(0.4), Inches(1.64), Inches(5.9), Inches(2.1),
    hdr_size=15, data_size=15,
    col_widths=[Inches(2.1), Inches(3.8)]
)

# 右：決定事項
add_rect(s7, Inches(6.7), Inches(1.12), Inches(6.2), Inches(0.44), fill=BLUE)
add_text(s7, "本日の決定事項", Inches(6.85), Inches(1.15), Inches(6.0), Inches(0.38),
         size=16, bold=True, color=WHITE)
decisions = [
    "□　役割分担（委員長・副委員長・その他担当）",
    "□　コンサル3社への連絡・プレゼン日程調整・会場確保の担当",
    "□　見積事業者選定プロセスを今設けるか後回しにするか",
    "□　省エネ改修・緊急修繕の優先確認事項",
    "□　次回開催日程",
]
y = Inches(1.64)
for i, d in enumerate(decisions):
    bg = GRAY if i % 2 == 0 else WHITE
    add_rect(s7, Inches(6.7), y, Inches(6.2), Inches(0.5), fill=bg, line=LIGHT)
    add_text(s7, d, Inches(6.88), y+Inches(0.06), Inches(5.9), Inches(0.42),
             size=14, color=DARK)
    y += Inches(0.5)

# 次回予定
add_rect(s7, Inches(0.4), Inches(3.95), Inches(12.5), Inches(0.46), fill=NAVY)
add_text(s7, "次回（7月）　コンサル3社プレゼン実施　→　業者選定に向けた評価・議論",
         Inches(0.55), Inches(3.98), Inches(12.2), Inches(0.4),
         size=16, bold=True, color=WHITE)
next_items = [
    "● コンサル3社（TDS・オフィスレコン・センターオフィス）のプレゼンを実施（各30分）",
    "● 各社の提案内容・金額・対応範囲を委員会内で評価・比較",
    "● 業者選定の方向性を議論し、8月の理事会決議へ向けて準備",
]
y = Inches(4.55)
for item in next_items:
    add_text(s7, item, Inches(0.6), y, Inches(12.2), Inches(0.55), size=15, color=DARK)
    y += Inches(0.58)

# ════════════════════════════════════════════════
out = "/home/user/Ryuichi_Ishikawa_Private/notes/mansion/修繕委員会_第1回定例_キックオフ.pptx"
prs.save(out)
print("saved:", out)
