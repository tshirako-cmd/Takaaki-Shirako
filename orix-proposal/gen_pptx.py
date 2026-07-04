# -*- coding: utf-8 -*-
"""M&A Architect ORIX提案ワンページャー PPTX生成（A4横・2スライド）"""
from pptx import Presentation
from pptx.util import Mm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

NAVY = RGBColor(0x1F, 0x38, 0x64)
NAVY2 = RGBColor(0x2E, 0x4A, 0x78)
NAVY3 = RGBColor(0x3B, 0x5A, 0x8C)
GOLD = RGBColor(0xD4, 0xA0, 0x17)
DGRAY = RGBColor(0x44, 0x54, 0x6A)
BODY = RGBColor(0x33, 0x33, 0x33)
LINE = RGBColor(0xD9, 0xDE, 0xE8)
BGSOFT = RGBColor(0xF5, 0xF7, 0xFA)
NAVYSOFT = RGBColor(0xEE, 0xF1, 0xF7)
RED = RGBColor(0xC0, 0x00, 0x00)
REDSOFT = RGBColor(0xF7, 0xEC, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ZEBRA = RGBColor(0xFA, 0xFB, 0xFD)

JP_FONT = "Yu Gothic"


def set_font(run, size, bold=False, color=BODY, italic=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = JP_FONT
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", JP_FONT)


def rect(slide, x, y, w, h, fill=None, line=None, line_pt=0.5):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(x), Mm(y), Mm(w), Mm(h))
    sh.shadow.inherit = False
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_pt)
    return sh


def txt(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
        space_after=0.4, line_spacing=1.15):
    """paras: list of either list-of-run-tuples or dict {runs, align, space_after, line_spacing}
    run tuple: (text, size, bold, color) or (text, size, bold, color, italic)"""
    tb = slide.shapes.add_textbox(Mm(x), Mm(y), Mm(w), Mm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, para in enumerate(paras):
        if isinstance(para, dict):
            runs = para["runs"]
            p_align = para.get("align", align)
            p_space = para.get("space_after", space_after)
            p_ls = para.get("line_spacing", line_spacing)
        else:
            runs = para
            p_align, p_space, p_ls = align, space_after, line_spacing
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = p_align
        p.space_after = Pt(p_space)
        p.line_spacing = p_ls
        for rt in runs:
            r = p.add_run()
            r.text = rt[0]
            set_font(r, rt[1], rt[2] if len(rt) > 2 else False,
                     rt[3] if len(rt) > 3 else BODY,
                     rt[4] if len(rt) > 4 else False)
    return tb


def bullet(slide, x, y, w, h, items, size=7.4, color=BODY):
    paras = []
    for it in items:
        paras.append([("● " if False else "・", size, False, NAVY)] and
                     [("・", size, True, NAVY), (it, size, False, color)])
    return txt(slide, x, y, w, h, paras, space_after=0.8, line_spacing=1.25)


def header(slide, title_runs, sub_runs, brand_sub):
    txt(slide, 200, 4.5, 86, 4.5,
        [[("Strictly Confidential", 7, True, RED)]], align=PP_ALIGN.RIGHT)
    txt(slide, 11, 10, 200, 14, [
        {"runs": title_runs, "space_after": 1.5},
        {"runs": sub_runs},
    ])
    txt(slide, 200, 10, 86, 13, [
        {"runs": [("DESIGN", 9, True, NAVY), ("・", 9, True, GOLD), ("MANAGEMENT", 9, True, NAVY)],
         "align": PP_ALIGN.RIGHT, "space_after": 1},
        {"runs": [(brand_sub, 7, False, DGRAY)], "align": PP_ALIGN.RIGHT},
    ], anchor=MSO_ANCHOR.BOTTOM)
    rect(slide, 11, 24.6, 275, 0.9, fill=NAVY)


def lead_band(slide, y, h, badge_lbl, badge_val, summary_runs, badge_w=52):
    rect(slide, 11, y, badge_w, h, fill=NAVY)
    rect(slide, 11 + badge_w, y, 1.2, h, fill=GOLD)
    txt(slide, 14, y + 1.5, badge_w - 6, h - 3, [
        {"runs": [(badge_lbl, 6.5, False, RGBColor(0xC5, 0xCE, 0xDD))], "space_after": 1},
        {"runs": [(badge_val, 12.5, True, WHITE)]},
    ], anchor=MSO_ANCHOR.MIDDLE)
    sx = 11 + badge_w + 1.2
    rect(slide, sx, y, 275 - badge_w - 1.2, h, fill=BGSOFT)
    txt(slide, sx + 3.5, y + 1.5, 275 - badge_w - 1.2 - 7, h - 3,
        [summary_runs], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)


def footer(slide, left_text, page_no):
    rect(slide, 11, 199.5, 275, 0.45, fill=NAVY)
    txt(slide, 11, 201, 200, 5, [[(left_text, 7, False, DGRAY)]])
    txt(slide, 216, 201, 70, 5,
        [[("© 2026 Design Management LLC　　" + page_no, 7, False, DGRAY)]],
        align=PP_ALIGN.RIGHT)


def h2(slide, x, y, text, w=120):
    rect(slide, x, y + 0.4, 1.2, 4.2, fill=GOLD)
    txt(slide, x + 2.4, y, w, 5, [[(text, 9.5, True, NAVY)]])


# ============================================================
prs = Presentation()
prs.slide_width = Mm(297)
prs.slide_height = Mm(210)
blank = prs.slide_layouts[6]

# ================= SLIDE 1: ビジネスマップ =================
s = prs.slides.add_slide(blank)

header(
    s,
    [("M&A Architect ビジネスマップ", 15.5, True, NAVY), ("  ／ Business Map", 10.5, True, DGRAY)],
    [("買い手のM&A機能を、設計する", 8, False, DGRAY), ("独自AIシステム", 8, True, BODY),
     ("。", 8, False, DGRAY), ("誰の・どんな課題を・どう解決するか", 8, True, BODY),
     ("——想定顧客と利用シーンで見る全体像。", 8, False, DGRAY)],
    "ORIX 事業投資部門 ご提案",
)

lead_band(
    s, 27.5, 13, "WHAT IT IS", "買い手M&A特化AI",
    [("代表 白子貴章が", 8.4, False, BODY), ("26年の当事者経験", 8.4, True, NAVY),
     ("で体系化した判断基準をAIに実装（", 8.4, False, BODY), ("5カテゴリ・26モジュール", 8.4, True, NAVY),
     ("、中核2機構は", 8.4, False, BODY), ("特許出願中", 8.4, True, NAVY),
     ("）。EDINET等の一次情報に直結し、", 8.4, False, BODY), ("翌営業日", 8.4, True, NAVY),
     ("で判断材料を構造化（判断系は代表レビュー）。", 8.4, False, BODY),
     ("人は最終判断に集中", 8.4, True, NAVY), ("する。", 8.4, False, BODY)],
)

# --- persona cards ---
personas = [
    ("中堅事業会社の経営企画", "M&A兼務1〜3名・専任チームなし",
     "仲介からの持ち込みが続くが検討が属人化し、初期評価に1ヶ月。小型案件はFA費用が見合わず検討見送り。",
     [("案件評価一式を翌営業日で構造化", 7.2, True, NAVY),
      ("。兼務でも回る検討体制になり、見送り判断も根拠つきで即答できる。", 7.2, False, BODY)]),
    ("初めてM&Aに取り組むオーナー企業・HD", "中計でM&A成長を掲げた段階",
     "目的・投資枠・ターゲット要件が言語化されておらず、持ち込み案件に個別対応で振り回される。",
     [("戦略策定エンジンが事業戦略から目的類型・投資枠・要件を逆算", 7.2, True, NAVY),
      ("。受け身の検討から能動的なソーシングへ。", 7.2, False, BODY)]),
    ("投資会社・事業投資部門", "PE／CVC／ORIX 事業投資部門",
     "パイプラインは多いが、初期スクリーニングと投資後のKPIモニタリングに人手を取られ、深い検討に時間を割けない。",
     [("11評価モジュール並列で一次評価を量産", 7.2, True, NAVY),
      ("し、PMI KPI監視で投資後管理を標準化。人は交渉・判断に集中。", 7.2, False, BODY)]),
]
py, ph, pw, gap = 43, 27, (275 - 4) / 3, 2
for i, (who, who_s, pain, sol_runs) in enumerate(personas):
    px = 11 + i * (pw + gap)
    rect(s, px, py, pw, ph, fill=WHITE, line=LINE, line_pt=0.5)
    rect(s, px, py, pw, 0.9, fill=GOLD)
    txt(s, px + 2.4, py + 2, pw - 4.8, ph - 3.5, [
        {"runs": [(who, 8.4, True, NAVY), ("  " + who_s, 6.4, False, DGRAY)], "space_after": 1.4},
        {"runs": [("課題", 6.4, True, RED), ("  " + pain, 7.2, False, BODY)],
         "space_after": 1.4, "line_spacing": 1.22},
        {"runs": [("解決", 6.4, True, NAVY), ("  ", 7.2, False, BODY)] + sol_runs,
         "line_spacing": 1.22},
    ])

# --- 3-layer ribbon ---
ly, lh = 72.5, 6.5
total, gaps = 275 - 4, 2
w1 = total * 2 / 5.15
w3 = total * 1.15 / 5.15
layers = [
    (11, w1, NAVY, "設計する", "VISION & DESIGN ｜ 戦略を生み、要件・体制を逆算"),
    (11 + w1 + gaps, w1, NAVY2, "見極める", "DISCERNING ｜ 案件と条件を見極める"),
    (11 + 2 * (w1 + gaps), w3, NAVY3, "経営する", "MANAGEMENT ｜ 買収後を経営"),
]
for lx, lw, col, jp, en in layers:
    rect(s, lx, ly, lw, lh, fill=col)
    txt(s, lx + 3, ly, lw - 6, lh,
        [[(jp, 9, True, WHITE), ("  " + en, 6.4, False, RGBColor(0xD5, 0xDC, 0xE8))]],
        anchor=MSO_ANCHOR.MIDDLE)

# --- 5 category cards ---
cats = [
    ("① STRATEGY", "戦略策定", "事業戦略からM&A戦略を導出・投資枠を逆算",
     "中計にM&Aを掲げたが、目的・投資枠・ターゲット要件が言語化できていない",
     "主なモジュール",
     ["戦略策定エンジン（3分岐×6ステップ）", "Series 5 投資枠算定", "EDINETスクリーニング／戦略フィットスコア"],
     "M&A戦略ナラティブレポート（Word）／投資枠算定モデル（Excel）"),
    ("② SOURCING", "ソーシング", "候補を網羅発掘しスコアで絞り込む",
     "候補が仲介持ち込み頼み。要件に合う相手を能動的に探せていない",
     "主なモジュール",
     ["ロングリスト発掘（Web×gBizINFO）", "ショートリスト化（発生確度×シナジー）", "3段階ターゲット統合・名寄せ"],
     "ロングリスト／ショートリスト（4象限マップ付）／統合マスターリスト"),
    ("③ DEAL ASSESSMENT", "案件評価", "多面評価で投資判断材料を構造化",
     "IMが持ち込まれ、短期間で一次回答が必要。だが専任チームがいない",
     "主なモジュール",
     ["ディール評価オーケストレーター（4ステップ）", "11評価モジュール並列（事業/市場/競争/財務/シナジー/リスク…）", "バリュエーション／4シナリオProjection"],
     "統合Excelモデル／IC統合評価レポート（Word）／IC資料（PPTX）"),
    ("④ DD・NEGOTIATION・DECISION", "DD・交渉・意思決定", "DD設計〜SPA条件〜クロージング管理",
     "DD報告書が分厚く、横断論点とSPA交渉材料を整理しきれない",
     "主なモジュール（3層）",
     ["DD準備（スコープ／QA・資料依頼シート）", "DD統合分析（Findings統合／論点抽出）", "意思決定・クロージング（取締役会起案／SPA）"],
     "DD Findings統合一覧／取締役会起案資料／SPA主要条件整理表"),
    ("⑤ PMI", "PMI", "買収後の統合を体系的に支援",
     "買収後のKPI管理・100日プランが属人化し、統合が計画倒れに",
     "主なモジュール",
     ["ガバナンス層定義", "月次KPIモニタリング（業種別プリセット）", "100日プラン生成／Earn-out横断ビュー"],
     "100日プラン（Excel）／月次KPI管理表／PMI計画書（Word）"),
]
cy, ch = 81, 82
cw = (275 - 4 * 3.2) / 5
for i, (num, nm, en, scene, mlbl, mods, out) in enumerate(cats):
    cx = 11 + i * (cw + 3.2)
    rect(s, cx, cy, cw, ch, fill=WHITE, line=LINE, line_pt=0.5)
    rect(s, cx, cy, cw, 0.9, fill=NAVY)
    hh = 15.5
    rect(s, cx, cy + 0.9, cw, hh, fill=NAVYSOFT)
    txt(s, cx + 2.2, cy + 2.2, cw - 4.4, hh - 2, [
        {"runs": [(num, 6.6, True, GOLD)], "space_after": 0.6},
        {"runs": [(nm, 9.5, True, NAVY)], "space_after": 0.6},
        {"runs": [(en, 6.2, False, DGRAY)]},
    ])
    by = cy + 0.9 + hh + 2
    scene_paras = [
        {"runs": [("利用シーン", 6.1, True, GOLD)], "space_after": 0.6},
        {"runs": [(scene, 7.1, False, DGRAY)], "space_after": 2, "line_spacing": 1.25},
        {"runs": [(mlbl, 6.3, True, NAVY)], "space_after": 0.8},
    ]
    for m in mods:
        scene_paras.append({"runs": [("・", 7.2, True, NAVY), (m, 7.2, False, BODY)],
                            "space_after": 0.7, "line_spacing": 1.2})
    txt(s, cx + 2.2, by, cw - 4.4, ch - hh - 16, scene_paras)
    oh = 12.5
    rect(s, cx + 1.6, cy + ch - oh - 1.6, cw - 3.2, oh, fill=BGSOFT)
    txt(s, cx + 3, cy + ch - oh - 0.6, cw - 6, oh - 2,
        [[("出力", 6.5, True, NAVY), ("｜" + out, 6.5, False, DGRAY)]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.25)
    if i < 4:
        txt(s, cx + cw - 0.4, cy + ch / 2 - 4, 4.5, 8,
            [[("›", 12, True, GOLD)]], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# --- bottom strip ---
sy, sh_, sw_total = 166, 16.5, 275 - 4
w_wide = sw_total * 1.5 / 3.5
w_norm = sw_total / 3.5
strip = [
    (11, w_wide, "横断機能 ｜ 全カテゴリを下支え",
     [("データ検証モジュール", 7.2, True, NAVY),
      ("：独立コンテキストで再取得・算式再計算・レンジチェック（作成者と監査人の分離）。", 7.2, False, BODY),
      ("マルチエージェント討議", 7.2, True, NAVY),
      ("：業界／バイサイド／売り手視点の並列討議で仮説検証。", 7.2, False, BODY)]),
    (11 + w_wide + 2, w_norm, "設計原則",
     [("頭はメイン、手足はサブエージェント", 7.2, True, NAVY),
      ("。判断・洞察・統合はメインが担い、データ取得・並列分析・独立検証を委任。一貫性とスピードを両立。", 7.2, False, BODY)]),
    (11 + w_wide + w_norm + 4, w_norm, "従来との違い",
     [("工程間の判断を", 7.2, False, BODY), ("自動継承", 7.2, True, NAVY),
      ("し組織に蓄積。属人化を排し", 7.2, False, BODY), ("再現可能な組織知", 7.2, True, NAVY),
      ("へ。仲介・FAを代替せず、買い手", 7.2, False, BODY), ("内側の検討", 7.2, True, NAVY),
      ("を支える。", 7.2, False, BODY)]),
]
for bx, bw, lbl, runs in strip:
    rect(s, bx, sy, bw, sh_, fill=WHITE, line=LINE, line_pt=0.5)
    txt(s, bx + 2.4, sy + 1.6, bw - 4.8, sh_ - 3, [
        {"runs": [(lbl, 6.4, True, GOLD)], "space_after": 1},
        {"runs": runs, "line_spacing": 1.25},
    ])

# --- note ---
rect(s, 11, 185, 275, 0.35, fill=LINE)
txt(s, 11, 186.5, 275, 12, [[(
    "※ M&A Architect は意思決定を自動化するものではなく、人間を主担当として判断材料の構造化と再現性を提供する。"
    "10論点はすべて人間が主担当（うちAI支援9論点、「推進検討体制」のみ対象外）。実装フレームワーク：M&A戦略3層構造／目的6類型／成功の10論点／失敗3タイプ ほか。"
    "案件情報はNDAのもと顧客ごとに論理分離・AI学習に不使用。", 6.6, False, DGRAY)]], line_spacing=1.3)

footer(s, "M&A Architect ／ 買い手のM&A機能を、設計する独自AIシステム（Invitation-only Beta・特許出願中）", "1 / 2")

# ================= SLIDE 2: サービス料金体系（仮） =================
s = prs.slides.add_slide(blank)

header(
    s,
    [("M&A Architect サービス料金体系", 15.5, True, NAVY), ("（仮）", 15.5, True, RED),
     ("  ／ Service Pricing (Draft)", 10.5, True, DGRAY)],
    [("全プラン共通の売り＝", 8, False, DGRAY), ("「早くて、質が良い」", 8, True, BODY),
     ("（判断系は必ず代表レビュー）。差別化は", 8, False, DGRAY), ("範囲", 8, True, BODY),
     ("と", 8, False, DGRAY), ("案件数", 8, True, BODY), ("。", 8, False, DGRAY),
     ("※本紙は仮案・最終条件は個別協議", 8, True, RED)],
    "ORIX 事業投資部門 ご提案",
)

lead_band(
    s, 27.5, 13, "PRICING MODEL", "月額 ＋ ポイント制",
    [("2コース（範囲）× 2プラン（同時案件数）＝ 4プラン", 8.4, True, NAVY),
     ("。月額に応じて毎月ポイントを付与し、スキル起動ごとに消化。", 8.4, False, BODY),
     ("専任チーム相当を、解約できる変動費で", 8.4, True, NAVY),
     ("。入口は", 8.4, False, BODY), ("3ヶ月パイロット（▲40%）", 8.4, True, NAVY),
     ("、最低継続6ヶ月（縛りはダウングレード/解約のみ、アップグレードは随時・日割り可）。", 8.4, False, BODY)],
)

LX, LW = 11, 143
RX, RW = 159, 127

# --- 左：プラン表 ---
h2(s, LX, 44.5, "プラン ｜ 2コース × 2プラン = 4プラン")
plan_rows = [
    ("ディレクター・ライト", "評価まで", "1案件", "55", "50 万円", "6ヶ月", False),
    ("ディレクター・スタンダード", "評価まで", "最大3案件", "110", "90 万円", "6ヶ月", False),
    ("パートナー・ライト", "実行まで（DD以降）", "1案件", "100", "110 万円", "6ヶ月", True),
    ("パートナー・スタンダード", "実行まで（DD以降）", "最大3案件", "200", "200 万円", "6ヶ月", True),
]
ty = 50.5
gf = s.shapes.add_table(5, 6, Mm(LX), Mm(ty), Mm(LW), Mm(38))
tbl = gf.table
tbl.first_row = False
tbl.horz_banding = False
widths = [40, 33, 20, 16, 20, 14]
for i, wmm in enumerate(widths):
    tbl.columns[i].width = Mm(wmm)
heads = ["プラン", "コース（範囲）", "同時案件数", "付与pt/月", "月額（税別）", "最低継続"]
for j, htxt in enumerate(heads):
    c = tbl.cell(0, j)
    c.fill.solid(); c.fill.fore_color.rgb = NAVY
    c.margin_left = c.margin_right = Mm(1.6)
    c.margin_top = c.margin_bottom = Mm(0.8)
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = c.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if j >= 2 else PP_ALIGN.LEFT
    r = p.add_run(); r.text = htxt; set_font(r, 7.2, True, WHITE)
for ri, (pn, course, dn, pt_, price, minm, is_p) in enumerate(plan_rows, start=1):
    bg = NAVYSOFT if is_p else (ZEBRA if ri % 2 == 0 else WHITE)
    vals = [pn, course, dn, pt_, price, minm]
    for j, v in enumerate(vals):
        c = tbl.cell(ri, j)
        c.fill.solid(); c.fill.fore_color.rgb = bg
        c.margin_left = c.margin_right = Mm(1.6)
        c.margin_top = c.margin_bottom = Mm(0.6)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if j >= 2 else PP_ALIGN.LEFT
        r = p.add_run(); r.text = v
        if j == 0:
            set_font(r, 8, True, NAVY)
        elif j == 3:
            set_font(r, 8.5, True, GOLD)
        elif j == 4:
            set_font(r, 8.5, True, NAVY)
        else:
            set_font(r, 7.6, False, BODY)

# --- 左：2軸ボックス ---
ay, ah, aw = 91, 15, (LW - 2) / 2
axes = [
    (LX, "軸1 ｜ コース＝範囲",
     [("評価まで（ディレクター）／DD以降の実行まで（パートナー）", 7.3, True, NAVY),
      ("。DD・IC・PMI・戦略策定はパートナーコース限定。", 7.3, False, BODY)]),
    (LX + aw + 2, "軸2 ｜ プラン＝同時案件数",
     [("ライト＝1案件／スタンダード＝最大3案件", 7.3, True, NAVY),
      ("（同時並行）。速さ・質は全プラン共通の土台。", 7.3, False, BODY)]),
]
for ax_, lbl, runs in axes:
    rect(s, ax_, ay, aw, ah, fill=WHITE, line=LINE, line_pt=0.5)
    rect(s, ax_, ay, aw, 0.9, fill=GOLD)
    txt(s, ax_ + 2.4, ay + 1.8, aw - 4.8, ah - 3, [
        {"runs": [(lbl, 6.5, True, GOLD)], "space_after": 1},
        {"runs": runs, "line_spacing": 1.28},
    ])

# --- 左：ポイント消化 ---
h2(s, LX, 109.5, "ポイント消化の仕組み")
my, mh, mw = 115.5, 30, (LW - 2) / 2
rect(s, LX, my, mw, mh, fill=WHITE, line=LINE, line_pt=0.5)
txt(s, LX + 2.4, my + 1.8, mw - 4.8, mh - 3, [
    {"runs": [("消化pt ＝ 委任Tier連動（1起動で消化）", 7, True, NAVY)], "space_after": 1.4},
    {"runs": [("1 pt", 7.2, True, GOLD), ("　完全委任（AIにすべて任せる）", 7.2, False, BODY)], "space_after": 0.8},
    {"runs": [("2 pt", 7.2, True, GOLD), ("　ドラフト委任・代行可", 7.2, False, BODY)], "space_after": 0.8},
    {"runs": [("5 pt", 7.2, True, GOLD), ("　対話・独立セッション実施", 7.2, False, BODY)], "space_after": 0.8},
    {"runs": [("7 pt", 7.2, True, GOLD), ("　相談（1時間あたり）", 7.2, False, BODY)]},
])
rect(s, LX + mw + 2, my, mw, mh, fill=WHITE, line=LINE, line_pt=0.5)
txt(s, LX + mw + 4.4, my + 1.8, mw - 4.8, mh - 3, [
    {"runs": [("品質保証・運用ルール", 7, True, NAVY)], "space_after": 1.4},
    {"runs": [("判断系", 7.1, True, NAVY), ("＝代表がレビュー（品質責任は代表）、", 7.1, False, BODY),
              ("分析系", 7.1, True, NAVY), ("＝AI生成・レビュー任意（バリュエーションは推奨）。", 7.1, False, BODY)],
     "space_after": 1.2, "line_spacing": 1.28},
    {"runs": [("ゲートは", 7.1, False, BODY), ("ハード", 7.1, True, NAVY),
              ("（コース外スキル不可）。繰越なし（月次リセット）。付与ptは上限であり消費保証ではない。", 7.1, False, BODY)],
     "line_spacing": 1.28},
])

# --- 左：フィーの非対称性 ---
fy, fh = 149, 14
rect(s, LX, fy, LW, fh, fill=BGSOFT)
txt(s, LX + 2.6, fy + 1.6, LW - 5.2, fh - 3, [
    {"runs": [("フィーの位置づけ", 6.5, True, GOLD)], "space_after": 1},
    {"runs": [("投資1件（エクイティ数十億円）の1%未満", 7.2, True, NAVY),
              ("＝“保険”と“コールオプション”を同時に買う非対称性。追加購入は割高（枠内消費・上位プランへ誘導）。", 7.2, False, BODY)],
     "line_spacing": 1.28},
])

# --- 右：スキル表 ---
h2(s, RX, 44.5, "スキル別 消化pt・利用可能コース（全25スキル）", w=RW)
skills = [
    ("戦略策定", [("戦略策定エンジン", True, "5", "パートナー"),
                  ("挑戦的問い", False, "1", "両コース")]),
    ("ソーシング", [("ロングリスト", False, "1", "両コース"),
                    ("ショートリスト", False, "1", "両コース"),
                    ("売却可能性調査", False, "1", "両コース"),
                    ("初期アプローチ提案", True, "2", "両コース")]),
    ("案件評価", [("案件評価・総合", True, "5", "両コース"),
                  ("バリュエーション（推奨）", False, "2", "両コース"),
                  ("4シナリオProjection", False, "2", "両コース"),
                  ("Pro Forma", False, "2", "両コース"),
                  ("財務分析", False, "1", "両コース"),
                  ("シナジー分析", False, "1", "両コース"),
                  ("市場分析", False, "1", "両コース"),
                  ("ビジネスモデル／マップ", False, "1", "両コース"),
                  ("競合分析", False, "1", "両コース"),
                  ("経営陣評価", False, "1", "両コース"),
                  ("リスク評価", False, "1", "両コース")]),
    ("論点討議", [("論点討議 α／本編", True, "1", "両コース")]),
    ("DD・交渉\n・意思決定", [("DD計画", True, "2", "パートナー"),
                              ("DD分析・統合", True, "5", "パートナー"),
                              ("SPA交渉支援", True, "2", "パートナー"),
                              ("IC・契約起案", True, "5", "パートナー")]),
    ("PMI", [("PMI・100日・買収後経営", True, "5", "パートナー")]),
    ("相談", [("相談（1時間あたり）", False, "7", "両コース")]),
]
n_rows = 1 + sum(len(v) for _, v in skills)
sty = 50.5
gf2 = s.shapes.add_table(n_rows, 4, Mm(RX), Mm(sty), Mm(RW), Mm(132))
t2 = gf2.table
t2.first_row = False
t2.horz_banding = False
for i, wmm in enumerate([20, 71, 13, 23]):
    t2.columns[i].width = Mm(wmm)
for j, htxt in enumerate(["フェーズ", "スキル", "消化pt", "コース"]):
    c = t2.cell(0, j)
    c.fill.solid(); c.fill.fore_color.rgb = NAVY
    c.margin_left = c.margin_right = Mm(1.4)
    c.margin_top = c.margin_bottom = Mm(0.5)
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = c.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if j >= 2 else PP_ALIGN.LEFT
    r = p.add_run(); r.text = htxt; set_font(r, 6.8, True, WHITE)
ri = 1
for phase, rows in skills:
    start = ri
    for (sk, judge, pt_, course) in rows:
        zebra = ZEBRA if ri % 2 == 0 else WHITE
        cph = t2.cell(ri, 0)
        cph.fill.solid(); cph.fill.fore_color.rgb = NAVYSOFT
        c1 = t2.cell(ri, 1)
        c1.fill.solid(); c1.fill.fore_color.rgb = zebra
        c1.margin_left = c1.margin_right = Mm(1.4)
        c1.margin_top = c1.margin_bottom = Mm(0.3)
        c1.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c1.text_frame.paragraphs[0]
        r = p.add_run(); r.text = sk; set_font(r, 6.9, False, BODY)
        if judge:
            r2 = p.add_run(); r2.text = " ◆判断"; set_font(r2, 6.4, True, RED)
        c2 = t2.cell(ri, 2)
        c2.fill.solid(); c2.fill.fore_color.rgb = zebra
        c2.margin_top = c2.margin_bottom = Mm(0.3)
        c2.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c2.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = pt_; set_font(r, 7, True, NAVY)
        c3 = t2.cell(ri, 3)
        c3.fill.solid(); c3.fill.fore_color.rgb = zebra
        c3.margin_top = c3.margin_bottom = Mm(0.3)
        c3.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c3.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        pcol = NAVY if course == "パートナー" else DGRAY
        r = p.add_run(); r.text = course; set_font(r, 6.4, course == "パートナー", pcol)
        ri += 1
    mc = t2.cell(start, 0)
    if len(rows) > 1:
        mc.merge(t2.cell(start + len(rows) - 1, 0))
    mc.margin_left = mc.margin_right = Mm(1.4)
    mc.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = mc.text_frame.paragraphs[0]
    for li, ln in enumerate(phase.split("\n")):
        pp = p if li == 0 else mc.text_frame.add_paragraph()
        r = pp.add_run(); r.text = ln; set_font(r, 6.8, True, NAVY)

# --- note ---
rect(s, 11, 186, 275, 0.35, fill=LINE)
note2 = ("※ ◆判断＝代表レビュー必須（差別化の核）。無印＝分析（AI生成・レビュー任意）。消化pt＝委任Tier（完全委任1／ドラフト委任2／対話・独立セッション5）連動。"
         "「評価まで」のスキルは両コース、DD以降＋戦略策定はパートナーコース限定。非課金：内部利用・凍結・別運用・個人運用・コンサル専用。")
txt(s, 11, 187.5, 275, 8, [
    {"runs": [(note2, 6.4, False, DGRAY)], "space_after": 0.8, "line_spacing": 1.3},
    {"runs": [("本紙の金額・付与pt・消化ptはすべて仮のものであり、最終条件は個別提案・協議のうえ確定します。", 6.6, True, RED)]},
])

footer(s, "出典：料金体系_v3_20260704.xlsx（仮案・プラン別／スキル別／前提・改定履歴）", "2 / 2")

OUT = "/home/user/Takaaki-Shirako/orix-proposal/MA-Architect_ORIX提案_ビジネスマップ_料金体系仮_20260704.pptx"
prs.save(OUT)
print("saved:", OUT)
