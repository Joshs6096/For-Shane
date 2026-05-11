#!/usr/bin/env python3
"""Render CUR-28 grading recommendations to a PDF that displays inline in browsers
(server natively supports application/pdf). Mirrors the xlsx content."""

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, KeepTogether)
from reportlab.lib.enums import TA_LEFT

OUT = "/Users/josh/.paperclip/instances/default/workspaces/bf3d8eaa-282f-494b-bfab-b370b8561669/agents/sheldon/CUR-28-grading-recommendations.pdf"

styles = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=16, spaceAfter=8,
                    textColor=colors.HexColor("#111827"))
H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13, spaceAfter=6,
                    textColor=colors.HexColor("#1F2937"))
H3 = ParagraphStyle("H3", parent=styles["Heading3"], fontSize=11, spaceAfter=4,
                    textColor=colors.HexColor("#374151"))
SUB = ParagraphStyle("Sub", parent=styles["Italic"], fontSize=9,
                     textColor=colors.HexColor("#6B7280"), spaceAfter=10)
BODY = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=9.5, leading=12,
                      alignment=TA_LEFT, spaceAfter=6)
CELL = ParagraphStyle("Cell", parent=BODY, fontSize=8.5, leading=10.5)
CELL_B = ParagraphStyle("CellB", parent=CELL, fontName="Helvetica-Bold")
CELL_SM = ParagraphStyle("CellSm", parent=CELL, fontSize=8, leading=10)

HEADER_BG = colors.HexColor("#1F2937")
HEADER_FG = colors.white
B1_BG = colors.HexColor("#D1FAE5")
B2_BG = colors.HexColor("#FEF3C7")
B3_BG = colors.HexColor("#FEE2E2")

doc = SimpleDocTemplate(OUT, pagesize=landscape(letter),
                        leftMargin=0.4 * inch, rightMargin=0.4 * inch,
                        topMargin=0.4 * inch, bottomMargin=0.4 * inch,
                        title="CUR-28 Grading Recommendations")
story = []


def p(s, style=CELL):
    return Paragraph(str(s), style)


def header_row_style(table, n_cols, row=0):
    return [
        ("BACKGROUND", (0, row), (n_cols - 1, row), HEADER_BG),
        ("TEXTCOLOR", (0, row), (n_cols - 1, row), HEADER_FG),
        ("FONTNAME", (0, row), (n_cols - 1, row), "Helvetica-Bold"),
        ("FONTSIZE", (0, row), (n_cols - 1, row), 9),
        ("ALIGN", (0, row), (n_cols - 1, row), "CENTER"),
        ("VALIGN", (0, row), (n_cols - 1, row), "MIDDLE"),
        ("BOTTOMPADDING", (0, row), (n_cols - 1, row), 6),
        ("TOPPADDING", (0, row), (n_cols - 1, row), 6),
    ]


def grid_style(n_rows, n_cols):
    return [
        ("GRID", (0, 0), (n_cols - 1, n_rows - 1), 0.4, colors.HexColor("#9CA3AF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]


# ----- Title page -----
story.append(Paragraph("CUR-28 — Grade-vs-Fee Triage on the 188oz Holdings", H1))
story.append(Paragraph(
    "Author: Sheldon (Numismatic Grader & Researcher) · Date: 2026-05-07 · "
    "Spot baseline: $74/oz silver · Source: coin-collector MCP (live, 2026-05-07 21:38 UTC)",
    SUB))

story.append(Paragraph("Recommendation (one line)", H2))
story.append(Paragraph(
    "<b>Submit ONE batch — 13 American Silver Eagle Proofs (mixed years 2006–2022) → PCGS Bulk Modern.</b> "
    "~$285 fees, ~$135 expected net profit (photo-gated). Everything else falls into Bucket 2 (photo-gated) "
    "or Bucket 3 (sell raw / hold raw). At $74 spot, modern raw bullion grading economics are broken — "
    "MS70 premium ($25–40/coin) is too thin to pay back grading friction.", BODY))

story.append(Paragraph("Headline finding — the spot-collapse trap", H2))
story.append(Paragraph(
    "At $74/oz silver, the MS70 premium ($25–40 over melt for common-date ASE) is too small a fraction of total "
    "realized value to justify grading friction. Walked example: 2014 ASE Bullion (qty 15, $26 cost) nets only ~$3/coin "
    "lift from grading vs. selling raw. This kills the v1 hypothesis of submitting modern bullion. "
    "Modern PROOFS still pencil (PR70 vs PR69 spread is $50–90/coin); CLASSIC US in MS condition still pencils (rarity premium).",
    BODY))

# Portfolio scope
scope = [
    ["Catalog entries", "416"],
    ["Total coins (set-expanded)", "2,574"],
    ["Total cost paid", "$17,877.57"],
    ["Total high-value (catalog)", "$77,866.07"],
    ["Total low-value (catalog)", "$45,230.91"],
    ["Unrealized profit", "$59,988.50  (+335.55%)"],
    ["Silver oz on hand", "188.033 oz / 339 entries / 2,365 coins"],
    ["Silver spot baseline", "$74/oz (CUR-29 W19; MCP spot field is null)"],
]
scope_data = [[p("Metric", CELL_B), p("Value", CELL_B)]] + [[p(k, CELL_B), p(v)] for k, v in scope]
t = Table(scope_data, colWidths=[2.5 * inch, 6 * inch])
t.setStyle(TableStyle(header_row_style(t, 2) + grid_style(len(scope_data), 2)))
story.append(Paragraph("Portfolio scope (live MCP)", H3))
story.append(t)
story.append(Spacer(1, 0.18 * inch))

# Bucket roll-up
story.append(Paragraph("Bucket roll-up", H3))
roll_hdr = ["Bucket", "Action", "Coins / Clusters", "Est. fees", "Expected net", "Confidence"]
roll_rows = [
    [p("<b>Bucket 1</b><br/>Submit Now", CELL),
     p("Submit one batch (Batch A: 13 ASE Proofs → PCGS Bulk Modern)", CELL),
     p("13 coins / 1 batch", CELL), p("~$285", CELL),
     p("~$135 (base case)", CELL), p("Medium (photo-gated)", CELL)],
    [p("<b>Bucket 2</b><br/>Photos required", CELL),
     p("Macro-photo audit by Elvis; Sheldon writes v3 go/no-go after", CELL),
     p("8 candidate clusters (~30+ coins)", CELL), p("Variable", CELL),
     p("Variable; high stakes on 1976-S Ike crackout", CELL),
     p("Medium (photo-dependent)", CELL)],
    [p("<b>Bucket 3</b><br/>Do not submit", CELL),
     p("Sell raw (modern bullion) / hold raw (junk silver) / sell intact (sets) / hold graded (existing slabs)", CELL),
     p("~80–100 modern bullion + class-wide junk silver / sets / slabs", CELL),
     p("$0", CELL),
     p("$0 grading lift; raw sale captures spot premium", CELL),
     p("High", CELL)],
]
data = [[p(h, CELL_B) for h in roll_hdr]] + roll_rows
t = Table(data, colWidths=[1.1 * inch, 2.6 * inch, 2 * inch, 0.9 * inch, 1.9 * inch, 1.5 * inch])
ts = TableStyle(header_row_style(t, len(roll_hdr)) + grid_style(len(data), len(roll_hdr)))
ts.add("BACKGROUND", (0, 1), (-1, 1), B1_BG)
ts.add("BACKGROUND", (0, 2), (-1, 2), B2_BG)
ts.add("BACKGROUND", (0, 3), (-1, 3), B3_BG)
t.setStyle(ts)
story.append(t)
story.append(Spacer(1, 0.18 * inch))

# Decisions
story.append(Paragraph("Decisions the board needs to make", H3))
dec_hdr = ["#", "Decision", "Default if no answer", "Owner"]
dec = [
    ("1", "Approve Batch A (13 ASE Proofs → PCGS Bulk Modern, ~$285 fees, ~$135 net) pending photo audit?",
     "Defer to v3", "Board → Elvis pulls photos"),
    ("2", "Crack one Bicentennial Silver Proof Set (1 of 2 sets, hedged) for the 1976-S Ike DCAM PR70 chase?",
     "Hold both intact", "Board → Sheldon executes if greenlit"),
    ("3", "Disclose contents of 2024 Mint Limited Edition Silver Proof Set so we can run crackout math?",
     "Hold intact", "Board / Elvis"),
    ("4", "Reconcile inventory `high_value_each` against eBay 30-day sold medians (catalog overstates real market on common-date bullion)?",
     "Sheldon flags but no action", "CEO"),
]
data = [[p(h, CELL_B) for h in dec_hdr]]
for n, q, d, o in dec:
    data.append([p(n, CELL), p(q, CELL), p(d, CELL), p(o, CELL)])
t = Table(data, colWidths=[0.3 * inch, 5.4 * inch, 2 * inch, 2.3 * inch])
t.setStyle(TableStyle(header_row_style(t, len(dec_hdr)) + grid_style(len(data), len(dec_hdr))))
story.append(t)
story.append(Spacer(1, 0.18 * inch))

# Gates
gates = [
    ("Gate 1 — Raw cost", "≤ $65/coin for modern bullion; ≤ 50% of MS65 retail for classic US"),
    ("Gate 2 — Service routing", "PCGS for modern bullion + modern proofs; NGC for classic US"),
    ("Gate 3 — Batch size", "n ≥ 5 of same series/window on a single submission form"),
    ("Gate 4 (new this rev)", "Modern bullion gate effectively closed at $74 spot — MS70 premium too thin; PROOFS only"),
    ("Express gate", "Never NGC Express on modern bullion (CUR-27 v2 lesson)"),
]
story.append(Paragraph("Economic gates (must clear all 3 to submit)", H3))
data = [[p("Gate", CELL_B), p("Threshold", CELL_B)]]
for k, v in gates:
    data.append([p(k, CELL_B), p(v, CELL)])
t = Table(data, colWidths=[2.2 * inch, 7.8 * inch])
t.setStyle(TableStyle(header_row_style(t, 2) + grid_style(len(data), 2)))
story.append(t)
story.append(PageBreak())

# ============================================================
# Bucket 1 — Submit Now
# ============================================================
story.append(Paragraph("Bucket 1 — Submit Now (Batch A)", H1))
story.append(Paragraph(
    "Single batch: 13 American Silver Eagle Proofs (mixed years 2006–2022) → PCGS Bulk Modern, single shipment. "
    "Skip First Strike / Advance Release labels (window long-closed for these years). "
    "Photo audit required before sub form is finalized.", BODY))

b1_hdr = ["ID", "Year", "Coin", "Qty", "Cost/ea", "Cat. high $",
          "Comp source", "PR70 retail", "PR69 fallback", "Lift/ea (PR70)", "Notes / justification"]
b1_rows = [
    (162, "2021", "American Eagle Silver Proof (Type 2)", 4, "$86.00", "$219.99",
     "eBay sold 2021-W PR70 DCAM, 2026 Q1–Q2",
     "$130–180", "$80–95", "+$30–55",
     "Type 2 reverse die premium real; PR69 close to raw cost"),
    (116, "2006", "American Eagle Silver Proof Bullion", 3, "$86.00", "$194.99",
     "PCGS CoinFacts + eBay sold 2026",
     "$150–200", "$90–110", "+$40–70",
     "20th-anniversary year — premium label window for FS / Advance Release strikes"),
    (148, "2014", "American Eagle Silver Proof Coin", 3, "$86.00", "$165.00",
     "eBay sold 2026 30-day median",
     "$130–160", "$90–110", "+$25–45",
     "Mid-vintage; 70-rate ~70% historically"),
    (155, "2020", "American Eagle Silver Proof", 1, "$86.10", "$399.00",
     "eBay sold 2020-W PR70 DCAM 2026",
     "$200–300 (W variants higher)", "$100–130", "+$60–120",
     "$399 catalog suggests 2020-W or Congratulations Set example. PHOTO AUDIT CRITICAL — "
     "large lift on FS variant; if standard issue, lift is in base-case band."),
    (167, "2022", "American Eagle Silver Proof", 1, "$86.10", "$126.00",
     "eBay sold 2026 30-day median",
     "$130–150", "$90–100", "+$15–25",
     "Recent year, modest lift; batches well"),
    (123, "2007", "American Eagle Silver Proof Coin", 1, "$86.10", "$140.00",
     "eBay sold 2026 30-day median",
     "$130–160", "$85–110", "+$20–35",
     "Mid-vintage; routine candidate; batches well"),
]
data = [[p(h, CELL_B) for h in b1_hdr]]
for r in b1_rows:
    data.append([p(str(c), CELL_SM) for c in r])
col_widths = [0.35, 0.45, 1.7, 0.3, 0.55, 0.65, 1.5, 0.95, 0.85, 0.85, 2.45]
t = Table(data, colWidths=[w * inch for w in col_widths], repeatRows=1)
ts = TableStyle(header_row_style(t, len(b1_hdr)) + grid_style(len(data), len(b1_hdr)))
ts.add("BACKGROUND", (0, 1), (-1, -1), B1_BG)
t.setStyle(ts)
story.append(t)
story.append(Spacer(1, 0.15 * inch))

# Aggregate economics
story.append(Paragraph("Aggregate economics (base case, 75% PR70 conversion)", H3))
econ = [
    ("Coins in batch", "13"),
    ("Total raw cost (already paid)", "$1,118.40"),
    ("PCGS Bulk Modern fee", "13 × $20 = $260"),
    ("Return shipping", "~$25"),
    ("Total fees", "~$285"),
    ("Expected gross lift @ 75% PR70 rate", "~9.75 coins × $50 avg lift = ~$485"),
    ("Less: 3.25 coins at PR69 (no upgrade)", "~$65 absorbed fees"),
    ("Net before / after fees", "~$420 net before fees / ~$135 net AFTER fees"),
    ("Realized portfolio value (post-grade)", "~$1,600–2,000 vs ~$1,100 raw"),
    ("Swing factor (upside)", "ID 155 (2020 ASE Proof) — if FS or 2020-W variant, batch lift jumps materially"),
]
data = [[p("Line item", CELL_B), p("Value", CELL_B)]]
for k, v in econ:
    data.append([p(k, CELL_B), p(v, CELL)])
t = Table(data, colWidths=[3 * inch, 7 * inch])
t.setStyle(TableStyle(header_row_style(t, 2) + grid_style(len(data), 2)))
story.append(t)
story.append(Spacer(1, 0.12 * inch))

# Pre-flight notes
story.append(Paragraph("Pre-flight and service notes", H3))
notes = [
    ("Recommended grading service", "PCGS, modern bulk submission tier, single shipment"),
    ("Do NOT split", "Across NGC and PCGS — defeats batch economics (Gate 3 violation)"),
    ("Skip", "First Strike / Advance Release label add-ons (window long-closed for these years)"),
    ("Elvis pre-flight", "Confirm physical possession of all 13 coins; pull macro-photos of every obverse + reverse for Sheldon's per-coin call before sub form is filled out"),
    ("Slab status", "All 13 = `confirmed_raw` (held in OGP from Mint)"),
    ("Mint mark detection", "ID 155 requires explicit mint-mark photo (W vs S) — drives variant call"),
]
data = [[p("Item", CELL_B), p("Note", CELL_B)]]
for k, v in notes:
    data.append([p(k, CELL_B), p(v, CELL)])
t = Table(data, colWidths=[2.5 * inch, 7.5 * inch])
t.setStyle(TableStyle(header_row_style(t, 2) + grid_style(len(data), 2)))
story.append(t)
story.append(PageBreak())

# ============================================================
# Bucket 2 — Photos required
# ============================================================
story.append(Paragraph("Bucket 2 — Submit Only If Photos Confirm", H1))
story.append(Paragraph(
    "Ranked by potential portfolio impact (highest stakes first). Each cluster gates on Elvis macro-photo audit before Sheldon writes v3 go/no-go.",
    BODY))

b2_hdr = ["#", "ID(s)", "Coin / Set", "Qty", "Cost/ea", "Cat. high",
          "Stakes", "Photo audit ask", "Recommendation if photos clear (service)"]

b2_rows = [
    ("B1", "172", "1776–1976 Bicentennial Silver Proof Set", "2 sets / 10 coins",
     "$45/set", "$190/set",
     "Very high — 1976-S Silver Ike DCAM PR70 = $649–730 recent eBay; Heritage record $6,900. Set intact = $190.",
     "Macro photos of the Ike's reverse: cameo contrast and any field marks. PR69→PR70 cliff is the most binary call in the entire portfolio.",
     "If photos clear: submit ONE Ike (1 of 2 sets) to PCGS Bulk Modern; keep second set sealed as hedge. PCGS."),
    ("B2", "169", "2024 Mint Limited Edition Silver Proof Set", "1 set / 5 coins",
     "$180", "$399.99",
     "Medium-high — set value $399.99; crackout depends on contents (board contents disclosure pending).",
     "Photo + contents disclosure required (which 5 coins).",
     "Default if no info: keep set intact. If contents include high-PR70-spread coins, submit individually to PCGS."),
    ("B3", "5", "1859 Indian Head Cent (Type 1 CN)", "17 coins",
     "$20", "$159.95",
     "Medium-high — first-year-of-issue. 1859/1858 DDO variety = $1,000+ in MS. Risk: ~40–50% body-bag rate on raw lots.",
     "All 17 coins, both sides: cleaning indicators + variety check for 1859/1858 DDO.",
     "If photos clear: submit only 8–10 cleanest at NGC Standard tier. Predicted ~$40 net + variety lottery upside. NGC."),
    ("B4", "228", "1982 Libertad", "3 coins",
     "$0 (gifted)", "$226.10",
     "Medium — first-year Libertad. NGC MS65 ~$300; MS68 ~$2,500; MS69 $8,000+. GATE 3 CONFLICT (n<3).",
     "Luster, contact marks, edge (1982 known for poor strike on eagle's wing).",
     "If MS66+ surfaces: submit single coin at NGC Standard ($40); rest raw. Don't blanket-submit. NGC."),
    ("B5", "232", "1996 ASE PCS PR69 DCAM (slabbed)", "1 coin",
     "$0 (gifted)", "$345",
     "Low-medium — PCS is low-tier; PCGS PR70 = $400–700 for 1996 (tougher year). Crossover at PCGS = $25.",
     "Existing slab photo to assess upgrade likelihood (haze, milk spots, hairlines).",
     "If pristine: PCGS crossover as one-off, add-on to Batch A. Risk: PR69 confirmed = ~$0 lift. PCGS."),
    ("B6", "416", "2026 Australian Swan Proof", "2 coins",
     "$97.50", "$275",
     "Medium — recently acquired; follow-on to CUR-2's NGC bullion-Swan submission.",
     "DEFER until CUR-2 returns 2026-06-16 → 06-30 to calibrate Perth-mint 70-rate.",
     "If 70-rate ≥70%: batch with other 2026 Perth proof bullion. If ≤70%: sell raw at premium retail. PCGS."),
    ("B7", "—", "WLH common-date short-set", "11 coins",
     "$75–140", "varies",
     "Low-medium — common dates raw at $75–140 most likely XF/AU. Only MS-grade pays back.",
     "Liberty's head-detail, skirt-line completeness, luster cartwheels.",
     "Likely 0–2 of 11 grade MS64+. Submit at most 2–3 if photos justify. NGC."),
    ("B8", "243, 249, 301", "1891 Morgan trio", "3 coins",
     "varies", "varies",
     "Medium IF a 1891-CC; otherwise low. 1891-CC NGC MS62 ~$700; MS63 ~$1,500.",
     "Mintmark check on reverse below eagle (CC, O, S, none).",
     "If any 1891-CC: submit at NGC Standard. Otherwise hold/sell raw at $200–275. NGC."),
]

data = [[p(h, CELL_B) for h in b2_hdr]]
for r in b2_rows:
    data.append([p(str(c), CELL_SM) for c in r])
col_widths = [0.3, 0.6, 1.7, 0.7, 0.5, 0.55, 1.6, 1.7, 2.65]
t = Table(data, colWidths=[w * inch for w in col_widths], repeatRows=1)
ts = TableStyle(header_row_style(t, len(b2_hdr)) + grid_style(len(data), len(b2_hdr)))
ts.add("BACKGROUND", (0, 1), (-1, -1), B2_BG)
t.setStyle(ts)
story.append(t)
story.append(PageBreak())

# ============================================================
# Bucket 3 — Don't submit
# ============================================================
story.append(Paragraph("Bucket 3 — Do Not Submit (gate-fail or worth more raw)", H1))
story.append(Paragraph(
    "Class-level standing rules from v1, accepted as policy by CEO 2026-05-06, applied to live MCP inventory.",
    BODY))

b3_hdr = ["Class", "Examples / Inv IDs", "Why no grading", "Action", "Trigger to revisit"]
b3_rows = [
    ("Modern silver bullion 2024–2026 acquired at $78–110/ea",
     "Koalas qty 6; Britannias qty 7; Maples qty 7+; Pandas qty 5; Queens Beasts qty 2; Libertads 2022–23 qty 6; raw ASE bullion outside 2014 cluster (~80–100 coins)",
     "SPOT-COLLAPSE TRAP. At $74 spot, MS70 premium ($25–40/coin) too thin to pay back grading friction. Cost basis already at/above raw market.",
     "SELL RAW at premium retail ($110–220/ea by series).",
     "Silver retraces below $50/oz."),
    ("2014 ASE Bullion (representative case)",
     "ID 146, qty 15, $26/ea",
     "Most-promising case in inventory; even at sub-melt cost, lift = $3/coin vs $38 raw profit. Headline data point for spot-collapse trap.",
     "SELL RAW. Add to CUR-10 SELL bucket at $80–90/ea.",
     "Spot moves >±10% from $74."),
    ("Circulated common-date silver dollars",
     "Morgans 1889–1926 + Peace dollars at $66 cost VF/EF, qty ~15 across IDs 6, 7, 8, 10, 12, 13, 14, 17, 19, 22",
     "Circulated common-date Morgans at VF/EF have flat slab demand. Slabbing fee exceeds grade lift.",
     "Sell raw at $115–250/ea per CUR-10.",
     "If any 1893-S, 1895, or 1889-CC surfaces (none seen)."),
    ("Junk silver bags (sub-MS65 raw)",
     "Roosevelt dimes (382, qty 99); Mercury dimes (381, 414, qty 103); Silver Quarters (383, 415, qty 85); Franklin Halves (385, qty 17); Kennedy Halves 1965–70 (174, qty 40); 1964 Kennedys (28, 380, 384, qty 10)",
     "Sells at/near spot; no rarity premium for circulated common-date silver.",
     "HOLD per CUR-11 spot-trigger ≥$95/oz.",
     "Spot ≥ $95/oz."),
    ("Already-graded coins (MS70/PR70 + PCS PR69 except crossover)",
     "IDs 333, 335, 231, 238, 268, 280, 279, 343, 354, 366, 413, 418",
     "Already at top retail tier. Re-grading risk > upside.",
     "HOLD or SELL graded.",
     "Specific tier-jump cases (CAC sticker on PCGS coin lacking one) — case-by-case."),
    ("Pre-2024 standard-issue Silver Proof / Mint Sets",
     "IDs 45, 60, 72, 82, 85, 87, 89, 103, 106, 107, 111, 112, 115, 120, 124, 143, 149, 150, 152, 154, 156, 160, 161, 164, 166, 364, others",
     "Crackout math fails on standard issues at current MS70 premiums; intact set holds packaging premium.",
     "HOLD INTACT, sell as sets.",
     "Specific high-stakes year (e.g., 1976-S Bicentennial in B1)."),
    ("Foreign single-coin oddities",
     "USSR Olympic Basketball; lone Filipina; Liberian Platinum; single Isle of Man; etc. (n<5 each)",
     "Gate 3 fail (batch <5). No grading lift.",
     "Sell raw to specialist buyers.",
     "Bundling opportunity into curated foreign lot (unlikely)."),
    ("NORFEDs / private mint silver rounds and bars",
     "Silver Round/Bar category, 19 entries / 40 coins",
     "No graded market — privately minted rounds/bars don't slab competitively.",
     "Sell raw at melt + small premium.",
     "Spot moves materially."),
    ("Cleaned, damaged, or low-grade cents",
     "ID 419 (1904 IHC); ID 420 (1905 IHC) at $2.50; ID 4 (1848 Braided Hair); ID 387 (1968 IGC slab MS67 RD)",
     "Low cost-basis cents fail Gate 1; IGC slab is third-tier; crossover only worthwhile if confirmed superior surfaces.",
     "Sell raw / hold IGC slabs intact unless photos suggest crossover.",
     "Photo audit on ID 387 confirms surfaces support PCGS upgrade."),
    ("Active SELL listings — already on the market raw",
     "CUR-3 through CUR-8 (6 active listings)",
     "Already in motion — no grading injection at this stage.",
     "Continue listed.",
     "Listing expires unsold + relisting strategy review."),
]

data = [[p(h, CELL_B) for h in b3_hdr]]
for r in b3_rows:
    data.append([p(str(c), CELL_SM) for c in r])
col_widths = [1.85, 2.55, 2.55, 1.55, 1.5]
t = Table(data, colWidths=[w * inch for w in col_widths], repeatRows=1)
ts = TableStyle(header_row_style(t, len(b3_hdr)) + grid_style(len(data), len(b3_hdr)))
ts.add("BACKGROUND", (0, 1), (-1, -1), B3_BG)
t.setStyle(ts)
story.append(t)
story.append(PageBreak())

# ============================================================
# Comp sources
# ============================================================
story.append(Paragraph("Comp sources cited", H1))
story.append(Paragraph(
    "Pricing decays — every comp carries a date and source. Refresh before the next submission greenlight.",
    BODY))
cs = [
    ("PCGS CoinFacts (pcgs.com/coinfacts)", "Series retail bands and 70-rate priors",
     "1976-S Silver Ike, 1859 IHC, 1982 Libertad, modern ASE Proof series. Authoritative for PCGS-graded comparables."),
    ("eBay sold listings (ebay.com, 30-day sold filter)", "Realized-price ground truth for modern bullion + ASE Proof",
     "2024–2026 modern ASE bullion ($44–55 MS70); ASE Proof ($130–180 PR70 DCAM); Bicentennial Silver Ike DCAM PR70 ($649–730 recent). Use 30-day median."),
    ("Heritage Auctions realized (ha.com)", "High-end / record comps",
     "1976-S Silver Ike DCAM PR70 record $6,900 anchor."),
    ("GreatCollections (greatcollections.com)", "Mid-market US auction comps",
     "Cross-check on classic US series."),
    ("NGC and PCGS fee schedules", "Submission cost calc",
     "NGC Bulk Modern $17; PCGS Bulk Modern $20. Always include return shipping ~$25 in batch math."),
    ("Internal Sheldon resources", "Durable lessons applied this pass",
     "grading-cliff-1976S-silver-ike; ngc-vs-pcgs-modern-bullion; modern-bullion-grading-economics; silver-spot-regime; spot-collapse-bullion-grading (new this revision)."),
    ("CUR-29 W19 dashboard", "Spot baseline",
     "$74/oz silver spot used throughout. MCP `spot_prices` field is null this run."),
    ("coin-collector MCP", "Inventory ground truth",
     "All cost basis, qty, catalog high-value, and inventory tallies in this report pulled live from MCP 2026-05-07 21:38 UTC. Reconciles to W19 dashboard."),
]
data = [[p(h, CELL_B) for h in ["Source", "Used for", "Notes"]]]
for s, u, n in cs:
    data.append([p(s, CELL), p(u, CELL), p(n, CELL)])
t = Table(data, colWidths=[3 * inch, 2.6 * inch, 4.4 * inch])
t.setStyle(TableStyle(header_row_style(t, 3) + grid_style(len(data), 3)))
story.append(t)

doc.build(story)
print(f"WROTE {OUT}")
