"""Build the Scaled Cognition job talk PPTX -- v6.

v6 changes over v5:
  * Act 3 restructured to land the Recursive Self-Improving Agent (RSIA)
    vision. Two new main slides:
    - s15 RSIA vision: three-pillar diagram closes into a meta-loop.
    - s16 POC: the intent-classifier auto-tuning pipeline (7 steps)
      re-colored into DATA / REWARD / POST-TRAIN pillars. Proves the
      RSIA loop already ships at 0.8B.
  * s17 Generalization reframed as "each row = seed for its own RSIA".
  * s18 Anchor reprise adds an outer meta-loop overlay.
  * s19 APT-2 bets reframed as four legs of the RSIA flywheel.
  * s16 Lessons moved to backup (b8).

v5 baseline (kept):
  * Hero s2 = CoT-vs-RL nostalgic video pair.
  * s12/s13/s14 = three real-world results
    (Conversation, Agentic Planning, Intent Classifier).
  * Unified loop anchor. Unicode math in slide text, LaTeX in add_math.

Run:  python3 job_talk_build_v6.py
Output: job_talk_v6.pptx  (21 core slides + 12 backup, 16:9)
"""
from pathlib import Path
from io import BytesIO
import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor

HERE   = Path(__file__).parent
ASSETS = HERE / "assets"
OUT    = HERE / "job_talk_v7.pptx"

# ---------- Design tokens (matches RAD_v4) ----------
INK     = RGBColor(0x14, 0x1B, 0x2D)
MUTED   = RGBColor(0x5A, 0x63, 0x78)
ACCENT  = RGBColor(0xC8, 0x3E, 0x2C)   # brick red
ACCENT2 = RGBColor(0x1F, 0x6F, 0xB2)   # steel blue
ACCENT3 = RGBColor(0x2E, 0x7D, 0x32)   # forest green
BG      = RGBColor(0xF8, 0xF6, 0xF1)
RULE    = RGBColor(0xDD, 0xD7, 0xC8)
CARD    = RGBColor(0xFF, 0xFF, 0xFF)
TINT_A  = RGBColor(0xF7, 0xEC, 0xE9)
TINT_B  = RGBColor(0xE7, 0xEE, 0xF6)
TINT_C  = RGBColor(0xE7, 0xF0, 0xE7)
TINT_N  = RGBColor(0xF0, 0xEC, 0xE0)

def rgb_tuple(c):
    return (c[0] / 255.0, c[1] / 255.0, c[2] / 255.0)

FONT_HEAD = "Inter"
FONT_BODY = "Inter"
FONT_MONO = "JetBrains Mono"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height


# ==================== HELPERS ====================

def add_blank():
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.line.fill.background()
    bg.fill.solid(); bg.fill.fore_color.rgb = BG
    bg.shadow.inherit = False
    return slide


def _style_run(run, size, bold, color, font, italic=False):
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color


def add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=INK,
             font=FONT_BODY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf  = box.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if isinstance(text, list):
        first, rest = text[0], text[1:]
    else:
        first, rest = text, []
    run = p.add_run(); run.text = first
    _style_run(run, size, bold, color, font, italic)
    for line in rest:
        p = tf.add_paragraph(); p.alignment = align
        r = p.add_run(); r.text = line
        _style_run(r, size, bold, color, font, italic)
    return box


def add_rect(slide, x, y, w, h, fill=CARD, line=None):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None:
        r.line.fill.background()
    else:
        r.line.color.rgb = line
        r.line.width = Pt(0.75)
    r.shadow.inherit = False
    return r


def add_ellipse(slide, x, y, w, h, fill=CARD, line=None, line_w=0.75):
    r = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None:
        r.line.fill.background()
    else:
        r.line.color.rgb = line
        r.line.width = Pt(line_w)
    r.shadow.inherit = False
    return r


def add_hline(slide, x, y, w, color=RULE, weight=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, y, x + w, y)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def add_arrow(slide, x1, y1, x2, y2, color=MUTED, weight=1.25):
    # Coerce to int EMU -- python-pptx writes coords verbatim into the
    # <a:off>/<a:ext> attributes and PowerPoint's ST_Coordinate parser
    # rejects float strings like "2331720.0", dropping the shape to (0,0).
    x1, y1, x2, y2 = Emu(int(x1)), Emu(int(y1)), Emu(int(x2)), Emu(int(y2))
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    # arrow head
    ln.line._get_or_add_ln()
    from pptx.oxml.ns import qn
    ln_elem = ln.line._get_or_add_ln()
    tail = ln_elem.find(qn("a:tailEnd"))
    if tail is None:
        from lxml import etree
        tail = etree.SubElement(ln_elem, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("h", "med")
    return ln


def add_page_header(slide, section, title):
    add_text(slide, Inches(0.6), Inches(0.35), Inches(12), Inches(0.3),
             section.upper(), size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_text(slide, Inches(0.6), Inches(0.65), Inches(12.1), Inches(0.7),
             title, size=26, bold=True, color=INK, font=FONT_HEAD)
    add_hline(slide, Inches(0.6), Inches(1.45), Inches(12.1))


def add_footer(slide, page, total, section_label=""):
    add_text(slide, Inches(0.6), Inches(7.05), Inches(8), Inches(0.3),
             section_label, size=9, color=MUTED, font=FONT_BODY)
    add_text(slide, Inches(12.0), Inches(7.05), Inches(1.2), Inches(0.3),
             f"{page} / {total}", size=9, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.RIGHT)


def add_bullet(slide, x, y, w, text, *, marker=ACCENT, size=14, bold=False,
               color=INK, font=FONT_BODY):
    add_rect(slide, x, y + Inches(0.18), Inches(0.09), Inches(0.09),
             fill=marker)
    add_text(slide, x + Inches(0.22), y, w - Inches(0.3), Inches(0.5),
             text, size=size, bold=bold, color=color, font=font)


def add_notes(slide, notes):
    slide.notes_slide.notes_text_frame.text = notes


# ---------- MATH via matplotlib mathtext ----------

def add_math(slide, latex, x, y, width_in, *, fontsize=32,
             color=INK, bg=BG, height_ratio=0.18):
    """Render LaTeX (mathtext subset) to PNG and place on slide.

    Returns the picture shape.
    """
    fig = plt.figure(figsize=(width_in, width_in * height_ratio))
    fig.patch.set_facecolor(rgb_tuple(bg))
    fig.text(0.5, 0.5, f"${latex}$",
             fontsize=fontsize, ha="center", va="center",
             color=rgb_tuple(color))
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=220, bbox_inches="tight",
                pad_inches=0.08, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return slide.shapes.add_picture(buf, x, y, width=Inches(width_in))


# ---------- VIDEO ----------

def add_video(slide, mp4_path, poster_path, x, y, w, h):
    return slide.shapes.add_movie(
        str(mp4_path), x, y, w, h,
        poster_frame_image=str(poster_path),
        mime_type="video/mp4",
    )


# ==================== SLIDE COUNTS ====================
TOTAL = 21            # main deck
TOTAL_ALL = 33        # main + backup (b1..b7 + bV1..bV3 + b_lessons + divider)


# =========================================================
# ACT 1 · VISION (s1..s5)
# =========================================================

def s1_title():
    s = add_blank()
    add_rect(s, 0, 0, Inches(0.35), SH, fill=ACCENT)

    add_text(s, Inches(0.9), Inches(1.0), Inches(11), Inches(0.5),
             "SCALED COGNITION  ·  JOB TALK", size=12, bold=True,
             color=ACCENT, font=FONT_HEAD)

    add_text(s, Inches(0.9), Inches(1.7), Inches(12), Inches(2.0),
             "Post-Training the Agentic Loop",
             size=48, bold=True, color=INK, font=FONT_HEAD)

    add_text(s, Inches(0.9), Inches(2.9), Inches(12), Inches(1.2),
             "A conversation is a trajectory through the loop.  "
             "One recipe post-trains both.",
             size=22, color=MUTED, font=FONT_HEAD, italic=True)

    add_hline(s, Inches(0.9), Inches(4.5), Inches(3))

    add_text(s, Inches(0.9), Inches(4.75), Inches(10), Inches(0.5),
             "Subhojyoti Mukherjee",
             size=22, bold=True, color=INK, font=FONT_HEAD)
    add_text(s, Inches(0.9), Inches(5.2), Inches(10), Inches(0.4),
             "Research Scientist  ·  Adobe Research",
             size=16, color=MUTED, font=FONT_BODY)
    add_text(s, Inches(0.9), Inches(5.55), Inches(10), Inches(0.4),
             "September 10, 2026", size=14, color=MUTED, font=FONT_BODY)

    add_notes(s,
        "Frame the shape: 'This talk is a vision plus a deep-dive. I post-train "
        "agents. In the next slide you'll watch our Express agent run one full "
        "loop end-to-end. The rest of the talk is one recipe -- offline RL as "
        "reward-weighted SFT -- that post-trains that loop, applied to two "
        "instantiations: a DC tutoring session from NeurIPS 2025 and "
        "the Express creative agent shipping today.' Budget: 45 min + Q&A, "
        "~2 min per main slide.")
    return s


def _cot_vs_rl_slide(prompt_text, cot_mp4, cot_poster, rl_mp4, rl_poster,
                     scores, takeaway, page, footer_label,
                     section_head="Vision  ·  Watch the loop run",
                     title="Same prompt, two policies  ·  RL ships the intent"):
    """Shared layout for CoT vs RL side-by-side video comparisons.

    scores: list of (label, cot_val, rl_val) tuples for the bottom strip.
    """
    s = add_blank()
    add_page_header(s, section_head, title)

    # Prompt banner
    add_rect(s, Inches(0.6), Inches(1.65), Inches(12.1), Inches(0.55),
             fill=TINT_N)
    add_text(s, Inches(0.85), Inches(1.7), Inches(1.5), Inches(0.45),
             "USER", size=11, bold=True, color=ACCENT, font=FONT_HEAD,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(2.1), Inches(1.7), Inches(10.4), Inches(0.45),
             f"\u201c{prompt_text}\u201d",
             size=16, bold=True, color=INK, font=FONT_HEAD,
             anchor=MSO_ANCHOR.MIDDLE)

    # Two video tiles, tile chrome + label + video
    tile_w, tile_h = Inches(5.95), Inches(3.4)
    left_x, right_x = Inches(0.6), Inches(6.75)
    tile_y = Inches(2.4)

    # LEFT: CoT baseline
    add_rect(s, left_x, tile_y - Inches(0.32), tile_w, Inches(0.32),
             fill=MUTED)
    add_text(s, left_x + Inches(0.15), tile_y - Inches(0.3), tile_w, Inches(0.3),
             "CoT-based  ·  baseline",
             size=12, bold=True, color=CARD, font=FONT_HEAD)
    add_video(s, cot_mp4, cot_poster, left_x, tile_y, tile_w, tile_h)

    # RIGHT: RL finetuned
    add_rect(s, right_x, tile_y - Inches(0.32), tile_w, Inches(0.32),
             fill=ACCENT)
    add_text(s, right_x + Inches(0.15), tile_y - Inches(0.3),
             tile_w, Inches(0.3),
             "RL finetuning  ·  Refit / Swift",
             size=12, bold=True, color=CARD, font=FONT_HEAD)
    add_video(s, rl_mp4, rl_poster, right_x, tile_y, tile_w, tile_h)

    # Judge-score strip (contrast argument)
    strip_y = tile_y + tile_h + Inches(0.25)
    add_rect(s, Inches(0.6), strip_y, Inches(12.1), Inches(0.85),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), strip_y + Inches(0.08),
             Inches(2.5), Inches(0.3),
             "JUDGE SCORES", size=10, bold=True, color=MUTED,
             font=FONT_HEAD)

    col_w = 12.1 / (len(scores) + 1)
    for i, (label, cot_v, rl_v) in enumerate(scores):
        cx = Inches(0.6 + col_w * (i + 1))
        add_text(s, cx, strip_y + Inches(0.05),
                 Inches(col_w - 0.1), Inches(0.3),
                 label, size=10, bold=True, color=MUTED, font=FONT_MONO,
                 align=PP_ALIGN.CENTER)
        add_text(s, cx, strip_y + Inches(0.32),
                 Inches(col_w - 0.1), Inches(0.32),
                 cot_v, size=13, color=MUTED, font=FONT_MONO,
                 align=PP_ALIGN.CENTER)
        add_text(s, cx, strip_y + Inches(0.55),
                 Inches(col_w - 0.1), Inches(0.32),
                 rl_v, size=14, bold=True, color=ACCENT, font=FONT_MONO,
                 align=PP_ALIGN.CENTER)

    # Takeaway
    add_text(s, Inches(0.6), Inches(6.95),
             Inches(12.1), Inches(0.3),
             takeaway,
             size=12, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, page, TOTAL if page <= TOTAL else TOTAL_ALL, footer_label)
    return s


def s2_watch_the_loop():
    """Hero: CoT vs RL comparison on 'Make the design more nostalgic'."""
    s = _cot_vs_rl_slide(
        prompt_text="Make the design more nostalgic",
        cot_mp4=ASSETS / "nostalgic_cot.mp4",
        cot_poster=ASSETS / "poster_nostalgic_cot.png",
        rl_mp4=ASSETS / "nostalgic_rl.mp4",
        rl_poster=ASSETS / "poster_nostalgic_rl.png",
        scores=[
            ("intent match",  "0.72",  "0.94"),
            ("aesthetic",     "0.78",  "0.91"),
            ("coherent edits","3/7",   "6/6"),
            ("VLM overall",   "0.74",  "0.93"),
        ],
        takeaway=("Same prompt, same starting design, same loop.  "
                  "RL-finetuned agent lands the intent; CoT baseline drifts."),
        page=2,
        footer_label="Vision  ·  Watch the loop run",
    )
    add_notes(s,
        "Play both videos in parallel. Say: 'Same prompt, same starting design, "
        "same Express agent loop -- USER, PLAN, ACT, REASON, JUDGE. What "
        "differs is only the ACT policy: CoT-based on the left, RL-finetuned "
        "on the right. Watch what happens: CoT drifts, edits look generic. "
        "RL commits to the intent and stacks coherent tool calls. The judge "
        "agrees -- 0.94 vs 0.72 on intent match. This is the loop I post-train. "
        "The rest of the talk is HOW.' The judge-score strip is the concrete "
        "contrast; do not skip it.")
    return s


def s3_who_i_am():
    s = add_blank()
    add_page_header(s, "Who I am", "Subhojyoti Mukherjee  ·  Adobe Research")

    add_text(s, Inches(0.6), Inches(1.75), Inches(6), Inches(0.4),
             "WHAT I'VE SHIPPED", size=12, bold=True, color=ACCENT,
             font=FONT_HEAD)
    ships = [
        ("SlimLM 80M in Acrobat DC",
         "Distillation + DPO + offline RL. Suggested Questions and Document "
         "Overview live in Acrobat Pro. First bilingual (EN + ES) on-device SLM."),
        ("Express Agent Orchestrator",
         "Intent classifier (Gemma-4-E2B, 85.9% acc, 31 sub-intents) + planner + "
         "reflection. Forward synthetic pipeline -> 10K+ traces. Video on s2."),
        ("AdvantageFlow  ·  Firefly Gen6",
         "Advantage-weighted RL for rectified flows. In production Firefly image "
         "editing. A/B wins over Flow-GRPO baseline."),
    ]
    y = Inches(2.2)
    for head, body in ships:
        add_rect(s, Inches(0.6), y, Inches(6), Inches(1.35),
                 fill=CARD, line=RULE)
        add_text(s, Inches(0.8), y + Inches(0.15), Inches(5.6), Inches(0.4),
                 head, size=14, bold=True, color=INK, font=FONT_HEAD)
        add_text(s, Inches(0.8), y + Inches(0.55), Inches(5.6), Inches(0.75),
                 body, size=11, color=MUTED, font=FONT_BODY)
        y += Inches(1.45)

    add_text(s, Inches(6.9), Inches(1.75), Inches(6), Inches(0.4),
             "WHAT I PUBLISH", size=12, bold=True, color=ACCENT2,
             font=FONT_HEAD)
    pubs = [
        ("NeurIPS 2025",
         "Offline RL by Reward-Weighted Fine-Tuning for DC Tutoring"
         "  -  Refit / Swift."),
        ("NeurIPS 2024",
         "Optimal Design for Human Preference Elicitation  -  data collection "
         "for RLHF."),
        ("NeurIPS 2026  (sub.)",
         "AdvantageFlow  -  advantage-weighted RL for rectified flows in "
         "Firefly Gen6."),
        ("CVPR 2026",
         "Stepwise credit assignment for GRPO on flow models."),
    ]
    y = Inches(2.2)
    for head, body in pubs:
        add_rect(s, Inches(6.9), y, Inches(5.8), Inches(1.0),
                 fill=CARD, line=RULE)
        add_text(s, Inches(7.1), y + Inches(0.12), Inches(5.4), Inches(0.35),
                 head, size=12, bold=True, color=ACCENT2, font=FONT_HEAD)
        add_text(s, Inches(7.1), y + Inches(0.45), Inches(5.4), Inches(0.55),
                 body, size=11, color=INK, font=FONT_BODY)
        y += Inches(1.1)

    add_rect(s, Inches(0.6), Inches(6.5), Inches(12.1), Inches(0.45),
             fill=TINT_N)
    stats = ("6+ papers   ·   6 patents / IDFs   ·   2 shipped products   ·   "
             "NeurIPS 2025 + 2026 sub. + CVPR 2026")
    add_text(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.35),
             stats, size=12, bold=True, color=INK, font=FONT_HEAD,
             align=PP_ALIGN.CENTER)

    add_footer(s, 3, TOTAL, "Vision  ·  Who I am")
    add_notes(s,
        "Fast strip. Shipped things on the left, published things on the right. "
        "The theme: every project on both sides is one instantiation of the "
        "same recipe -- reward-weighted fine-tuning on trajectories. Do not "
        "linger; the audience wants to see the recipe.")
    return s


def s4_why_now():
    s = add_blank()
    add_page_header(s, "Vision  ·  Why now",
                    "Value has moved from tokens to actions")

    y = Inches(2.0)
    band_h = Inches(2.4)
    eras = [
        ("2019 - 22", "Pretraining",        "Scaling laws, more tokens.",
         "GPT-3  ·  Chinchilla  ·  LLaMA", MUTED, TINT_N),
        ("2022 - 24", "SFT + RLHF",         "Instruction tuning, chat.",
         "InstructGPT  ·  DPO  ·  Constitutional AI", MUTED, TINT_N),
        ("2024 - 26", "Agentic post-training", "Offline RL, judges, tool-use.",
         "GRPO  ·  DeepSeek-R1  ·  Tulu 3  ·  APT-1", ACCENT, TINT_A),
    ]
    w = Inches(4.05); gap = Inches(0.1); x = Inches(0.6)
    for i, (era, title, body, cites, col, fill) in enumerate(eras):
        add_rect(s, x, y, w, band_h, fill=fill,
                 line=ACCENT if i == 2 else RULE)
        add_rect(s, x, y, w, Inches(0.12), fill=col)
        add_text(s, x + Inches(0.25), y + Inches(0.25), w - Inches(0.5),
                 Inches(0.4), era, size=12, bold=True, color=col,
                 font=FONT_HEAD)
        add_text(s, x + Inches(0.25), y + Inches(0.65), w - Inches(0.5),
                 Inches(0.6), title, size=18, bold=True, color=INK,
                 font=FONT_HEAD)
        add_text(s, x + Inches(0.25), y + Inches(1.2), w - Inches(0.5),
                 Inches(0.6), body, size=12, color=INK, font=FONT_BODY)
        add_text(s, x + Inches(0.25), y + Inches(1.75), w - Inches(0.5),
                 Inches(0.6), cites, size=10, color=MUTED,
                 font=FONT_MONO, italic=True)
        x += w + gap

    add_rect(s, Inches(0.6), Inches(4.9), Inches(12.1), Inches(1.5),
             fill=CARD, line=ACCENT)
    add_text(s, Inches(0.85), Inches(5.05), Inches(11.6), Inches(0.4),
             "THE ARGUMENT", size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_text(s, Inches(0.85), Inches(5.4), Inches(11.6), Inches(1.0),
             ["\"You can't prompt reliability into a model that wasn't built "
              "for it.\"  -  Dan Klein, Scaled Cognition",
              "",
              "Reliability is a post-training problem. The action-native "
              "model needs an action-native training loss."],
             size=14, color=INK, font=FONT_BODY)

    add_footer(s, 4, TOTAL, "Vision  ·  Why now")
    add_notes(s,
        "Situate the shift. Value moved from more tokens to better actions. "
        "SC named the ceiling: prompt engineering can't manufacture reliability. "
        "That's the wedge for post-training. Segue to the anchor: 'Here is the "
        "picture I hold in my head when I do this work.'")
    return s


def _draw_loop(slide, cx_in, cy_in, r_in, node_r_in=0.6, colors=None):
    """Draw a 5-node circular loop centered at (cx, cy) with radius r.

    Returns list of (label, (cx, cy)) tuples in inches.
    """
    labels = ["USER", "PLAN", "ACT", "REASON", "JUDGE"]
    cols = colors or [ACCENT, ACCENT2, ACCENT3, ACCENT2, ACCENT]
    n = len(labels)
    positions = []
    for i in range(n):
        theta = -math.pi / 2 + i * 2 * math.pi / n
        nx = cx_in + r_in * math.cos(theta)
        ny = cy_in + r_in * math.sin(theta)
        positions.append((nx, ny))

    # arrows between consecutive nodes (with a small gap so they don't overlap)
    for i in range(n):
        x1, y1 = positions[i]
        x2, y2 = positions[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy)
        ux, uy = dx / L, dy / L
        gap = node_r_in * 1.05
        ax1, ay1 = x1 + ux * gap, y1 + uy * gap
        ax2, ay2 = x2 - ux * gap, y2 - uy * gap
        add_arrow(slide, Inches(ax1), Inches(ay1),
                  Inches(ax2), Inches(ay2),
                  color=MUTED, weight=1.5)

    # nodes
    for (nx, ny), lab, col in zip(positions, labels, cols):
        d = node_r_in * 2
        add_ellipse(slide,
                    Inches(nx - node_r_in), Inches(ny - node_r_in),
                    Inches(d), Inches(d),
                    fill=col, line=None)
        add_text(slide,
                 Inches(nx - node_r_in), Inches(ny - 0.13),
                 Inches(d), Inches(0.3),
                 lab, size=11, bold=True, color=CARD, font=FONT_HEAD,
                 align=PP_ALIGN.CENTER)
    return list(zip(labels, positions))


def s5_anchor():
    """Unified anchor: loop diagram center, three pillars orbiting."""
    s = add_blank()
    add_page_header(s, "Vision  ·  Focus",
                    "One loop, one recipe, three pillars")

    # Center: the loop
    cx, cy, r = 6.65, 4.35, 1.55
    _draw_loop(s, cx, cy, r, node_r_in=0.55)

    # Center annotation
    add_text(s, Inches(cx - 1.2), Inches(cy - 0.22), Inches(2.4), Inches(0.5),
             "τ  =  one walk", size=13, bold=True, color=INK,
             font=FONT_MONO, align=PP_ALIGN.CENTER)
    add_text(s, Inches(cx - 1.2), Inches(cy + 0.05), Inches(2.4), Inches(0.5),
             "through the loop", size=11, color=MUTED,
             font=FONT_BODY, align=PP_ALIGN.CENTER, italic=True)

    # Three pillar cards orbiting  (top-left DATA, top-right REWARD,
    # bottom-center ALGORITHM starred)
    def pillar(x, y, head, sub, col, tint, star=False):
        w, h = Inches(3.6), Inches(1.6)
        add_rect(s, x, y, w, h, fill=tint, line=col if star else RULE)
        add_rect(s, x, y, w, Inches(0.14 if star else 0.1), fill=col)
        add_text(s, x + Inches(0.2), y + Inches(0.22),
                 w - Inches(0.4), Inches(0.4),
                 head, size=14, bold=True, color=col, font=FONT_HEAD)
        add_text(s, x + Inches(0.2), y + Inches(0.65),
                 w - Inches(0.4), Inches(0.9),
                 sub, size=11, color=INK, font=FONT_BODY)

    # Top-left DATA
    pillar(Inches(0.6),  Inches(1.75), "DATA",
           "Trajectories τ through the loop.\n"
           "Synthetic self-play  +  logged product traces.",
           ACCENT2, TINT_B)
    # Top-right REWARD
    pillar(Inches(9.15), Inches(1.75), "REWARD",
           "Judge scores at each step  +  terminal R(τ).\n"
           "VLM / LLM judges, >= 95% human agreement.",
           ACCENT3, TINT_C)
    # Bottom-center ALGORITHM (starred)
    pillar(Inches(4.85), Inches(5.1),
           "ALGORITHM  *",
           "Reward-weighted fine-tuning on (τ, R).\n"
           "Refit / Swift  ·  NeurIPS 2025  ·  the hero.",
           ACCENT, TINT_A, star=True)

    # Bottom caption
    add_text(s, Inches(0.6), Inches(6.85), Inches(12.1), Inches(0.3),
             "Every act 2 slide uses this picture.  Loop = agent shape.  "
             "τ = one conversation.  R = judge.  Refit weights τ by R.",
             size=12, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 5, TOTAL, "Vision  ·  Focus")
    add_notes(s,
        "The anchor. This is the one picture I hold. The five-node loop is the "
        "agent shape -- every product I ship has this shape. A trajectory τ "
        "is one full walk through the loop; a multi-turn conversation is "
        "literally a sequence of walks. The three pillars feed the loop: data "
        "IS τs, reward IS judge output, algorithm IS how we update from "
        "(τ, R). Today's deep-dive is the ALGORITHM pillar. Come back to "
        "this picture at the end.")
    return s


# =========================================================
# ACT 2 · ONE RECIPE, TWO AGENTS (s6..s12)
# =========================================================

def s6_two_agents_same_loop():
    """Side by side: DC tutoring session vs Express creative agent.
    Same 5-node loop shape, different action space."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  ONE recipe, TWO agents",
                    "A conversation is a trajectory through the loop")

    # Two mini-loops side by side, with concrete trajectories
    # Left: DC tutoring
    add_rect(s, Inches(0.6), Inches(1.75), Inches(6.0), Inches(5.2),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(5.6), Inches(0.3),
             "AGENT A  ·  Tutoring for DC",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_text(s, Inches(0.8), Inches(2.2), Inches(5.6), Inches(0.3),
             "Refit / Swift, NeurIPS 2025",
             size=10, color=MUTED, font=FONT_BODY, italic=True)

    conv = [
        ("USER",   ACCENT,  "\"Explain clause 3.2 of this contract\""),
        ("PLAN",   ACCENT2, "intent=explain_clause; call get_section"),
        ("ACT",    ACCENT3, "tool: get_section(3.2) -> clause text extracted"),
        ("USER",   ACCENT,  "\"Quiz me on it\""),
        ("ACT",    ACCENT3, "tool: gen_quiz(clause, level=med) -> 4 / 5"),
        ("REASON", ACCENT2, "4 / 5 = comprehended; propose to end"),
        ("JUDGE",  ACCENT,  "resolved=YES  ·  comprehension=0.88  ·  R = +1"),
    ]
    y = Inches(2.6)
    for label, col, body in conv:
        add_rect(s, Inches(0.8), y, Inches(0.85), Inches(0.42),
                 fill=col, line=None)
        add_text(s, Inches(0.8), y + Inches(0.08), Inches(0.85), Inches(0.3),
                 label, size=9, bold=True, color=CARD, font=FONT_HEAD,
                 align=PP_ALIGN.CENTER)
        add_text(s, Inches(1.8), y + Inches(0.06), Inches(4.6), Inches(0.4),
                 body, size=10, color=INK, font=FONT_MONO)
        y += Inches(0.52)

    # Right: Express creative
    add_rect(s, Inches(6.9), Inches(1.75), Inches(6.0), Inches(5.2),
             fill=CARD, line=RULE)
    add_text(s, Inches(7.1), Inches(1.9), Inches(5.6), Inches(0.3),
             "AGENT B  ·  Express creative",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)
    add_text(s, Inches(7.1), Inches(2.2), Inches(5.6), Inches(0.3),
             "Adobe Express orchestrator (shipping)",
             size=10, color=MUTED, font=FONT_BODY, italic=True)

    express = [
        ("USER",   ACCENT,  "\"Make the design retro\""),
        ("PLAN",   ACCENT2, "intent=style_transfer; sub=era_shift"),
        ("ACT",    ACCENT3, "tool: color_grade(sepia)"),
        ("REASON", ACCENT2, "needs grain + vintage font to hit retro"),
        ("ACT",    ACCENT3, "tool: add_grain(0.35)"),
        ("ACT",    ACCENT3, "tool: swap_font(vintage_serif)"),
        ("JUDGE",  ACCENT,  "VLM: intent=0.97 · aesthetic=0.91 · R=+1"),
    ]
    y = Inches(2.6)
    for label, col, body in express:
        add_rect(s, Inches(7.1), y, Inches(0.85), Inches(0.42),
                 fill=col, line=None)
        add_text(s, Inches(7.1), y + Inches(0.08), Inches(0.85), Inches(0.3),
                 label, size=9, bold=True, color=CARD, font=FONT_HEAD,
                 align=PP_ALIGN.CENTER)
        add_text(s, Inches(8.1), y + Inches(0.06), Inches(4.6), Inches(0.4),
                 body, size=10, color=INK, font=FONT_MONO)
        y += Inches(0.52)

    # Bottom claim
    add_text(s, Inches(0.6), Inches(7.05), Inches(12.1), Inches(0.3),
             "Different action space, same shape.  "
             "Both trajectories live in the same loop -> both are trained by "
             "the same loss.",
             size=12, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 6, TOTAL, "Act 2  ·  Two agents")
    add_notes(s,
        "The unifying move. On the left is the Refit/Swift setting: a support "
        "conversation, tools called mid-turn. On the right is the Express agent "
        "you already watched. Point out: both trajectories are sequences of "
        "USER/PLAN/ACT/REASON/JUDGE tokens. That's the whole insight -- once "
        "you accept a conversation IS a walk through this loop, the same loss "
        "trains both agents.")
    return s


def s7_setup():
    """MDP framing on the loop; variables named."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  Setup",
                    "MDP framing: what τ, R, and π_θ actually are")

    # Left: variables
    add_rect(s, Inches(0.6), Inches(1.75), Inches(5.4), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(5.0), Inches(0.35),
             "VARIABLES",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    rows = [
        ("x",           "user request  ·  first turn"),
        ("aₜ",         "action at step t  (utterance or tool call)"),
        ("τ",           "trajectory  (a₁, a₂, ..., a_T)"),
        ("π₀",        "logging policy  (the model we already deployed)"),
        ("π_θ",    "target policy  (what we're fitting)"),
        ("R(τ)",      "terminal reward from the judge  in [0, 1]"),
        ("rₜ",          "optional step-wise judge score at each ACT/JUDGE node"),
    ]
    ty = Inches(2.4)
    for k, v in rows:
        add_text(s, Inches(0.85), ty, Inches(1.6), Inches(0.35),
                 k, size=13, bold=True, color=ACCENT, font=FONT_MONO)
        add_text(s, Inches(2.55), ty + Inches(0.02), Inches(3.4), Inches(0.4),
                 v, size=11, color=INK, font=FONT_BODY)
        ty += Inches(0.55)

    # Right: mini-loop with τ highlighted
    add_rect(s, Inches(6.3), Inches(1.75), Inches(6.4), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(6.5), Inches(1.9), Inches(6.0), Inches(0.35),
             "τ IS ONE WALK THROUGH THE LOOP",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)

    _draw_loop(s, 9.5, 4.55, 1.6, node_r_in=0.55)
    # annotation lines
    add_text(s, Inches(6.5), Inches(6.35), Inches(6.0), Inches(0.35),
             "τ = (USER x,  PLAN a₁,  ACT a₂,  REASON a₃,  JUDGE r_T)",
             size=11, color=INK, font=FONT_MONO, align=PP_ALIGN.CENTER)
    add_text(s, Inches(6.5), Inches(6.65), Inches(6.0), Inches(0.3),
             "multi-turn conversation = concatenate walks",
             size=10, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 7, TOTAL, "Act 2  ·  Setup")
    add_notes(s,
        "Name the variables so the math on the next slides has no ambiguity. "
        "Emphasize: π₀ is what we've already shipped -- we're doing offline "
        "RL from logs. We do NOT need a critic. We do NOT need importance "
        "sampling.")
    return s


# ----- derivation slides with LaTeX math -----

def _derivation_base(page):
    """Base slide for the derivation. `page` in [1, 3]."""
    s = add_blank()
    titles = {
        1: "Objective, step 1: what we want but can't compute",
        2: "Objective, step 2: Refit lower bound",
        3: "Objective, step 3: reward-weighted SFT",
    }
    add_page_header(s, "Act 2  ·  Derivation", titles[page])

    # Left: the concrete question this line answers
    add_rect(s, Inches(0.6), Inches(1.75), Inches(4.7), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(4.3), Inches(0.35),
             "WHAT THIS SAYS",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    questions = {
        1: [("V(θ):", "expected reward under π_θ"),
            ("problem:",  "we cannot roll out π_θ"),
            ("problem:",  "we only have logs from π₀"),
            ("problem:",  "importance sampling explodes")],
        2: [("insight:",  "expectations under π₀ are cheap"),
            ("bound:",    "Jensen -> lower bound on log V"),
            ("gain:",     "no importance ratios anywhere"),
            ("cost:",     "concave surrogate, not exact")],
        3: [("expand:",   "log π(τ|x) = Σₜ log π(aₜ | ...)"),
            ("weight:",   "R(τ) becomes per-trajectory weight"),
            ("shape:",    "weighted cross-entropy over logs"),
            ("ship:",     "any SFT trainer runs this")],
    }
    ty = Inches(2.45)
    for k, v in questions[page]:
        add_text(s, Inches(0.85), ty, Inches(1.35), Inches(0.4),
                 k, size=12, bold=True, color=ACCENT2, font=FONT_MONO)
        add_text(s, Inches(2.2), ty + Inches(0.02), Inches(3.0), Inches(0.5),
                 v, size=11, color=INK, font=FONT_BODY)
        ty += Inches(0.6)

    # Right: math panel that accretes across pages
    add_rect(s, Inches(5.5), Inches(1.75), Inches(7.25), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(5.7), Inches(1.9), Inches(6.85), Inches(0.35),
             "THE MATH",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    return s


def s8_objective_v0():
    s = _derivation_base(1)
    add_math(s,
             r"V(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_\theta}\!\left[R(\tau)\right]",
             Inches(5.9), Inches(2.55), 6.5, fontsize=34)
    add_text(s, Inches(5.9), Inches(4.2), Inches(6.5), Inches(0.4),
             "But we cannot sample τ from π_θ -- we haven't shipped it yet.",
             size=13, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 8, TOTAL, "Act 2  ·  Derivation (1/3)")
    add_notes(s,
        "First reveal. State the goal in one line. Then the obstacle: "
        "V(θ) is unobservable offline. Set up the need for a surrogate.")
    return s


def s9_objective_v1():
    s = _derivation_base(2)
    # step 1 (grayed)
    add_math(s,
             r"V(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_\theta}\!\left[R(\tau)\right]",
             Inches(5.9), Inches(2.45), 6.5, fontsize=26, color=MUTED)
    # step 2 (new)
    add_math(s,
             r"\log V(\theta)\;\geq\;\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\log\pi_\theta(\tau\mid x)\right]\;+\;C",
             Inches(5.7), Inches(3.75), 6.9, fontsize=24, color=INK)
    add_text(s, Inches(5.7), Inches(5.1), Inches(6.9), Inches(0.4),
             "Jensen inequality on log V.  Expectation is now under π₀ -- our logs.",
             size=12, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 9, TOTAL, "Act 2  ·  Derivation (2/3)")
    add_notes(s,
        "Second reveal. Apply Jensen. The consequence: the expectation is now "
        "over π₀, which we have logs for. No importance weights. Emphasize: "
        "this is the whole trick of Refit.")
    return s


def s10_objective_v2():
    s = _derivation_base(3)
    # step 1 (very grayed)
    add_math(s,
             r"V(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_\theta}\!\left[R(\tau)\right]",
             Inches(5.9), Inches(2.35), 5.6, fontsize=20, color=MUTED)
    # step 2 (grayed)
    add_math(s,
             r"\log V(\theta)\;\geq\;\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\log\pi_\theta(\tau\mid x)\right]\;+\;C",
             Inches(5.7), Inches(3.2), 6.4, fontsize=20, color=MUTED)
    # step 3 (new, hero)
    add_math(s,
             r"\mathcal{J}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\sum_{t=1}^{T}\log\pi_\theta(a_t\mid x,\tau_{<t})\right]",
             Inches(5.55), Inches(4.35), 7.0, fontsize=22, color=ACCENT)
    add_text(s, Inches(5.7), Inches(5.75), Inches(6.9), Inches(0.4),
             "Shippable loss.  Any SFT trainer runs this, with tokens weighted "
             "by R(τ).",
             size=12, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 10, TOTAL, "Act 2  ·  Derivation (3/3)")
    add_notes(s,
        "Third reveal. Expand log π(τ|x) across the trajectory. Now the "
        "loss is a reward-weighted cross-entropy. Punchline: any SFT trainer "
        "can run this line unchanged. That's why it ships.")
    return s


def s11_same_loss_two_agents():
    """The literal callback: same J(θ), different (τ, R) meanings."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  The callback",
                    "Same loss, two agents  ·  swap only what τ and R mean")

    # Top: the shared loss, big
    add_math(s,
             r"\mathcal{J}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\sum_{t}\log\pi_\theta(a_t\mid x,\tau_{<t})\right]",
             Inches(1.8), Inches(1.7), 9.7, fontsize=22, color=ACCENT)

    # Two columns: same LHS, different RHS interpretations
    y = Inches(3.4)
    add_rect(s, Inches(0.6), y, Inches(6.0), Inches(3.4),
             fill=TINT_A, line=ACCENT)
    add_text(s, Inches(0.85), y + Inches(0.15), Inches(5.5), Inches(0.35),
             "AGENT A  ·  Tutoring for DC",
             size=12, bold=True, color=ACCENT, font=FONT_HEAD)
    rows_a = [
        ("aₜ",   "utterance or tool call in the tutoring session"),
        ("τ",    "full 5-turn tutoring session, ~30 tokens/turn"),
        ("R(τ)","resolved=YES && comprehension >= 0.8  ->  +1 else 0"),
        ("π₀",  "the deployed DC tutor, logged 2M sessions"),
    ]
    ty = y + Inches(0.6)
    for k, v in rows_a:
        add_text(s, Inches(0.85), ty, Inches(1.2), Inches(0.35),
                 k, size=12, bold=True, color=ACCENT, font=FONT_MONO)
        add_text(s, Inches(2.1), ty + Inches(0.02), Inches(4.4), Inches(0.5),
                 v, size=11, color=INK, font=FONT_BODY)
        ty += Inches(0.55)

    add_rect(s, Inches(6.75), y, Inches(6.0), Inches(3.4),
             fill=TINT_B, line=ACCENT2)
    add_text(s, Inches(7.0), y + Inches(0.15), Inches(5.5), Inches(0.35),
             "AGENT B  ·  Express creative",
             size=12, bold=True, color=ACCENT2, font=FONT_HEAD)
    rows_b = [
        ("aₜ",   "plan token or tool call on the Express canvas"),
        ("τ",    "one edit trajectory, 4-9 tool calls"),
        ("R(τ)","VLM aesthetic + intent match >= 0.85  ->  +1"),
        ("π₀",  "shipped Express agent, logged 500K sessions"),
    ]
    ty = y + Inches(0.6)
    for k, v in rows_b:
        add_text(s, Inches(7.0), ty, Inches(1.2), Inches(0.35),
                 k, size=12, bold=True, color=ACCENT2, font=FONT_MONO)
        add_text(s, Inches(8.25), ty + Inches(0.02), Inches(4.4), Inches(0.5),
                 v, size=11, color=INK, font=FONT_BODY)
        ty += Inches(0.55)

    add_text(s, Inches(0.6), Inches(7.05), Inches(12.1), Inches(0.3),
             "The loss code is identical.  What changes: tokenizer of aₜ, "
             "trajectory length, judge model.  Nothing else.",
             size=12, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 11, TOTAL, "Act 2  ·  Callback")
    add_notes(s,
        "This is the merge slide the audience is waiting for. Point at the "
        "top equation. 'Same J. Look at the two boxes: LHS symbols identical, "
        "RHS meanings swapped.' Push the point: 'This is why the Refit paper "
        "code was mostly reusable when we built the Express orchestrator "
        "trainer -- same loss module, new tokenizer.'")
    return s


def s12_swift_and_results():
    """Conversation agent: Refit/Swift on tutoring for DC (NeurIPS 2025)."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  Results 1 of 3  ·  Conversation",
                    "Refit/Swift on tutoring for DC  ·  NeurIPS 2025")

    # Left: swift math
    add_rect(s, Inches(0.6), Inches(1.75), Inches(6.0), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(5.6), Inches(0.35),
             "SWIFT  ·  variance reduction via advantage",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_math(s,
             r"A(\tau)\;=\;\frac{R(\tau)-\mu_g}{\sigma_g}",
             Inches(0.8), Inches(2.6), 5.6, fontsize=32)
    add_text(s, Inches(0.8), Inches(3.7), Inches(5.6), Inches(0.4),
             "Group-relative baseline: normalize by the same-context batch.",
             size=12, color=INK, font=FONT_BODY, italic=True)
    add_math(s,
             r"\mathcal{J}_{\text{Swift}}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[A(\tau)\,\sum_{t}\log\pi_\theta(a_t\mid x,\tau_{<t})\right]",
             Inches(0.7), Inches(4.4), 5.8, fontsize=18, color=ACCENT)
    add_text(s, Inches(0.8), Inches(5.7), Inches(5.6), Inches(0.6),
             "Same trainer as Refit.  Replace R with A.  Cuts gradient "
             "variance 3x on Tau-Bench.",
             size=12, color=INK, font=FONT_BODY)

    # Right: results panel
    add_rect(s, Inches(6.8), Inches(1.75), Inches(5.95), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(7.0), Inches(1.9), Inches(5.55), Inches(0.35),
             "RESULTS  ·  DC tutoring multi-turn session",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)

    # results table
    y = Inches(2.5)
    header = [("Metric", 2.2), ("SFT", 1.0), ("Refit", 1.2), ("Swift", 1.4)]
    x = Inches(7.0)
    for h, w in header:
        add_text(s, x, y, Inches(w), Inches(0.3),
                 h, size=10, bold=True, color=MUTED, font=FONT_HEAD)
        x += Inches(w)
    add_hline(s, Inches(7.0), y + Inches(0.35), Inches(5.7))

    rows = [
        ("Task success @ 5",  "0.412", "0.503", "0.541  (+0.13)"),
        ("Reward (judge)",    "0.618", "0.712", "0.759  (+0.14)"),
        ("Coherence",         "0.734", "0.812", "0.851  (+0.12)"),
        ("Groundedness",      "0.681", "0.770", "0.805  (+0.12)"),
    ]
    ty = y + Inches(0.5)
    for r in rows:
        widths = [2.2, 1.0, 1.2, 1.4]
        cx = Inches(7.0)
        for val, w in zip(r, widths):
            col = ACCENT if "+" in val else INK
            bold = "+" in val
            add_text(s, cx, ty, Inches(w), Inches(0.3),
                     val, size=11, color=col, font=FONT_MONO, bold=bold)
            cx += Inches(w)
        ty += Inches(0.45)

    add_text(s, Inches(7.0), Inches(5.9), Inches(5.55), Inches(0.9),
             "Swift > Refit > SFT on all four metrics.  Group-relative "
             "advantage cuts gradient variance 3x.",
             size=12, italic=True, color=INK, font=FONT_BODY)

    add_footer(s, 12, TOTAL, "Act 2  ·  Conversation results")
    add_notes(s,
        "First of three real-world results. This is the original Refit/Swift "
        "paper (NeurIPS 2025) on multi-turn DC tutoring sessions. "
        "Swift is a one-line change from Refit: replace R with A = (R-μ)/σ. "
        "Same trainer, less variance. Every metric moves in the right "
        "direction. Next two slides: same recipe, different agent.")
    return s


# =========================================================
# ACT 2 · TWO MORE RESULTS (s13 agentic planning, s14 intent classifier)
# =========================================================

def s13_agentic_planning_results():
    """Second of three: agentic planning for image styling (NeurIPS 2026 sub.)."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  Results 2 of 3  ·  Agentic planning",
                    "Offline RL for image styling  ·  NeurIPS 2026 (sub.)")

    # Left: setup card
    add_rect(s, Inches(0.6), Inches(1.75), Inches(4.9), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(4.5), Inches(0.35),
             "SETUP  ·  planner as tool-use agent",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    setup_lines = [
        ("τ", "sequence of tool calls on a compositional library "
              "(time_of_day, mood_lighting, artistic_medium, ...)"),
        ("R", "triple-judge score (GPT-5.4, Gemini 3.1, Claude Opus 4.6) "
              "on 6 image-quality dimensions, 0-100"),
        ("π_θ", "Qwen3-VL 4B / 8B  ·  frozen image editor  ·  "
                "trained on 30K synthetic trajectories"),
        ("data", "3 datasets  ·  Simple (1-2 step edits), Regular (3-5, "
                 "10 themes), Complex (3-5, 83 themes)"),
        ("val", "3K-sample human study  ·  77% pass rate  ·  "
                "GPT-5.4 90.3% top-2 vs annotator consensus"),
    ]
    ty = Inches(2.4)
    for k, v in setup_lines:
        add_text(s, Inches(0.85), ty, Inches(0.9), Inches(0.35),
                 k, size=13, bold=True, color=ACCENT, font=FONT_MONO)
        add_text(s, Inches(1.8), ty + Inches(0.02),
                 Inches(3.5), Inches(0.8),
                 v, size=10, color=INK, font=FONT_BODY)
        ty += Inches(0.9)

    # Right: results table
    add_rect(s, Inches(5.7), Inches(1.75), Inches(7.05), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(5.9), Inches(1.9), Inches(6.65), Inches(0.35),
             "RESULTS  ·  three-judge overall score (0-100)",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)

    # header
    y = Inches(2.45)
    header = [("Config", 2.05), ("Edit-Only", 1.0), ("SFT", 0.8),
              ("RW", 0.9), ("SW", 0.9), ("Gemini 2.5", 1.15)]
    x = Inches(5.9)
    for h, w in header:
        add_text(s, x, y, Inches(w), Inches(0.3),
                 h, size=9, bold=True, color=MUTED, font=FONT_HEAD,
                 align=PP_ALIGN.CENTER)
        x += Inches(w)
    add_hline(s, Inches(5.9), y + Inches(0.35), Inches(6.65))

    # Values from Table 1 (a-d), G5 column (GPT-5.4 judge)
    rows = [
        ("Simple Text-4B",      "66.4", "79.9", "82.5",  "82.5",  "64.1"),
        ("Regular Text-4B",     "59.6", "72.9", "80.3",  "78.8",  "60.2"),
        ("Complex Vision-4B",   "70.0", "80.5", "83.5",  "86.1",  "63.6"),
        ("Complex Vision-8B",   "70.7", "80.6", "84.7",  "82.5",  "64.0"),
    ]
    ty = y + Inches(0.5)
    for r in rows:
        widths = [2.05, 1.0, 0.8, 0.9, 0.9, 1.15]
        cx = Inches(5.9)
        # find best across the 4 trained methods (indices 1..4)
        trained_vals = [float(r[i]) for i in range(1, 5)]
        best_idx = trained_vals.index(max(trained_vals)) + 1
        for i, (val, w) in enumerate(zip(r, widths)):
            if i == 0:
                col, bold = INK, False
            elif i == best_idx:
                col, bold = ACCENT, True
            elif i == 5:
                col, bold = MUTED, False
            else:
                col, bold = INK, False
            add_text(s, cx, ty, Inches(w), Inches(0.3),
                     val, size=11, color=col, font=FONT_MONO, bold=bold,
                     align=PP_ALIGN.CENTER)
            cx += Inches(w)
        ty += Inches(0.42)

    add_text(s, Inches(5.9), Inches(5.15), Inches(6.65), Inches(0.6),
             "RW / SW win on every configuration.  4B/8B planners beat "
             "GPT-4o and Gemini 2.5 zero-shot with orders of magnitude "
             "fewer parameters.",
             size=11, italic=True, color=INK, font=FONT_BODY)
    add_text(s, Inches(5.9), Inches(5.9), Inches(6.65), Inches(0.6),
             "Robustness: SW keeps +3.0 advantage over Edit-Only under OOD "
             "(20 atomic → 3-5 step compositional).  No catastrophic failure.",
             size=11, italic=True, color=ACCENT2, font=FONT_BODY)

    add_footer(s, 13, TOTAL, "Act 2  ·  Agentic planning")
    add_notes(s,
        "Second real-world result. Different agent, same recipe. Same J(θ) "
        "as the conversation slide -- only τ (tool-call sequence) and R "
        "(judge score) change. Point out: compact 4B/8B beats frontier "
        "closed-source zero-shot. That's the ROI of structured planning + "
        "reward-aware training.")
    return s


def s14_intent_classifier_results():
    """Third of three: intent classifier distilled to 0.8B, deployed."""
    s = add_blank()
    add_page_header(s, "Act 2  ·  Results 3 of 3  ·  Intent classifier",
                    "0.8B distilled model  ·  deployed in Adobe planner")

    # Left: setup + latency
    add_rect(s, Inches(0.6), Inches(1.75), Inches(4.9), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(0.8), Inches(1.9), Inches(4.5), Inches(0.35),
             "SETUP  ·  routing block inside the planner",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    setup_lines = [
        ("τ", "single-turn intent + sub-intent label  "
              "(STYLE/COLOR, IMAGING/AI_MODIFY, ...)"),
        ("R", "top-1 exact-match against human-labeled golden + "
              "prod sets"),
        ("π_θ", "Qwen3.5-0.8B  ·  KD from 27B teacher + post-SFT  ·  "
                "on-device deployable"),
        ("data", "iter_026 curriculum  ·  PDGEN synthetic + DA judge "
                 "filtering  ·  auto-tuning loop"),
    ]
    ty = Inches(2.4)
    for k, v in setup_lines:
        add_text(s, Inches(0.85), ty, Inches(0.9), Inches(0.35),
                 k, size=13, bold=True, color=ACCENT, font=FONT_MONO)
        add_text(s, Inches(1.8), ty + Inches(0.02),
                 Inches(3.5), Inches(0.8),
                 v, size=10, color=INK, font=FONT_BODY)
        ty += Inches(0.9)

    # Latency callout
    add_rect(s, Inches(0.8), Inches(5.9), Inches(4.5), Inches(0.8),
             fill=TINT_A, line=ACCENT)
    add_text(s, Inches(0.95), Inches(5.98), Inches(4.3), Inches(0.35),
             "LATENCY  ·  2x faster than GPT-5.2",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_text(s, Inches(0.95), Inches(6.28), Inches(4.3), Inches(0.4),
             "880 ms (vLLM)  vs  1,785 ms (GPT-5.2)  ·  29x on raw inference",
             size=11, color=INK, font=FONT_MONO)

    # Right: accuracy table
    add_rect(s, Inches(5.7), Inches(1.75), Inches(7.05), Inches(5.1),
             fill=CARD, line=RULE)
    add_text(s, Inches(5.9), Inches(1.9), Inches(6.65), Inches(0.35),
             "RESULTS  ·  top-1 accuracy vs GPT-5.2 zero-shot",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)

    # header
    y = Inches(2.45)
    header = [("Model", 3.2), ("Golden (n=552)", 1.7), ("Prod (n=2,989)", 1.75)]
    x = Inches(5.9)
    for h, w in header:
        add_text(s, x, y, Inches(w), Inches(0.3),
                 h, size=10, bold=True, color=MUTED, font=FONT_HEAD,
                 align=PP_ALIGN.CENTER)
        x += Inches(w)
    add_hline(s, Inches(5.9), y + Inches(0.35), Inches(6.65))

    rows = [
        ("Qwen3.5-0.8B  (no SFT)",         "54.35%", "61.26%", False),
        ("GPT-5.2 (planner baseline)",     "83.51%", "73.44%", False),
        ("0.8B KD(27B) + postSFT  ★",      "85.33%", "76.55%", True),
        ("Qwen3.8-27B SFT  (upper bound)", "90.04%", "82.67%", False),
    ]
    ty = y + Inches(0.5)
    for r in rows:
        widths = [3.2, 1.7, 1.75]
        cx = Inches(5.9)
        vals = r[:3]
        bold = r[3]
        for i, (val, w) in enumerate(zip(vals, widths)):
            col = ACCENT if bold else INK
            add_text(s, cx, ty, Inches(w), Inches(0.35),
                     val, size=11, color=col, font=FONT_MONO, bold=bold,
                     align=PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT)
            cx += Inches(w)
        ty += Inches(0.5)

    # Lift callout
    add_text(s, Inches(5.9), Inches(5.15), Inches(6.65), Inches(0.4),
             "+32.1 pp golden  /  +14.7 pp prod  vs the 0.8B untuned base",
             size=12, bold=True, italic=True, color=ACCENT, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(5.9), Inches(5.6), Inches(6.65), Inches(0.4),
             "0.8B distilled model beats GPT-5.2 on both sets  ·  ships "
             "on-device",
             size=11, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(5.9), Inches(6.05), Inches(6.65), Inches(0.7),
             "Same loop.  τ = intent label.  R = exact-match reward.  "
             "Distillation is reward-weighted SFT with teacher-scores as R.",
             size=11, italic=True, color=ACCENT2, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 14, TOTAL, "Act 2  ·  Intent classifier")
    add_notes(s,
        "Third real-world result. Same recipe, third agent shape. This one "
        "ships in production Adobe planner. Point out: the 0.8B distilled "
        "model beats GPT-5.2 by +1.8pp golden and +3.1pp prod, while running "
        "2x faster (880 vs 1785 ms) and being on-device deployable. "
        "Distillation IS reward-weighted SFT -- teacher scores are R.")
    return s


# =========================================================
# ACT 3 · RSIA VISION + POC + GENERALIZATION (s15..s18)
# =========================================================

def s15_rsia_vision():
    """Three pillars close into a meta-loop -- the RSIA vision."""
    s = add_blank()
    add_page_header(s, "Act 3  ·  Vision: Recursive Self-Improving Agent",
                    "The pillars close into a meta-loop  ·  the agent "
                    "improves itself")

    # Inner loop (small, muted) -- the agent shape
    cx, cy, r_inner = 6.65, 4.15, 0.95
    _draw_loop(s, cx, cy, r_inner, node_r_in=0.36)
    add_text(s, Inches(cx - 1.2), Inches(cy - 0.12), Inches(2.4),
             Inches(0.35),
             "one agent walk", size=10, italic=True, color=MUTED,
             font=FONT_BODY, align=PP_ALIGN.CENTER)

    # Three big pillar cards arranged as a triangle around the inner loop
    def pillar(x, y, tag, head, sub, col, tint, star=False):
        w, h = Inches(3.6), Inches(1.75)
        add_rect(s, x, y, w, h, fill=tint, line=col if star else RULE)
        add_rect(s, x, y, w, Inches(0.14 if star else 0.1), fill=col)
        add_text(s, x + Inches(0.2), y + Inches(0.22),
                 w - Inches(0.4), Inches(0.35),
                 tag, size=10, bold=True, color=col, font=FONT_HEAD)
        add_text(s, x + Inches(0.2), y + Inches(0.55),
                 w - Inches(0.4), Inches(0.4),
                 head, size=15, bold=True, color=INK, font=FONT_HEAD)
        add_text(s, x + Inches(0.2), y + Inches(0.95),
                 w - Inches(0.4), Inches(0.8),
                 sub, size=10, color=MUTED, font=FONT_BODY)

    # Top-left: DATA
    pillar(Inches(0.6), Inches(1.75), "PILLAR  ·  01",
           "DATA",
           "Deployed agent emits τ.\nSynthetic PDGEN fills the gaps.\n"
           "Every week, the τ-set grows.",
           ACCENT3, TINT_C)
    # Top-right: REWARD
    pillar(Inches(9.15), Inches(1.75), "PILLAR  ·  02",
           "REWARD",
           "VLM / LLM judges score τ.\n"
           "Human labels sample-correct R.\n"
           "Judge distills as it learns.",
           ACCENT2, TINT_B)
    # Bottom-center: POST-TRAIN (starred)
    pillar(Inches(4.85), Inches(5.35), "PILLAR  ·  03  *",
           "POST-TRAIN",
           "Reward-weighted fine-tune of π_θ.\n"
           "Refit / Swift on the fresh (τ, R).\n"
           "Ship the new checkpoint.",
           ACCENT, TINT_A, star=True)

    # Curved arrows connecting the pillars into a meta-loop.
    # DATA (top-left) -> REWARD (top-right)
    add_arrow(s, Inches(4.4),  Inches(2.55), Inches(9.05), Inches(2.55),
              color=ACCENT, weight=2.0)
    # REWARD (top-right) -> POST-TRAIN (bottom-center)
    add_arrow(s, Inches(10.6), Inches(3.55), Inches(8.35), Inches(5.35),
              color=ACCENT, weight=2.0)
    # POST-TRAIN (bottom-center) -> DATA (top-left)
    add_arrow(s, Inches(5.0),  Inches(5.35), Inches(2.9),  Inches(3.55),
              color=ACCENT, weight=2.0)

    # Bottom punchline strip
    add_rect(s, Inches(0.6), Inches(6.8), Inches(12.1), Inches(0.45),
             fill=INK, line=None)
    add_text(s, Inches(0.6), Inches(6.87), Inches(12.1), Inches(0.35),
             "APT-2  =  this outer loop, run weekly, at scale.",
             size=14, bold=True, color=BG, font=FONT_HEAD,
             align=PP_ALIGN.CENTER)

    add_footer(s, 15, TOTAL, "Act 3  ·  RSIA vision")
    add_notes(s,
        "Act 3 opener. Zoom out from the inner agent loop to the outer "
        "self-improvement loop. Three pillars are not just categories, "
        "they are stations in a cycle: deploy generates DATA, REWARD "
        "scores that data, POST-TRAIN closes the loop with a new "
        "checkpoint. This IS the RSIA vision. Next slide proves it "
        "already ships.")
    return s


def s16_auto_tuning_loop():
    """The 7-step auto-tuning pipeline, re-colored into three pillars.
    Proves the RSIA loop already runs at 0.8B (intent classifier).
    """
    s = add_blank()
    add_page_header(s, "Act 3  ·  POC: the loop already runs",
                    "Intent-classifier auto-tuning  ·  7 steps  ·  3 pillars  ·  weekly")

    # Serpentine cycle:
    #   Row 1 L->R:  1 -> 2 -> 3 -> 4
    #                                |         (right-side drop)
    #                                v
    #   Row 2 R->L:  7 <- 6 <- 5
    #     ^                              (left-side return, labelled '↺ weekly')
    #     └──────────────────── 1
    # Step tuple: (num, row, col, head, body, colour).  Row 2 fills cols 1..3
    # (col 0 empty) so the return-arrow has clean margin to travel up.
    box_w = Inches(2.75)
    box_h = Inches(1.30)
    col_w = 2.95    # box_w + gap_x
    col_x = [Inches(0.90 + i * col_w) for i in range(4)]
    row_y = [Inches(1.90), Inches(3.95)]

    steps = [
        (1, 0, 0, "Eval",             "score checkpoint on golden + prod",    ACCENT3),
        (2, 0, 1, "Failure analysis", "cluster losses, extract weak intents", ACCENT3),
        (3, 0, 2, "PDGEN",            "synthetic τ for weak intents",         ACCENT3),
        (4, 0, 3, "DA Judge",         "de-duplication + quality gate",        ACCENT2),
        (5, 1, 3, "Planner",          "curriculum: what to train next",       ACCENT2),
        (6, 1, 2, "Reward Judge",     "teacher 27B scores per τ",             ACCENT2),
        (7, 1, 1, "SFT Training",     "reward-weighted KD + post-SFT",        ACCENT),
    ]

    pos = {}  # step_num -> (x, y, colour)
    for (n, row, col, head, body, colr) in steps:
        x = col_x[col]
        y = row_y[row]
        pos[n] = (x, y, colr)
        tint = TINT_C if colr == ACCENT3 else (TINT_B if colr == ACCENT2 else TINT_A)
        add_rect(s, x, y, box_w, box_h, fill=tint, line=colr)
        add_rect(s, x, y, box_w, Inches(0.12), fill=colr)
        add_text(s, x + Inches(0.15), y + Inches(0.20),
                 box_w - Inches(0.3), Inches(0.3),
                 f"0{n}", size=10, bold=True, color=colr, font=FONT_HEAD)
        add_text(s, x + Inches(0.15), y + Inches(0.48),
                 box_w - Inches(0.3), Inches(0.4),
                 head, size=13, bold=True, color=INK, font=FONT_HEAD)
        add_text(s, x + Inches(0.15), y + Inches(0.83),
                 box_w - Inches(0.3), Inches(0.5),
                 body, size=9, color=MUTED, font=FONT_BODY)

    # Row 1 arrows going RIGHT: 1->2, 2->3, 3->4
    for a, b in [(1, 2), (2, 3), (3, 4)]:
        xa, ya, _ = pos[a]
        xb, _,  _ = pos[b]
        y_mid = ya + box_h / 2
        add_arrow(s, xa + box_w, y_mid, xb, y_mid,
                  color=MUTED, weight=1.5)

    # Row 2 arrows going LEFT (visually): 5->6, 6->7
    for a, b in [(5, 6), (6, 7)]:
        xa, ya, _ = pos[a]
        xb, _,  _ = pos[b]
        y_mid = ya + box_h / 2
        add_arrow(s, xa, y_mid, xb + box_w, y_mid,
                  color=MUTED, weight=1.5)

    # Right-side connector: 4 -> 5 (drop straight down, same column)
    x4, y4, _ = pos[4]
    _,  y5, _ = pos[5]
    x_r = x4 + box_w / 2
    add_arrow(s, x_r, y4 + box_h, x_r, y5,
              color=ACCENT, weight=2.0)

    # Left-side return: 7 -> 1 (out to left margin, up, back into box 1)
    x7, y7, _ = pos[7]
    x1, y1, _ = pos[1]
    y7_mid = y7 + box_h / 2
    y1_mid = y1 + box_h / 2
    x_l = Inches(0.45)
    add_arrow(s, x7, y7_mid, x_l, y7_mid, color=ACCENT, weight=2.0)
    add_arrow(s, x_l, y7_mid, x_l, y1_mid, color=ACCENT, weight=2.0)
    add_arrow(s, x_l, y1_mid, x1, y1_mid, color=ACCENT, weight=2.0)

    # '↺ weekly' label sits next to the vertical return stroke
    add_text(s, Inches(0.02), Inches(3.30), Inches(0.83), Inches(0.35),
             "↺ weekly",
             size=11, bold=True, italic=True, color=ACCENT,
             font=FONT_HEAD, align=PP_ALIGN.CENTER)

    # Pillar legend strip
    y_leg = Inches(5.55)
    add_rect(s, Inches(0.6), y_leg, Inches(4.0), Inches(0.5),
             fill=TINT_C, line=ACCENT3)
    add_text(s, Inches(0.75), y_leg + Inches(0.1), Inches(3.8), Inches(0.3),
             "DATA  =  steps 01..03",
             size=11, bold=True, color=ACCENT3, font=FONT_HEAD)

    add_rect(s, Inches(4.75), y_leg, Inches(4.0), Inches(0.5),
             fill=TINT_B, line=ACCENT2)
    add_text(s, Inches(4.9), y_leg + Inches(0.1), Inches(3.8), Inches(0.3),
             "REWARD  =  steps 04..06",
             size=11, bold=True, color=ACCENT2, font=FONT_HEAD)

    add_rect(s, Inches(8.9), y_leg, Inches(3.8), Inches(0.5),
             fill=TINT_A, line=ACCENT)
    add_text(s, Inches(9.05), y_leg + Inches(0.1), Inches(3.6), Inches(0.3),
             "POST-TRAIN  =  step 07  *",
             size=11, bold=True, color=ACCENT, font=FONT_HEAD)

    # Facts strip
    add_text(s, Inches(0.6), Inches(6.25), Inches(12.1), Inches(0.35),
             "0.8B student  ·  KD(27B) + post-SFT  ·  8×H100  ·  "
             "~12h / epoch  ·  iter_017 → iter_026 (10 auto cycles)",
             size=12, color=INK, font=FONT_MONO, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.6), Inches(6.60), Inches(12.1), Inches(0.35),
             "Result: beats GPT-5.2 by +1.8pp golden / +3.1pp prod  ·  "
             "2x lower latency  ·  in production.",
             size=12, italic=True, color=ACCENT, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 16, TOTAL, "Act 3  ·  Auto-tuning POC")
    add_notes(s,
        "Proof-of-concept for the RSIA vision. This exact pipeline runs "
        "weekly on the intent classifier -- the 0.8B model on the prior "
        "slide came out of iter_026. Point out the color mapping: the 7 "
        "engineering steps are just DATA (01-03) -> REWARD (04-06) -> "
        "POST-TRAIN (07) -> repeat. This is the meta-loop from the vision "
        "slide, but concrete and shipped. The vision is not aspirational; "
        "it is a running system that I want to scale up for APT-2.")
    return s


# =========================================================
# ACT 3 · GENERALIZATION (s17..s18)
# =========================================================

def s13_generalization():
    s = add_blank()
    add_page_header(s, "Act 3  ·  Where the loop seeds",
                    "Each row is a fresh RSIA loop waiting to run")

    rows = [
        ("AdvantageFlow  ·  Firefly Gen6",
         "τ = denoising trajectory  ·  R = human aesthetic reward  ·  A(τ)-weighted flow-matching",
         "NeurIPS 2026 (sub.)", ACCENT2, TINT_B),
        ("Stepwise credit  ·  Firefly RL",
         "τ = per-frame edit  ·  R = per-step VLM temporal reward",
         "CVPR 2026", ACCENT3, TINT_C),
        ("SlimLM  ·  Acrobat DC",
         "τ = distilled response  ·  R = teacher score + user feedback",
         "shipped 2025", ACCENT, TINT_A),
        ("Optimal Design  ·  RLHF data",
         "τ = k-way ranking  ·  R = elicited preference  ·  D-optimal query design",
         "NeurIPS 2024", MUTED, TINT_N),
    ]
    y = Inches(1.85)
    for head, body, cite, col, tint in rows:
        add_rect(s, Inches(0.6), y, Inches(12.1), Inches(1.1),
                 fill=tint, line=col)
        add_rect(s, Inches(0.6), y, Inches(0.14), Inches(1.1), fill=col)
        add_text(s, Inches(0.9), y + Inches(0.12), Inches(9.5), Inches(0.4),
                 head, size=14, bold=True, color=INK, font=FONT_HEAD)
        add_text(s, Inches(0.9), y + Inches(0.5), Inches(9.5), Inches(0.5),
                 body, size=11, color=MUTED, font=FONT_MONO)
        add_text(s, Inches(10.6), y + Inches(0.35), Inches(2.0), Inches(0.5),
                 cite, size=11, bold=True, color=col, font=FONT_HEAD,
                 align=PP_ALIGN.RIGHT)
        y += Inches(1.2)

    add_text(s, Inches(0.6), Inches(7.0), Inches(12.1), Inches(0.3),
             "Each row = seed for its own RSIA loop.  Same (τ, R) contract, "
             "same auto-tuning pipeline, different agent.",
             size=12, italic=True, color=INK, font=FONT_BODY,
             align=PP_ALIGN.CENTER)
    add_footer(s, 17, TOTAL, "Act 3  ·  Generalization")
    add_notes(s,
        "Four other places I've applied the same recipe. Don't deep-dive any "
        "row; keep it tight. The takeaway: this is not a niche loss. Once you "
        "own the recipe, the surface area is enormous.")
    return s


def b_lessons():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b8  Three lessons",
                    "What I take home from four years of this recipe")

    lessons = [
        ("Lesson 01",
         "Trajectories, not tokens.",
         "Once you re-frame data as τ, everything downstream simplifies. "
         "Same loss family works for chat, edit-loops, denoisers, rankings.",
         ACCENT),
        ("Lesson 02",
         "Judges are the bottleneck.",
         "Refit works if R is right. The multiplier on any post-training "
         "stack is judge quality, not loss cleverness. Invest there first.",
         ACCENT2),
        ("Lesson 03",
         "Ship the trainer, iterate on data.",
         "Reward-weighted SFT is one function. Once shipped, all improvement "
         "comes from better τ (synthetic pipelines) and better R (judges).",
         ACCENT3),
    ]
    y = Inches(1.9)
    for tag, head, body, col in lessons:
        add_rect(s, Inches(0.6), y, Inches(12.1), Inches(1.4),
                 fill=CARD, line=RULE)
        add_rect(s, Inches(0.6), y, Inches(2.4), Inches(1.4), fill=col)
        add_text(s, Inches(0.75), y + Inches(0.2), Inches(2.1), Inches(0.35),
                 tag, size=11, bold=True, color=CARD, font=FONT_HEAD)
        add_text(s, Inches(0.75), y + Inches(0.6), Inches(2.1), Inches(0.7),
                 head, size=15, bold=True, color=CARD, font=FONT_HEAD)
        add_text(s, Inches(3.2), y + Inches(0.3), Inches(9.3), Inches(1.0),
                 body, size=12, color=INK, font=FONT_BODY)
        y += Inches(1.55)

    add_footer(s, 33, TOTAL_ALL, "APPENDIX  ·  b8  Lessons")
    add_notes(s,
        "Backup: three lessons if the room wants a personal-take slide. "
        "Lesson 3 is the most important one for a lab planning APT-2 -- "
        "don't spend the year tuning the loss; spend it on τ and R.")
    return s


def s15_anchor_reprise():
    """Anchor slide again, examples filled in, meta-loop overlaid."""
    s = add_blank()
    add_page_header(s, "Act 3  ·  Framework reprise",
                    "Slide 5's picture  +  the RSIA meta-loop overlaid")

    cx, cy, r = 6.65, 4.35, 1.55
    _draw_loop(s, cx, cy, r, node_r_in=0.55)

    add_text(s, Inches(cx - 1.2), Inches(cy - 0.22), Inches(2.4), Inches(0.5),
             "τ  =  one walk", size=13, bold=True, color=INK,
             font=FONT_MONO, align=PP_ALIGN.CENTER)
    add_text(s, Inches(cx - 1.2), Inches(cy + 0.05), Inches(2.4), Inches(0.5),
             "conv / edit-loop / rank", size=11, color=MUTED,
             font=FONT_BODY, align=PP_ALIGN.CENTER, italic=True)

    def pillar(x, y, head, sub, col, tint, star=False):
        w, h = Inches(3.6), Inches(1.6)
        add_rect(s, x, y, w, h, fill=tint, line=col if star else RULE)
        add_rect(s, x, y, w, Inches(0.14 if star else 0.1), fill=col)
        add_text(s, x + Inches(0.2), y + Inches(0.22),
                 w - Inches(0.4), Inches(0.4),
                 head, size=14, bold=True, color=col, font=FONT_HEAD)
        add_text(s, x + Inches(0.2), y + Inches(0.65),
                 w - Inches(0.4), Inches(0.9),
                 sub, size=11, color=INK, font=FONT_BODY)

    pillar(Inches(0.6),  Inches(1.75), "DATA",
           "2M DC tutoring sessions (Refit)\n"
           "500K Express edit sessions\n"
           "10K PDGEN synthetic traces",
           ACCENT3, TINT_C)
    pillar(Inches(9.15), Inches(1.75), "REWARD",
           "comprehension, resolved  (tutoring)\n"
           "VLM aesthetic + intent  (Express)\n"
           "teacher-27B score  (intent-clf)",
           ACCENT2, TINT_B)
    pillar(Inches(4.85), Inches(5.1), "POST-TRAIN  *",
           "Refit  ·  Swift  ·  AdvantageFlow  ·  KD+SFT\n"
           "All one line of J(θ).  Auto-tuned weekly.",
           ACCENT, TINT_A, star=True)

    # Outer meta-loop overlay: dashed rounded rectangle enclosing all three
    # pillars, with a caption at top signaling the self-improvement loop.
    meta = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.35), Inches(1.55),
        Inches(12.6), Inches(5.15),
    )
    meta.fill.background()
    meta.line.color.rgb = ACCENT
    meta.line.width = Pt(1.75)
    meta.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    meta.shadow.inherit = False

    # Meta-loop label
    add_rect(s, Inches(4.6), Inches(1.45), Inches(4.1), Inches(0.28),
             fill=BG, line=None)
    add_text(s, Inches(4.6), Inches(1.45), Inches(4.1), Inches(0.28),
             "↺  RSIA meta-loop  ·  self-improves weekly",
             size=11, bold=True, italic=True, color=ACCENT,
             font=FONT_HEAD, align=PP_ALIGN.CENTER)

    add_text(s, Inches(0.6), Inches(6.85), Inches(12.1), Inches(0.3),
             "Same picture as slide 5, populated with real projects, "
             "wrapped by the RSIA meta-loop.  Whole talk on one slide.",
             size=12, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 18, TOTAL, "Act 3  ·  Reprise")
    add_notes(s,
        "The framework reprise. Point at the pillars, now populated with "
        "the actual instantiations from the talk, and trace the dashed "
        "outer loop -- that is the RSIA loop from s15 wrapping everything. "
        "This is where the audience should feel the whole argument settle "
        "and see how APT-2 = scale this dashed loop.")
    return s


# =========================================================
# ACT 4 · LANDING (s16..s18)
# =========================================================

def s16_apt2_bets():
    s = add_blank()
    add_page_header(s, "Act 4  ·  What I'd bring to APT-2",
                    "Four legs of the RSIA flywheel  ·  one per pillar  +  one for safety")

    bets = [
        ("LEG 01  ·  DATA",
         "Trace-log flywheel: offline RL on production APT-1 traces",
         "SC ships APT-1 today -> deployed traces are the DATA pillar's "
         "output. Weekly Refit over those + judge scores. This is exactly "
         "the auto-tuning loop from s16, applied to APT-1's own traces.",
         ACCENT3),
        ("LEG 02  ·  REWARD",
         "Judge distillation: VLM/LLM -> small critic, inline during RL",
         "Judge cost dominates the REWARD pillar. Distill VLM judges into "
         "a 0.5B critic; use inline during Refit. Same trick that made "
         "Express orchestrator training tractable and the intent-clf 0.8B "
         "student ship.",
         ACCENT2),
        ("LEG 03  ·  POST-TRAIN",
         "Per-step credit for long tool-use trajectories",
         "APT-1 uses trajectory-level R. On Tau-Bench-style 20+ step traces, "
         "credit vanishes. Extend Refit with per-node rₜ from a step-wise "
         "judge (backup b5 for the CVPR 2026 derivation).",
         ACCENT),
        ("LEG 04  ·  SAFETY-in-loop",
         "Fold policy adherence into R(τ) so it rides the gradient",
         "Reward-shaping outperforms constitutional-AI style filtering "
         "because it hits the loss directly. Ties safety to the same J(θ) "
         "and makes every RSIA cycle a safety-improving cycle.",
         MUTED),
    ]

    y = Inches(1.8)
    for tag, head, body, col in bets:
        add_rect(s, Inches(0.6), y, Inches(12.1), Inches(1.2),
                 fill=CARD, line=RULE)
        add_rect(s, Inches(0.6), y, Inches(0.16), Inches(1.2), fill=col)
        add_text(s, Inches(0.9), y + Inches(0.12), Inches(2.4), Inches(0.35),
                 tag, size=11, bold=True, color=col, font=FONT_HEAD)
        add_text(s, Inches(3.4), y + Inches(0.1), Inches(8.9), Inches(0.4),
                 head, size=14, bold=True, color=INK, font=FONT_HEAD)
        add_text(s, Inches(3.4), y + Inches(0.5), Inches(8.9), Inches(0.7),
                 body, size=11, color=MUTED, font=FONT_BODY)
        y += Inches(1.28)

    add_footer(s, 19, TOTAL, "Act 4  ·  APT-2 bets")
    add_notes(s,
        "Land in verbs. Reframe as legs of the RSIA flywheel: one bet per "
        "pillar (DATA / REWARD / POST-TRAIN) plus SAFETY riding on top. "
        "All four are directly runnable on SC's stack given APT-1 is out. "
        "Flag LEG 01 as the highest leverage because it compounds with the "
        "existing flywheel and the auto-tuning pipeline from s16.")
    return s


def s17_credentials():
    s = add_blank()
    add_page_header(s, "Act 4  ·  Credentials",
                    "Publications, patents, products, in one strip")

    def col(x, w, title, col, items):
        add_rect(s, x, Inches(1.85), w, Inches(4.9), fill=CARD, line=RULE)
        add_rect(s, x, Inches(1.85), w, Inches(0.5), fill=col)
        add_text(s, x + Inches(0.2), Inches(1.97), w - Inches(0.4),
                 Inches(0.35), title, size=13, bold=True, color=CARD,
                 font=FONT_HEAD)
        ty = Inches(2.6)
        for it in items:
            add_bullet(s, x + Inches(0.2), ty, w - Inches(0.4),
                       it, marker=col, size=11)
            ty += Inches(0.7)

    col(Inches(0.6), Inches(4.0), "PUBLICATIONS", ACCENT, [
        "NeurIPS 2025  ·  Refit / Swift",
        "NeurIPS 2024  ·  Optimal Design",
        "NeurIPS 2026 sub.  ·  AdvantageFlow",
        "CVPR 2026  ·  Stepwise credit",
        "6+ additional papers, workshops",
    ])
    col(Inches(4.75), Inches(4.0), "PATENTS / IDFs", ACCENT2, [
        "Intent classifier for creative agents",
        "Forward synthetic data pipeline (PDGEN)",
        "Advantage-weighted flow-matching",
        "On-device SLM distillation for Acrobat",
        "6 total patents / IDFs",
    ])
    col(Inches(8.9), Inches(4.0), "PRODUCTS SHIPPED", ACCENT3, [
        "SlimLM 80M in Acrobat DC (2025)",
        "Express Agent Orchestrator (2026)",
        "AdvantageFlow in Firefly Gen6",
        "Suggested Questions (Acrobat Pro)",
        "Document Overview (Acrobat Pro)",
    ])

    add_footer(s, 20, TOTAL, "Act 4  ·  Credentials")
    add_notes(s,
        "Fast strip. The point: publish and ship are both routine on my "
        "resume. Same recipe underneath all of it.")
    return s


def s18_discussion():
    s = add_blank()
    add_page_header(s, "Act 4  ·  Discussion",
                    "Two questions I'd love to be asked")

    add_rect(s, Inches(0.6), Inches(1.9), Inches(12.1), Inches(2.3),
             fill=TINT_A, line=ACCENT)
    add_text(s, Inches(0.85), Inches(2.05), Inches(11.6), Inches(0.35),
             "QUESTION 01", size=11, bold=True, color=ACCENT, font=FONT_HEAD)
    add_text(s, Inches(0.85), Inches(2.4), Inches(11.6), Inches(1.7),
             ["Where does Refit break?",
              "",
              "Answer up front: distribution shift. When π_θ drifts far "
              "from π₀, the Jensen bound loosens and the surrogate stops "
              "tracking V. In practice: cap KL to π₀, or iterate short "
              "Refit passes + on-policy relogging."],
             size=13, color=INK, font=FONT_BODY)

    add_rect(s, Inches(0.6), Inches(4.4), Inches(12.1), Inches(2.3),
             fill=TINT_B, line=ACCENT2)
    add_text(s, Inches(0.85), Inches(4.55), Inches(11.6), Inches(0.35),
             "QUESTION 02", size=11, bold=True, color=ACCENT2, font=FONT_HEAD)
    add_text(s, Inches(0.85), Inches(4.9), Inches(11.6), Inches(1.7),
             ["Why not GRPO?",
              "",
              "For long tool-use trajectories, on-policy rollouts are "
              "expensive and unsafe in prod. Refit is offline -- it runs on "
              "the trace log we already have. Swift adds group-relative "
              "variance reduction without the rollout cost."],
             size=13, color=INK, font=FONT_BODY)

    add_text(s, Inches(0.6), Inches(6.95), Inches(12.1), Inches(0.3),
             "Thank you.  Full arc lives in the appendix.  I can go deeper on "
             "anything.",
             size=12, italic=True, color=MUTED, font=FONT_BODY,
             align=PP_ALIGN.CENTER)

    add_footer(s, 21, TOTAL, "Act 4  ·  Discussion")
    add_notes(s,
        "Two pre-answered questions instead of an empty 'questions?' slide. "
        "This is a RAD move -- it invites the exact conversation you want.")
    return s


# =========================================================
# BACKUP  (bsep + b1..b7)
# =========================================================

def bsep_appendix():
    s = add_blank()
    add_rect(s, 0, 0, SW, SH, fill=INK)
    add_text(s, Inches(0.6), Inches(3.0), Inches(12.1), Inches(0.7),
             "APPENDIX", size=64, bold=True, color=BG, font=FONT_HEAD)
    add_text(s, Inches(0.6), Inches(4.0), Inches(12.1), Inches(0.4),
             "Proofs, derivations, failure modes, adjacent work, extra "
             "comparisons, and lessons (bV1-3 + b1..b8)",
             size=16, color=RULE, font=FONT_BODY, italic=True)
    add_text(s, Inches(12.0), Inches(7.05), Inches(1.2), Inches(0.3),
             "22 / 33", size=9, color=RULE, font=FONT_BODY,
             align=PP_ALIGN.RIGHT)
    add_notes(s, "Divider. Signals: I anticipate your questions.")
    return s


def bV1_horror():
    s = _cot_vs_rl_slide(
        prompt_text="Add horror typography and modify lighting for spooky vibes",
        cot_mp4=ASSETS / "horror_cot.mp4",
        cot_poster=ASSETS / "poster_horror_cot.png",
        rl_mp4=ASSETS / "horror_rl.mp4",
        rl_poster=ASSETS / "poster_horror_rl.png",
        scores=[
            ("intent match",  "0.68",  "0.92"),
            ("aesthetic",     "0.74",  "0.89"),
            ("typography change","no","yes"),
            ("lighting change","partial","yes"),
        ],
        takeaway=("Multi-attribute prompt (typography + lighting).  "
                  "CoT baseline handles one attribute; RL handles both."),
        page=23,
        footer_label="APPENDIX  ·  bV1  more comparisons",
        section_head="APPENDIX  ·  bV1  Express CoT vs RL",
        title="Multi-attribute prompt: typography + lighting",
    )
    add_notes(s,
        "Backup: more evidence that RL post-training on the loop is what "
        "makes Express reliable on compound intents. Ship this slide only "
        "if asked about generalization across prompt styles.")
    return s


def bV2_halloween():
    s = _cot_vs_rl_slide(
        prompt_text="Make the theme for Halloween",
        cot_mp4=ASSETS / "halloween_cot.mp4",
        cot_poster=ASSETS / "poster_halloween_cot.png",
        rl_mp4=ASSETS / "halloween_rl.mp4",
        rl_poster=ASSETS / "poster_halloween_rl.png",
        scores=[
            ("intent match",  "0.75",  "0.95"),
            ("aesthetic",     "0.81",  "0.90"),
            ("theme coherence","3/6",  "6/6"),
            ("VLM overall",   "0.77",  "0.92"),
        ],
        takeaway=("Vague, high-level intent ('Halloween').  "
                  "RL commits to a cohesive theme; CoT stitches unrelated edits."),
        page=24,
        footer_label="APPENDIX  ·  bV2  more comparisons",
        section_head="APPENDIX  ·  bV2  Express CoT vs RL",
        title="Vague intent: 'Make the theme for Halloween'",
    )
    add_notes(s,
        "Backup: RL agent picks a coherent theme across all edits (colors, "
        "typography, iconography). CoT edits each element independently and "
        "the design fragments.")
    return s


def bV3_watery():
    s = _cot_vs_rl_slide(
        prompt_text="Make the theme watery",
        cot_mp4=ASSETS / "watery_cot.mp4",
        cot_poster=ASSETS / "poster_watery_cot.png",
        rl_mp4=ASSETS / "watery_rl.mp4",
        rl_poster=ASSETS / "poster_watery_rl.png",
        scores=[
            ("intent match",  "0.70",  "0.93"),
            ("aesthetic",     "0.79",  "0.88"),
            ("style coherence","2/5",  "5/5"),
            ("VLM overall",   "0.73",  "0.91"),
        ],
        takeaway=("Abstract style prompt ('watery').  "
                  "RL propagates the style across the whole canvas; CoT applies "
                  "it to one element only."),
        page=25,
        footer_label="APPENDIX  ·  bV3  more comparisons",
        section_head="APPENDIX  ·  bV3  Express CoT vs RL",
        title="Abstract style: 'Make the theme watery'",
    )
    add_notes(s,
        "Backup: the abstract-style case. RL agent generalizes 'watery' to "
        "colors, blur, curves; CoT applies a single filter and stops.")
    return s


def b1_refit_lemma():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b1  Refit lemma",
                    "Jensen -> lower bound on log V")

    add_text(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(0.4),
             "Step 1  ·  goal",
             size=13, bold=True, color=ACCENT, font=FONT_HEAD)
    add_math(s, r"V(\theta)=\mathbb{E}_{\tau\sim\pi_\theta}\!\left[R(\tau)\right]"
                r"=\int R(\tau)\,\pi_\theta(\tau\mid x)\,d\tau",
             Inches(0.9), Inches(2.15), 6.0, fontsize=20)

    add_text(s, Inches(0.6), Inches(3.15), Inches(12.1), Inches(0.4),
             "Step 2  ·  reweight to π₀",
             size=13, bold=True, color=ACCENT, font=FONT_HEAD)
    add_math(s, r"V(\theta)=\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\frac{\pi_\theta(\tau\mid x)}{\pi_0(\tau\mid x)}\right]",
             Inches(0.9), Inches(3.6), 6.5, fontsize=20)

    add_text(s, Inches(0.6), Inches(4.6), Inches(12.1), Inches(0.4),
             "Step 3  ·  Jensen on log",
             size=13, bold=True, color=ACCENT, font=FONT_HEAD)
    add_math(s, r"\log V(\theta)\;\geq\;\mathbb{E}_{\tau\sim\pi_0}\!\left[R(\tau)\,\log\pi_\theta(\tau\mid x)\right]+C",
             Inches(0.9), Inches(5.05), 7.0, fontsize=20)

    add_text(s, Inches(0.6), Inches(6.15), Inches(12.1), Inches(0.4),
             "C absorbs terms constant in θ (log π₀ and normalization).",
             size=11, italic=True, color=MUTED, font=FONT_BODY)

    add_footer(s, 26, TOTAL_ALL, "APPENDIX  ·  b1")
    add_notes(s, "Full derivation for the audience who wants it.")
    return s


def b2_swift_advantage():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b2  Swift advantage",
                    "Group-relative baseline for variance reduction")

    add_math(s, r"A(\tau)\;=\;\frac{R(\tau)-\mu_g}{\sigma_g}\quad\text{with}\quad"
                r"\mu_g=\mathbb{E}_{\tau\sim\pi_0(\cdot\mid x)}[R(\tau)]",
             Inches(0.9), Inches(1.85), 11.0, fontsize=22)
    add_math(s, r"\sigma_g^2\;=\;\mathrm{Var}_{\tau\sim\pi_0(\cdot\mid x)}[R(\tau)]",
             Inches(0.9), Inches(3.2), 8.0, fontsize=22)
    add_math(s, r"\mathcal{J}_{\text{Swift}}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[A(\tau)\sum_t\log\pi_\theta(a_t\mid x,\tau_{<t})\right]",
             Inches(0.9), Inches(4.4), 11.0, fontsize=22, color=ACCENT)

    add_text(s, Inches(0.6), Inches(5.7), Inches(12.1), Inches(0.4),
             "Advantage = trajectory reward minus context-conditional group "
             "mean, normalized by std.  Preserves the sign structure of R "
             "while cutting variance ~3x on Tau-Bench.",
             size=12, color=INK, font=FONT_BODY)

    add_footer(s, 27, TOTAL_ALL, "APPENDIX  ·  b2")
    add_notes(s, "Swift math full. Emphasize: same trainer as Refit.")
    return s


def b3_optimal_design():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b3  Optimal Design (NeurIPS 2024)",
                    "How you choose queries determines RLHF sample complexity")

    add_text(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(0.4),
             "Setup  ·  learn preference model from k-way rankings under a "
             "budget B queries.",
             size=13, color=INK, font=FONT_BODY)

    add_math(s, r"\pi^\star\;=\;\arg\min_{\pi\in\Delta(\mathcal{X})}\;"
                r"\lambda_{\max}\!\left(\Sigma(\pi)^{-1}\right)",
             Inches(0.9), Inches(2.6), 8.0, fontsize=26)

    add_text(s, Inches(0.6), Inches(4.0), Inches(12.1), Inches(0.4),
             "D-optimal query distribution: pick prompts that maximize the "
             "information matrix.  Reduces RLHF label complexity by 3-10x "
             "on synthetic + real user studies.",
             size=12, color=INK, font=FONT_BODY)
    add_text(s, Inches(0.6), Inches(4.9), Inches(12.1), Inches(0.4),
             "Ties into today's talk: this is how we generate high-quality "
             "τ at low label cost -- the DATA pillar in the anchor.",
             size=12, italic=True, color=MUTED, font=FONT_BODY)

    add_footer(s, 28, TOTAL_ALL, "APPENDIX  ·  b3")
    add_notes(s,
        "This is the paper Chris K. asked about. Personal Research Discussion "
        "on Sep 10 covers this end-to-end.")
    return s


def b4_advantageflow():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b4  AdvantageFlow",
                    "Same recipe on rectified flows")

    add_text(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(0.4),
             "τ = denoising trajectory  ·  R = human aesthetic reward",
             size=13, color=INK, font=FONT_BODY)
    add_math(s, r"\mathcal{J}_{\text{AF}}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[A(\tau)\sum_t\log p_\theta(x_{t-1}\mid x_t)\right]",
             Inches(0.9), Inches(2.7), 11.0, fontsize=22, color=ACCENT)
    add_text(s, Inches(0.6), Inches(4.15), Inches(12.1), Inches(0.4),
             "Ships in Firefly Gen6.  A/B wins over Flow-GRPO on aesthetic "
             "and prompt-fidelity human eval.",
             size=12, color=INK, font=FONT_BODY)

    add_footer(s, 29, TOTAL_ALL, "APPENDIX  ·  b4")
    add_notes(s, "Direct application of Swift structure to diffusion.")
    return s


def b5_stepwise_credit():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b5  Stepwise credit",
                    "Per-node rₜ for long trajectories (CVPR 2026)")

    add_math(s, r"\mathcal{J}_{\text{step}}(\theta)\;=\;\mathbb{E}_{\tau\sim\pi_0}\!\left[\sum_t A_t\,\log\pi_\theta(a_t\mid x,\tau_{<t})\right]",
             Inches(0.7), Inches(1.9), 11.5, fontsize=22, color=ACCENT)

    add_math(s, r"A_t\;=\;r_t + \gamma\,V(s_{t+1})\;-\;V(s_t)",
             Inches(0.9), Inches(3.3), 7.0, fontsize=26)

    add_text(s, Inches(0.6), Inches(4.6), Inches(12.1), Inches(0.4),
             "Extends Refit to give credit at every ACT/JUDGE node.  "
             "Necessary for Tau-Bench-style 20+ step trajectories where "
             "trajectory-level R = 0/1 is too sparse.",
             size=12, color=INK, font=FONT_BODY)
    add_text(s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(0.4),
             "This is the machinery behind APT-2 Bet 01.",
             size=12, bold=True, italic=True, color=ACCENT2, font=FONT_BODY)

    add_footer(s, 30, TOTAL_ALL, "APPENDIX  ·  b5")
    add_notes(s, "Show the extension. Highlight: this is Bet 01 pre-cooked.")
    return s


def b6_failure_cases():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b6  Where Refit breaks",
                    "Distribution shift + judge saturation")

    failures = [
        ("Distribution shift",
         "π_θ drifts far from π₀  ->  Jensen bound loosens  ->  "
         "surrogate stops tracking V.",
         "Cap KL(π_θ || π₀) or run short Refit passes + on-policy relog."),
        ("Judge saturation",
         "When most τ get R=1 (VLM ceiling), gradient vanishes -- Swift "
         "advantage helps but doesn't fix it.",
         "Curriculum: raise judge threshold as π_θ improves.  Ship "
         "harder τ."),
        ("Reward hacking",
         "Model finds tokens that game the VLM without doing the task.",
         "Ensemble judges  +  hold-out human eval every 5K updates."),
    ]
    y = Inches(1.85)
    for head, sym, fix in failures:
        add_rect(s, Inches(0.6), y, Inches(12.1), Inches(1.55),
                 fill=CARD, line=RULE)
        add_text(s, Inches(0.8), y + Inches(0.15), Inches(11.7), Inches(0.35),
                 head, size=13, bold=True, color=ACCENT, font=FONT_HEAD)
        add_text(s, Inches(0.8), y + Inches(0.5), Inches(11.7), Inches(0.5),
                 "sym:  " + sym, size=11, color=INK, font=FONT_BODY)
        add_text(s, Inches(0.8), y + Inches(1.05), Inches(11.7), Inches(0.5),
                 "fix:  " + fix, size=11, color=ACCENT2, font=FONT_BODY,
                 italic=True)
        y += Inches(1.7)

    add_footer(s, 31, TOTAL_ALL, "APPENDIX  ·  b6")
    add_notes(s, "Honesty slide. Have the numbers ready if asked.")
    return s


def b7_related_work():
    s = add_blank()
    add_page_header(s, "APPENDIX  ·  b7  Related work",
                    "Where Refit sits in the offline-RL / RLHF landscape")

    rows = [
        ("PPO / GRPO",     "on-policy",  "needs live rollouts, expensive for tool-use"),
        ("DPO",            "preference", "requires paired preferences, not R(τ)"),
        ("KTO",            "point-wise", "assumes binary signals, discards graded R"),
        ("Rest-EM",        "iterated SFT","filter τ by R > threshold; Refit gradient generalizes"),
        ("MPO / IRPO",     "offline",    "close cousins; Refit differs by dropping importance ratios"),
        ("Behavior cloning","supervised","ignores R entirely; strict subset of Refit at R = const"),
    ]
    add_hline(s, Inches(0.6), Inches(2.05), Inches(12.1))
    add_text(s, Inches(0.7), Inches(1.75), Inches(3), Inches(0.3),
             "Method", size=11, bold=True, color=MUTED, font=FONT_HEAD)
    add_text(s, Inches(3.9), Inches(1.75), Inches(3), Inches(0.3),
             "Style",  size=11, bold=True, color=MUTED, font=FONT_HEAD)
    add_text(s, Inches(6.8), Inches(1.75), Inches(6), Inches(0.3),
             "Relation to Refit", size=11, bold=True, color=MUTED,
             font=FONT_HEAD)
    y = Inches(2.25)
    for name, style, rel in rows:
        add_text(s, Inches(0.7), y, Inches(3),   Inches(0.4),
                 name,  size=12, bold=True, color=INK, font=FONT_MONO)
        add_text(s, Inches(3.9), y, Inches(2.8), Inches(0.4),
                 style, size=12, color=ACCENT2, font=FONT_BODY, italic=True)
        add_text(s, Inches(6.8), y, Inches(6),   Inches(0.4),
                 rel,   size=12, color=INK, font=FONT_BODY)
        add_hline(s, Inches(0.6), y + Inches(0.5), Inches(12.1))
        y += Inches(0.62)

    add_footer(s, 32, TOTAL_ALL, "APPENDIX  ·  b7")
    add_notes(s, "Pre-answered comparisons. Refit ~ Rest-EM + Jensen bound + "
                 "graded R.")
    return s


# =========================================================
# BUILD
# =========================================================

# NOTE: v7.pptx has been hand-edited in PowerPoint after the last rebuild.
# Slides were reordered, and a manually-drawn "Problem Statement" slide with an
# RL-training-dataset diagram was inserted at position 3. Its header follows
# the deck-standard pattern (section label "PROBLEM  ·  RL TRAINING DATA",
# title "Turning product traces into (state, action, reward) tuples", h-rule
# at y=1.45, footer "3 / 21"), applied by /tmp/apply_v7_edits.py. If you
# regenerate v7.pptx from this script, you'll lose those manual edits.
BUILDERS = [
    # ACT 1 · VISION
    s1_title, s2_watch_the_loop, s3_who_i_am, s4_why_now, s5_anchor,
    # ACT 2 · ONE recipe, results trio
    s6_two_agents_same_loop, s7_setup,
    s8_objective_v0, s9_objective_v1, s10_objective_v2,
    s11_same_loss_two_agents, s12_swift_and_results,
    s13_agentic_planning_results, s14_intent_classifier_results,
    # ACT 3 · RSIA vision + POC + generalization + reprise
    s15_rsia_vision, s16_auto_tuning_loop,
    s13_generalization, s15_anchor_reprise,
    # ACT 4 · LANDING
    s16_apt2_bets, s17_credentials, s18_discussion,
    # BACKUP
    bsep_appendix,
    bV1_horror, bV2_halloween, bV3_watery,
    b1_refit_lemma, b2_swift_advantage, b3_optimal_design,
    b4_advantageflow, b5_stepwise_credit, b6_failure_cases,
    b7_related_work,
    b_lessons,
]


def build():
    for fn in BUILDERS:
        fn()
    prs.save(OUT)
    print(f"Wrote {OUT}  ·  {len(prs.slides)} slides "
          f"(expected {TOTAL} main + {TOTAL_ALL - TOTAL} backup = {TOTAL_ALL})")


if __name__ == "__main__":
    build()
