"""
PPT Maker Script - Fully Offline PowerPoint Generator
Uses python-pptx to create presentations from a JSON config.

Usage:
    python script.py config.json output.pptx
    python script.py --json '{"title":"...","slides":[...]}' output.pptx

Install:
    pip install python-pptx
"""

import json
import sys
import argparse
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ─── THEMES ────────────────────────────────────────────────────────────────────

THEMES = {
    "dark": {
        "bg":          RGBColor(0x1E, 0x27, 0x61),   # deep navy
        "title_text":  RGBColor(0xFF, 0xFF, 0xFF),   # white
        "body_text":   RGBColor(0xE0, 0xE8, 0xFF),   # light blue-white
        "accent":      RGBColor(0x00, 0xC8, 0xFF),   # cyan accent
        "slide_bg":    RGBColor(0xF4, 0xF6, 0xFF),   # very light lavender
        "slide_text":  RGBColor(0x1A, 0x1A, 0x2E),   # near black
        "bullet_dot":  RGBColor(0x00, 0xC8, 0xFF),   # cyan bullet
    },
    "light": {
        "bg":          RGBColor(0xFF, 0xFF, 0xFF),
        "title_text":  RGBColor(0x1A, 0x1A, 0x2E),
        "body_text":   RGBColor(0x22, 0x22, 0x22),
        "accent":      RGBColor(0x21, 0x5C, 0xBF),
        "slide_bg":    RGBColor(0xF9, 0xF9, 0xF9),
        "slide_text":  RGBColor(0x1A, 0x1A, 0x2E),
        "bullet_dot":  RGBColor(0x21, 0x5C, 0xBF),
    },
    "blue": {
        "bg":          RGBColor(0x06, 0x5A, 0x82),
        "title_text":  RGBColor(0xFF, 0xFF, 0xFF),
        "body_text":   RGBColor(0xD0, 0xEC, 0xFF),
        "accent":      RGBColor(0x02, 0xC3, 0x9A),
        "slide_bg":    RGBColor(0xEA, 0xF4, 0xFF),
        "slide_text":  RGBColor(0x06, 0x2A, 0x40),
        "bullet_dot":  RGBColor(0x02, 0xC3, 0x9A),
    },
}


# ─── HELPERS ───────────────────────────────────────────────────────────────────

def set_slide_background(slide, color: RGBColor):
    """Fill slide background with a solid color."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=RGBColor(0, 0, 0),
                 align=PP_ALIGN.LEFT, word_wrap=True):
    """Add a styled text box to the slide."""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txBox


def add_bullet_points(slide, bullets, left, top, width, height,
                      font_size=16, text_color=RGBColor(0, 0, 0),
                      dot_color=RGBColor(0, 120, 200)):
    """Add bullet points with colored dots."""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.space_before = Pt(4)

        # Colored bullet dot
        dot_run = p.add_run()
        dot_run.text = "● "
        dot_run.font.size = Pt(font_size - 2)
        dot_run.font.color.rgb = dot_color
        dot_run.font.name = "Calibri"

        # Bullet text
        text_run = p.add_run()
        text_run.text = bullet
        text_run.font.size = Pt(font_size)
        text_run.font.color.rgb = text_color
        text_run.font.name = "Calibri"


def add_accent_bar(slide, left, top, width, height, color: RGBColor):
    """Add a colored rectangle accent bar."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()  # no border


# ─── SLIDE BUILDERS ────────────────────────────────────────────────────────────

def build_title_slide(prs, slide_data, theme):
    """Title slide: big title + subtitle on themed background."""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, theme["bg"])

    # Accent bar on left edge
    add_accent_bar(slide, 0, 0, 0.12, 7.5, theme["accent"])

    # Decorative bottom strip
    add_accent_bar(slide, 0, 6.8, 10, 0.12, theme["accent"])

    # Title
    add_text_box(
        slide,
        slide_data.get("title", "Untitled"),
        left=0.5, top=2.2, width=9, height=1.8,
        font_size=44, bold=True,
        color=theme["title_text"],
        align=PP_ALIGN.CENTER
    )

    # Subtitle
    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        add_text_box(
            slide,
            subtitle,
            left=0.5, top=4.1, width=9, height=1.0,
            font_size=22, bold=False,
            color=theme["body_text"],
            align=PP_ALIGN.CENTER
        )

    return slide


def build_content_slide(prs, slide_data, theme):
    """Content slide: heading + bullet points."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, theme["slide_bg"])

    # Top header bar
    add_accent_bar(slide, 0, 0, 10, 1.2, theme["bg"])

    # Slide title in header
    add_text_box(
        slide,
        slide_data.get("title", ""),
        left=0.3, top=0.1, width=9.4, height=1.0,
        font_size=28, bold=True,
        color=theme["title_text"],
        align=PP_ALIGN.LEFT
    )

    # Accent side bar
    add_accent_bar(slide, 0, 1.2, 0.08, 6.0, theme["accent"])

    # Bullets
    bullets = slide_data.get("bullets", [])
    add_bullet_points(
        slide, bullets,
        left=0.4, top=1.4, width=9.2, height=5.6,
        font_size=17,
        text_color=theme["slide_text"],
        dot_color=theme["bullet_dot"]
    )

    return slide


def build_two_column_slide(prs, slide_data, theme):
    """Two-column slide: heading + left/right content."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, theme["slide_bg"])

    # Header bar
    add_accent_bar(slide, 0, 0, 10, 1.2, theme["bg"])

    # Title
    add_text_box(
        slide,
        slide_data.get("title", ""),
        left=0.3, top=0.1, width=9.4, height=1.0,
        font_size=28, bold=True,
        color=theme["title_text"],
        align=PP_ALIGN.LEFT
    )

    # Divider line in middle
    add_accent_bar(slide, 4.88, 1.3, 0.06, 5.8, theme["accent"])

    # Left column label
    add_text_box(
        slide, "◀  Left",
        left=0.3, top=1.3, width=4.4, height=0.4,
        font_size=13, bold=True,
        color=theme["accent"]
    )

    # Right column label
    add_text_box(
        slide, "Right  ▶",
        left=5.1, top=1.3, width=4.5, height=0.4,
        font_size=13, bold=True,
        color=theme["accent"]
    )

    # Left bullets
    left_bullets = slide_data.get("left", [])
    add_bullet_points(
        slide, left_bullets,
        left=0.3, top=1.8, width=4.4, height=5.0,
        font_size=16,
        text_color=theme["slide_text"],
        dot_color=theme["bullet_dot"]
    )

    # Right bullets
    right_bullets = slide_data.get("right", [])
    add_bullet_points(
        slide, right_bullets,
        left=5.1, top=1.8, width=4.5, height=5.0,
        font_size=16,
        text_color=theme["slide_text"],
        dot_color=theme["bullet_dot"]
    )

    return slide


def build_closing_slide(prs, slide_data, theme):
    """Closing/thank-you slide."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, theme["bg"])

    # Accent bars
    add_accent_bar(slide, 0, 0, 10, 0.12, theme["accent"])
    add_accent_bar(slide, 0, 7.38, 10, 0.12, theme["accent"])

    # Main message
    add_text_box(
        slide,
        slide_data.get("title", "Thank You"),
        left=0.5, top=2.5, width=9, height=1.5,
        font_size=48, bold=True,
        color=theme["title_text"],
        align=PP_ALIGN.CENTER
    )

    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        add_text_box(
            slide,
            subtitle,
            left=0.5, top=4.2, width=9, height=1.0,
            font_size=20, bold=False,
            color=theme["body_text"],
            align=PP_ALIGN.CENTER
        )

    return slide


# ─── SLIDE DISPATCHER ──────────────────────────────────────────────────────────

SLIDE_BUILDERS = {
    "title":      build_title_slide,
    "content":    build_content_slide,
    "two_column": build_two_column_slide,
    "closing":    build_closing_slide,
}


# ─── MAIN ──────────────────────────────────────────────────────────────────────

def create_presentation(config: dict, output_path: str):
    """Build the full presentation from config dict."""

    theme_name = config.get("theme", "dark").lower()
    theme = THEMES.get(theme_name, THEMES["dark"])

    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)

    slides_data = config.get("slides", [])

    if not slides_data:
        print("⚠️  No slides found in config. Add at least one slide.")
        sys.exit(1)

    for i, slide_data in enumerate(slides_data):
        slide_type = slide_data.get("type", "content").lower()
        builder = SLIDE_BUILDERS.get(slide_type, build_content_slide)
        builder(prs, slide_data, theme)
        print(f"  ✅ Slide {i+1}/{len(slides_data)}: [{slide_type}] {slide_data.get('title','')}")

    prs.save(output_path)
    print(f"\n🎉 Presentation saved: {output_path}")
    print(f"   Slides: {len(slides_data)}  |  Theme: {theme_name}")


def main():
    parser = argparse.ArgumentParser(
        description="Create a PowerPoint presentation from a JSON config."
    )
    parser.add_argument(
        "config_or_flag",
        nargs="?",
        help="Path to JSON config file, OR use --json flag"
    )
    parser.add_argument(
        "output",
        nargs="?",
        default="output.pptx",
        help="Output .pptx filename (default: output.pptx)"
    )
    parser.add_argument(
        "--json",
        dest="json_str",
        help="Inline JSON config string"
    )

    args = parser.parse_args()

    # Load config
    if args.json_str:
        try:
            config = json.loads(args.json_str)
        except json.JSONDecodeError as e:
            print(f"❌ Invalid JSON: {e}")
            sys.exit(1)
        output_path = args.config_or_flag or args.output
    elif args.config_or_flag:
        try:
            with open(args.config_or_flag, "r", encoding="utf-8") as f:
                config = json.load(f)
        except FileNotFoundError:
            print(f"❌ Config file not found: {args.config_or_flag}")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"❌ Invalid JSON in config file: {e}")
            sys.exit(1)
        output_path = args.output
    else:
        print("❌ Provide a config file path or use --json '...'")
        parser.print_help()
        sys.exit(1)

    print(f"\n🚀 Building: {config.get('title', 'Presentation')}")
    print(f"   Theme: {config.get('theme', 'dark')}  |  Slides: {len(config.get('slides', []))}\n")

    create_presentation(config, output_path)


if __name__ == "__main__":
    main()
