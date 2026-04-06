
## 2. `scripts/ppt_generator.py` (Python Script for Gemma 4 ADK)

```python
#!/usr/bin/env python3
"""
PPT Generator Script for Gemma 4 ADK Skill
Creates professional PowerPoint presentations using python-pptx
"""

import argparse
import json
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from typing import List, Dict, Any, Optional, Tuple


class PPTMaker:
    """Professional PowerPoint presentation generator"""
    
    THEMES = {
        "modern": {
            "primary": (30, 58, 138),
            "secondary": (59, 130, 246),
            "accent": (245, 158, 11),
            "bg": (255, 255, 255),
            "text": (31, 41, 55),
            "light_bg": (243, 244, 246)
        },
        "dark": {
            "primary": (17, 24, 39),
            "secondary": (55, 65, 81),
            "accent": (16, 185, 129),
            "bg": (249, 250, 251),
            "text": (17, 24, 39),
            "light_bg": (229, 231, 235)
        },
        "creative": {
            "primary": (124, 58, 237),
            "secondary": (236, 72, 153),
            "accent": (251, 191, 36),
            "bg": (255, 255, 255),
            "text": (31, 41, 55),
            "light_bg": (254, 243, 199)
        }
    }
    
    def __init__(self, theme: str = "modern"):
        self.prs = Presentation()
        self.theme = self.THEMES.get(theme, self.THEMES["modern"])
        self._setup_slide_size()
    
    def _setup_slide_size(self):
        """Set 16:9 widescreen format"""
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
    
    def _rgb_color(self, rgb: Tuple[int, int, int]) -> RgbColor:
        """Create RGB color from tuple"""
        return RgbColor(rgb[0], rgb[1], rgb[2])
    
    def _add_shape(self, slide, shape_type, left, top, width, height, 
                   fill_color: Optional[Tuple] = None, line_color: Optional[Tuple] = None):
        """Add shape with optional fill and line colors"""
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self._rgb_color(fill_color)
        else:
            shape.fill.background()
            
        if line_color:
            shape.line.color.rgb = self._rgb_color(line_color)
        else:
            shape.line.fill.background()
            
        return shape
    
    def _add_textbox(self, slide, left, top, width, height, 
                     text: str, font_size: int = 18, 
                     bold: bool = False, color: Optional[Tuple] = None,
                     align: PP_ALIGN = PP_ALIGN.LEFT, font_name: str = "Calibri"):
        """Add styled text box"""
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = text
        p.alignment = align
        
        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.name = font_name
            run.font.color.rgb = self._rgb_color(color or self.theme["text"])
        
        return box
    
    def create_title_slide(self, title: str, subtitle: Optional[str] = None, 
                          presenter: Optional[str] = None):
        """Create opening title slide"""
        slide_layout = self.prs.slide_layouts[6]  # Blank
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Header bar
        self._add_shape(slide, MSO_SHAPE.RECTANGLE, 
                       Inches(0), Inches(0), 
                       Inches(13.333), Inches(1.2), 
                       self.theme["primary"])
        
        # Accent line
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                       Inches(0), Inches(1.2),
                       Inches(13.333), Inches(0.1),
                       self.theme["accent"])
        
        # Title
        self._add_textbox(slide, Inches(0.5), Inches(2.5), 
                         Inches(12.333), Inches(1.5),
                         title, font_size=44, bold=True, 
                         color=self.theme["primary"], align=PP_ALIGN.CENTER)
        
        # Subtitle
        if subtitle:
            self._add_textbox(slide, Inches(0.5), Inches(4.2),
                             Inches(12.333), Inches(1),
                             subtitle, font_size=24, 
                             color=self.theme["secondary"], align=PP_ALIGN.CENTER)
        
        # Presenter
        if presenter:
            self._add_textbox(slide, Inches(0.5), Inches(6),
                             Inches(12.333), Inches(0.8),
                             f"Presented by: {presenter}", 
                             font_size=16, align=PP_ALIGN.CENTER)
        
        return slide
    
    def add_content_slide(self, title: str, content_items: List[str]):
        """Add bullet point content slide"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Header
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                       Inches(0), Inches(0),
                       Inches(13.333), Inches(1.3),
                       self.theme["primary"])
        
        # Title
        self._add_textbox(slide, Inches(0.5), Inches(0.3),
                         Inches(12.333), Inches(0.8),
                         title, font_size=32, bold=True, 
                         color=(255, 255, 255))
        
        # Content background
        self._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       Inches(0.5), Inches(1.6),
                       Inches(12.333), Inches(5.5),
                       self.theme["light_bg"])
        
        # Content
        content_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.9), Inches(11.733), Inches(5))
        tf = content_box.text_frame
        tf.word_wrap = True
        
        for i, item in enumerate(content_items):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            p.text = f"• {item}"
            p.level = 0
            p.font.size = Pt(20)
            p.font.name = "Calibri"
            p.font.color.rgb = self._rgb_color(self.theme["text"])
            p.space_after = Pt(12)
        
        return slide
    
    def add_two_column_slide(self, title: str, 
                            left_title: str, left_items: List[str],
                            right_title: str, right_items: List[str]):
        """Create side-by-side content layout"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Header
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                       Inches(0), Inches(0),
                       Inches(13.333), Inches(1.2),
                       self.theme["primary"])
        
        # Main title
        self._add_textbox(slide, Inches(0.5), Inches(0.25),
                         Inches(12.333), Inches(0.7),
                         title, font_size=28, bold=True, 
                         color=(255, 255, 255))
        
        # Left column
        self._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       Inches(0.5), Inches(1.5),
                       Inches(6), Inches(5.5),
                       (255, 255, 255))
        
        self._add_textbox(slide, Inches(0.7), Inches(1.7),
                         Inches(5.6), Inches(0.6),
                         left_title, font_size=22, bold=True, 
                         color=self.theme["primary"])
        
        left_content = slide.shapes.add_textbox(
            Inches(0.7), Inches(2.4), Inches(5.6), Inches(4.4))
        tf = left_content.text_frame
        tf.word_wrap = True
        
        for item in left_items:
            p = tf.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(16)
            p.space_after = Pt(8)
        
        # Right column
        self._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                       Inches(6.8), Inches(1.5),
                       Inches(6), Inches(5.5),
                       self.theme["light_bg"])
        
        self._add_textbox(slide, Inches(7), Inches(1.7),
                         Inches(5.6), Inches(0.6),
                         right_title, font_size=22, bold=True, 
                         color=self.theme["secondary"])
        
        right_content = slide.shapes.add_textbox(
            Inches(7), Inches(2.4), Inches(5.6), Inches(4.4))
        tf = right_content.text_frame
        tf.word_wrap = True
        
        for item in right_items:
            p = tf.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(16)
            p.space_after = Pt(8)
        
        return slide
    
    def add_section_divider(self, section_title: str, 
                           section_number: Optional[int] = None):
        """Create section transition slide"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Full background
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                       Inches(0), Inches(0),
                       Inches(13.333), Inches(7.5),
                       self.theme["primary"])
        
        # Decorative circle
        self._add_shape(slide, MSO_SHAPE.OVAL,
                       Inches(10), Inches(-2),
                       Inches(5), Inches(5),
                       self.theme["secondary"])
        
        # Section number
        if section_number:
            num_text = f"0{section_number}" if section_number < 10 else str(section_number)
            self._add_textbox(slide, Inches(0.5), Inches(2.5),
                             Inches(2), Inches(1.5),
                             num_text, font_size=72, bold=True, 
                             color=self.theme["accent"])
        
        # Section title
        self._add_textbox(slide, Inches(0.5), Inches(4),
                         Inches(12.333), Inches(1.5),
                         section_title, font_size=48, bold=True, 
                         color=(255, 255, 255))
        
        # Accent line
        self._add_shape(slide, MSO_SHAPE.RECTANGLE,
                       Inches(0.5), Inches(5.5),
                       Inches(3), Inches(0.15),
                       self.theme["accent"])
        
        return slide
    
    def save(self, filename: str = "presentation.pptx") -> str:
        """Save presentation to file"""
        # Ensure output directory exists
        output_dir = os.path.join(os.path.dirname(__file__), "..", "output")
        os.makedirs(output_dir, exist_ok=True)
        
        output_path = os.path.join(output_dir, filename)
        self.prs.save(output_path)
        return output_path


def create_presentation_from_config(config: Dict[str, Any]) -> str:
    """
    Create presentation from configuration dictionary
    
    Config format:
    {
        "title": "Presentation Title",
        "theme": "modern|dark|creative",
        "presenter": "Optional Name",
        "slides": [
            {"type": "title", "subtitle": "..."},
            {"type": "content", "title": "...", "content": ["...", "..."]},
            {"type": "two_column", "title": "...", "left": {...}, "right": {...}},
            {"type": "section", "title": "...", "number": 1}
        ]
    }
    """
    maker = PPTMaker(theme=config.get("theme", "modern"))
    
    slides = config.get("slides", [])
    
    # First slide should be title
    if slides and slides[0].get("type") == "title":
        title_slide = slides[0]
        maker.create_title_slide(
            title=config["title"],
            subtitle=title_slide.get("subtitle"),
            presenter=config.get("presenter")
        )
        slides = slides[1:]
    
    # Process remaining slides
    for slide in slides:
        slide_type = slide.get("type", "content")
        
        if slide_type == "content":
            maker.add_content_slide(
                title=slide["title"],
                content_items=slide.get("content", [])
            )
        elif slide_type == "two_column":
            maker.add_two_column_slide(
                title=slide["title"],
                left_title=slide["left"]["title"],
                left_items=slide["left"]["content"],
                right_title=slide["right"]["title"],
                right_items=slide["right"]["content"]
            )
        elif slide_type == "section":
            maker.add_section_divider(
                section_title=slide["title"],
                section_number=slide.get("number")
            )
    
    filename = config.get("filename", "presentation.pptx")
    return maker.save(filename)


def main():
    """CLI entry point"""
    parser = argparse.ArgumentParser(description="Generate PowerPoint presentations")
    parser.add_argument("--config", type=str, required=True,
                       help="JSON configuration string or path to JSON file")
    parser.add_argument("--output", type=str, default="presentation.pptx",
                       help="Output filename")
    
    args = parser.parse_args()
    
    # Parse config
    config_str = args.config
    if os.path.exists(config_str):
        with open(config_str, 'r') as f:
            config = json.load(f)
    else:
        config = json.loads(config_str)
    
    # Add output filename if not specified
    if "filename" not in config:
        config["filename"] = args.output
    
    # Generate presentation
    output_path = create_presentation_from_config(config)
    print(f"Presentation saved to: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
