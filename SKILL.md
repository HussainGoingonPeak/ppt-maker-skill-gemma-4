---
name: ppt-maker
description: Creates PowerPoint presentations using Gemma 4. Use this skill when the user wants to make slides, a deck, or a presentation.
---

# PPT Maker
## Overview

This skill enables the agent to generate professional PowerPoint presentations
using the python-pptx library. It supports multiple themes, slide layouts,
charts, images, and structured content conversion.

## When to Use

Use this skill when:
- User requests a PowerPoint presentation
- Converting text outlines to slides
- Creating data visualizations in PPT format
- Generating templates or reports
- Batch slide creation from structured data

## Workflow

### 1. Analyze Requirements
- Determine presentation topic and audience
- Identify number of slides needed
- Choose appropriate theme (modern, dark, creative)
- Plan slide structure (title, content, two-column, charts, images)

### 2. Design Structure
Create a logical flow:
- **Slide 1**: Title slide (title, subtitle, presenter)
- **Slide 2**: Agenda/Overview
- **Slides 3-N**: Content sections
- **Final Slide**: Conclusion/Thank you/Q&A

### 3. Generate Presentation
Use the `scripts/ppt_generator.py` script to create the PPT file.

### 4. Validate Output
Ensure:
- Consistent formatting across slides
- Proper text wrapping and alignment
- Color scheme adherence
- File saved successfully

## Available Themes

| Theme | Primary | Secondary | Accent | Use Case |
|-------|---------|-----------|--------|----------|
| modern | Deep Blue (#1E3A8A) | Bright Blue (#3B82F6) | Amber (#F59E0B) | Corporate/Business |
| dark | Near Black (#111827) | Dark Gray (#374151) | Emerald (#10B981) | Technical/Developer |
| creative | Violet (#7C3AED) | Pink (#EC4899) | Amber (#FBBF24) | Marketing/Creative |

## Slide Types

1. **Title Slide**: Main title, subtitle, optional presenter name
2. **Content Slide**: Title + bullet point list
3. **Two-Column Slide**: Side-by-side comparison or related content
4. **Section Divider**: Visual break between major sections
5. **Image Slide**: Title + image with optional caption
6. **Chart Slide**: Data visualization (bar, line, pie charts)

## Best Practices

- **Font Sizes**: Titles 32-44pt, Headers 24-28pt, Body 18-20pt
- **Color Contrast**: Ensure text is readable against backgrounds
- **White Space**: Don't overcrowd slides; use margins
- **Consistency**: Use same theme throughout presentation
- **Bullet Points**: Max 6 points per slide, max 2 lines per point

## Script Usage

Call the PPT generator script with JSON configuration:

```bash
python scripts/ppt_generator.py --config '{"title": "My Presentation", "theme": "modern", "slides": [...]}'
