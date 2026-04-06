# PPT Maker

Create professional PowerPoint presentations using python-pptx library with intelligent content structuring and modern design aesthetics.

## Usage

Use this skill when the user wants to:
- Create a new PowerPoint presentation from scratch
- Convert text content, outlines, or structured data into slides
- Design presentations with specific themes, colors, or layouts
- Generate charts, diagrams, or visual elements in slides
- Batch create slides from structured content (JSON, CSV, etc.)

## Requirements

- python-pptx library for PowerPoint generation
- Pillow (PIL) for image processing if needed
- Standard Python libraries (os, json, etc.)

## Best Practices

1. **Content Structure**: Always organize content hierarchically - Title Slide → Agenda → Content Sections → Conclusion/Thank You
2. **Visual Hierarchy**: Use font sizes strategically (Titles 32-44pt, Headers 24-28pt, Body 18-20pt)
3. **Color Schemes**: Apply consistent color palettes (primary, secondary, accent colors)
4. **Slide Layouts**: Match layout to content type (Title, Title and Content, Two Content, Blank for custom)
5. **Charts**: Use python-pptx chart capabilities for data visualization
6. **Images**: Handle image insertion with proper aspect ratio preservation
7. **Spacing**: Maintain consistent margins and text box positioning

## Design Templates

### Modern Corporate
- Primary: #1E3A8A (Deep Blue)
- Secondary: #3B82F6 (Bright Blue)  
- Accent: #F59E0B (Amber)
- Background: #FFFFFF
- Text: #1F2937

### Minimal Dark
- Primary: #111827 (Near Black)
- Secondary: #374151 (Dark Gray)
- Accent: #10B981 (Emerald)
- Background: #F9FAFB
- Text: #111827

### Creative Gradient
- Primary: #7C3AED (Violet)
- Secondary: #EC4899 (Pink)
- Accent: #FBBF24 (Amber)
- Background: gradient effect
- Text: #1F2937

## Core Functions

### create_presentation(title, subtitle=None, theme="modern")
Initialize a new presentation with theme settings.

### add_title_slide(prs, title, subtitle, presenter=None)
Create opening slide with branding.

### add_content_slide(prs, title, content_list, layout="title_and_content")
Add bullet point content slides.

### add_two_column_slide(prs, title, left_content, right_content)
Side-by-side content layout.

### add_image_slide(prs, title, image_path, caption=None)
Insert image with optional text overlay.

### add_chart_slide(prs, title, chart_data, chart_type="bar")
Create data visualization slides.

### add_section_header(prs, section_title)
Divider slides between sections.

### save_presentation(prs, filename)
Export to .pptx file.

## Example Workflow

1. Parse user content/requirements
2. Design slide outline and structure
3. Initialize presentation with appropriate theme
4. Generate slides sequentially
5. Apply consistent formatting
6. Save and provide download link

## Output Format

Always return:
- Python script using python-pptx
- Brief explanation of slide structure
- Instructions for customization
