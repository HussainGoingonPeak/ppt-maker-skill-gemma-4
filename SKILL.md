---
name: ppt-maker
description: Creates PowerPoint presentations fully offline using python-pptx. Use this skill when the user wants to make slides, a presentation, a pitch deck, or any .pptx file from a text prompt. Trigger whenever the user mentions slides, presentation, deck, PowerPoint, or pptx.
---

# PPT Maker Skill

Creates fully offline PowerPoint (.pptx) presentations from a user prompt using the `python-pptx` library.

## How to Use This Skill

1. Read the user's prompt carefully — extract: **topic, number of slides, style/tone**
2. Plan slide content (title, bullet points, layout per slide)
3. Run `script.py` with a JSON config describing the slides
4. Return the generated `.pptx` file to the user

---

## Step 1 — Install Dependency

```bash
pip install python-pptx
```

---

## Step 2 — Prepare the JSON Config

Build a JSON object with this structure:

```json
{
  "title": "Presentation Title",
  "author": "Your Name",
  "theme": "dark",
  "slides": [
    {
      "type": "title",
      "title": "Main Title",
      "subtitle": "Subtitle or tagline here"
    },
    {
      "type": "content",
      "title": "Slide Heading",
      "bullets": [
        "First point",
        "Second point",
        "Third point"
      ]
    },
    {
      "type": "two_column",
      "title": "Comparison Slide",
      "left": ["Left point 1", "Left point 2"],
      "right": ["Right point 1", "Right point 2"]
    },
    {
      "type": "closing",
      "title": "Thank You",
      "subtitle": "Questions? Contact: example@email.com"
    }
  ]
}
```

### Slide Types

| Type | Description |
|------|-------------|
| `title` | Opening title slide with subtitle |
| `content` | Heading + bullet points |
| `two_column` | Two side-by-side content columns |
| `closing` | Final thank-you or closing slide |

### Themes

| Theme | Style |
|-------|-------|
| `dark` | Dark navy background, white text (professional) |
| `light` | White background, dark text (clean/minimal) |
| `blue` | Corporate blue gradient feel |

---

## Step 3 — Run the Script

```bash
python script.py config.json output.pptx
```

Or pass JSON inline:

```bash
python script.py --json '{"title":"My Talk","slides":[...]}' output.pptx
```

---

## Design Rules

- **Title slides**: Large bold title, smaller subtitle, strong background color
- **Content slides**: Max 5-6 bullets per slide — keep it readable
- **Two-column**: Use for comparisons, pros/cons, before/after
- **Closing slides**: Simple, clean — just a thank you + contact
- **Fonts**: Calibri for body, Calibri Bold for headings
- **Never** put more than 6 lines of text on one slide

---

## Full Example Workflow

User says: *"Make a 5-slide presentation about AI in healthcare"*

```bash
pip install python-pptx
python script.py --json '{
  "title": "AI in Healthcare",
  "author": "Claude",
  "theme": "dark",
  "slides": [
    {"type": "title", "title": "AI in Healthcare", "subtitle": "Transforming Patient Care"},
    {"type": "content", "title": "What is AI in Healthcare?", "bullets": ["Machine learning for diagnosis", "Predictive analytics", "Medical imaging analysis", "Drug discovery acceleration"]},
    {"type": "content", "title": "Key Benefits", "bullets": ["Faster diagnosis", "Reduced human error", "Lower costs", "Personalized treatment"]},
    {"type": "two_column", "title": "AI vs Traditional Methods", "left": ["Slower diagnosis", "Limited data processing", "Manual record keeping"], "right": ["Real-time analysis", "Millions of data points", "Automated EHR systems"]},
    {"type": "closing", "title": "Thank You", "subtitle": "The future of healthcare is intelligent."}
  ]
}' ai_healthcare.pptx
```
