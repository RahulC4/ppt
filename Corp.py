from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Create a new presentation
prs = Presentation()

# --- CORPORATE COLOR PALETTE ---
PRIMARY_COLOR = RGBColor(0, 102, 204)   # Corporate Blue
SECONDARY_COLOR = RGBColor(240, 240, 240)
TEXT_COLOR = RGBColor(40, 40, 40)

# -------------------------------
# SLIDE 0 — TITLE SLIDE
# -------------------------------
slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(slide_layout)

title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Corporate Presentation"
subtitle.text = "Professional Business Template"

for paragraph in title.text_frame.paragraphs:
    for run in paragraph.runs:
        run.font.size = Pt(40)
        run.font.bold = True
        run.font.color.rgb = PRIMARY_COLOR

for paragraph in subtitle.text_frame.paragraphs:
    for run in paragraph.runs:
        run.font.size = Pt(20)
        run.font.color.rgb = TEXT_COLOR

# -------------------------------
# SLIDE 1 — SECTION HEADER
# -------------------------------
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)

slide.shapes.title.text = "Section Title"
body = slide.placeholders[1]
body.text = "Use this slide for section introductions."

# -------------------------------
# SLIDE 2 — CONTENT + IMAGE
# -------------------------------
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)

slide.shapes.title.text = "Key Business Topic"

tf = slide.placeholders[1].text_frame
tf.text = "First key business point"

p = tf.add_paragraph()
p.text = "Second strategic point"
p.level = 1

p = tf.add_paragraph()
p.text = "Third supporting point"
p.level = 1

# -------------------------------
# SLIDE 3 — TWO COLUMN CONTENT
# -------------------------------
slide_layout = prs.slide_layouts[3]
slide = prs.slides.add_slide(slide_layout)

slide.shapes.title.text = "Comparison Slide"

left = slide.placeholders[1]
right = slide.placeholders[2]

left.text = "Advantages\n• Growth\n• Stability\n• Scale"
right.text = "Risks\n• Cost\n• Competition\n• Timing"

# -------------------------------
# SLIDE 4 — THANK YOU
# -------------------------------
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)

slide.shapes.title.text = "Thank You"
slide.placeholders[1].text = "Questions & Discussion"

# Save template
output_path = "templates/corporate_template.pptx"
prs.save(output_path)

print("✅ Corporate template generated at:", output_path)
