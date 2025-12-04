title_tf = agenda_slide.shapes.title.text_frame
p = title_tf.paragraphs[0]
p.font.size = Pt(32)
p.font.color.rgb = RGBColor(0, 102, 204)  # Corporate Blue
p.alignment = PP_ALIGN.LEFT

title_tf = slide.shapes.title.text_frame
p = title_tf.paragraphs[0]
p.font.size = Pt(32)
p.font.color.rgb = RGBColor(0, 102, 204)  # Corporate Blue
p.alignment = PP_ALIGN.LEFT

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

p.font.bold = True

title_shape = title_slide.shapes.title
title_shape.text = slides[0]["title"]

for p in title_shape.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(0, 102, 204)  # Corporate Blue

thank_shape = thank_slide.shapes.title
thank_shape.text = "Thank You"

for p in thank_shape.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(0, 102, 204)  # Corporate Blue
    p.font.bold = True


from pptx.dml.color import RGBColor

def apply_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(234, 242, 251)  # #EAF2FB


apply_background(slide)

title_slide = prs.slides.add_slide(prs.slide_layouts[0])
apply_background(title_slide)

agenda_slide = prs.slides.add_slide(prs.slide_layouts[1])
apply_background(agenda_slide)


slide = prs.slides.add_slide(prs.slide_layouts[1])
apply_background(slide)

thank_slide = prs.slides.add_slide(prs.slide_layouts[1])
apply_background(thank_slide)


def extract_title_from_ppt(ppt_path):
    try:
        prs = Presentation(ppt_path)
        if prs.slides and prs.slides[0].shapes.title:
            return prs.slides[0].shapes.title.text.strip()
    except Exception:
        pass
    return "Generated Presentation"

ppt_title = extract_title_from_ppt(ppt_path)
display_name = f"{ppt_title}.pptx"

st.session_state["generated_ppts"].insert(
    0,   # ✅ keeps newest on top
    {"path": ppt_path, "name": display_name}
)


ppt_title = extract_title_from_ppt(ppt_path)

timestamp = datetime.now().strftime("%d_%b_%H-%M")
display_name = f"{ppt_title}_{timestamp}.pptx"
