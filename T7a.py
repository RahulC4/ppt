def add_agenda_slide(prs, titles, image_required):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background_color(slide)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(8), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Agenda"
    p.font.size = Pt(32)
    p.font.bold = True

    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6), Inches(4.5))
    tf = body.text_frame
    tf.word_wrap = True

    for t in titles:
        p = tf.add_paragraph()
        p.text = t
        p.font.size = Pt(20)
        p.level = 0   # ✅ ✅ ✅ BULLET FIX

    if image_required:
        img = generate_visual_image("agenda corporate business")
        if img:
            slide.shapes.add_picture(img, Inches(6.8), Inches(2), width=Inches(2.5))
