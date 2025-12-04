# ✅ FINAL AUTO PIPELINE (ONLY TITLE FIX CHANGED)
def generate_presentation_auto(
    prompt,
    requested_num_slides=5,
    tag_filters=None,
    template_style=None,
    image_required=False,
):
    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan_auto(prompt, reference_text, num_slides)

    slides = []
    for sp in plan:
        img_path = generate_visual_image(sp["title"]) if image_required else None
        slides.append({
            "title": sp["title"],
            "bullets": sp["bullets"],
            "image_path": img_path,
        })

    prs = Presentation()

    # ✅ ✅ ✅ TITLE NOW COMES FROM FIRST GENERATED SLIDE
    add_title_slide(prs, slides[0]["title"])

    # ✅ AGENDA
    agenda_titles = [s["title"] for s in slides]
    add_agenda_slide(prs, agenda_titles, image_required)

    # ✅ CONTENT
    for sp in slides:
        build_content_slide(prs, sp)

    # ✅ THANK YOU
    add_thank_you(prs)

    out_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out_path)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides) + 3,
        "ppt_file": fname,
        "error": False,
        "image_required": image_required,
        "template_style": template_style,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
