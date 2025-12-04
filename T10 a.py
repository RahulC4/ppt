def generate_presentation_auto(
    prompt,
    requested_num_slides=5,
    tag_filters=None,
    template_style=None,
    image_required=False,
):
    # ✅ 1. Run semantic search
    raw_refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

    # ✅ 2. Apply SAME similarity threshold as preview
    SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

    refs = [
        r for r in raw_refs
        if r.get("score") is None or r["score"] <= SIMILARITY_THRESHOLD
    ]

    # ✅ 3. HARD STOP if no relevant dataset match
    if not refs:
        return None, {
            "error": True,
            "message": (
                "No relevant content found in your uploaded PPTs for this prompt. "
                "Please rephrase using dataset-related topics."
            ),
        }

    # ✅ 4. Build reference text ONLY from valid refs
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected_slides = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected_slides or 5

    # ✅ 5. LLM is now forced to use ONLY dataset content
    plan = call_llm_plan_auto(prompt, reference_text, num_slides)

    if not plan:
        return None, {
            "error": True,
            "message": "Not enough relevant dataset content to generate slides.",
        }

    slides = []
    for sp in plan:
        img_path = generate_visual_image(sp["title"]) if image_required else None
        slides.append({
            "title": sp["title"],
            "bullets": sp["bullets"],
            "image_path": img_path,
        })

    agenda_titles = [s["title"] for s in slides]

    out_path = build_ppt(slides, agenda_titles, image_required)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "error": False,
        "image_required": image_required,
        "template_style": template_style,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
