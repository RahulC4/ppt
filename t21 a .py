def generate_presentation(
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

    # ✅ 3. HARD STOP if dataset does not match
    if not refs:
        return None, {
            "error": True,
            "message": (
                "No relevant content found in your uploaded PPTs for this prompt. "
                "Please rephrase using dataset-related topics."
            ),
        }

    # ✅ 4. Build reference text only from VALID refs
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    # ✅ 5. LLM is now forced to stay inside dataset
    plan = call_llm_plan(prompt, reference_text, num_slides)

    if not plan:
        return None, {
            "error": True,
            "message": "Not enough relevant dataset content to generate slides.",
        }

    ppt_path = build_corporate_ppt(plan, image_required=image_required)
    total_slides = len(plan) + 3

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(ppt_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": total_slides,
        "ppt_file": fname,
        "image_required": image_required,
        "template_style": template_style,
        "error": False,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode("utf-8"), f"logs/{fname}.json")
    return ppt_path, log
