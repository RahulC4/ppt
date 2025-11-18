best = refs[0] if refs else None
design_meta = None

if best:
    try:
        matched_idx = int(best.get("slide_index"))
    except:
        matched_idx = 0

    design_meta = extract_design_for_slide(
        design_context,
        best.get("ppt_name"),
        matched_idx
    )

----------------
best = refs[0] if refs else None
design_meta = None

if best:
    logger.info(f"[MATCH] Trying PPT match: best.ppt_name={best.get('ppt_name')}, slide_index={best.get('slide_index')}")

    try:
        matched_idx = int(best.get("slide_index"))
    except:
        matched_idx = 0

    design_meta = extract_design_for_slide(
        design_context,
        best.get("ppt_name"),
        matched_idx
    )

    logger.info(f"[MATCH] Result design_meta: {design_meta is not None}")


---------------------
def extract_design_for_slide(design_context, matched_ppt_name, matched_slide_idx):


----------

def extract_design_for_slide(design_context, matched_ppt_name, matched_slide_idx):

    logger.info(f"[EXTRACT] Searching for PPT: {os.path.basename(matched_ppt_name)}, slide={matched_slide_idx}")

    for design in design_context:
        logger.info(f"[EXTRACT] Checking JSON ppt_name={design['ppt_name']} → basename={os.path.basename(design['ppt_name'])}")

        if os.path.basename(design["ppt_name"]) == os.path.basename(matched_ppt_name):
            logger.info("[EXTRACT] PPT name matched!")

            for slide in design["slides"]:
                if slide["index"] == matched_slide_idx:
                    logger.info(f"[EXTRACT] Slide index matched! Returning design meta.")
                    return slide

            logger.info("[EXTRACT] Slide index NOT matched")

    logger.info("[EXTRACT] No matching PPT found in design context")
    return None
  
