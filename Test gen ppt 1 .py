"""
Test script to validate generate_ppt.py end-to-end.
"""

import os
from generate_ppt import generate_presentation
from pptx import Presentation

def test_generate():
    print("\n=== TEST: generate_presentation() ===")

    prompt = "Create a 3-slide corporate presentation about AI in healthcare."
    print(f"Prompt: {prompt}")

    try:
        ppt_path, log = generate_presentation(
            prompt=prompt,
            style="Auto",
            requested_num_slides=None,   # let LLM detect
            theme=None,                  # let LLM decide
            tag_filters=None
        )
    except Exception as e:
        print("\n❌ ERROR: generate_presentation crashed\n", e)
        return

    print("\n✔ generate_presentation() completed")
    print("Generated file:", ppt_path)
    print("Log:", log)

    # Validate file exists
    if not os.path.exists(ppt_path):
        print("\n❌ PPT file was NOT created.")
        return

    print("✔ PPT file exists.")

    # Validate slide count
    try:
        prs = Presentation(ppt_path)
        print(f"Slides generated: {len(prs.slides)}")
    except Exception as e:
        print("\n❌ Could not open generated PPT:", e)
        return
    
    print("\n=== TEST COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    test_generate()
