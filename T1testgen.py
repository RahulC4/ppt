# test_generate_ppt.py
import os
from generate_ppt import generate_presentation
from utils import logger

if __name__ == "__main__":
    logger.info("🧪 Starting PPT generation test...")

    try:
        prompt = "Create a healthcare migration deck with 5 slides focused on design and testing."

        output_path, metadata = generate_presentation(
            prompt=prompt,
            style="Professional",
            requested_num_slides=None,     # let auto-detect override
            theme=None,                    # let auto-detect handle
            tag_filters=None               # no tag filtering for test
        )

        print("\n================= TEST RESULT =================")
        print(f"PPT generated at local path: {output_path}")
        print("Metadata:")
        for k, v in metadata.items():
            print(f"  {k}: {v}")

        print("\nIf upload succeeded, PPT should also appear in your Azure Blob:")
        print(f"Blob Path: {metadata.get('ppt_file')}")

    except Exception as e:
        logger.exception("❌ Test failed:")
        print(f"Error: {e}")
