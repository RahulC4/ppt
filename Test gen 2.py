# test_generate.py
import os
from generate_ppt import generate_presentation
from utils import logger

print("\n==============================")
print(" TEST: generate_presentation()")
print("==============================\n")

try:
    out_path, log = generate_presentation(
        prompt="Create a 5-slide corporate presentation about AI in healthcare architecture and migration.",
        style="Design",       
        tag_filters=["Design"]
    )

    print("\n--- SUCCESS ---")
    print("PPT generated at:", out_path)
    print("Metadata log:\n", log)

except Exception as e:
    print("\n--- FAILED ---")
    logger.exception("Error while generating PPT")
    print("Error:", e)
