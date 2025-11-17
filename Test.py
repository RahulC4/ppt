import os
import base64
import requests
from openai import AzureOpenAI
from utils import get_env


# ------------------------------------------------------------
# Load environment variables
# ------------------------------------------------------------
TEXT_ENDPOINT = get_env("OPENAI_API_BASE", required=True)
TEXT_KEY = get_env("OPENAI_API_KEY", required=True)

IMAGE_ENDPOINT = get_env("IMAGE_API_BASE", required=True)
IMAGE_KEY = get_env("IMAGE_API_KEY", required=True)

API_VERSION = get_env("OPENAI_API_VERSION", "2024-05-01-preview")

CHAT_MODEL = get_env("CHAT_MODEL", "gpt-5-mini")
EMBED_MODEL = get_env("EMBEDDING_MODEL", "text-embedding-3-large")
IMAGE_MODEL = get_env("DALLE_MODEL", "gpt-image-1-mini")


# ------------------------------------------------------------
# Create clients
# ------------------------------------------------------------
text_client = AzureOpenAI(
    azure_endpoint=TEXT_ENDPOINT,
    api_key=TEXT_KEY,
    api_version=API_VERSION
)

image_client = AzureOpenAI(
    azure_endpoint=IMAGE_ENDPOINT,
    api_key=IMAGE_KEY,
    api_version=API_VERSION
)


print("\n===============================")
print(" TEST 1 — TEXT CHAT COMPLETION ")
print("===============================\n")

try:
    resp = text_client.chat.completions.create(
        model=CHAT_MODEL,
        messages=[{"role": "user", "content": "Say hello in one short sentence."}],
        max_tokens=20
    )
    print("✔ Chat model responded successfully:")
    print("Response:", resp.choices[0].message.content)
except Exception as e:
    print("❌ Chat model FAILED:", e)


print("\n===============================")
print(" TEST 2 — EMBEDDING GENERATION ")
print("===============================\n")

try:
    emb = text_client.embeddings.create(
        model=EMBED_MODEL,
        input="This is a test sentence for embeddings."
    )
    vec = emb.data[0].embedding
    print("✔ Embedding generated successfully.")
    print("Embedding length:", len(vec))
except Exception as e:
    print("❌ Embedding model FAILED:", e)


print("\n===============================")
print(" TEST 3 — IMAGE GENERATION ")
print("===============================\n")

try:
    img_resp = image_client.images.generate(
        model=IMAGE_MODEL,
        prompt="Simple blue square icon",
        size="512x512"
    )

    # Check URL or b64 output
    img_item = img_resp.data[0]

    if "url" in img_item and img_item["url"]:
        print("✔ Image generated successfully (URL mode).")
        print("Image URL:", img_item["url"])

        # Optional: download for local verification
        r = requests.get(img_item["url"])
        out_path = "test_image.png"
        with open(out_path, "wb") as f:
            f.write(r.content)
        print("Image saved as:", out_path)

    elif "b64_json" in img_item:
        print("✔ Image generated successfully (Base64 mode).")
        b64 = img_item["b64_json"]
        data = base64.b64decode(b64)
        out_path = "test_image.png"
        with open(out_path, "wb") as f:
            f.write(data)
        print("Image saved as:", out_path)

    else:
        print("❌ No usable image returned:", img_item)

except Exception as e:
    print("❌ Image model FAILED:", e)


print("\n===============================")
print("   TESTING COMPLETE ")
print("===============================\n")
