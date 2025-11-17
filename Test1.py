import os
from openai import AzureOpenAI
from dotenv import load_dotenv

load_dotenv()

print("\n========== TEST 1: GPT CHAT COMPLETION ==========\n")

try:
    chat_client = AzureOpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        api_version=os.getenv("OPENAI_API_VERSION"),
        azure_endpoint=os.getenv("OPENAI_API_BASE")
    )

    response = chat_client.chat.completions.create(
        model=os.getenv("CHAT_MODEL"),
        messages=[{"role": "user", "content": "Say hello"}],
        max_completion_tokens=50
    )

    print("Chat Response:", response.choices[0].message["content"])
    print("TEST 1 PASSED\n")

except Exception as e:
    print("TEST 1 FAILED:", e)
    print("\n")


print("\n========== TEST 2: EMBEDDING GENERATION ==========\n")

try:
    embedding_client = AzureOpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        api_version=os.getenv("OPENAI_API_VERSION"),
        azure_endpoint=os.getenv("OPENAI_API_BASE")
    )

    # IMPORTANT: No max_tokens for embeddings
    embedding = embedding_client.embeddings.create(
        model=os.getenv("EMBEDDING_MODEL"),
        input="hello world"
    )

    print("Embedding length:", len(embedding.data[0].embedding))
    print("TEST 2 PASSED\n")

except Exception as e:
    print("TEST 2 FAILED:", e)
    print("\n")


print("\n========== TEST 3: IMAGE GENERATION ==========\n")

try:
    image_client = AzureOpenAI(
        api_key=os.getenv("IMAGE_API_KEY"),
        api_version=os.getenv("IMAGE_API_VERSION"),
        azure_endpoint=os.getenv("IMAGE_API_BASE")
    )

    img = image_client.images.generate(
        model=os.getenv("IMAGE_MODEL"),
        prompt="A blue parrot sitting on a branch"
    )

    print("Image generated successfully")
    print("TEST 3 PASSED\n")

except Exception as e:
    print("TEST 3 FAILED:", e)
    print("\n")

print("========== TESTING COMPLETE ==========\n")
