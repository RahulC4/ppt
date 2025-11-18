from utils import image_client
resp = image_client.images.generate(
    model="gpt-image-1-mini",
    prompt="test",
    size="512x512"
)
print(resp)
