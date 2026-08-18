from PIL import Image, ImageDraw, ImageFont

SIZES = [16, 32, 64, 80, 128]
BG = (0, 90, 158)      # blue
FG = (255, 255, 255)   # white

for size in SIZES:
    img = Image.new("RGBA", (size, size), BG + (255,))
    draw = ImageDraw.Draw(img)
    text = "S"
    try:
        font = ImageFont.truetype("arialbd.ttf", int(size * 0.62))
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((size - w) / 2 - bbox[0], (size - h) / 2 - bbox[1]), text, fill=FG, font=font)
    img.save(f"assets/icon-{size}.png")

print("done")
