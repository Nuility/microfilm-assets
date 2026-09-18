from pathlib import Path
from PIL import Image, ImageOps, ImageDraw

root = Path(__file__).resolve().parent / "docx_render"
pages = sorted(root.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))
for batch_idx in range(0, len(pages), 6):
    batch = pages[batch_idx:batch_idx + 6]
    thumbs = []
    for p in batch:
        img = Image.open(p).convert("RGB")
        img.thumbnail((500, 650))
        canvas = Image.new("RGB", (520, 700), "white")
        x = (520 - img.width) // 2
        canvas.paste(img, (x, 30))
        draw = ImageDraw.Draw(canvas)
        draw.text((12, 8), p.stem, fill="black")
        canvas = ImageOps.expand(canvas, border=2, fill="#B7B7B7")
        thumbs.append(canvas)
    sheet = Image.new("RGB", (1044, 2106), "#E5E5E5")
    for i, img in enumerate(thumbs):
        sheet.paste(img, ((i % 2) * 522, (i // 2) * 702))
    out = root / f"contact_{batch_idx // 6 + 1:02d}.jpg"
    sheet.save(out, quality=90)
    print(out)
