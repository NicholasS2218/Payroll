from PIL import Image, ImageDraw, ImageFont

img = Image.new("RGB", (400, 200), "#f4640d")
d = ImageDraw.Draw(img)
try:
    font = ImageFont.truetype("arialbd.ttf", 28)
except OSError:
    font = ImageFont.load_default()
d.text((200, 80), "Payroll Report Generator", fill="white", font=font, anchor="mm")
img.save("splash.png")