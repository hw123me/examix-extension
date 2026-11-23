from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "store-promo"
OUT.mkdir(exist_ok=True)

WIDTH, HEIGHT = 440, 280
NAVY = "#073b66"
GREEN = "#08c995"
WHITE = "#ffffff"
FONT_REGULAR = r"C:\Windows\Fonts\tahoma.ttf"
FONT_BOLD = r"C:\Windows\Fonts\tahomabd.ttf"


def arabic(draw, xy, value, size, fill, bold=False, anchor=None):
    font = ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size)
    reshaped = arabic_reshaper.reshape(str(value))
    displayed = get_display(reshaped)
    draw.text(xy, displayed, font=font, fill=fill, anchor=anchor)


canvas = Image.new("RGB", (WIDTH, HEIGHT), NAVY)
draw = ImageDraw.Draw(canvas)

draw.rounded_rectangle((20, 20, 420, 260), radius=24, fill="#ffffff")
icon = Image.open(ROOT / "icons" / "icon128.png").convert("RGBA")
icon = icon.resize((106, 106), Image.Resampling.LANCZOS)
canvas.paste(icon, (167, 45), icon)

arabic(draw, (220, 168), "إكزامكس", 30, NAVY, True, "mm")
arabic(draw, (220, 205), "رصد الدرجات في نظام نور", 20, "#172b4d", True, "mm")
draw.rounded_rectangle((98, 230, 342, 248), radius=9, fill=GREEN)
arabic(draw, (220, 239), "استيراد الدرجات بسهولة", 13, WHITE, True, "mm")

canvas.save(OUT / "promo-small-440x280.png", format="PNG", optimize=True)
print(OUT / "promo-small-440x280.png")
