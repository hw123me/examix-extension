from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display


ROOT = Path(__file__).resolve().parent
OUT = Path(r"C:\examix-store-screenshots")
OUT.mkdir(exist_ok=True)

WIDTH, HEIGHT = 1280, 800
NAVY = "#073b66"
GREEN = "#08c995"
WHITE = "#ffffff"
LIGHT = "#f4f7fb"
TEXT = "#172b4d"
MUTED = "#6b7a90"
BORDER = "#dce4ef"

FONT_REGULAR = r"C:\Windows\Fonts\tahoma.ttf"
FONT_BOLD = r"C:\Windows\Fonts\tahomabd.ttf"


def font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size)


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def paste_icon(canvas, box):
    icon = Image.open(ROOT / "icons" / "icon128.png").convert("RGBA")
    icon.thumbnail((box[2] - box[0], box[3] - box[1]), Image.Resampling.LANCZOS)
    x = box[0] + (box[2] - box[0] - icon.width) // 2
    y = box[1] + (box[3] - box[1] - icon.height) // 2
    canvas.alpha_composite(icon, (x, y))


def text(draw, xy, value, size=22, fill=TEXT, bold=False, anchor=None):
    reshaped = arabic_reshaper.reshape(str(value))
    displayed = get_display(reshaped)
    draw.text(xy, displayed, font=font(size, bold), fill=fill, anchor=anchor)


def base_canvas():
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), (255, 255, 255, 255))
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, 92), fill=NAVY)
    text(draw, (1240, 46), "نظام نور", 26, WHITE, True, "rm")
    rounded(draw, (70, 26, 118, 74), 12, GREEN)
    text(draw, (94, 50), "ن", 32, NAVY, True, "mm")
    text(draw, (40, 124), "رصد الدرجات", 30, TEXT, True, "lm")
    draw.line((40, 170, 1240, 170), fill=BORDER, width=2)
    return canvas, draw


def toolbar(canvas, draw, compact=False):
    box = (820, 188, 1210, 700) if not compact else (830, 190, 1210, 610)
    rounded(draw, box, 18, WHITE, BORDER, 2)
    rounded(draw, (840, 205, 1190, 260), 12, NAVY)
    paste_icon(canvas, (852, 208, 900, 256))
    text(draw, (1180, 232), "إكزامكس — رصد نور", 18, WHITE, True, "rm")
    return box


def table(canvas, draw):
    left, top = 40, 200
    rounded(draw, (left, top, 790, 700), 12, WHITE, BORDER, 2)
    headers = ["الاسم", "رقم الهوية", "الاختبار", "الدرجة"]
    widths = [280, 190, 150, 130]
    x = left
    for i, header in enumerate(headers):
        rounded(draw, (x, top, x + widths[i], top + 54), 8, "#edf3f8")
        text(draw, (x + widths[i] - 18, top + 27), header, 17, TEXT, True, "rm")
        x += widths[i]

    names = ["محمد أحمد العتيبي", "سارة خالد الدوسري", "عبدالله محمد القحطاني", "نورة سعد الحربي", "فيصل عمر الشمري"]
    ids = ["١٠٤٢٣٨١٢٩٥", "١١٧٤٩٢٥٥٠٣", "١٠٧٨٣٤٦١٢٢", "١٢٢١٤٤٣٨٨١", "١١٤٠٨٧۶٩٢٣"]
    scores = ["18", "20", "16", "19", "17"]
    y = top + 54
    for row in range(5):
        if row % 2:
            draw.rectangle((left, y, 790, y + 82), fill="#fafcff")
        text(draw, (left + widths[0] - 18, y + 41), names[row], 17, TEXT, False, "rm")
        text(draw, (left + widths[0] + 18, y + 41), ids[row], 16, MUTED, False, "lm")
        text(draw, (left + widths[0] + widths[1] + widths[2] - 18, y + 41), "اختبار الوحدة", 15, TEXT, False, "rm")
        rounded(draw, (left + 654, y + 22, left + 760, y + 62), 8, "#e8fbf5", GREEN, 1)
        text(draw, (left + 707, y + 42), scores[row], 18, "#087b61", True, "mm")
        y += 82


def screenshot_1():
    canvas, draw = base_canvas()
    table(canvas, draw)
    box = toolbar(canvas, draw)
    text(draw, (1190, 292), "استيراد درجات الاختبار", 18, TEXT, True, "rm")
    text(draw, (1180, 330), "ارفع ملف Excel أو الصق البيانات", 15, MUTED, False, "rm")
    rounded(draw, (845, 350, 1185, 430), 12, "#f8fafc", BORDER, 2)
    text(draw, (1015, 390), "اختر ملف Excel", 16, NAVY, True, "mm")
    rounded(draw, (845, 448, 1185, 510), 12, GREEN)
    text(draw, (1015, 479), "بدء الاستيراد", 17, WHITE, True, "mm")
    text(draw, (1190, 548), "تم تجهيز 5 طلاب", 14, "#087b61", True, "rm")
    text(draw, (1190, 575), "جاهز للمراجعة قبل الإرسال", 14, MUTED, False, "rm")
    rounded(draw, (845, 625, 1185, 668), 12, "#eef3f8")
    text(draw, (1015, 646), "إلغاء", 15, MUTED, True, "mm")
    canvas.convert("RGB").save(OUT / "01-import-grades.png", quality=95)


def screenshot_2():
    canvas, draw = base_canvas()
    table(canvas, draw)
    box = toolbar(canvas, draw)
    text(draw, (1190, 292), "ربط الطالب بالدرجات", 18, TEXT, True, "rm")
    text(draw, (1190, 328), "مطابقة دقيقة برقم الهوية أو الاسم", 15, MUTED, False, "rm")
    rows = [("محمد أحمد العتيبي", "18"), ("سارة خالد الدوسري", "20"), ("عبدالله محمد القحطاني", "16")]
    y = 360
    for name, score in rows:
        rounded(draw, (845, y, 1185, y + 72), 10, "#f8fafc", BORDER, 1)
        text(draw, (1172, y + 24), name, 15, TEXT, True, "rm")
        rounded(draw, (860, y + 18, 930, y + 54), 8, "#e8fbf5")
        text(draw, (895, y + 36), score, 16, "#087b61", True, "mm")
        text(draw, (1172, y + 54), "تمت المطابقة", 12, "#087b61", False, "rm")
        y += 82
    rounded(draw, (845, 620, 1185, 682), 12, GREEN)
    text(draw, (1015, 651), "تنفيذ الرصد", 17, WHITE, True, "mm")
    canvas.convert("RGB").save(OUT / "02-match-students.png", quality=95)


def screenshot_3():
    canvas, draw = base_canvas()
    table(canvas, draw)
    box = toolbar(canvas, draw)
    text(draw, (1190, 292), "نتائج الاستيراد", 18, TEXT, True, "rm")
    rounded(draw, (845, 320, 1185, 420), 14, "#f0fdf8", "#b8ead8", 2)
    text(draw, (1015, 357), "تم رصد الدرجات", 20, "#087b61", True, "mm")
    text(draw, (1015, 392), "5 طلاب • 5 مطابقة ناجحة", 14, MUTED, False, "mm")
    text(draw, (1190, 460), "ملخص العملية", 17, TEXT, True, "rm")
    text(draw, (1168, 498), "الدرجات المضافة", 15, MUTED, False, "rm")
    text(draw, (860, 498), "5", 17, GREEN, True, "lm")
    text(draw, (1168, 540), "القيم غير المطابقة", 15, MUTED, False, "rm")
    text(draw, (860, 540), "0", 17, GREEN, True, "lm")
    rounded(draw, (845, 620, 1185, 682), 12, NAVY)
    text(draw, (1015, 651), "الرجوع إلى رصد الدرجات", 16, WHITE, True, "mm")
    canvas.convert("RGB").save(OUT / "03-import-results.png", quality=95)


def screenshot_4():
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), (244, 247, 251, 255))
    draw = ImageDraw.Draw(canvas)
    rounded(draw, (55, 55, 1225, 745), 28, WHITE)
    paste_icon(canvas, (470, 105, 810, 445))
    text(draw, (640, 470), "إكزامكس — رصد الدرجات في نظام نور", 34, NAVY, True, "mm")
    text(draw, (640, 520), "استيراد الدرجات من Excel أو النسخ المباشر", 23, TEXT, False, "mm")
    text(draw, (640, 562), "مطابقة الطلاب تلقائياً • مراجعة البيانات • إكمال الرصد", 19, MUTED, False, "mm")
    rounded(draw, (480, 610, 800, 680), 14, GREEN)
    text(draw, (640, 645), "تعمل فقط داخل صفحة رصد الدرجات بنظام نور", 17, WHITE, True, "mm")
    canvas.convert("RGB").save(OUT / "04-feature-overview.png", quality=95)


if __name__ == "__main__":
    screenshot_1()
    screenshot_2()
    screenshot_3()
    screenshot_4()
    print("\n".join(str(path) for path in sorted(OUT.glob("*.png"))))
