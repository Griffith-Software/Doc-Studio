"""File-type icons for documents Doc Studio opens (build/filetypes/<ext>.ico, shipped in resources\\filetypes).

Each icon is a white page with the Doc Studio logo and a coloured band naming the type (DOCX blue, XLSX green,
PPTX orange, PDF red, MD dark grey...). Generic colours and the extension only: no other company's logos.
At 32 px and below the band has no text (it couldn't be read), just its colour.

    python tools/make-filetype-icons.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'build', 'filetypes')
LOGO = Image.open(os.path.join(ROOT, 'build', 'icon.png')).convert('RGBA')
FONT = 'C:/Windows/Fonts/segoeuib.ttf'

BLUE, GREEN, ORANGE, RED, SLATE, GREY, TEAL, AMBER, INDIGO = '#2a5bd7', '#1e8e3e', '#e8710a', '#d93025', '#3c4043', '#6b7280', '#00897b', '#c77700', '#5e35b1'
TYPES = {
    'docx': BLUE, 'docm': BLUE, 'dotx': BLUE, 'doc': BLUE, 'dot': BLUE, 'rtf': BLUE, 'odt': BLUE, 'ott': BLUE,
    'xlsx': GREEN, 'xlsm': GREEN, 'xlsb': GREEN, 'xls': GREEN, 'xltx': GREEN, 'xltm': GREEN, 'ods': GREEN, 'fods': GREEN,
    'numbers': GREEN, 'csv': GREEN, 'tsv': GREEN,
    'pptx': ORANGE, 'pptm': ORANGE, 'ppsx': ORANGE, 'ppt': ORANGE, 'pps': ORANGE, 'odp': ORANGE,
    'pdf': RED, 'epub': TEAL,
    'md': SLATE, 'markdown': SLATE, 'txt': GREY, 'log': GREY,
    'html': AMBER, 'htm': AMBER, 'json': INDIGO, 'xml': INDIGO,
}
LABEL = {'markdown': 'MD', 'numbers': 'NUM'}


def page(ext, color, size):
    """One size of the icon, drawn at 4x and scaled down for smooth edges."""
    s = size * 4
    k = s / 256
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1, fold = 40 * k, 10 * k, 216 * k, 246 * k, 46 * k
    # The page with a folded top-right corner, and a soft shadow.
    shadow = [(x0 + 4 * k, y0 + 6 * k), (x1 - fold + 4 * k, y0 + 6 * k), (x1 + 4 * k, y0 + fold + 6 * k), (x1 + 4 * k, y1 + 6 * k), (x0 + 4 * k, y1 + 6 * k)]
    d.polygon(shadow, fill=(0, 0, 0, 40))
    body = [(x0, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1), (x0, y1)]
    d.polygon(body, fill='#ffffff', outline='#c9ced6', width=max(1, round(3 * k)))
    d.polygon([(x1 - fold, y0), (x1 - fold, y0 + fold), (x1, y0 + fold)], fill='#e3e6eb', outline='#c9ced6', width=max(1, round(3 * k)))
    # The Doc Studio logo, top left.
    lg = round(78 * k)
    img.alpha_composite(LOGO.resize((lg, lg), Image.LANCZOS), (round(x0 + 16 * k), round(y0 + 22 * k)))
    # A few text lines.
    for i, w in enumerate((70, 58, 104, 92)):
        y = y0 + (30 + i * 18) * k if i < 2 else y0 + (112 + (i - 2) * 18) * k
        x = x0 + (106 if i < 2 else 18) * k
        d.rounded_rectangle([x, y, x + w * k, y + 8 * k], radius=4 * k, fill='#d5d9e0')
    # The coloured band naming the type.
    by0, by1 = y0 + 156 * k, y0 + 216 * k
    d.rectangle([x0 + 2 * k, by0, x1 - 1 * k, by1], fill=color)
    if size > 32:
        label = LABEL.get(ext, ext.upper())
        fs = 50 * k if len(label) <= 3 else 42 * k
        font = ImageFont.truetype(FONT, round(fs))
        tw = d.textlength(label, font=font)
        while tw > (x1 - x0) - 20 * k:
            fs -= 2 * k
            font = ImageFont.truetype(FONT, round(fs))
            tw = d.textlength(label, font=font)
        d.text(((x0 + x1) / 2, (by0 + by1) / 2), label, font=font, fill='#ffffff', anchor='mm')
    return img.resize((size, size), Image.LANCZOS)


def main():
    os.makedirs(OUT, exist_ok=True)
    for ext, color in TYPES.items():
        sizes = [256, 64, 48, 40, 32, 24, 20, 16]
        frames = [page(ext, color, n) for n in sizes]
        # Small sizes as BMP entries, 256 as PNG (what Explorer expects).
        frames[0].save(os.path.join(OUT, f'{ext}.ico'), format='ICO', sizes=[(n, n) for n in sizes], append_images=frames[1:], bitmap_format='bmp')
    page('docx', BLUE, 256).save(os.path.join(OUT, '_preview-docx.png'))
    sheet = Image.new('RGBA', (8 * 136, 5 * 136), (40, 42, 48, 255))
    for i, (ext, color) in enumerate(TYPES.items()):
        sheet.alpha_composite(page(ext, color, 128), ((i % 8) * 136 + 4, (i // 8) * 136 + 4))
    sheet.save(os.path.join(OUT, '_preview-all.png'))
    print(f'{len(TYPES)} icons -> build/filetypes')


main()
