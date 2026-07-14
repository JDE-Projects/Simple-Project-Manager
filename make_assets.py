"""
Generate the Simple Project Manager icon, window PNG, and splash.

Glyph: checklist (concept A) on a JDE teal plate.
Run from the repo root:  python make_assets.py
Outputs:
  simple_project_manager.png          (256px live-window icon)
  simple_project_manager.ico          (multi-size; plain teal square at 16/24)
  simple_project_manager-splash.png   (460x280 splash)
"""
from PIL import Image, ImageDraw, ImageFont

# --- JDE dark teal tokens used here ---
ACCENT   = (77, 214, 193)   # #4dd6c1  teal plate
DARKTEAL = (4, 43, 39)      # #042b27  glyph on plate
BG       = (10, 14, 20)     # #0a0e14  splash background
PANEL    = (18, 24, 35)     # #121823  splash panel
BORDER   = (32, 41, 58)     # #20293a  splash panel border
TEXT     = (223, 231, 242)  # #dfe7f2
TEXTDIM  = (138, 152, 176)  # #8a98b0
TEXTFNT  = (90, 102, 120)   # #5a6678

SS = 4  # supersample factor for crisp downscaling
FONTS = "fonts"


def _rrect(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def glyph(size, plate=ACCENT):
    """Rounded teal plate with a 3-row checklist, rendered at `size`px."""
    s = size * SS
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # plate fills the icon, rounded-square corners
    _rrect(d, (0, 0, s - 1, s - 1), radius=int(s * 0.22), fill=plate)

    # checklist geometry (proportions from the approved SVG, 0..1 of size)
    box_x, box_w = 0.20, 0.135          # checkbox left + width
    line_x, line_end = 0.42, 0.78       # line start/end x
    rows_y = (0.255, 0.45, 0.645)       # row centers
    checked = (True, True, False)       # first two ticked
    stroke = max(2, int(s * 0.028))
    bw = box_w * s

    for cy, tick in zip(rows_y, checked, strict=True):
        y = cy * s
        bx = box_x * s
        # checkbox outline
        d.rounded_rectangle((bx, y - bw / 2, bx + bw, y + bw / 2),
                            radius=int(bw * 0.22), outline=DARKTEAL, width=stroke)
        if tick:
            d.line([(bx + bw * 0.22, y + bw * 0.05),
                    (bx + bw * 0.45, y + bw * 0.30),
                    (bx + bw * 0.82, y - bw * 0.28)],
                   fill=DARKTEAL, width=stroke, joint="curve")
        # the task line
        d.line([(line_x * s, y), (line_end * s, y)],
               fill=DARKTEAL, width=stroke)

    return img.resize((size, size), Image.LANCZOS)


def plain_square(size, plate=ACCENT):
    """Plain teal rounded square for tiny icon sizes (16/24)."""
    s = size * SS
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    _rrect(d, (0, 0, s - 1, s - 1), radius=int(s * 0.18), fill=plate)
    return img.resize((size, size), Image.LANCZOS)


def make_png():
    glyph(256).save("simple_project_manager.png")
    print("  simple_project_manager.png")


def make_ico():
    imgs = [plain_square(16), plain_square(24),
            glyph(32), glyph(48), glyph(64), glyph(128), glyph(256)]
    imgs[-1].save("simple_project_manager.ico", format="ICO",
                  sizes=[(i.width, i.height) for i in imgs],
                  append_images=imgs[:-1])
    print("  simple_project_manager.ico  (16,24,32,48,64,128,256)")


def make_splash():
    W, H = 460, 280
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    # centered dark panel
    m = 18
    d.rounded_rectangle((m, m, W - m - 1, H - m - 1), radius=14,
                        fill=PANEL, outline=BORDER, width=1)

    # glyph plate
    g = glyph(76)
    gx = (W - 76) // 2
    img.paste(g, (gx, 52), g)

    sora_bold = ImageFont.truetype(f"{FONTS}/Sora-Bold.ttf", 26)
    mono = ImageFont.truetype(f"{FONTS}/JetBrainsMono-Regular.ttf", 13)
    sora_reg = ImageFont.truetype(f"{FONTS}/Sora-Regular.ttf", 13)

    def center(text, y, font, fill):
        w = d.textlength(text, font=font)
        d.text(((W - w) / 2, y), text, font=font, fill=fill)

    center("Simple Project Manager", 150, sora_bold, TEXT)
    center("github.com/JDE-Projects", 188, mono, TEXTDIM)
    center("starting…", 214, sora_reg, TEXTFNT)

    img.save("simple_project_manager-splash.png")
    print("  simple_project_manager-splash.png")


if __name__ == "__main__":
    print("Generating assets:")
    make_png()
    make_ico()
    make_splash()
    print("Done.")
