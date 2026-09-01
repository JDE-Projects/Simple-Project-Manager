"""
Generate the Simple Project Manager icon and window PNG.

Glyph: checklist (concept A) on a JDE teal plate.
Run from the repo root:  python make_assets.py
Outputs:
  simple_project_manager.png          (256px live-window icon)
  simple_project_manager.ico          (multi-size; plain teal square at 16/24)
"""
from PIL import Image, ImageDraw

# --- JDE dark teal tokens used here ---
ACCENT   = (77, 214, 193)   # #4dd6c1  teal plate
DARKTEAL = (4, 43, 39)      # #042b27  glyph on plate

SS = 4  # supersample factor for crisp downscaling


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


if __name__ == "__main__":
    print("Generating assets:")
    make_png()
    make_ico()
    print("Done.")
