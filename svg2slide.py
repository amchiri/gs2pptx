"""Conversion d'une page SVG du viewer Google Slides en slide python-pptx éditable.

Structure observée dans le SVG du viewer publié :
  - images : <g transform="matrix(...)"><clipPath/><image xlink:href=... width height/></g>
  - texte  : vectorisé en <path> (glyphes, 1 path par ligne / run), MAIS précédé d'un
             <g id="a11y-..." role="img" aria-label="TEXTE"> contenant un <path> = la zone
  - formes : <path fill=... stroke=...> (les fill-opacity="0" sont des boîtes invisibles)

Repère : viewBox 960x540 -> 1 unité SVG = 1/96 in = 9525 EMU (slide 10 x 5.625 in).
"""
import io
import math
import re
import xml.etree.ElementTree as ET

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Emu, Pt

SVG = "{http://www.w3.org/2000/svg}"
XLINK = "{http://www.w3.org/1999/xlink}"

# ---------------------------------------------------------------- matrices
IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def mat_mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


def apply(m, x, y):
    a, b, c, d, e, f = m
    return a * x + c * y + e, b * x + d * y + f


def parse_transform(s):
    m = IDENT
    if not s:
        return m
    for name, args in re.findall(r"(\w+)\s*\(([^)]*)\)", s):
        v = [float(x) for x in re.split(r"[\s,]+", args.strip()) if x]
        if name == "matrix":
            t = tuple(v)
        elif name == "translate":
            t = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif name == "scale":
            t = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif name == "rotate":
            r = math.radians(v[0])
            t = (math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0, 0)
            if len(v) == 3:
                t = mat_mul(mat_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        else:
            continue
        m = mat_mul(m, t)
    return m


# ---------------------------------------------------------------- paths
_TOK = re.compile(r"[MmLlHhVvCcSsQqTtAaZz]|-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")


def parse_path(d, curve_steps=8):
    """Retourne une liste de contours [[(x,y),...], closed] (courbes aplaties)."""
    toks = _TOK.findall(d)
    i, cmd = 0, None
    x = y = sx = sy = 0.0
    last_ctrl = None
    contours, cur = [], None

    def num():
        nonlocal i
        v = float(toks[i])
        i += 1
        return v

    def bez(p0, p1, p2, p3):
        return [
            (
                (1 - t) ** 3 * p0[0] + 3 * (1 - t) ** 2 * t * p1[0] + 3 * (1 - t) * t * t * p2[0] + t ** 3 * p3[0],
                (1 - t) ** 3 * p0[1] + 3 * (1 - t) ** 2 * t * p1[1] + 3 * (1 - t) * t * t * p2[1] + t ** 3 * p3[1],
            )
            for t in (k / curve_steps for k in range(1, curve_steps + 1))
        ]

    while i < len(toks):
        if toks[i].isalpha():
            cmd = toks[i]
            i += 1
        rel = cmd.islower()
        C = cmd.upper()
        ox, oy = (x, y) if rel else (0.0, 0.0)

        if C == "Z":
            if cur:
                cur[1] = True
            x, y = sx, sy
            last_ctrl = None
            continue
        if C == "M":
            x, y = ox + num(), oy + num()
            sx, sy = x, y
            cur = [[(x, y)], False]
            contours.append(cur)
            cmd = "l" if rel else "L"  # coordonnées suivantes = lineto
            last_ctrl = None
            continue
        if cur is None:
            cur = [[(x, y)], False]
            contours.append(cur)
        if C == "L":
            x, y = ox + num(), oy + num()
            cur[0].append((x, y))
            last_ctrl = None
        elif C == "H":
            x = ox + num()
            cur[0].append((x, y))
            last_ctrl = None
        elif C == "V":
            y = oy + num()
            cur[0].append((x, y))
            last_ctrl = None
        elif C == "C":
            p1 = (ox + num(), oy + num())
            p2 = (ox + num(), oy + num())
            p3 = (ox + num(), oy + num())
            cur[0] += bez((x, y), p1, p2, p3)
            last_ctrl, (x, y) = p2, p3
        elif C == "S":
            p1 = (2 * x - last_ctrl[0], 2 * y - last_ctrl[1]) if last_ctrl else (x, y)
            p2 = (ox + num(), oy + num())
            p3 = (ox + num(), oy + num())
            cur[0] += bez((x, y), p1, p2, p3)
            last_ctrl, (x, y) = p2, p3
        elif C in "QT":
            if C == "Q":
                q = (ox + num(), oy + num())
            else:
                q = (2 * x - last_ctrl[0], 2 * y - last_ctrl[1]) if last_ctrl else (x, y)
            p3 = (ox + num(), oy + num())
            p1 = (x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y))
            p2 = (p3[0] + 2 / 3 * (q[0] - p3[0]), p3[1] + 2 / 3 * (q[1] - p3[1]))
            cur[0] += bez((x, y), p1, p2, p3)
            last_ctrl, (x, y) = q, p3
        elif C == "A":  # arc : approximé par une ligne droite
            for _ in range(5):
                num()
            x, y = ox + num(), oy + num()
            cur[0].append((x, y))
            last_ctrl = None
    return contours


def bbox(points):
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def path_bbox(d, m):
    pts = [apply(m, *p) for c in parse_path(d, 4) for p in c[0]]
    return bbox(pts) if pts else None


# ---------------------------------------------------------------- styles
def hex_rgb(s):
    if not s or s == "none" or not s.startswith("#"):
        return None
    s = s[1:]
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    return RGBColor.from_string(s.upper())


def visible_fill(el):
    f = el.get("fill")
    if not f or f == "none":
        return None
    if float(el.get("fill-opacity", "1")) == 0:
        return None
    return hex_rgb(f)


def visible_stroke(el):
    s = el.get("stroke")
    if not s or s == "none":
        return None
    if float(el.get("stroke-opacity", "1")) == 0:
        return None
    return hex_rgb(s)


# ---------------------------------------------------------------- conversion
class SlideBuilder:
    def __init__(self, slide, fetch_image, emu_per_unit=9525):
        self.slide = slide
        self.fetch = fetch_image  # url -> bytes | None
        self.k = emu_per_unit
        self.text_boxes = []  # (bbox, dict) zones texte a11y
        self.clips = {}

    def E(self, v):
        return Emu(int(round(v * self.k)))

    # ---- passe 1 : collecter clipPaths et zones de texte
    def collect(self, el, m):
        m = mat_mul(m, parse_transform(el.get("transform")))
        tag = el.tag.replace(SVG, "")
        if tag == "clipPath":
            p = el.find(SVG + "path")
            if p is not None:
                self.clips[el.get("id")] = p.get("d")
            return
        if tag == "g" and el.get("aria-label") and (el.get("id") or "").startswith("a11y-"):
            p = el.find(SVG + "path")
            if p is not None and el.get("aria-label").strip():
                bb = path_bbox(p.get("d"), m)
                if bb:
                    self.text_boxes.append({"bbox": bb, "text": el.get("aria-label"), "glyphs": []})
            return
        for ch in el:
            self.collect(ch, m)

    def text_box_for(self, bb):
        """Zone de texte contenant (avec tolérance) la bbox d'un path de glyphes."""
        x0, y0, x1, y1 = bb
        best = None
        for tb in self.text_boxes:
            X0, Y0, X1, Y1 = tb["bbox"]
            tol = 6  # le path doit être entièrement dans la zone (± tolérance)
            if X0 - tol <= x0 and x1 <= X1 + tol and Y0 - tol <= y0 and y1 <= Y1 + tol:
                area = (X1 - X0) * (Y1 - Y0)
                if best is None or area < best[0]:
                    best = (area, tb)
        return best[1] if best else None

    # ---- passe 2 : émettre images / formes, rattacher les glyphes aux zones de texte
    def emit(self, el, m):
        m = mat_mul(m, parse_transform(el.get("transform")))
        tag = el.tag.replace(SVG, "")
        if tag in ("clipPath", "defs"):
            return
        if tag in ("g", "svg"):
            if (el.get("id") or "").startswith("a11y-"):
                return  # boîte invisible de la zone texte
            for ch in el:
                self.emit(ch, m)
        elif tag == "image":
            self.add_image(el, m)
        elif tag == "path":
            self.add_path(el, m)

    def add_image(self, el, m):
        url = el.get(XLINK + "href") or el.get("href")
        data = self.fetch(url) if url else None
        if not data:
            return
        x, y = float(el.get("x", 0)), float(el.get("y", 0))
        w, h = float(el.get("width")), float(el.get("height"))

        # rectangle visible = image ∩ clipPath (clip dans le même repère que l'image)
        vx0, vy0, vx1, vy1 = x, y, x + w, y + h
        clip_id = re.match(r"url\(#(.+)\)", el.get("clip-path") or "")
        if clip_id and clip_id.group(1) in self.clips:
            cb = path_bbox(self.clips[clip_id.group(1)], IDENT)
            if cb:
                vx0, vy0 = max(vx0, cb[0]), max(vy0, cb[1])
                vx1, vy1 = min(vx1, cb[2]), min(vy1, cb[3])

        X0, Y0 = apply(m, vx0, vy0)
        X1, Y1 = apply(m, vx1, vy1)
        rot = math.degrees(math.atan2(m[1], m[0]))
        sx = math.hypot(m[0], m[1])
        sy = math.hypot(m[2], m[3])
        W, H = (vx1 - vx0) * sx, (vy1 - vy0) * sy
        # position = coin haut-gauche non tourné autour du centre
        cx, cy = (X0 + X1) / 2, (Y0 + Y1) / 2
        if abs(rot) < 0.01:
            left, top = min(X0, X1), min(Y0, Y1)
        else:
            left, top = cx - W / 2, cy - H / 2

        pic = self.slide.shapes.add_picture(io.BytesIO(data), self.E(left), self.E(top), self.E(W), self.E(H))
        if w and h:
            pic.crop_left = (vx0 - x) / w
            pic.crop_top = (vy0 - y) / h
            pic.crop_right = (x + w - vx1) / w
            pic.crop_bottom = (y + h - vy1) / h
        if abs(rot) >= 0.01:
            pic.rotation = rot

    def add_path(self, el, m):
        fill, stroke = visible_fill(el), visible_stroke(el)
        if fill is None and stroke is None:
            return  # boîte invisible

        d = el.get("d")
        bb = path_bbox(d, m)
        if not bb:
            return

        # glyphes de texte -> rattachés à leur zone, pas dessinés
        tb = self.text_box_for(bb) if fill is not None and stroke is None else None
        if tb is not None:
            tb["glyphs"].append({"bbox": bb, "color": fill})
            return

        contours = [([apply(m, *p) for p in pts], closed) for pts, closed in parse_path(d)]
        contours = [c for c in contours if len(c[0]) >= 2]
        if not contours:
            return

        ff = self.slide.shapes.build_freeform(
            self.E(contours[0][0][0][0]), self.E(contours[0][0][0][1]), scale=1.0
        )
        first = True
        for pts, closed in contours:
            if not first:
                ff.move_to(self.E(pts[0][0]), self.E(pts[0][1]))
            ff.add_line_segments([(self.E(px), self.E(py)) for px, py in pts[1:]], close=closed)
            first = False
        shp = ff.convert_to_shape()

        if fill is not None:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        else:
            shp.fill.background()
        if stroke is not None:
            shp.line.color.rgb = stroke
            shp.line.width = Pt(float(el.get("stroke-width", "1")) * 0.75 * math.hypot(m[0], m[1]))
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False

    def emit_text(self):
        for tb in self.text_boxes:
            X0, Y0, X1, Y1 = tb["bbox"]
            lines = tb["text"].rstrip("\n").split("\n")
            glyphs = tb["glyphs"]

            # lignes de glyphes = paths regroupés par bas de bbox (≈ baseline+descente)
            rows = []
            for g in sorted(glyphs, key=lambda g: g["bbox"][3]):
                if rows and abs(g["bbox"][3] - rows[-1]["y1"]) < 3:
                    r = rows[-1]
                    r["y0"] = min(r["y0"], g["bbox"][1])
                else:
                    rows.append({"y0": g["bbox"][1], "y1": g["bbox"][3], "color": g["color"]})

            color = glyphs[0]["color"] if glyphs else None
            # taille : interligne mesuré (pas entre lignes non vides) sinon hauteur de ligne
            # interligne = écart premier/dernier rang / nb de lignes (blanches incluses)
            pitch = None
            nonempty = [i for i, l in enumerate(lines) if l.strip()]
            if len(rows) >= 2 and len(rows) == len(nonempty):
                pitch = (rows[-1]["y1"] - rows[0]["y1"]) / (nonempty[-1] - nonempty[0])
            elif len(rows) >= 2:
                diffs = sorted(b["y1"] - a["y1"] for a, b in zip(rows, rows[1:]))
                pitch = diffs[len(diffs) // 2]
            if pitch:
                font_px = pitch / 1.2
            elif rows:
                font_px = (rows[0]["y1"] - rows[0]["y0"]) / 0.95
            else:
                font_px = 18 / 0.75
            font_pt = round(font_px * 0.75 * 2) / 2

            box = self.slide.shapes.add_textbox(self.E(X0), self.E(Y0), self.E(X1 - X0), self.E(Y1 - Y0))
            tf = box.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.vertical_anchor = MSO_ANCHOR.TOP
            # alignement : centré si les glyphes sont centrés dans la zone
            align = None
            if glyphs:
                gx0 = min(g["bbox"][0] for g in glyphs)
                gx1 = max(g["bbox"][2] for g in glyphs)
                left_gap, right_gap = gx0 - X0, X1 - gx1
                if abs(left_gap - right_gap) < 6 and left_gap > 15:
                    from pptx.enum.text import PP_ALIGN
                    align = PP_ALIGN.CENTER

            for i, line in enumerate(lines):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                if pitch:
                    p.line_spacing = Pt(pitch * 0.75)
                if align is not None:
                    p.alignment = align
                run = p.add_run()
                run.text = line
                run.font.size = Pt(font_pt)
                if color is not None:
                    run.font.color.rgb = color


def svg_to_slide(svg_code, slide, fetch_image):
    # le parseur XML normalise les \n des attributs en espaces -> on les protège
    svg_code = re.sub(
        r'aria-label="([^"]*)"',
        lambda mm: 'aria-label="' + mm.group(1).replace("\r", "").replace("\n", "&#10;") + '"',
        svg_code,
    )
    root = ET.fromstring(svg_code)
    vb = [float(v) for v in re.split(r"[\s,]+", root.get("viewBox", "0 0 960 540").strip())]
    b = SlideBuilder(slide, fetch_image)
    m0 = (1, 0, 0, 1, -vb[0], -vb[1])
    b.collect(root, m0)
    b.emit(root, m0)
    b.emit_text()
    return b
