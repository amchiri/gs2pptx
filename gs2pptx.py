"""Google Slides publié (/d/e/2PACX-.../embed) -> PPTX éditable.

Chaque slide est convertie depuis le SVG du viewer :
  images -> vraies images, texte -> vraies zones de texte, formes -> formes libres.
Mettre EDITABLE = False pour revenir à l'ancien mode capture d'écran.
"""
from pathlib import Path
import hashlib
import html
import img2pdf
import re

from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches

from svg2slide import svg_to_slide

EDITABLE = False  # True = texte éditable, False = capture d'écran
SAVE_SVG = True  # écrit mes_diapos/svg/<n>.svg avec images locales dans svg/images/

HREF_RE = re.compile(r'((?:xlink:)?href)="(https?://[^"]+)"')


def ext_image(data):
    """Extension d'après les premiers octets (Google sert souvent du JPEG/WebP)."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpg"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    if data[:4] == b"GIF8":
        return "gif"
    return "png"


class ImageManager:
    """Télécharge et stocke chaque image UNE seule fois.

    Deux niveaux de dédoublonnage :
      - par URL      : une URL déjà vue n'est pas re-téléchargée
      - par contenu  : Google génère une URL différente par slide pour la même image
                       (ex. le fond commun) -> on identifie l'image par son SHA-1,
                       et le fichier s'appelle <sha1>.<ext>, écrit une seule fois.
    """

    def __init__(self, page, dossier_images: Path):
        self.page = page
        self.dossier = dossier_images
        self.dossier.mkdir(parents=True, exist_ok=True)
        self._url_vers_hash = {}   # url -> sha1 (None si échec)
        self._hash_vers_data = {}  # sha1 -> bytes
        self._hash_vers_nom = {}   # sha1 -> nom de fichier
        self.telechargements = 0

    def _charger(self, url):
        if url in self._url_vers_hash:
            return self._url_vers_hash[url]
        h = None
        try:
            resp = self.page.request.get(url)
            if resp.ok and resp.headers.get("content-type", "").startswith("image/"):
                data = resp.body()
                self.telechargements += 1
                h = hashlib.sha1(data).hexdigest()
                self._hash_vers_data.setdefault(h, data)
            else:
                print(f"    image ignorée : HTTP {resp.status} {resp.headers.get('content-type')}")
        except Exception as e:
            print(f"    image ignorée : {e}")
        self._url_vers_hash[url] = h
        return h

    def get_bytes(self, url):
        """Contenu de l'image (pour le PPTX). None si indisponible."""
        h = self._charger(url)
        return self._hash_vers_data.get(h) if h else None

    def get_fichier(self, url):
        """Nom du fichier local (écrit au premier appel). None si indisponible."""
        h = self._charger(url)
        if not h:
            return None
        if h not in self._hash_vers_nom:
            data = self._hash_vers_data[h]
            nom = f"{h[:16]}.{ext_image(data)}"
            chemin = self.dossier / nom
            if not chemin.exists():  # aussi réutilisé d'une exécution à l'autre
                chemin.write_bytes(data)
            self._hash_vers_nom[h] = nom
        return self._hash_vers_nom[h]

    def localiser_svg(self, svg_code, prefixe="images"):
        """Remplace href/xlink:href distants par le chemin local relatif."""
        def remplacer(m):
            attr, url_xml = m.group(1), m.group(2)
            nom = self.get_fichier(html.unescape(url_xml))  # &amp; -> &
            return f'{attr}="{prefixe}/{nom}"' if nom else m.group(0)

        return HREF_RE.sub(remplacer, svg_code)

    def stats(self):
        return (f"{len(self._url_vers_hash)} URL(s), {self.telechargements} téléchargement(s), "
                f"{len(self._hash_vers_data)} image(s) unique(s)")


def main():
    
    
    link = input("Enter your link: ").strip()

    prs = Presentation()
    captures = []
    # viewBox Google = 960x540 -> 10 x 5.625 in (1 unité SVG = 9525 EMU)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(
            viewport={"width": 1920, "height": 1080},
            device_scale_factor=2,
        )



        page.goto(link)
        page.wait_for_function("() => window.viewerData && window.viewerData.docData")
        diction = page.evaluate("() => window.viewerData.docData")[1]
        base = link.split("/embed")[0] + "/embed?start=false&loop=false"
        
        name = page.title()
        dossier = Path(f"data/{name}")
        (dossier / "svg").mkdir(parents=True, exist_ok=True)
        (dossier / "svg" / "images").mkdir(parents=True, exist_ok=True)
        output_dir = Path("output")
        (output_dir).mkdir(parents=True, exist_ok=True)
        images = ImageManager(page, "data" /dossier / "svg" / "images")


        for cpt, liste in enumerate(diction, start=1):
            slide_id = liste[0]
            print(f"Diapositive {cpt} (id {slide_id})")

            page.goto(f"{base}#slide=id.{slide_id}")
            page.wait_for_timeout(500)

            navbar = page.locator(".punch-viewer-navbar")
            if navbar.count() > 0:
                navbar.evaluate("el => el.style.display = 'none'")

            svg = page.locator(".punch-viewer-svgpage-svgcontainer svg:visible").last
            svg.wait_for(state="visible")

            slide = prs.slides.add_slide(prs.slide_layouts[6])

            if EDITABLE:
                # XMLSerializer : xmlns corrects, \n des aria-label conservés
                svg_code = svg.evaluate("el => new XMLSerializer().serializeToString(el)")
                # PPTX : python-pptx dédoublonne aussi les images identiques (SHA-1) dans le fichier
                b = svg_to_slide(svg_code, slide, images.get_bytes)

                if SAVE_SVG:
                    svg_local = images.localiser_svg(svg_code)
                    (dossier / "svg" / f"{cpt}.svg").write_text(
                        '<?xml version="1.0" encoding="UTF-8"?>\n' + svg_local,
                        encoding="utf-8",
                    )

                print(f"  {len(b.text_boxes)} zone(s) de texte, {len(slide.shapes)} objet(s)")
            else:
                filename = dossier / f"{cpt}.png"
                svg.screenshot(path=str(filename))
                captures.append(filename)
                slide.shapes.add_picture(
                    str(filename), 0, 0,
                    width=prs.slide_width, height=prs.slide_height,
                )

        print("Images :", images.stats())
        browser.close()

    prs.save(output_dir / f"{name}.pptx")
    if captures:
        layout = img2pdf.get_layout_fun((img2pdf.in_to_pt(10), img2pdf.in_to_pt(5.625)))
        with open(output_dir / f"{name}.pdf", "wb") as f:
            f.write(img2pdf.convert([str(c) for c in captures], layout_fun=layout))
        print("PDF créé : presentation.pdf")
    print("PowerPoint créé : presentation.pptx")


if __name__ == "__main__":
    main()