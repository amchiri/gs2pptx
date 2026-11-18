from playwright.sync_api import sync_playwright
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches


link = input("Enter your link: ")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto(link)
    page.wait_for_timeout(2000)

    diction = page.evaluate("() => window.viewerData.docData")
    diction = diction[1]

    # Création du dossier des images
    dossier = Path("mes_diapos")
    dossier.mkdir(exist_ok=True)

    # Création de la présentation PowerPoint
    prs = Presentation()

    # Format 16:9
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)

    for cpt, liste in enumerate(diction[3:], start=1):
        
        slide_id = liste[0]

        url = (
            link.split("/embed")[0]
            + "/embed?start=false&loop=false"
            + "#slide=id." + slide_id
        )

        page.goto(url)
        page.wait_for_timeout(300)

        # Masquer la barre Google Slides
        navbar = page.locator(".punch-viewer-navbar")

        if navbar.count() > 0:
            navbar.evaluate(
                "element => element.style.display = 'none'"
            )

        # Capture du SVG
        # svg = page.locator("svg").nth(0)
        # svg.wait_for(state="visible")
        
        # svgs = page.locator("svg")
        # for i in range(svgs.count()):
        #     svg = svgs.nth(i)

        #     infos = svg.evaluate("""element => ({
        #         classe: element.getAttribute("class"),
        #         taille: element.outerHTML.length,
        #         viewBox: element.getAttribute("viewBox"),
        #         enfants: element.children.length
        #     })""")

        #     print(f"SVG {i} :", infos)
        
        # for i in range(page.locator("svg").count()):
        #     svg = page.locator("svg").nth(i)

        #     infos = svg.evaluate("""element => ({
        #     taille: element.outerHTML.length,
        #     parent: element.parentElement?.tagName,
        #     classeParent: element.parentElement?.getAttribute("class"),
        #     visible: element.getBoundingClientRect().width > 0
        #     })""")

        #     print(f"SVG {i} :", infos)
        
        svg = page.locator(
            ".punch-viewer-svgpage-svgcontainer svg:visible"
        ).last

        svg_code = svg.evaluate("element => element.outerHTML")

        print("Taille du SVG :", len(svg_code))
            
        svg_code = svg.evaluate("element => element.outerHTML")
        print(svg_code)
        open("svg_code.html", "w", encoding="utf-8").write(svg_code)
        exit()

        filename = dossier / f"{cpt}.png"
        svg.screenshot(path=str(filename))

        print(f"Diapositive {cpt} capturée")

        # Ajouter une diapositive PowerPoint
        slide = prs.slides.add_slide(
            prs.slide_layouts[6]
        )

        # Insérer l'image sur toute la diapositive
        slide.shapes.add_picture(
            str(filename),
            0,
            0,
            width=prs.slide_width,
            height=prs.slide_height
        )

    # Enregistrer le PowerPoint
    prs.save("presentation.pptx")

    browser.close()

print("PowerPoint créé : presentation.pptx")
