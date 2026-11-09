"""Testdeck für den PPTX-Rundlauf: sechs Inhaltsfolien, wie ein Mensch sie in PowerPoint baut.

    python bench/testdeck.py [ziel.pptx]

Icons: PPTX_ICONS=<ordner> mit Paaren name.svg/name.png (z. B. eine Icon-Bibliothek); ohne den
Ordner zeichnet das Skript neutrale Ersatz-Icons (Kreis mit Initiale). PPTX_LOGO=<datei.png>
setzt ein Logo unten rechts. Schriften: das Theme nennt "Sparkasse Head/Rg/Medium" - wer sie
nicht hat, setzt PPTX_FONT_HEAD/PPTX_FONT_BODY/PPTX_FONT_MEDIUM (z. B. Arial).

Bewusst „fremd" gebaut (python-pptx + rohes DrawingML), nicht über den Export von Office
Creator: native Vorlageformen mit Anpassern, SVG-Icons mit PNG-Ersatzbild (so speichert
PowerPoint eingefügte Symbole), Schatten, Verläufe, Theme-Farben und -Schriften, Gruppen,
Verbinder mit Pfeilspitzen, ein Tabellenstil, runde Fotos mit Beschnitt, Aufzählungen.

Koordinaten in Pixeln der 1920x1080-Bühne (1 px = 6350 EMU), Schriftgrößen in Pixeln
(1 pt = 2 px) - so lassen sich Folie und importiertes HTML direkt vergleichen.
"""
from __future__ import annotations

import copy
import io
import pathlib
import sys

from lxml import etree
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import os

ROOT = pathlib.Path(__file__).resolve().parents[1]
GFX = pathlib.Path(os.environ["PPTX_ICONS"]) if os.environ.get("PPTX_ICONS") else None
LOGO = pathlib.Path(os.environ["PPTX_LOGO"]) if os.environ.get("PPTX_LOGO") else None
ZIEL = ROOT / "bench" / "out" / "testdeck.pptx"
F_HEAD = os.environ.get("PPTX_FONT_HEAD", "Sparkasse Head")
F_BODY = os.environ.get("PPTX_FONT_BODY", "Sparkasse Rg")
F_MED = os.environ.get("PPTX_FONT_MEDIUM", "Sparkasse Medium")

ROT, TINTE, GRAU, HELL, WEISS = "EE0000", "1F252C", "66707C", "F2F4F7", "FFFFFF"
GRUEN, BLAU, GOLD = "2BB673", "0099E6", "F2B01E"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


def px(v: float) -> Emu:
    return Emu(int(round(v * 6350)))


# ---------------------------------------------------------------------- Theme
def theme_anpassen(prs: Presentation) -> None:
    """Sparkassen-Farben als Akzente, Sparkassen-Schriften als Theme-Schriften."""
    master = prs.slide_masters[0]
    theme_part = next(r.target_part for r in master.part.rels.values() if r.reltype == RT.THEME)
    th = etree.fromstring(theme_part.blob)
    farben = {"dk2": "1F252C", "lt2": "F2F4F7", "accent1": ROT, "accent2": "1F252C",
              "accent3": GRAU, "accent4": BLAU, "accent5": GRUEN, "accent6": GOLD}
    for name, wert in farben.items():
        el = th.find(f".//{{{A}}}clrScheme/{{{A}}}{name}")
        for c in list(el):
            el.remove(c)
        etree.SubElement(el, f"{{{A}}}srgbClr", val=wert)
    th.find(f".//{{{A}}}majorFont/{{{A}}}latin").set("typeface", F_HEAD)
    th.find(f".//{{{A}}}minorFont/{{{A}}}latin").set("typeface", F_BODY)
    theme_part._blob = etree.tostring(th, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------- Grundbausteine
def _spPr(shape):
    return shape._element.spPr


def schatten(shape, blur=28, dist=8, alpha=18, farbe="1F252C", winkel=90):
    sp = _spPr(shape)
    eff = sp.find(qn("a:effectLst"))
    if eff is None:
        eff = etree.SubElement(sp, qn("a:effectLst"))
    o = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(int(blur * 6350)), dist=str(int(dist * 6350)),
                         dir=str(winkel * 60000), algn="t", rotWithShape="0")
    c = etree.SubElement(o, qn("a:srgbClr"), val=farbe)
    etree.SubElement(c, qn("a:alpha"), val=str(alpha * 1000))


def verlauf(shape, stops, winkel=90):
    """stops: [(pos 0..100, 'RRGGBB'), ...]; winkel in Grad (0 = links -> rechts)."""
    sp = _spPr(shape)
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for el in sp.findall(qn(tag)):
            sp.remove(el)
    gf = etree.Element(qn("a:gradFill"), rotWithShape="1")
    lst = etree.SubElement(gf, qn("a:gsLst"))
    for pos, farbe in stops:
        gs = etree.SubElement(lst, qn("a:gs"), pos=str(int(pos * 1000)))
        etree.SubElement(gs, qn("a:srgbClr"), val=farbe)
    etree.SubElement(gf, qn("a:lin"), ang=str(int(winkel * 60000)), scaled="0")
    geom = sp.find(qn("a:prstGeom"))
    geom.addnext(gf)


def form(slide, prst, x, y, w, h, fill=None, linie=None, linienbreite=0, adj=None, theme_fill=None,
         rot=0, ohne_text=True):
    s = slide.shapes.add_shape(prst, px(x), px(y), px(w), px(h))
    if adj:
        for i, v in enumerate(adj):
            s.adjustments[i] = v
    if theme_fill:
        sp = _spPr(s)
        for el in sp.findall(qn("a:solidFill")):
            sp.remove(el)
        sf = etree.Element(qn("a:solidFill"))
        etree.SubElement(sf, qn("a:schemeClr"), val=theme_fill)
        sp.find(qn("a:prstGeom")).addnext(sf)
    elif fill:
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        s.fill.background()
    if linie:
        s.line.color.rgb = RGBColor.from_string(linie)
        s.line.width = px(linienbreite or 2)
    else:
        s.line.fill.background()
    if rot:
        s.rotation = rot
    # python-pptx setzt einen Stil (lnRef/fillRef/fontRef) - der bleibt, wie bei echten Folien.
    if ohne_text:
        s.text_frame.text = ""
    return s


def text(slide_or_shape, x=None, y=None, w=None, h=None, absaetze=(), anker="t", raender=None,
         umbruch=True, autofit=None):
    """absaetze: [dict(runs=[(text, {font,size,bold,italic,color,theme_color,spc})], align, space_before,
    space_after, line, bullet, level, marL, indent)]"""
    if x is not None:
        tb = slide_or_shape.shapes.add_textbox(px(x), px(y), px(w), px(h))
    else:
        tb = slide_or_shape
    tf = tb.text_frame
    tf.word_wrap = umbruch
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "ctr": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anker]
    if raender is not None:
        l, t, r, b = raender
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = px(l), px(t), px(r), px(b)
    bp = tf._txBody.find(qn("a:bodyPr"))
    if autofit == "norm":
        for el in list(bp):
            bp.remove(el)
        etree.SubElement(bp, qn("a:normAutofit"))
    for i, a in enumerate(absaetze):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if a.get("align"):
            p.alignment = {"l": PP_ALIGN.LEFT, "ctr": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT,
                           "just": PP_ALIGN.JUSTIFY}[a["align"]]
        if a.get("space_before") is not None:
            p.space_before = Pt(a["space_before"] / 2)
        if a.get("space_after") is not None:
            p.space_after = Pt(a["space_after"] / 2)
        if a.get("line"):
            p.line_spacing = a["line"]
        pPr = p._p.get_or_add_pPr()
        if a.get("bullet"):
            pPr.set("marL", str(int(a.get("marL", 30) * 6350)))
            pPr.set("indent", str(int(a.get("indent", -30) * 6350)))
            if a.get("bullet_color"):
                bc = etree.SubElement(pPr, qn("a:buClr"))
                etree.SubElement(bc, qn("a:srgbClr"), val=a["bullet_color"])
            etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
            etree.SubElement(pPr, qn("a:buChar"), char=a["bullet"])
        for txt, st in a.get("runs", []):
            r = p.add_run()
            r.text = txt
            f = r.font
            if st.get("font"):
                f.name = st["font"]
            if st.get("size"):
                f.size = Pt(st["size"] / 2)
            if st.get("bold"):
                f.bold = True
            if st.get("italic"):
                f.italic = True
            if st.get("color"):
                f.color.rgb = RGBColor.from_string(st["color"])
            if st.get("theme_color"):
                rPr = r._r.get_or_add_rPr()
                sf = etree.SubElement(rPr, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:schemeClr"), val=st["theme_color"])
            if st.get("spc"):
                r._r.get_or_add_rPr().set("spc", str(int(st["spc"] * 100)))
            if st.get("theme_font"):
                rPr = r._r.get_or_add_rPr()
                etree.SubElement(rPr, qn("a:latin"), typeface=st["theme_font"])
    return tb


def run(t, **st):
    return (t, st)


# ---------------------------------------------------------------------- Bilder und Icons
_svg_n = [0]


def _gefaerbtes_png(png: pathlib.Path, farbe: str, groesse: int) -> bytes:
    im = Image.open(png).convert("RGBA")
    alpha = im.getchannel("A")
    voll = Image.new("RGBA", im.size, "#" + farbe)
    voll.putalpha(alpha)
    voll = voll.resize((groesse, groesse), Image.LANCZOS)
    b = io.BytesIO()
    voll.save(b, "PNG")
    return b.getvalue()


def _ersatz_icon(name: str, farbe: str) -> tuple[str, bytes]:
    """Neutrales Icon ohne Bibliothek: Ring mit Initiale, als SVG und PNG."""
    ch = name[:1].upper()
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 36 36">'
           f'<circle cx="18" cy="18" r="15" fill="none" stroke="#{farbe}" stroke-width="3"/>'
           f'<rect x="12" y="11" width="12" height="14" rx="2" fill="#{farbe}"/></svg>')
    im = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    col = "#" + farbe
    d.ellipse([22, 22, 234, 234], outline=col, width=21)
    d.rounded_rectangle([85, 78, 171, 178], 14, fill=col)
    b = io.BytesIO()
    im.save(b, "PNG")
    return svg, b.getvalue()


def icon(slide_or_group, prs, name, x, y, size, farbe):
    """Ein Symbol wie aus „Einfügen > Piktogramme": SVG plus PNG-Ersatzbild, beide eingefärbt."""
    if GFX and (GFX / f"{name}.svg").exists():
        svg = (GFX / f"{name}.svg").read_text(encoding="utf-8")
        for alt in ("#666666", "#666", "#000000", "#000"):
            svg = svg.replace(f'fill="{alt}"', f'fill="#{farbe}"')
        png = _gefaerbtes_png(GFX / f"{name}.png", farbe, 256)
    else:
        svg, png = _ersatz_icon(name, farbe)
    pic = slide_or_group.shapes.add_picture(io.BytesIO(png), px(x), px(y), px(size), px(size))
    part = slide_or_group.part if hasattr(slide_or_group, "part") else pic.part
    _svg_n[0] += 1
    svg_part = Part(PackURI(f"/ppt/media/icon_svg{_svg_n[0]}.svg"), "image/svg+xml", prs.part.package,
                    svg.encode("utf-8"))
    rid = part.relate_to(svg_part, RT.IMAGE)
    blip = pic._element.find(".//" + qn("a:blip"))
    ext_lst = etree.SubElement(blip, qn("a:extLst"))
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri=SVG_EXT)
    sb = etree.SubElement(ext, f"{{{SVG_NS}}}svgBlip", nsmap={"asvg": SVG_NS})
    sb.set(qn("r:embed"), rid)
    pic.name = f"Grafik {name}"
    return pic


def foto(farbe_a, farbe_b, motiv, groesse=(900, 600)) -> bytes:
    """Ein „Foto": weicher Verlauf mit einer Silhouette, genug Struktur für den Beschnitt."""
    w, h = groesse
    im = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(im)
    a, b = Image.new("RGB", (1, 1), "#" + farbe_a).getpixel((0, 0)), Image.new("RGB", (1, 1), "#" + farbe_b).getpixel((0, 0))
    for yy in range(h):
        u = yy / h
        d.line([(0, yy), (w, yy)], fill=tuple(int(a[i] + (b[i] - a[i]) * u) for i in range(3)))
    if motiv == "person":
        d.ellipse([w * .36, h * .16, w * .64, h * .56], fill=(245, 222, 200))
        d.ellipse([w * .22, h * .58, w * .78, h * 1.25], fill=(40, 52, 66))
        d.pieslice([w * .35, h * .12, w * .65, h * .42], 180, 360, fill=(70, 50, 40))
    else:
        for i in range(6):
            d.rectangle([w * (.08 + i * .15), h * (.35 - (i % 3) * .08), w * (.18 + i * .15), h * .9],
                        fill=(255, 255, 255, 120))
    im = im.filter(ImageFilter.GaussianBlur(1.2))
    bio = io.BytesIO()
    im.save(bio, "JPEG", quality=90)
    return bio.getvalue()


def verbinder(slide, x1, y1, x2, y2, farbe, breite=3, strich=None, pfeil=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(x1), px(y1), px(x2), px(y2))
    c.line.color.rgb = RGBColor.from_string(farbe)
    c.line.width = px(breite)
    ln = c.line._get_or_add_ln()
    if strich:
        etree.SubElement(ln, qn("a:prstDash"), val=strich)
    if pfeil:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return c


# ---------------------------------------------------------------------- Rahmen jeder Folie
def rahmen(prs, slide, nr, kicker, titel):
    form(slide, MSO_SHAPE.RECTANGLE, 0, 0, 16, 1080, fill=ROT)
    text(slide, 96, 52, 1300, 40, [dict(runs=[run(kicker.upper(), font=F_BODY, size=24,
                                                   color=GRAU, spc=1.5)])], raender=(0, 0, 0, 0))
    text(slide, 96, 92, 1500, 90, [dict(runs=[run(titel, theme_font="+mj-lt", size=64, theme_color="accent1")])],
         raender=(0, 0, 0, 0))
    verbinder(slide, 96, 1010, 1824, 1010, "D3D9E0", 2)
    text(slide, 96, 1022, 900, 36, [dict(runs=[run("pptx-to-dom · Testdeck", font=F_BODY,
                                                    size=20, color=GRAU)])], raender=(0, 0, 0, 0))
    text(slide, 1500, 1022, 120, 36, [dict(align="r", runs=[run(str(nr), font=F_BODY, size=20, color=GRAU)])],
         raender=(0, 0, 0, 0))
    if LOGO and LOGO.exists():
        slide.shapes.add_picture(str(LOGO), px(1652), px(1016), height=px(48))


# ---------------------------------------------------------------------- Folien
def folie_bento(prs, s):
    rahmen(prs, s, 1, "Jahresbilanz 2026", "Unser Jahr in Zahlen")
    # große rote Kachel
    k = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, 96, 230, 620, 740, fill=ROT, adj=[0.06])
    schatten(k, 30, 10, 22, "990000")
    icon(s, prs, "personen-drei_1", 150, 290, 110, WEISS)
    text(s, 150, 560, 520, 150, [dict(runs=[run("1,2 Mio.", font=F_HEAD, size=120, color=WEISS)])],
         raender=(0, 0, 0, 0))
    text(s, 150, 720, 520, 200, [
        dict(runs=[run("Kundinnen und Kunden", font=F_MED, size=40, color=WEISS)], space_after=12),
        dict(runs=[run("vertrauen uns ihre Finanzen an - so viele wie nie zuvor.", font=F_BODY, size=28,
                       color="FFE3E3")], line=1.1)], raender=(0, 0, 0, 0))
    # vier Kacheln rechts, 2x2, eine davon dunkel, eine mit Verlauf
    kacheln = [(756, 230, "sparschwein_1", "Spareinlagen", "+8,4 %", "18,6 Mrd. Euro", "w"),
               (1300, 230, "smartphone_1", "Digital aktiv", "74 %", "nutzen App und Online-Banking", "d"),
               (756, 610, "wohnhaus_1", "Baufinanzierung", "2.870", "neue Finanzierungen", "g"),
               (1300, 610, "rakete_1", "Start-ups", "112", "Gründungen begleitet", "w")]
    for x, y, ic, label, zahl, unter, art in kacheln:
        fill = {"w": WEISS, "d": TINTE, "g": None}[art]
        k = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 524, 360, fill=fill or WEISS, adj=[0.08],
                 linie="E3E7EC" if art == "w" else None, linienbreite=2)
        if art == "g":
            verlauf(k, [(0, "FFF1CC"), (100, "FFD66B")], 45)
        if art != "d":
            schatten(k, 24, 6, 14)
        vorder = WEISS if art == "d" else TINTE
        kreis = form(s, MSO_SHAPE.OVAL, x + 40, y + 40, 96, 96, fill=ROT if art != "d" else WEISS)
        icon(s, prs, ic, x + 60, y + 60, 56, WEISS if art != "d" else ROT)
        text(s, x + 160, y + 52, 330, 70, [dict(runs=[run(label, font=F_MED, size=32, color=vorder)])],
             raender=(0, 0, 0, 0), anker="ctr")
        text(s, x + 40, y + 170, 450, 110, [dict(runs=[run(zahl, font=F_HEAD, size=88,
                                                          color=ROT if art != "d" else WEISS)])], raender=(0, 0, 0, 0))
        text(s, x + 40, y + 280, 450, 50, [dict(runs=[run(unter, font=F_BODY, size=26,
                                                         color=GRAU if art != "d" else "C6CED7")])],
             raender=(0, 0, 0, 0))


def folie_kpi(prs, s):
    rahmen(prs, s, 2, "Steuerung", "Kennzahlen im Überblick")
    band = form(s, MSO_SHAPE.RECTANGLE, 96, 220, 1728, 12, fill=ROT)
    verlauf(band, [(0, ROT), (100, "FF8A8A")], 0)
    karten = [("balkendiagramm_1", "Betriebsergebnis", "412 Mio. €", "+6,1 %", True),
              ("kurs-steigend_1", "Kreditvolumen", "24,9 Mrd. €", "+3,8 %", True),
              ("tacho_1", "Cost-Income-Ratio", "61,2 %", "-2,4 Pp.", True),
              ("kurs-fallend_1", "Risikovorsorge", "38 Mio. €", "+12 %", False)]
    for i, (ic, label, wert, trend, gut) in enumerate(karten):
        x = 96 + i * 444
        k = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 280, 408, 470, fill=WEISS, adj=[0.05], linie="E3E7EC")
        schatten(k, 30, 8, 16)
        form(s, MSO_SHAPE.OVAL, x + 144, 320, 120, 120, theme_fill="accent1")
        icon(s, prs, ic, x + 172, 348, 64, WEISS)
        text(s, x + 20, 470, 368, 50, [dict(align="ctr", runs=[run(label, font=F_MED, size=30,
                                                                    color=TINTE)])], raender=(0, 0, 0, 0))
        text(s, x + 20, 530, 368, 100, [dict(align="ctr", runs=[run(wert, font=F_HEAD, size=72,
                                                                     color=TINTE)])], raender=(0, 0, 0, 0))
        farbe = GRUEN if gut else ROT
        pille = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 114, 660, 180, 54, fill="E8F7EF" if gut else "FFE3E3",
                     adj=[0.5])
        form(s, MSO_SHAPE.ISOSCELES_TRIANGLE, x + 134, 674, 26, 24, fill=farbe,
             rot=0 if trend.startswith("+") else 180)
        text(s, x + 168, 660, 120, 54, [dict(runs=[run(trend, font=F_BODY, size=26, bold=True, color=farbe)])],
             raender=(0, 0, 0, 0), anker="ctr")
    # Hinweisbox mit Akzentleiste
    form(s, MSO_SHAPE.RECTANGLE, 96, 800, 1728, 170, fill=HELL)
    form(s, MSO_SHAPE.RECTANGLE, 96, 800, 10, 170, fill=ROT)
    icon(s, prs, "gluehbirne_1", 146, 845, 80, ROT)
    text(s, 260, 815, 1500, 140, [
        dict(runs=[run("Einordnung: ", font=F_BODY, size=30, bold=True, color=TINTE),
                   run("Das Betriebsergebnis wächst trotz höherer Risikovorsorge. Treiber sind das Zinsergebnis "
                       "und eine konsequente Kostendisziplin.", font=F_BODY, size=30, color=TINTE)],
             line=1.15)], anker="ctr", raender=(0, 0, 0, 0))


def folie_zeitstrahl(prs, s):
    rahmen(prs, s, 3, "Roadmap", "Meilensteine 2026")
    verbinder(s, 140, 600, 1790, 600, "C6CED7", 8)
    verbinder(s, 140, 600, 1130, 600, ROT, 8)
    punkte = [("kalender_1", "Jan", "Strategie-Kick-off", "Ziele und Leitplanken mit dem Vorstand festgelegt", True),
              ("personen-zwei_1", "Mär", "Family & Friends", "Erste Testgruppe nutzt die neue App", False),
              ("chat_1", "Mai", "Nutzerbefragung", "1.400 Antworten, Note 1,8", True),
              ("rakete_1", "Sep", "Go-live", "Start für alle Privatkundinnen und -kunden", False),
              ("siegerehrung_1", "Dez", "Bilanz", "Wirkung messen und nachsteuern", True)]
    for i, (ic, monat, titel, unter, fertig) in enumerate(punkte):
        cx = 220 + i * 370
        oben = i % 2 == 0
        kreis = form(s, MSO_SHAPE.OVAL, cx - 60, 540, 120, 120, fill=ROT if fertig else WEISS,
                     linie=ROT if fertig else "C6CED7", linienbreite=6)
        schatten(kreis, 16, 4, 20)
        icon(s, prs, ic, cx - 32, 568, 64, WEISS if fertig else GRAU)
        pille = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx - 54, 680 if oben else 466, 108, 44,
                     fill=TINTE, adj=[0.5])
        text(pille, absaetze=[dict(align="ctr", runs=[run(monat, font=F_MED, size=24, color=WEISS)])],
             anker="ctr", raender=(0, 0, 0, 0))
        ty = 290 if oben else 760
        verbinder(s, cx, 470 if oben else 724, cx, 410 if oben else 760, "C6CED7", 3, strich="dash")
        karte = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx - 160, ty if oben else 770, 320, 170,
                     fill=WEISS, adj=[0.1], linie="E3E7EC")
        schatten(karte, 20, 4, 12)
        text(karte, absaetze=[dict(align="ctr", runs=[run(titel, font=F_MED, size=30, color=ROT)],
                                    space_after=8),
                               dict(align="ctr", runs=[run(unter, font=F_BODY, size=24, color=GRAU)],
                                    line=1.05)], anker="ctr", raender=(18, 12, 18, 12))


def folie_prozess(prs, s):
    rahmen(prs, s, 4, "Vorgehen", "So arbeiten wir")
    schritte = [("lupe_1", "Analyse", ["Bedarf erheben", "Daten sichten"]),
                ("gluehbirne_1", "Ideen", ["Workshops", "Prototypen"]),
                ("zahnrad_1", "Umsetzung", ["Agile Sprints", "Tests mit Kunden"]),
                ("checkliste_1", "Freigabe", ["Compliance", "Datenschutz"]),
                ("pfeile-kreislauf_1", "Betrieb", ["Monitoring", "Verbesserung"])]
    for i, (ic, titel, punkte) in enumerate(schritte):
        x = 96 + i * 346
        prst = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        ch = form(s, prst, x, 430, 360, 130, fill=ROT if i < 3 else (TINTE if i == 3 else GRAU), adj=[0.3])
        text(ch, absaetze=[dict(align="ctr", runs=[run(f"{i + 1}  ", font=F_HEAD, size=40, color="FFD0D0"),
                                                    run(titel, font=F_MED, size=32, color=WEISS)])],
             anker="ctr", raender=(40 if i else 16, 0, 30, 0))
        box = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 110, 250, 130, 130, fill=HELL, adj=[0.2])
        icon(s, prs, ic, x + 135, 275, 80, ROT)
        if i < 4:
            verbinder(s, x + 240, 315, x + 456, 315, "C6CED7", 3, strich="sysDot", pfeil=True)
        text(s, x + 20, 600, 310, 300, [dict(runs=[run(p, font=F_BODY, size=28, color=TINTE)], bullet="•",
                                             bullet_color=ROT, marL=34, indent=-34, space_after=10) for p in punkte],
             raender=(10, 10, 10, 10))
    # Callout „Wichtig"
    c = form(s, MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, 1190, 800, 620, 150, fill="FFF1CC", linie=GOLD,
             linienbreite=3, adj=[-0.3, 0.75, 0.16667])
    text(c, absaetze=[dict(runs=[run("Wichtig: ", font=F_BODY, size=28, bold=True, color="7A5600"),
                                 run("Jeder Schritt endet mit einer Freigabe durch die Fachseite.",
                                     font=F_BODY, size=28, color="7A5600")])], anker="ctr",
         raender=(24, 12, 24, 12))


def folie_tabelle(prs, s):
    rahmen(prs, s, 5, "Angebot", "Kontomodelle im Vergleich")
    zeilen = [["Leistung", "Giro Basis", "Giro Plus", "Giro Premium"],
              ["Monatlicher Preis", "4,90 €", "8,90 €", "14,90 €"],
              ["Girocard", "inklusive", "inklusive", "inklusive"],
              ["Kreditkarte", "-", "Basis", "Gold"],
              ["Bargeld weltweit", "-", "4x pro Jahr", "unbegrenzt"],
              ["Persönliche Beratung", "Filiale", "Filiale + Video", "Private Banking"]]
    gf = s.shapes.add_table(len(zeilen), 4, px(96), px(240), px(1100), px(600))
    t = gf.table
    for c, b in enumerate((380, 240, 240, 240)):
        t.columns[c].width = px(b)
    for r, zeile in enumerate(zeilen):
        t.rows[r].height = px(100)
        for c, wert in enumerate(zeile):
            zelle = t.cell(r, c)
            zelle.text = wert
            p = zelle.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            f = p.runs[0].font
            f.name = F_MED if r == 0 else F_BODY
            f.size = Pt(13 if r else 14)
            zelle.vertical_anchor = MSO_ANCHOR.MIDDLE
    # Seitenleiste mit Häkchen-Liste
    pan = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1250, 240, 574, 600, fill=HELL, adj=[0.05])
    text(s, 1290, 270, 500, 60, [dict(runs=[run("Darum lohnt sich Plus", font=F_MED, size=34,
                                                color=TINTE)])], raender=(0, 0, 0, 0))
    vorteile = ["Kreditkarte inklusive", "Bargeld auch im Urlaub", "Videoberatung am Abend",
                "Versicherungsschutz für Ihr Smartphone"]
    for i, v in enumerate(vorteile):
        y = 360 + i * 110
        grp = s.shapes.add_group_shape()
        form(grp, MSO_SHAPE.OVAL, 1290, y, 64, 64, fill=GRUEN)
        icon(grp, prs, "checkmark_1", 1304, y + 14, 36, WEISS)
        text(grp, 1374, y - 6, 420, 76, [dict(runs=[run(v, font=F_BODY, size=28, color=TINTE)])],
             anker="ctr", raender=(0, 0, 0, 0))
    hinweis = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, 96, 880, 1728, 90, fill=WEISS, linie=ROT, linienbreite=3, adj=[0.5])
    text(hinweis, absaetze=[dict(align="ctr", runs=[
        run("Tipp: ", font=F_BODY, size=28, bold=True, color=ROT),
        run("Für junge Leute bis 25 ist Giro Plus kostenlos.", font=F_BODY, size=28, color=TINTE)])],
         anker="ctr")


def folie_team(prs, s):
    rahmen(prs, s, 6, "Ansprechpartner", "Ihr Team vor Ort")
    personen = [("Anna Becker", "Leiterin Privatkunden", "0221 226-1234", "6FA8DC", "CFE2F3"),
                ("Murat Yilmaz", "Firmenkundenberater", "0221 226-2345", "93C47D", "D9EAD3"),
                ("Lea Hoffmann", "Digitalberaterin", "0221 226-3456", "E69138", "FCE5CD")]
    for i, (name, rolle, tel, fa, fb) in enumerate(personen):
        x = 96 + i * 400
        karte = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 240, 370, 560, fill=WEISS, adj=[0.06], linie="E3E7EC")
        schatten(karte, 30, 10, 16)
        bild = s.shapes.add_picture(io.BytesIO(foto(fa, fb, "person")), px(x + 85), px(280), px(200), px(200))
        bild.crop_left, bild.crop_right = 0.16, 0.16           # 900x600 -> quadratisch
        bild.auto_shape_type = MSO_SHAPE.OVAL
        bild.line.color.rgb = RGBColor.from_string(WEISS)
        bild.line.width = px(6)
        text(s, x + 20, 500, 330, 50, [dict(align="ctr", runs=[run(name, font=F_BODY, size=34, bold=True,
                                                                     color=TINTE)])], raender=(0, 0, 0, 0))
        text(s, x + 20, 552, 330, 40, [dict(align="ctr", runs=[run(rolle, font=F_BODY, size=26,
                                                                     color=GRAU)])], raender=(0, 0, 0, 0))
        verbinder(s, x + 40, 620, x + 330, 620, "E3E7EC", 2)
        for j, (ic, wert) in enumerate((("telefon_1", tel), ("brief_1", name.split()[0].lower() + "@example.org"))):
            grp = s.shapes.add_group_shape()
            icon(grp, prs, ic, x + 50, 650 + j * 64, 40, ROT)
            text(grp, x + 104, 646 + j * 64, 250, 48, [dict(runs=[run(wert, font=F_BODY, size=24,
                                                                       color=TINTE)])], anker="ctr", raender=(0, 0, 0, 0))
    # Zitat und Abzeichen rechts
    z = form(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1310, 240, 514, 400, fill=TINTE, adj=[0.06])
    text(s, 1340, 230, 120, 140, [dict(runs=[run("„", font=F_HEAD, size=180, color=ROT)])],
         raender=(0, 0, 0, 0))
    text(s, 1350, 360, 440, 260, [
        dict(runs=[run("Wir sind da, wo unsere Kundinnen und Kunden sind - in der Filiale und digital.",
                       font=F_BODY, size=30, italic=True, color=WEISS)], line=1.2, space_after=16),
        dict(runs=[run("Anna Becker", font=F_MED, size=26, color="FF8A8A")])], raender=(0, 0, 0, 0))
    stern = form(s, MSO_SHAPE.STAR_5_POINT, 1660, 160, 150, 150, fill=GOLD, rot=12)
    text(stern, absaetze=[dict(align="ctr", runs=[run("Top", font=F_HEAD, size=32, color=TINTE)])],
         anker="ctr", raender=(0, 0, 0, 0))
    # Fortschritt als Bogen + Kreis
    form(s, MSO_SHAPE.DONUT, 1360, 690, 280, 280, fill="E3E7EC", adj=[0.12])
    form(s, MSO_SHAPE.BLOCK_ARC, 1360, 690, 280, 280, fill=ROT, adj=[-90, 180, 0.12])
    text(s, 1400, 770, 200, 120, [dict(align="ctr", runs=[run("75 %", font=F_HEAD, size=56, color=TINTE)]),
                                   dict(align="ctr", runs=[run("Terminquote", font=F_BODY, size=22,
                                                               color=GRAU)])], raender=(0, 0, 0, 0), anker="ctr")
    text(s, 1660, 740, 170, 200, [dict(runs=[run("Termine innerhalb von 48 Stunden", font=F_BODY, size=24,
                                                 color=GRAU)], line=1.1)], raender=(0, 0, 0, 0), anker="ctr")


def bauen(ziel: pathlib.Path = ZIEL) -> pathlib.Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = px(1920), px(1080)
    theme_anpassen(prs)
    leer = prs.slide_layouts[6]
    for f in (folie_bento, folie_kpi, folie_zeitstrahl, folie_prozess, folie_tabelle, folie_team):
        f(prs, prs.slides.add_slide(leer))
    prs.core_properties.title = "Testdeck Rundlauf"
    ziel.parent.mkdir(parents=True, exist_ok=True)
    prs.save(ziel)
    return ziel


if __name__ == "__main__":
    print(bauen(pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else ZIEL))
