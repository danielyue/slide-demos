# /// script
# dependencies = ["python-pptx", "lxml"]
# ///
"""Build the test deck: native slide, live simulation, stepped demo, MP4 fallback, native slide.

Each live slide embeds a sideloaded content add-in (addin/manifest-*.xml) by writing the
webextension parts python-pptx has no API for. If PowerPoint does not pick it up, the
slide shows the snapshot picture and the add-in can be inserted by hand instead.

Usage: uv run tools/build_test_deck.py TEMPLATE.pptx ASSETS_DIR OUT.pptx
ASSETS_DIR holds network-tipping.png, tipping-steps.png and network-tipping.mp4.
"""
import copy
import sys
import uuid
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent.parent
ADDINS = {  # <Id> and <DisplayName> in addin/manifest-*.xml
    "network-tipping": ("8c390e32-0980-46f6-8ad3-1bf81103aa77", "Slide Demo: Network tipping"),
    "tipping-steps": ("3315f584-a0eb-49c6-94da-c480cac51e2d", "Slide Demo: Tipping in steps"),
    "viewer": ("157ce4b8-984b-4227-9259-e47e662996de", "Slide Demos (viewer)"),  # the one to use from now on
}


def viewer_slide(prs, snapshot, url, n, notes=""):
    """A live slide through the Slide Demos viewer: the page address is stored in the slide itself."""
    return live_slide(prs, snapshot, "viewer", n, notes, props={"url": url})
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
# Which PowerPoint resolves the add-in: "desktop" (Mac sideload, storeType Registry) or "web" (uploaded in
# the browser, uploadfiledevcatalog). Each honours only its own reference when it is listed first (tested
# 2026-11-24), so a deck is built for one or the other; tools/retarget.py flips an existing deck.
TARGET = "desktop"
RT_WEBEXT = "http://schemas.microsoft.com/office/2011/relationships/webextension"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_WEBEXT = "application/vnd.ms-office.webextension+xml"


def drop_all_slides(prs):
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        prs.part.drop_rel(sld_id.get(f"{{{NS['r']}}}id"))
        sld_ids.remove(sld_id)


def layout(prs, name):
    return next(l for l in prs.slide_layouts if l.name == name)


def text_slide(prs, title, bullets):
    s = prs.slides.add_slide(layout(prs, "TITLE_AND_BODY"))
    s.shapes.title.text_frame.text = title
    body = s.placeholders[1].text_frame
    for i, b in enumerate(bullets):
        para = body.paragraphs[0] if i == 0 else body.add_paragraph()
        run = para.add_run()
        run.text = b
        run.font.size = Pt(16)
    return s


def live_slide(prs, snapshot, demo, n, notes, props=None):
    """props: add-in settings stored in the slide (Office.context.document.settings), e.g. {"url": ...}.
    Two references: "Registry" is how desktop PowerPoint names a sideloaded add-in, "uploadfiledevcatalog"
    how PowerPoint for the web names one uploaded in the browser; listing both lets one file work in each."""
    addin_id, addin_name = ADDINS[demo]
    first, second = ("uploadfiledevcatalog", "Registry") if TARGET == "web" else ("Registry", "uploadfiledevcatalog")
    import json
    from xml.sax.saxutils import quoteattr
    props_xml = "".join(f'<we:property name={quoteattr(k)} value={quoteattr(json.dumps(v))}/>' for k, v in (props or {}).items())
    s = prs.slides.add_slide(layout(prs, "BLANK"))
    w, h = prs.slide_width, prs.slide_height
    img_part, img_rid = s.part.get_or_add_image_part(snapshot)

    # The webextension part: which add-in, plus a snapshot picture for viewers without it.
    webext_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" xmlns:r="{NS['r']}" id="{{{str(uuid.uuid4()).upper()}}}">
  <we:reference id="{addin_id}" version="1.0.0.0" store="developer" storeType="{first}"/>
  <we:alternateReferences>
    <we:reference id="{addin_id}" version="1.0.0.0" store="developer" storeType="{second}"/>
    <we:reference id="{addin_id}" version="1.0.0.0" store="" storeType="omex"/>
  </we:alternateReferences>
  <we:properties>{props_xml}</we:properties>
  <we:bindings/>
  <we:snapshot r:embed="rId1"/>
</we:webextension>""".encode()
    webext = Part(PackURI(f"/ppt/webextensions/webextension{n}.xml"), CT_WEBEXT, s.part.package, webext_xml)
    assert webext.relate_to(img_part, RT_IMAGE) == "rId1"
    we_rid = s.part.relate_to(webext, RT_WEBEXT)

    frame = f"""
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:p="{NS['p']}" xmlns:a="{NS['a']}" xmlns:r="{NS['r']}">
  <mc:Choice xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" Requires="we">
    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="10" name="Live demo: {demo}"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">
          <we:webextensionref r:id="{we_rid}"/>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </mc:Choice>
  <mc:Fallback>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="10" name="Live demo: {demo}"/>
        <p:cNvPicPr><a:picLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill><a:blip r:embed="{img_rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
  </mc:Fallback>
</mc:AlternateContent>"""
    s.shapes._spTree.append(etree.fromstring(frame))
    s.notes_slide.notes_text_frame.text = notes + (
        "\n\nIf this slide shows a still picture, insert the add-in by hand: "
        f"Insert > Add-ins > My Add-ins > {addin_name}."
    )
    return s


def video_slide(prs, video, poster):
    s = prs.slides.add_slide(layout(prs, "BLANK"))
    s.shapes.add_movie(video, 0, 0, prs.slide_width, prs.slide_height,
                       poster_frame_image=poster, mime_type="video/mp4")
    s.notes_slide.notes_text_frame.text = "Fallback: the same demo as a 20-second video."
    return s


def main(template, assets, out):
    a = Path(assets)
    prs = Presentation(template)
    drop_all_slides(prs)
    text_slide(prs, "Network effects", [
        "The value of joining depends on who else joins",
        "Telephones, payment systems, programming languages",
    ])
    live_slide(prs, str(a / "network-tipping.png"), "network-tipping", 1,
               "Live simulation. Move the slider; Restart reruns it.")
    live_slide(prs, str(a / "tipping-steps.png"), "tipping-steps", 2,
               "Click the arrow (or, once the slide has focus, the clicker) for each step:\n"
               "1. Early users choose on quality, so most pick B\n"
               "2. A buys a lead with a launch promotion\n"
               "3. New users follow the installed base\n"
               "4. The market tips; some B users switch\n"
               "One more click moves to the next slide.")
    video_slide(prs, str(a / "network-tipping.mp4"), str(a / "network-tipping.png"))
    text_slide(prs, "When does the better product lose?", [
        "Strong network effects and an early lead",
        "What would let B catch up?",
    ])
    prs.save(out)
    print(out)


if __name__ == "__main__":
    main(*sys.argv[1:4])
