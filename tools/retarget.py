# /// script
# dependencies = []
# ///
"""Switch a deck's live slides between desktop PowerPoint and PowerPoint for the web.

  uv run tools/retarget.py deck.pptx --to web       before opening it from OneDrive in a browser
  uv run tools/retarget.py deck.pptx --to desktop   before presenting from the Mac app
  uv run tools/retarget.py deck.pptx                just report which target each live slide has

Why: each add-in on a slide names how to find it. Desktop PowerPoint finds a sideloaded add-in under
storeType "Registry"; PowerPoint for the web finds one uploaded in the browser under "uploadfiledevcatalog".
Each honours only the reference listed first (tested 2026-11-24), so this rewrites that order. The deck is
edited in place; a copy is kept as deck.pptx.bak. Close the deck in PowerPoint first.
"""
import argparse, re, shutil, zipfile
from pathlib import Path

FIRST = {"desktop": "Registry", "web": "uploadfiledevcatalog"}


def refs_xml(addin_id, version, first):
    second = "uploadfiledevcatalog" if first == "Registry" else "Registry"
    return (f'<we:reference id="{addin_id}" version="{version}" store="developer" storeType="{first}"/>'
            f'<we:alternateReferences><we:reference id="{addin_id}" version="{version}" store="developer" storeType="{second}"/>'
            f'<we:reference id="{addin_id}" version="{version}" store="" storeType="omex"/></we:alternateReferences>')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck"); ap.add_argument("--to", choices=["desktop", "web"])
    a = ap.parse_args()
    deck = Path(a.deck)
    with zipfile.ZipFile(deck) as z:
        items = {n: z.read(n) for n in z.namelist()}
    changed = 0
    for name in sorted(n for n in items if re.match(r"ppt/webextensions/webextension\d+\.xml$", n)):
        xml = items[name].decode("utf-8")
        m = re.search(r'<we:reference id="([^"]+)" version="([^"]+)"[^>]*storeType="([^"]+)"\s*/>', xml)
        if not m:
            continue
        addin_id, version, store_type = m.groups()
        now = "desktop" if store_type == "Registry" else "web" if store_type == "uploadfiledevcatalog" else store_type
        print(f"{name}: add-in {addin_id[:8]}... currently {now}")
        if a.to and now != a.to:
            xml = re.sub(r"<we:reference [^>]*/>\s*(<we:alternateReferences\s*/>|<we:alternateReferences>.*?</we:alternateReferences>)",
                         refs_xml(addin_id, version, FIRST[a.to]), xml, count=1, flags=re.S)
            items[name] = xml.encode("utf-8"); changed += 1
    if changed:
        shutil.copy2(deck, str(deck) + ".bak")
        tmp = deck.with_suffix(".tmp.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n, data in items.items():
                z.writestr(n, data)
        tmp.replace(deck)
        print(f"retargeted {changed} live slide(s) to {a.to}; backup at {deck}.bak")


if __name__ == "__main__":
    main()
