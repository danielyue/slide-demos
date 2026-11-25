# slide-demos

Live web demos for teaching slides, served by GitHub Pages at https://danielyue.github.io/slide-demos/ and shown inside PowerPoint through one content add-in, the **Slide Demos viewer**. How to build, embed, host and test them: the `/slides` skill (`~/.claude/skills/slides/live/`).

| Path | What it is |
|:--|:--|
| `viewer/` | The add-in page. Each slide's copy stores the page it shows (setting `url`); "Change page" appears while editing. Demos ask it to change slide with `postMessage({type: 'slide-demos:nav', dir})`. Preview outside PowerPoint with `viewer/?url=...` |
| `addin/manifest-viewer.xml` | The one manifest to install (desktop: copy to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`, restart PowerPoint; web: Home > Add-ins > Advanced > Upload My Add-in) |
| `addin/manifest-*.xml` (others) | Older one-add-in-per-demo manifests, kept so the Nov 24 test decks still open |
| `network-tipping/`, `tipping-steps/` | Example demos (slider; stepped with arrows) |
| `templates/stepped-explainer/` | Starter for a slide that builds one click at a time |
| `demos.json` | The list the viewer offers |
| `tools/new_demo.sh` | New demo from the template (public, or `--private` path for unpublished work) |
| `tools/capture.py` | Still picture and MP4 fallback of a demo, recorded in WebKit |
| `tools/build_test_deck.py` | Builds .pptx files with live slides (`viewer_slide(prs, snapshot, url, n, notes)`) |

Private demos (unpublished papers) live in their own repos and are served from the MacBook at `https://mdr.dny/raw/<absolute path>`; they work only on that machine.
