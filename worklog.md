# Worklog

Last updated: 2026-11-24

## 2026-11-24

- Created the repo and GitHub Pages site with one demo (network tipping) and a sideloadable PowerPoint content add-in manifest, to test live HTML in PowerPoint slideshow and Presenter View on the Mac.
- Added `tipping-steps/`: five-step version advanced by on-screen arrows (and the clicker once focused); past the last step it calls `goToByIdAsync` to move PowerPoint on. Second add-in manifest; test deck v2 with both demos.
- Fixed: in WebKit (PowerPoint for Mac add-ins) the canvas grew its grid row and pushed the arrows off the frame; canvases now sit absolutely inside a box.
- Fixed: slideshow mode made text huge and stretched the canvas (PowerPoint reports a much taller window than is shown). Pages now lay out at 1600x900 and scale to the width, using the height only when it looks 16:9. D toggles a size readout.
- Added the Slide Demos viewer (one add-in for every live slide, URL stored per slide, nav via postMessage), demos.json, template, new_demo.sh, capture.py, dual desktop/web references in the builder. Office.js now loads only outside the viewer (in a frame it blanked the page).
- Desktop and web PowerPoint each honour only the first add-in reference (Registry vs uploadfiledevcatalog): builder TARGET switch and tools/retarget.py.
