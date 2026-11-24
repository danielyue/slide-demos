# Worklog

Last updated: 2026-11-24

## 2026-11-24

- Created the repo and GitHub Pages site with one demo (network tipping) and a sideloadable PowerPoint content add-in manifest, to test live HTML in PowerPoint slideshow and Presenter View on the Mac.
- Added `tipping-steps/`: five-step version advanced by on-screen arrows (and the clicker once focused); past the last step it calls `goToByIdAsync` to move PowerPoint on. Second add-in manifest; test deck v2 with both demos.
- Fixed: in WebKit (PowerPoint for Mac add-ins) the canvas grew its grid row and pushed the arrows off the frame; canvases now sit absolutely inside a box.
