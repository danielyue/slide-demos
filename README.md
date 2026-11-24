# slide-demos

Small live web demos for teaching slides, served by GitHub Pages at https://danielyue.github.io/slide-demos/ and shown inside PowerPoint as a content add-in.

- `network-tipping/`: network effects and tipping; `?s=0.5` sets the starting strength.
- `addin/manifest.xml`: sideloadable PowerPoint content add-in that shows `network-tipping/`. On a Mac, copy it to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`, restart PowerPoint, then Insert > Add-ins > My Add-ins.

Started 2026-11-24 as a test of live HTML in PowerPoint; the plan is in `~/Code/admin/docs/research-tools/2026-11-24-slide-writer-plan.md`.
