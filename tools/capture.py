# /// script
# dependencies = ["playwright==1.64.0"]
# ///
"""Snapshot and video of a demo page, for the slide's still picture and an MP4 fallback slide.

  uv run tools/capture.py URL --png still.png [--step 4]          still at a given step (?step=N)
  uv run tools/capture.py URL --mp4 demo.mp4 --clicks 4 [--dwell 3]  click › through the steps and record
  uv run tools/capture.py URL --mp4 demo.mp4 --seconds 20          record a page that animates on its own

Uses WebKit, the engine PowerPoint for Mac uses for add-ins, at 1920x1080. Needs ffmpeg for --mp4.
"""
import argparse, asyncio, shutil, subprocess, tempfile
from pathlib import Path
from playwright.async_api import async_playwright


async def main(a):
    async with async_playwright() as p:
        b = await p.webkit.launch()
        if a.png:
            pg = await b.new_page(viewport={"width": 1920, "height": 1080})
            sep = "&" if "?" in a.url else "?"
            await pg.goto(a.url + (f"{sep}step={a.step}" if a.step is not None else ""))
            await pg.wait_for_timeout(2000)
            await pg.screenshot(path=a.png)
            print("wrote", a.png)
        if a.mp4:
            tmp = tempfile.mkdtemp()
            ctx = await b.new_context(viewport={"width": 1920, "height": 1080}, record_video_dir=tmp,
                                      record_video_size={"width": 1920, "height": 1080})
            pg = await ctx.new_page()
            await pg.goto(a.url)
            await pg.wait_for_timeout(int(a.dwell * 1000))
            for _ in range(a.clicks):
                await pg.click("#next")
                await pg.wait_for_timeout(int(a.dwell * 1000))
            if a.seconds:
                await pg.wait_for_timeout(int(a.seconds * 1000))
            await ctx.close()
            webm = next(Path(tmp).glob("*.webm"))
            subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-ss", "1", "-i", str(webm), "-c:v", "libx264",
                            "-pix_fmt", "yuv420p", "-crf", "20", "-movflags", "+faststart", a.mp4], check=True)
            shutil.rmtree(tmp)
            print("wrote", a.mp4)
        await b.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("url"); ap.add_argument("--png"); ap.add_argument("--step", type=int)
    ap.add_argument("--mp4"); ap.add_argument("--clicks", type=int, default=0)
    ap.add_argument("--dwell", type=float, default=3.0); ap.add_argument("--seconds", type=float, default=0)
    asyncio.run(main(ap.parse_args()))
