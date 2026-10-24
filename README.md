# IMX: animated PowerPoint maker

IMX builds animated presentations in the browser and saves them as **PowerPoint (.pptx)** files. Every page has a moving background, and its text appears with PowerPoint animations. Everything plays by itself in Slide Show (F5).

## How to use

1. Open `index.html` in **Chrome or Microsoft Edge**.
2. Pick page styles in the gallery. Every style you click becomes a page.
3. Select a page on the left and fill in its text and pictures on the right (**Matn / Content** tab).
4. Open **Dizayn / Design** to choose a colour theme, fonts, text size, text animation and background image for the whole deck. The **Ko'rinish / Look** tab changes these for one page only.
5. Press **YARATISH · PPTX / CREATE · PPTX**. The file downloads to your Downloads folder.

**Save** keeps a `.imx.json` project you can **Open** later to keep editing.

## Page styles

| # | Style | Good for |
|---|---|---|
| 9 | Split Screen | text beside a picture |
| 10 | Key Points | 3–5 main ideas as cards |
| 11 | Puzzle | 4–6 parts of one idea |
| 12 | Polaroid Table | photo gallery |
| 13 | Domino Process | step-by-step process |
| 14 | Rocket Stats | key numbers |
| 15 | Balance Scale | before/after, pros/cons |
| 16 | Window Blinds | one big photo with a caption |
| 17 | Spotlight Stage | team, people, experts |
| 18 | Growing Tree | ideas, mind map |
| 19 | Device Mockup | app, website or video demo |
| 20 | Quiz Cards | audience quiz |
| 21 | Glassmorphism | text and numbers on glass panels |
| 22 | Kinetic Type | slogan or short sentence |
| 23 | Split-flap Display | title, date, slogan |

Styles 1–8 (Curtain Title, Reja Hub, Open Book, Gears & Magnifier, Ball & Walls, Lego Builder, Closing Book, Particle Finale) are being rebuilt for PPTX. The old HTML version is kept in `src/legacy/`.

## What the PPTX contains

- **Background animation:** an MP4 video per page, made in the browser (Chrome/Edge). If a browser cannot make MP4, IMX uses an animated GIF instead.
- **Text:** real, editable PowerPoint text with entrance animations (Fade, Rise, Fly in, Wipe, Split, Zoom, Typewriter, Word by word) that start automatically after the background.
- In **Kinetic Type** and **Split-flap Display** the words are part of the animation, so change them in IMX and create the file again.
- Fonts come with Windows and Office and support Uzbek Latin and Cyrillic, so the file looks the same on other computers.
- Videos need **PowerPoint 2016 or newer**.

## For developers

- `src/styles.js`: the page styles (background drawing and text boxes)
- `src/pptx.js`: video encoding and the PowerPoint writer
- `src/editor.html`: the editor interface
- `src/template/blank.pptx`: the empty 16:9 PowerPoint template
- `vendor/`: JSZip, mp4-muxer, gifenc (with their licences)

After changing any of these, rebuild `index.html`:

```
node build.js
```
