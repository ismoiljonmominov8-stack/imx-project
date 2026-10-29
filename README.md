# IMX: animated PowerPoint maker

IMX builds animated presentations in the browser and saves them as **PowerPoint (.pptx)** files. Every page has a moving background, and its text appears with PowerPoint animations. Everything plays by itself in Slide Show (F5).

## How to use

1. Open `index.html` in **Chrome or Microsoft Edge**.
2. Pick page styles in the gallery. Every style you click becomes a page.
3. Select a page on the left and fill in its text and pictures on the right (**Matn / Content** tab).
4. Open **Dizayn / Design** to choose a colour theme, fonts, text size, text animation and background image for the whole deck. The **Ko'rinish / Look** tab changes these for one page only.
5. Press **YARATISH · PPTX / CREATE · PPTX**. The file downloads to your Downloads folder.

**Pictures:** after you upload a picture, sliders appear under it: **Visibility**, **Zoom**, **Left ↔ right** and **Up ↕ down**, so you can choose how much of it shows and which part stays in view. The background picture (Design → background image) has its own visibility slider and now shows on every page style, including the Morph templates. Pictures are kept in the browser (IndexedDB), so a big project survives a page reload.

**Videos:** any picture place also takes a video (MP4 is best; WebM and MOV are converted to MP4 when the file is created). It plays by itself, muted and looping, in the preview and in Slide Show. On the Morph templates (24–33) it is a real PowerPoint video; on styles 9–23 the video is placed as a real PowerPoint video exactly where the picture ends up, after the page animation. Videos are saved inside the `.imx.json` project.

**Text size:** Design → *Heading size (pt)* and *Text size (pt)* take any number from 6 to 200, like Word. *If text does not fit* chooses between shrinking it to fit its box or keeping the exact size. Each page can have its own sizes in the *Look* tab.

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

### Morph templates (24–33)

These pages are built from real PowerPoint shapes, not video. One page becomes several slides, and PowerPoint's **Morph** transition moves the shapes between them on each click (PowerPoint 2019 / 365; older versions show a fade). Every word, shape and picture stays editable. The number of items, steps, parts and cards can be changed in IMX.

| # | Template | Good for |
|---|---|---|
| 24 | Glass Title Card | company or topic name |
| 25 | Accordion List | 2–6 items that open one at a time |
| 26 | Section Divider | big picture and section name |
| 27 | Contact / Closing Card | final page with contacts |
| 28 | Steps Hub | 3–8 steps as colour tiles |
| 29 | Step Photo Card | one slide per step with photo and tip |
| 30 | Wheel | 3–8 part wheel, one detail slide per part |
| 31 | Flip Cards | 2–6 numbered cards opening one by one |
| 32 | Photo-in-Text Title | title letters filled with your photo |
| 33 | Wave Photo | your photo, a wave and two text columns |

Styles 1–8 (Curtain Title, Reja Hub, Open Book, Gears & Magnifier, Ball & Walls, Lego Builder, Closing Book, Particle Finale) are being rebuilt for PPTX. The old HTML version is kept in `src/legacy/`.

## What the PPTX contains

- **Background animation:** an MP4 video per page, made in the browser (Chrome/Edge). If a browser cannot make MP4, IMX uses an animated GIF instead.
- **Text:** real, editable PowerPoint text with entrance animations (Fade, Rise, Fly in, Wipe, Split, Zoom, Typewriter, Word by word) that start automatically after the background.
- In **Kinetic Type** and **Split-flap Display** the words are part of the animation, so change them in IMX and create the file again.
- Fonts come with Windows and Office and support Uzbek Latin and Cyrillic, so the file looks the same on other computers.
- Videos need **PowerPoint 2016 or newer**.

## For developers

- `src/styles.js`: the page styles (background drawing and text boxes)
- `src/morph.js`: the Morph templates 24–33 and their preview
- `src/pptx.js`: video encoding and the PowerPoint writer
- `src/editor.html`: the editor interface
- `src/template/blank.pptx`: the empty 16:9 PowerPoint template
- `vendor/`: JSZip, mp4-muxer, gifenc (with their licences)

After changing any of these, rebuild `index.html`:

```
node build.js
```
