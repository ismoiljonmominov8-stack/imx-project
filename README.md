# IMX — interactive presentation maker

IMX builds animated, interactive presentations in the browser. You fill in a sidebar, watch a live preview, and press **CREATE** to download one `.html` file. Anyone can open that file in a browser and present it: no installation, no internet needed (except for the fonts).

## How to use

1. Download this repository (**Code → Download ZIP**) and unzip it.
2. Open `index.html` in Chrome or Microsoft Edge.
3. Fill in the sidebar. The topic and your full name are required; the **CREATE** button stays locked until both are filled.
4. Press **CREATE**. Your presentation is saved to your Downloads folder as `<topic>.html`.
5. Send that file to anyone. They double-click it, and it opens full-screen in their browser.

Use **Save project** to keep a `.imx.json` file you can reopen later with **Open project** and keep editing. Text and pictures are also kept automatically in the browser between visits; videos are kept only in saved projects.

## What the sidebar contains

| Section | What it does |
|---|---|
| Presentation | Topic (required) and the presentation language (Uzbek, Russian, English) |
| Presenter | Full name (required), position, photo |
| Supervisor or partner | Optional name, academic title, photo |
| University | Optional name and logo |
| Menu (pages) | Turn pages on or off, rename them, reorder them, and add each page's text, pictures, chart and video |
| Final page | Thank-you word and the line underneath |
| Voice and animation | Automatic animation, voice guide that reads the slides, voice control (next, back, play, stop) |
| Settings | Heading colour, background colour, background image and its visibility |

## Pages and their animations

1. **Title** — topic and university drop to the centre, presenter and supervisor panels sweep in.
2. **Outline (Reja)** — a dot draws a circle, the word is written inside, page rows arrive from both sides.
3. **Introduction** — a book opens, the title rises; text, picture and a live chart (line, bar, key figures).
4. **Problems** — gears jam and smoke, a magnifier zooms in, the page splits open; slideshow of up to 10 pictures.
5. **Plans** — a ball races and climbs walls; four plan panels slam together.
6. **Solutions** — a stickman completes a Lego wall that spells the title; rotating 3D cards (add as many as you like).
7. **Conclusion** — a book closes, the title lifts off the cover; same layout as Introduction.
8. **Thank you & questions** — particles form the thank-you word; buttons jump back to any page during Q&A.

Any page can start with your own video instead of its animation.

## Presenting

- **→ / Space / Enter**: next page · **←**: back · **F**: full screen
- The control bar appears when you move the mouse.
- **Voice guide** reads each page aloud. The most natural voices are in Microsoft Edge.
- **Voice control** listens for "next / keyingi / далее", "back / orqaga / назад", "play / ijro / играй" and "stop / to'xta / стоп". It works in Chrome and Edge and asks for microphone permission.

## For developers

The source lives in `src/`:

- `src/player.html` — the presentation player. It reads its content from the `__IMX_DATA__` JSON slot.
- `src/editor.html` — the editor UI. It embeds the player at the `__PLAYER__` slot.

After changing either file, rebuild `index.html`:

```
node build.js
```
