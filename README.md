# BUILD MODE — Strategic Brand & Business Strategy Deck

A 12-slide, 16:9 strategic planning presentation for BUILD MODE, generated with
[PptxGenJS](https://gitbrent.github.io/PptxGenJS/). Every slide element is a native,
editable PowerPoint object (text, shapes, lines) — nothing is flattened into an image.

| File | Purpose |
|---|---|
| `build_mode_strategy.pptx` | The finished deck |
| `src/build.js` | Deck source: theme, layouts and all 12 slides |
| `src/apply-theme.js` | Writes the brand palette into the PowerPoint theme |
| `assets/fonts/` | Inter Regular and Bold (SIL Open Font License) |
| `preview/` | JPEG render of each slide, for quick review |
| `scripts/render.sh` | Re-renders `preview/` from the deck |

## Regenerate the deck

Requires Node.js 18 or later.

```bash
npm install
```

```bash
npm run build
```

This overwrites `build_mode_strategy.pptx` in the project root.

## Fonts

The deck is set in **Inter**. Install `assets/fonts/Inter-Regular.otf` and
`assets/fonts/Inter-Bold.otf` on any machine that opens or presents the deck;
otherwise PowerPoint substitutes a default font and line breaks may shift.

To use a different typeface (for example Aptos or Arial), change the `FONT`
constant at the top of `src/build.js` and rebuild. Font and colours are theme-level,
so they can also be changed inside PowerPoint via Design > Variants.

## Editing

- **Palette:** the `THEME.colors` object in `src/build.js` (navy `101827`, dark blue
  `18263A`, white, soft grey `F2F4F7`, blue-grey `91A0B6`, accent `477BFF`).
- **Content:** each slide is a clearly labelled block in `src/build.js`; text lives in
  plain arrays at the top of each block.
- **Layouts:** four slide layouts (Cover, Light, Dark, Statement) hold the title and
  eyebrow placeholders, footer and slide number. The cover has no slide number.

## Re-render the previews (optional)

Requires LibreOffice (`soffice`) and Poppler (`pdftoppm`) on the PATH.

```bash
npm run render
```

## Evidence status

This is a planning deck. It contains no market statistics, competitor figures or
customer research findings. Audience tiers (slide 4), segment ambitions (slide 3),
differentiation (slide 9) and all five validation gates (slide 11) are labelled on
the slides as hypotheses still to be validated.
