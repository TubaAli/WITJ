# LinkedIn日本代表から学ぶAIを活用した交渉術｜WITJ

Slide deck for the Women in Technology Japan (WITJ) × LinkedIn event, published as a single-page site.

- Live: https://tubaali.github.io/WITJ/
- Navigate with ← → keys, the on-screen buttons, or swipe on a phone. On a phone in portrait, tap「横向きで表示」to view in landscape even with rotation lock on.
- Add `?all` to the URL to show every build step at once with no animation (e.g. for printing).

## PowerPoint export

The **PPT** button at the bottom of the screen downloads the deck as a `.pptx` built from native PowerPoint elements, so slides can be copied straight into Google Slides or PowerPoint and edited there:

- text → text boxes with editable runs (font, size, weight, colour, line breaks preserved)
- cards, chips, pills, bullets, dividers → shapes with fills, borders, rounded corners and shadows
- photos and logos → pictures (cropped and rounded natively)
- icons and illustrations (inline SVG) → transparent PNG pictures; text inside them stays editable text

The export runs entirely in the browser (`pptx-export.js`, JSZip from cdnjs); nothing is uploaded. The deck uses the Noto Sans JP font, which Google Slides provides and PowerPoint substitutes if it is not installed.
