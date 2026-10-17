# Agen document preview fork

Upstream: https://github.com/VolodymyrBaydalka/docxjs (Apache-2.0).
Fork: https://github.com/simon100500/docxjs, branch `codex/word-layout-fidelity`.

## Exact document output

`renderAsync(data, body, styles, {renderMode:'pdf', pdfAdapter})` accepts an
application-owned authoritative converter and PDF page rasterizer. The adapter
implements `convert(data): Promise<Uint8Array>` and
`renderPages(pdf): Promise<Array<{width,height,imageUrl}>>`. Dimensions are CSS
pixels; image URLs are data images. The result contains `pdfBytes` and `pageCount`.
PDF page boundaries and all document typography come from the converter. Errors
are propagated; there is no silent fallback to approximate HTML. No converter,
worker, network service or PDF package is bundled into this library.

The optional adapter API is retained for consumers that supply a converter.
Agen currently uses HTML for both filled acts and template field editing;
it does not install or invoke a Microsoft Word converter or PDF.js rasterizer.
HTML is not claimed to reproduce every Word layout feature.

## HTML fixes

- Resolve tab positions from the text-column origin, including positive and
  negative paragraph indents. Indented attachment-list tabs stay on the item
  line instead of overflowing and wrapping onto a separate underline-only line.
- Decorate the entire tab advance with native font underline metrics, including
  empty fields and separate tab runs. An invisible generated glyph with measured
  letter spacing carries the same text decoration as the run; no border or
  manually positioned line is used. Selectable text retains the original tab.
  Covered at 50%, 100% and 125% CSS zoom.
- Read paragraph-mark font size/family to establish the correct line strut.
- Zero the font/line height of empty bookmark anchors.
- Measure tab stops after fonts are ready, before `renderAsync` resolves.
- Handle empty trailing right/center tab ranges without using the viewport origin.
- Correct tab coordinates for CSS zoom and transforms; use explicit inline width.
- Connect detached containers temporarily for layout, removing the old 500ms timer.

Build: `npm run build-prod`. Browser regressions and upstream fixtures:
`npm run e2e -- --browsers ChromeHeadless` (`CHROME_BIN` may be required).
Built distribution files are committed in this fork so applications can pin a
Git commit without build/install hooks.
