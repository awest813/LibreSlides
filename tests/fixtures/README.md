# Test fixtures

## mtx-test.fntdata

An embedded font part as PowerPoint stores it: Embedded OpenType (version
0x00020002, flags `TTEMBED_TTCOMPRESSED`) holding MicroType Express (MTX)
compressed font data. Used by `tests/unit/mtx.test.js`.

The font inside, "MTX Test", was drawn for these tests (rectangles,
triangles and a few curves; no third-party font data), so the file is part of
LibreSlides under its MIT license. It was made with `make-mtx-fixture.py`
(needs Python and fontTools), which builds the font, turns its glyphs into
MTX's Compact Table Format, LZCOMP-compresses the three blocks and wraps the
result in EOT. It covers what a decoder must handle: off-curve points, every
coordinate triplet size, a glyph whose bounding box isn't its points', a
600-point contour, composite glyphs with and without instructions, push data
with words, negative numbers and hops, a cvt table, and LZCOMP literals,
near and far copies, the one-byte repeats and run-length encoding.

A self-made font was chosen over a real PowerPoint sample because the real
samples available carry commercial or third-party fonts. The file was
checked against libeot's `EOT2ttf_buffer` (an independent MTX decoder, used
as a black box): it unpacks it into the same font as the original – every
table but glyf and loca byte for byte, and the same glyph outlines and
instructions – except for the composite glyph's instructions, which libeot
gets wrong (it does so for PowerPoint's own files too). LibreSlides' decoder
gives back the original's tables byte for byte, bar the head table's
checkSumAdjustment.

## pictures/

Pictures in the formats LibreSlides converts when it opens a PowerPoint file
(see `src/io/pictures.js`), used by `tests/unit/pictures.test.js`. All were
made for these tests by `make-picture-fixtures.py` (needs Pillow, ImageMagick
and LibreOffice), so they're part of LibreSlides under its MIT license:

- `*.tif`: one 40 × 30 test picture (coloured quarters and a gradient that's
  half see-through at the right) as TIFF in the ways the decoder handles:
  LZW and Deflate with the horizontal predictor, PackBits (big-endian),
  RGBA, 16-bit samples, bilevel, a 4-bit palette, tiles and planar
  samples. `tiff-pixels.json` has some of each one's pixels as Pillow reads
  them.
- `drawing.emf`, `drawing.wmf`: an SVG drawing (a rectangle, an ellipse, a
  curve, a triangle, the word "Metafile" and a small bitmap) saved as EMF and
  WMF by LibreOffice. The WMF also carries an EMF copy in its comments, as
  Office's do; the tests read it both ways.
