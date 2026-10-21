"""Make the picture fixtures in tests/fixtures/pictures (see README.md).

Needs Python with Pillow, ImageMagick (`convert`) and LibreOffice (`soffice`).
Run from the repository root: python3 tests/fixtures/make-picture-fixtures.py
"""
import json, os, subprocess, tempfile
from PIL import Image

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pictures')
os.makedirs(OUT, exist_ok=True)
W, H = 40, 30


def color(x, y):
    """The test picture: red, blue and green quarters, and a gradient (half see-through at the right)."""
    if x < 20 and y < 15:
        return (255, 0, 0, 255)
    if y < 15:
        return (0, 0, 255, 255)
    if x < 20:
        return (0, 160, 0, 255)
    return (x * 6, y * 8, 128, 128 if x > 30 else 255)


tmp = tempfile.mkdtemp()
im = Image.new('RGBA', (W, H))
im.putdata([color(x, y) for y in range(H) for x in range(W)])
im.save(f'{tmp}/src.png')
im.convert('RGB').save(f'{tmp}/src-rgb.png')

TIFFS = {
    'rgb-lzw-pred': ['src-rgb.png', '-compress', 'LZW', '-define', 'tiff:predictor=2'],
    'rgb-zip-pred': ['src-rgb.png', '-compress', 'Zip', '-define', 'tiff:predictor=2'],
    'rgb-packbits-msb': ['src-rgb.png', '-compress', 'RLE', '-endian', 'MSB'],
    'rgba-lzw': ['src.png', '-compress', 'LZW'],
    'rgb16-lzw-pred': ['src-rgb.png', '-depth', '16', '-compress', 'LZW', '-define', 'tiff:predictor=2'],
    'gray1-packbits': ['src-rgb.png', '-colorspace', 'Gray', '-threshold', '50%', '-type', 'Bilevel', '-compress', 'RLE'],
    'palette4': ['src-rgb.png', '-colors', '8', '-type', 'Palette', '-depth', '4', '-compress', 'None'],
    'tiled': ['src-rgb.png', '-compress', 'LZW', '-define', 'tiff:tile-geometry=16x16'],
    'planar': ['src-rgb.png', '-compress', 'LZW', '-interlace', 'Plane'],
}
# Some pixels of each, as Pillow reads them (RGBA), to check the decoder against.
SAMPLES = [(0, 0), (19, 14), (20, 0), (39, 14), (0, 15), (19, 29), (25, 20), (35, 25), (39, 29), (10, 22), (30, 5), (21, 16)]
expected = {}
for name, (src, *args) in TIFFS.items():
    path = os.path.join(OUT, f'{name}.tif')
    subprocess.run(['convert', f'{tmp}/{src}', *args, path], check=True)
    ref = Image.open(path).convert('RGBA')
    expected[name] = [[x, y, *ref.getpixel((x, y))] for x, y in SAMPLES]
json.dump(expected, open(os.path.join(OUT, 'tiff-pixels.json'), 'w'), indent=0)

# A drawing, made into EMF and WMF by LibreOffice: shapes, a curve, text and a bitmap.
png = Image.new('RGB', (8, 8))
png.putdata([(255, 0, 0) if (x + y) % 2 else (0, 0, 255) for y in range(8) for x in range(8)])
png.save(f'{tmp}/tile.png')
import base64
tile = base64.b64encode(open(f'{tmp}/tile.png', 'rb').read()).decode()
svg = f'''<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="400" height="300" viewBox="0 0 400 300">
<rect x="20" y="20" width="160" height="100" fill="#3366cc" stroke="#000000" stroke-width="4"/>
<ellipse cx="290" cy="70" rx="80" ry="50" fill="#ff9900"/>
<path d="M20,200 C80,140 140,260 200,200" fill="none" stroke="#009933" stroke-width="6"/>
<polygon points="250,160 380,160 315,280" fill="#cc0033"/>
<text x="30" y="270" font-family="Liberation Sans" font-size="28" fill="#222222">Metafile</text>
<image x="220" y="190" width="40" height="40" xlink:href="data:image/png;base64,{tile}"/>
</svg>'''
open(f'{tmp}/drawing.svg', 'w').write(svg)
for ext in ('emf', 'wmf'):
    subprocess.run(['soffice', f'-env:UserInstallation=file://{tmp}/lo', '--headless', '--convert-to', ext, '--outdir', OUT, f'{tmp}/drawing.svg'], check=True, capture_output=True)
print('made', sorted(os.listdir(OUT)))
