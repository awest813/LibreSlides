"""
Make tests/fixtures/mtx-test.fntdata: a small TrueType font drawn for the
tests ("MTX Test"), MicroType Express compressed and wrapped in EOT the way
PowerPoint stores embedded fonts (EOT version 0x00020002, flags
TTEMBED_TTCOMPRESSED).

Usage: python3 tests/fixtures/make-mtx-fixture.py tests/fixtures/mtx-test.fntdata [ORIGINAL.ttf]
Needs fontTools. The font only has simple shapes; it exists to exercise the
decoder: off-curve points, every kind of coordinate triplet, a glyph whose
bounding box isn't its points', a contour with more than 506 points,
composite glyphs (with and without instructions), push data with words,
negative numbers and hops, a cvt table, and LZCOMP literals, copies (near and
far), the one-byte repeats and run-length encoding.
"""
import array, io, struct, sys
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables import ttProgram
from fontTools.ttLib.tables._g_l_y_f import Glyph, GlyphComponent

# ---------------------------------------------------------------- the font

def program(data):
    p = ttProgram.Program(); p.fromBytecode(bytes(data)); return p

def push_bytes(values):
    """Canonical push instructions (as the decoder writes them)."""
    out = bytearray(); i = 0
    while i < len(values):
        word = not (0 <= values[i] <= 255)
        j = i
        while j < len(values) and j - i < 255 and (not (0 <= values[j] <= 255)) == word: j += 1
        n = j - i
        out += bytes([(0xb8 if word else 0xb0) + n - 1]) if n < 8 else bytes([0x41 if word else 0x40, n])
        for v in values[i:j]: out += struct.pack('>h', v) if word else bytes([v])
        i = j
    return bytes(out)

def make_font():
    names = ['.notdef', 'space', 'A', 'B', 'C', 'D', 'acute', 'Aacute', 'Adieresis', 'E']
    fb = FontBuilder(1000, isTTF=True)
    fb.setupGlyphOrder(names)
    fb.setupCharacterMap({0x20: 'space', 0x41: 'A', 0x42: 'B', 0x43: 'C', 0x44: 'D', 0x45: 'E', 0xB4: 'acute', 0xC1: 'Aacute', 0xC4: 'Adieresis'})
    glyphs = {}

    pen = TTGlyphPen(None)
    for x0, y0, x1, y1 in ((50, 0, 450, 700), (100, 50, 400, 650)):
        pen.moveTo((x0, y0)); pen.lineTo((x0, y1)); pen.lineTo((x1, y1)); pen.lineTo((x1, y0)); pen.closePath()
    g = pen.glyph()
    g.program = program(push_bytes([5, 6, 2, 1, 4, 7, 3, 0, 5, 4, 2, 3, 6, 7, 1, 0]) + bytes([0x2f, 0x3c, 0xdc, 0x3c] * 4))
    glyphs['.notdef'] = g

    glyphs['space'] = Glyph()

    # A: curves (off-curve points) and moves of every size: tiny, 4-bit, byte, 12-bit and 16-bit.
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0)); pen.lineTo((300, 700)); pen.qCurveTo((310, 712), (330, 700)); pen.lineTo((620, 0))
    pen.lineTo((500, 0)); pen.qCurveTo((480, 60), (420, 230), (310, 240)); pen.lineTo((300, -1)); pen.lineTo((299, 3))
    pen.lineTo((1299, 3)); pen.lineTo((1299, 5003)); pen.lineTo((-4000, 5003)); pen.lineTo((-4000, -600)); pen.closePath()
    pen.moveTo((200, 300)); pen.lineTo((250, 450)); pen.lineTo((350, 450)); pen.lineTo((400, 300)); pen.closePath()
    g = pen.glyph()
    g.program = program(push_bytes([1, 300, -2, 7]) + bytes([0x2b, 0x00, 0x01]))
    glyphs['A'] = g

    # B: a bounding box bigger than its points (needs the explicit-bbox record).
    pen = TTGlyphPen(None)
    pen.moveTo((100, 0)); pen.lineTo((100, 700)); pen.lineTo((500, 350)); pen.closePath()
    g = pen.glyph(); g.recalcBounds(None)
    g.xMin, g.yMin, g.xMax, g.yMax = 50, -20, 560, 760
    g.program = program(b'')
    glyphs['B'] = g

    # C: one contour of 600 points (255UShort's longest forms), with long runs of the same flag.
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0))
    for i in range(1, 300): pen.lineTo((i * 2, (i % 2) * 3))
    for i in range(300): pen.lineTo((600 - i * 2, 400))
    pen.closePath()
    g = pen.glyph()
    g.program = program(b'')
    glyphs['C'] = g

    # D: instructions whose pushes have words, negatives, the 255Short's longer forms and hop patterns.
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0)); pen.lineTo((0, 500)); pen.lineTo((400, 500)); pen.closePath()
    g = pen.glyph()
    pushes = [9, 1, 9, 2, 9, 3, 9, 251, 252, 253, 260, 499, 510, 755, 4, 0, 4, 8, 4, 0, 4,
              2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]
    g.program = program(push_bytes(pushes) + push_bytes([-1, -250, -32768, 32767, 1000]) + bytes([0x4b, 0x2b, 0x1c]))
    glyphs['D'] = g

    pen = TTGlyphPen(None)
    pen.moveTo((200, 750)); pen.lineTo((300, 900)); pen.lineTo((360, 880)); pen.lineTo((240, 740)); pen.closePath()
    g = pen.glyph(); g.program = program(push_bytes([3]) + bytes([0x2b]))
    glyphs['acute'] = g

    # Composites: one with instructions (pushes of a byte, a word and a byte, like ttfautohint's), one scaled.
    g = Glyph(); g.numberOfContours = -1
    c1 = GlyphComponent(); c1.glyphName = 'A'; c1.x, c1.y = 0, 0; c1.flags = 0x204
    c2 = GlyphComponent(); c2.glyphName = 'acute'; c2.x, c2.y = 60, -20; c2.flags = 0x4
    g.components = [c1, c2]
    g.program = program(push_bytes([1, 1]) + push_bytes([-20]) + push_bytes([53]) + bytes([0x2b]))
    glyphs['Aacute'] = g
    g = Glyph(); g.numberOfContours = -1
    c1 = GlyphComponent(); c1.glyphName = 'A'; c1.x, c1.y = 0, 0; c1.flags = 0x204
    c2 = GlyphComponent(); c2.glyphName = 'acute'; c2.x, c2.y = 400, 300; c2.flags = 0x4; c2.transform = [[0.5, 0], [0, 0.75]]
    g.components = [c1, c2]
    glyphs['Adieresis'] = g

    # E: a plain glyph repeated from A's second contour (so LZCOMP finds a copy).
    pen = TTGlyphPen(None)
    pen.moveTo((200, 300)); pen.lineTo((250, 450)); pen.lineTo((350, 450)); pen.lineTo((400, 300)); pen.closePath()
    g = pen.glyph(); g.program = program(b'')
    glyphs['E'] = g

    fb.setupGlyf(glyphs)
    glyf = fb.font['glyf']
    for n in names:
        if n != 'B': glyf[n].recalcBounds(glyf)
    fb.setupHorizontalMetrics({n: (600, getattr(glyf[n], 'xMin', 0)) for n in names})
    fb.setupHorizontalHeader(ascent=800, descent=-200)
    fb.setupNameTable({'familyName': 'MTX Test', 'styleName': 'Regular', 'version': 'Version 1.000', 'fullName': 'MTX Test Regular', 'psName': 'MTXTest-Regular', 'uniqueFontIdentifier': 'MTX Test Regular'})
    fb.setupOS2(sTypoAscender=800, usWinAscent=900, usWinDescent=250)
    fb.setupPost()
    fb.setupMaxp()
    fb.font['maxp'].maxSizeOfInstructions = max(len(glyphs[n].program.getBytecode()) if hasattr(glyphs[n], 'program') else 0 for n in names)
    cvt = [0, 0, 24, 24, 237, 238, 1200, 0, 870, -870, -290, 12000, -12000, 5, 5, 5, 300, 800]
    fb.font['cvt '] = newTable('cvt ')
    fb.font['cvt '].values = array.array('h', cvt)
    for tag, code in (('fpgm', bytes([0xb0, 0x00, 0x2c, 0x01, 0x2d])), ('prep', bytes([0xb1, 0x01, 0x02, 0x1d, 0x1e]))):
        t = newTable(tag)
        t.program = program(code)
        fb.font[tag] = t
    buf = io.BytesIO(); fb.save(buf)
    # Then give B its bigger bounding box, which saving would otherwise work out afresh.
    font = TTFont(io.BytesIO(buf.getvalue()), recalcBBoxes=False, recalcTimestamp=False)
    font['head'].created = font['head'].modified = 3870000000  # a fixed date, so the file comes out the same each time
    b = font['glyf']['B']; b.expand(font['glyf'])
    b.xMin, b.yMin, b.xMax, b.yMax = 50, -20, 560, 760
    buf = io.BytesIO(); font.save(buf)
    return buf.getvalue()

# ---------------------------------------------------------------- CTF

def u255(v):
    if v < 253: return bytes([v])
    if v < 506: return bytes([255, v - 253])
    if v < 759: return bytes([254, v - 506])
    return bytes([253]) + struct.pack('>H', v)

def s255(v):
    if -756 < v < 0: return bytes([250]) + s255(-v)
    if v < 0: return bytes([253]) + struct.pack('>h', v)
    if v < 250: return bytes([v])
    if v < 500: return bytes([255, v - 250])
    if v < 756: return bytes([254, v - 500])
    return bytes([253]) + struct.pack('>h', v)

def encode_pushes(values):
    """Push data with HOP3/HOP4 where the pattern allows."""
    out = bytearray(); i = 0
    while i < len(values):
        if i >= 2 and i + 4 < len(values) and values[i] == values[i - 2] and values[i + 2] == values[i - 2] and values[i + 4] == values[i - 2]:
            out += bytes([252]) + s255(values[i + 1]) + s255(values[i + 3]); i += 5
        elif i >= 2 and i + 2 < len(values) and values[i] == values[i - 2] and values[i + 2] == values[i - 2]:
            out += bytes([251]) + s255(values[i + 1]); i += 3
        else:
            out += s255(values[i]); i += 1
    return bytes(out)

def split_pushes(bc):
    vals = []; p = 0
    while p < len(bc):
        op = bc[p]
        if op == 0x40: n = bc[p + 1]; vals += list(bc[p + 2:p + 2 + n]); p += 2 + n
        elif op == 0x41: n = bc[p + 1]; vals += [struct.unpack('>h', bc[p + 2 + 2 * i:p + 4 + 2 * i])[0] for i in range(n)]; p += 2 + 2 * n
        elif 0xb0 <= op <= 0xb7: n = op - 0xaf; vals += list(bc[p + 1:p + 1 + n]); p += 1 + n
        elif 0xb8 <= op <= 0xbf: n = op - 0xb7; vals += [struct.unpack('>h', bc[p + 1 + 2 * i:p + 3 + 2 * i])[0] for i in range(n)]; p += 1 + 2 * n
        else: break
    return vals, bc[p:]

def triplet(dx, dy, on):
    """The smallest triplet form for a move."""
    ax, ay = abs(dx), abs(dy)
    sx, sy = dx >= 0, dy >= 0
    if dx == 0 and ay < 1280:
        f = (ay >> 8) * 2 + (1 if sy else 0); data = bytes([ay & 255])
    elif dy == 0 and ax < 1280:
        f = 10 + (ax >> 8) * 2 + (1 if sx else 0); data = bytes([ax & 255])
    elif 1 <= ax <= 64 and 1 <= ay <= 64:
        bx, by = (ax - 1) >> 4, (ay - 1) >> 4
        f = 20 + bx * 16 + by * 4 + (1 if sx else 0) + (2 if sy else 0)
        data = bytes([(((ax - 1) & 15) << 4) | ((ay - 1) & 15)])
    elif 1 <= ax <= 768 and 1 <= ay <= 768:
        qx, qy = (ax - 1) >> 8, (ay - 1) >> 8
        f = 84 + (qx * 3 + qy) * 4 + (1 if sx else 0) + (2 if sy else 0)
        data = bytes([(ax - 1) & 255, (ay - 1) & 255])
    elif ax < 4096 and ay < 4096:
        f = 120 + (1 if sx else 0) + (2 if sy else 0)
        data = bytes([ax >> 4, ((ax & 15) << 4) | (ay >> 8), ay & 255])
    else:
        f = 124 + (1 if sx else 0) + (2 if sy else 0)
        data = struct.pack('>HH', ax, ay)
    return (f | (0 if on else 0x80)), data

def ctf_glyf(font):
    glyf = font['glyf']; out = bytearray(); push = bytearray(); code = bytearray()
    for name in font.getGlyphOrder():
        g = glyf[name]
        def instr(g):
            bc = g.program.getBytecode() if hasattr(g, 'program') else b''
            vals, rest = split_pushes(bc)
            push.extend(encode_pushes(vals)); code.extend(rest)
            return u255(len(vals)) + u255(len(rest))
        if g.numberOfContours == 0:
            out += struct.pack('>h', 0)
        elif g.isComposite():
            # As TrueType has it, but with the instructions stored like a simple glyph's.
            body = bytearray(struct.pack('>hhhhh', -1, g.xMin, g.yMin, g.xMax, g.yMax))
            instructed = hasattr(g, 'program') and len(g.program.getBytecode()) > 0
            for c in g.components:
                last = c is g.components[-1]
                body += c.compile(not last, instructed and last, glyf)
            out += body
            if hasattr(g, 'program') and len(g.program.getBytecode()): out += instr(g)
        else:
            coords = list(g.coordinates); flags = list(g.flags)
            xs = [c[0] for c in coords]; ys = [c[1] for c in coords]
            if (min(xs), min(ys), max(xs), max(ys)) != (g.xMin, g.yMin, g.xMax, g.yMax):
                out += struct.pack('>hhhhhh', 0x7fff, g.numberOfContours, g.xMin, g.yMin, g.xMax, g.yMax)
            else:
                out += struct.pack('>h', g.numberOfContours)
            prev = -1
            for i, e in enumerate(g.endPtsOfContours):
                out += u255(e if i == 0 else e - prev); prev = e
            fl = bytearray(); data = bytearray(); px = py = 0
            for (x, y), f in zip(coords, flags):
                t, d = triplet(x - px, y - py, f & 1); px, py = x, y
                fl.append(t); data += d
            out += fl + data + instr(g)
    return bytes(out), bytes(push), bytes(code)

def ctf_cvt(values):
    out = bytearray(struct.pack('>H', len(values))); prev = 0
    for v in values:
        d = v - prev; prev = v
        if 0 <= d < 238: out.append(d)
        elif -2141 <= d < 0 and (-d) // 238 <= 8:
            out += bytes([239 + (-d) // 238, (-d) % 238])
        elif 238 <= d < 238 * 9 and d // 238 <= 8 and d % 238 < 256:
            out += bytes([247 + d // 238, d % 238])
        else: out += bytes([238]) + struct.pack('>h', d)
    return bytes(out)

def sfnt(tables):
    tags = sorted(tables); n = len(tags); off = 12 + 16 * n
    hdr = struct.pack('>IHHHH', 0x10000, n, 128, 3, n * 16 - 128)
    dirs = b''; body = b''
    for t in tags:
        d = tables[t]
        dirs += struct.pack('>4sIII', t.encode('latin1'), 0, (off + len(body)) if d else 0, len(d))
        body += d + b'\0' * (-len(d) % 4)
    return hdr + dirs + body

def ctf_font(ttf):
    font = TTFont(io.BytesIO(ttf))
    raw = lambda tag: (font.reader.file.seek(font.reader.tables[tag].offset), font.reader.file.read(font.reader.tables[tag].length))[1]
    tables = {t: raw(t) for t in font.reader.keys()}
    glyf, push, code = ctf_glyf(font)
    tables['glyf'] = glyf; tables['loca'] = b''
    tables['cvt '] = ctf_cvt(list(font['cvt '].values))
    return sfnt(tables), push, code

# ---------------------------------------------------------------- LZCOMP

class Huffman:
    """The same adaptive Huffman code the decoder keeps (see src/io/mtx.js)."""
    def __init__(self, n, w0):
        self.n = n; N = 2 * n
        self.up = [i // 2 for i in range(N)]; self.left = [0] * N; self.right = [0] * N
        self.sym = [-1] * N; self.w = [0] * N; self.leaf = [0] * n
        for i in range(1, n): self.left[i] = 2 * i; self.right[i] = 2 * i + 1
        for s in range(n): self.sym[n + s] = s; self.w[n + s] = w0; self.leaf[s] = n + s
        for i in range(n - 1, 0, -1): self.w[i] = self.w[2 * i] + self.w[2 * i + 1]
    def swap(self, a, b):
        for arr in (self.left, self.right, self.sym, self.w): arr[a], arr[b] = arr[b], arr[a]
        for x in (a, b):
            if self.sym[x] >= 0: self.leaf[self.sym[x]] = x
            else: self.up[self.left[x]] = self.up[self.right[x]] = x
    def update(self, s):
        a = self.leaf[s]
        while a != 1:
            f = a
            while self.w[f - 1] == self.w[a]: f -= 1
            if f != a: self.swap(a, f); a = f
            self.w[a] += 1; a = self.up[a]
        self.w[1] += 1
    def write(self, s, bits):
        a = self.leaf[s]; code = []
        while a != 1:
            p = self.up[a]; code.append(1 if self.right[p] == a else 0); a = p
        bits.extend(reversed(code)); self.update(s)

def preload():
    p = bytearray()
    for a in range(32):
        for b in range(96): p += bytes([a, b])
    for a in range(256): p += bytes([a]) * 4
    return bytes(p)

def run_length(data):
    counts = [0] * 256
    for b in data: counts[b] += 1
    esc = counts.index(min(counts))
    out = bytearray([esc]); i = 0
    while i < len(data):
        j = i
        while j < len(data) and data[j] == data[i] and j - i < 255: j += 1
        if j - i >= 4: out += bytes([esc, j - i, data[i]]); i = j
        elif data[i] == esc: out += bytes([esc, 0]); i += 1
        else: out.append(data[i]); i += 1
    return bytes(out)

def lzcomp(data, rle=False):
    if rle: data = run_length(data)
    L = len(data)
    pieces = 1
    while 8 ** pieces < L: pieces += 1
    dups = 256 + 8 * pieces
    sym = Huffman(dups + 3, 1)
    for s in [256, 257] + [dups] * 12 + [dups + 1] * 6: sym.update(s)
    lens, dists = Huffman(8, 3), Huffman(8, 3)
    bits = [1 if rle else 0] + [(L >> (23 - i)) & 1 for i in range(24)]
    buf = preload() + data; base = len(preload())
    index = {}
    for i in range(len(buf) - 2):
        if i < base: index.setdefault(buf[i:i + 3], []).append(i)
    pos = base
    while pos < len(buf):
        best = (0, 0)
        for s in reversed(index.get(buf[pos:pos + 3], [])[-64:]):
            n = 0
            while pos + n < len(buf) and s + n < pos and buf[s + n] == buf[pos + n] and n < 300: n += 1
            # the copy must end before pos; distance counts back from its last byte
            dist = pos - s - n + 1
            if n >= 2 and dist >= 1 and (dist < 512 or n >= 3) and dist <= 8 ** pieces and n > best[0]: best = (n, s)
        if best[0] >= 2:
            n, s = best; dist = pos - s - n + 1
            ncode = n - 1 if dist >= 512 else n
            v = ncode - 2; digits = []
            while True:
                digits.append(v & 3); v >>= 2
                if not v: break
            digits.reverse(); chunks = [d | 4 for d in digits[:-1]] + [digits[-1]]
            d = dist - 1; r = 1
            while d >> (3 * r): r += 1
            sym.write(256 + 8 * (r - 1) + chunks[0], bits)
            for c in chunks[1:]: lens.write(c, bits)
            for i in range(r - 1, -1, -1): dists.write((d >> (3 * i)) & 7, bits)
            step = n
        elif any(buf[pos] == buf[pos - k] for k in (2, 4, 6)):
            k = next(k for k in (2, 4, 6) if buf[pos] == buf[pos - k])
            sym.write(dups + k // 2 - 1, bits); step = 1
        else:
            sym.write(buf[pos], bits); step = 1
        for i in range(pos, pos + step):
            if i + 3 <= len(buf): index.setdefault(buf[i:i + 3], []).append(i)
        pos += step
    bits += [0] * (-len(bits) % 8)
    return bytes(int(''.join(map(str, bits[i:i + 8])), 2) for i in range(0, len(bits), 8))

def mtx(block1, push, code):
    b1, b2, b3 = lzcomp(block1), lzcomp(push, rle=True), lzcomp(code)
    o2 = 10 + len(b1); o3 = o2 + len(b2)
    return bytes([3]) + len(block1).to_bytes(3, 'big') + o2.to_bytes(3, 'big') + o3.to_bytes(3, 'big') + b1 + b2 + b3

# ---------------------------------------------------------------- EOT

def eot(font_data, ttf):
    f = TTFont(io.BytesIO(ttf)); os2 = f['OS/2']; name = f['name']
    def s(t):
        b = t.encode('utf-16-le') + b'\0\0'
        return b'\0\0' + struct.pack('<H', len(b)) + b
    names = s(name.getDebugName(1)) + s(name.getDebugName(2)) + s(name.getDebugName(5)) + s(name.getDebugName(4))
    tail = b'\0\0' + struct.pack('<H', 0) + struct.pack('<I', 0x50475342) + struct.pack('<I', 1252) + b'\0\0' + struct.pack('<H', 0) + struct.pack('<I', 0) + struct.pack('<I', 0)
    fixed = struct.pack('<IIII', 0, len(font_data), 0x00020002, 0x4)
    fixed += bytes(os2.panose.compile(f) if hasattr(os2.panose, 'compile') else bytes(10))
    fixed += bytes([1, 0]) + struct.pack('<I', os2.usWeightClass) + struct.pack('<H', os2.fsType) + struct.pack('<H', 0x504c)
    fixed += struct.pack('<IIII', os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4)
    fixed += struct.pack('<II', os2.ulCodePageRange1, os2.ulCodePageRange2) + struct.pack('<I', f['head'].checkSumAdjustment) + bytes(16)
    assert len(fixed) == 80
    out = bytearray(fixed + names + tail + font_data)
    struct.pack_into('<I', out, 0, len(out))
    return bytes(out)

if __name__ == '__main__':
    ttf = make_font()
    block1, push, code = ctf_font(ttf)
    data = eot(mtx(block1, push, code), ttf)
    open(sys.argv[1], 'wb').write(data)
    if len(sys.argv) > 2: open(sys.argv[2], 'wb').write(ttf)
    print('original', len(ttf), 'bytes; fixture', len(data), 'bytes; blocks', len(block1), len(push), len(code))
