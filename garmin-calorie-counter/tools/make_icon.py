#!/usr/bin/env python3
"""Generate the launcher icon PNG without any third-party imaging library.

Draws a green ring with a plus sign in the middle, supersampled 4x so the
edges stay smooth once the watch scales the bitmap down.
"""
import struct
import zlib
import sys

SIZE = 64
SS = 4  # supersampling factor
BG = (0, 0, 0, 0)
RING = (60, 214, 130, 255)
PLUS = (60, 214, 130, 255)


def build_pixels():
    n = SIZE * SS
    c = (n - 1) / 2.0
    r_outer = n * 0.47
    r_inner = n * 0.35
    plus_arm = n * 0.20
    plus_half = n * 0.055

    rows = []
    for y in range(n):
        row = []
        for x in range(n):
            dx = x - c
            dy = y - c
            dist = (dx * dx + dy * dy) ** 0.5
            if r_inner <= dist <= r_outer:
                row.append(RING)
            elif (abs(dx) <= plus_half and abs(dy) <= plus_arm) or (
                abs(dy) <= plus_half and abs(dx) <= plus_arm
            ):
                row.append(PLUS)
            else:
                row.append(BG)
        rows.append(row)
    return rows


def downsample(rows):
    out = []
    for y in range(SIZE):
        row = []
        for x in range(SIZE):
            r = g = b = a = 0
            for sy in range(SS):
                for sx in range(SS):
                    pr, pg, pb, pa = rows[y * SS + sy][x * SS + sx]
                    # premultiply so transparent pixels don't bleed colour
                    r += pr * pa
                    g += pg * pa
                    b += pb * pa
                    a += pa
            if a == 0:
                row.append((0, 0, 0, 0))
            else:
                row.append((r // a, g // a, b // a, a // (SS * SS)))
        out.append(row)
    return out


def write_png(path, rows):
    raw = bytearray()
    for row in rows:
        raw.append(0)  # filter type: none
        for r, g, b, a in row:
            raw += bytes((r, g, b, a))

    def chunk(tag, data):
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", SIZE, SIZE, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as fh:
        fh.write(png)


if __name__ == "__main__":
    dest = sys.argv[1] if len(sys.argv) > 1 else "resources/drawables/launcher_icon.png"
    write_png(dest, downsample(build_pixels()))
    print("wrote " + dest)
