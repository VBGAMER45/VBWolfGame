#!/usr/bin/env python3
# This file is part of VBWolf, a Visual Basic 6 port of the Wolf3D Browser
# Version GPL Source Code. Licensed under the GNU GPL version 2 or later;
# see COPYING.txt.
#
# Builds packs/example, a small content pack that shows the format: PNG and
# BMP images, transparency, a weapon frame and a sound. Everything is drawn
# here, so the pack is original.
#   python tools/make_example_pack.py packs/example

import math
import os
import struct
import sys
import zlib

MAGENTA = (255, 0, 255, 255)


def write_png(path, pixels, w=64, h=64):
    raw = b''.join(b'\0' + b''.join(bytes(pixels[y][x]) for x in range(w)) for y in range(h))

    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n')
        f.write(chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)))
        f.write(chunk(b'IDAT', zlib.compress(raw, 9)))
        f.write(chunk(b'IEND', b''))


def write_bmp(path, pixels, w=64, h=64):
    rows = []
    for y in range(h - 1, -1, -1):                      # bottom-up
        rows.append(b''.join(bytes((p[2], p[1], p[0])) for p in pixels[y]))
    data = b''.join(rows)
    with open(path, 'wb') as f:
        f.write(b'BM' + struct.pack('<IHHI', 54 + len(data), 0, 0, 54))
        f.write(struct.pack('<IiiHHIIiiII', 40, w, h, 1, 24, 0, len(data), 2835, 2835, 0, 0))
        f.write(data)


def write_wav(path, samples, rate):
    data = bytes(max(0, min(255, int(128 + s * 120))) for s in samples)     # 8-bit unsigned
    with open(path, 'wb') as f:
        f.write(b'RIFF' + struct.pack('<I', 36 + len(data)) + b'WAVE')
        f.write(b'fmt ' + struct.pack('<IHHIIHH', 16, 1, 1, rate, rate, 1, 8))
        f.write(b'data' + struct.pack('<I', len(data)) + data)


def blank(fill=(0, 0, 0, 0)):
    return [[fill for _ in range(64)] for _ in range(64)]


def main():
    out = sys.argv[1]
    for sub in ('walls', 'sprites', 'weapons', 'sounds'):
        os.makedirs(os.path.join(out, sub), exist_ok=True)
    with open(os.path.join(out, 'pack.txt'), 'w', newline='\r\n') as f:
        f.write('; VBWolf content pack. See PACKS.md for the folder layout.\n@name Example Pack\n')

    # walls/0.png: teal tiles in place of the stone wall (its dark face is made automatically)
    img = blank()
    for y in range(64):
        for x in range(64):
            lx, ly = x % 16, y % 16
            if lx == 0 or ly == 0:
                img[y][x] = (40, 60, 60, 255)
            else:
                v = 1.0 + (0.12 if lx < 3 or ly < 3 else 0) - (0.15 if lx > 13 or ly > 13 else 0)
                img[y][x] = (int(40 * v), int(150 * v), int(140 * v), 255)
    write_png(os.path.join(out, 'walls', '0.png'), img)

    # walls/98.bmp: a hazard-striped door, as a BMP
    img = blank()
    for y in range(64):
        for x in range(64):
            stripe = ((x + y) // 8) % 2 == 0
            img[y][x] = (230, 160, 30, 255) if stripe else (40, 40, 44, 255)
            if 20 <= y <= 43 and 12 <= x <= 51:
                img[y][x] = (90, 95, 105, 255)
    write_bmp(os.path.join(out, 'walls', '98.bmp'), img)

    # sprites/3.png: a red crate in place of the barrel, with a transparent background
    img = blank()
    for y in range(34, 64):
        for x in range(16, 48):
            edge = x in (16, 17, 46, 47) or y in (34, 35, 62, 63) or y in (44, 45, 53, 54)
            img[y][x] = (90, 25, 20, 255) if edge else (180, 50, 40, 255)
    write_png(os.path.join(out, 'sprites', '3.png'), img)

    # sprites/22.bmp: a different gold key, magenta marking the see-through parts
    img = blank(MAGENTA)
    for y in range(48, 62):
        for x in range(18, 30):
            if (x - 24) ** 2 + (y - 55) ** 2 <= 36 and (x - 24) ** 2 + (y - 55) ** 2 > 6:
                img[y][x] = (240, 200, 60, 255)
    for x in range(29, 46):
        for y in range(54, 57):
            img[y][x] = (220, 180, 50, 255)
    for x, y in ((40, 57), (40, 58), (44, 57), (44, 58), (44, 59)):
        img[y][x] = (220, 180, 50, 255)
    write_bmp(os.path.join(out, 'sprites', '22.bmp'), img)

    # weapons/1_0.png: a bronze pistol for the idle frame
    img = blank()
    for y in range(36, 64):
        for x in range(22, 42):
            if y < 47:
                img[y][x] = (170, 110, 50, 255) if x > 23 else (210, 150, 80, 255)
            elif 26 <= x <= 37:
                img[y][x] = (60, 40, 30, 255)
    write_png(os.path.join(out, 'weapons', '1_0.png'), img)

    # sounds/ammo.wav: a rising two-note chirp, 8-bit at 11025 Hz
    rate = 11025
    samples = []
    for i in range(int(rate * 0.18)):
        t = i / rate
        f = 900 if t < 0.08 else 1350
        samples.append(math.sin(2 * math.pi * f * t) * math.exp(-8 * t))
    write_wav(os.path.join(out, 'sounds', 'ammo.wav'), samples, rate)
    print('wrote', out)


if __name__ == '__main__':
    main()
