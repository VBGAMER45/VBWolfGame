#!/usr/bin/env python3
# This file is part of VBWolf, a Visual Basic 6 port of the Wolf3D Browser
# Version GPL Source Code. Licensed under the GNU GPL version 2 or later;
# see COPYING.txt.
#
# Compiles a text level (maps/*.txt) into the .map format read by
# modLevel.bas (the format of the browser version's maps/wNN.map files):
#   python tools/mkmap.py maps/w00.txt maps/w00.map
#
# Text format: one character per tile, north at the top. Lines starting with
# ";" are comments; lines starting with "@" set a header value:
#   @name <text>   @music <text>   @par MM:SS
#   @ceiling R G B   @floor R G B
#   @enemy <char> <guard|officer|ss|dog|mutant> <stand|patrol> <N|E|S|W> [medium|hard] [ambush]
#   @enemy <char> wreck
#   @turn <char> <E|NE|N|NW|W|SW|S|SE>      (patrol waypoint: walkers turn here)
# A "~" floor tile marks a secret elevator room: its elevator leads to the
# episode's secret floor.
# Tiles beyond the text are solid wall (code 1). See the tables below.

import struct
import sys
from collections import deque

WALLS = {'#': 1, 'V': 2, 'R': 3, 'P': 4, 'W': 5, 'F': 6, 'T': 7, 'M': 8, 'E': 21}
PUSHWALLS = {'*': 1, '&': 3, '%': 5, '$': 8}            # wall code, plus the secret marker
DOORS = {'D': (0x5A, 0x5B), 'G': (0x5C, 0x5D), 'S': (0x5E, 0x5F), 'L': (0x64, 0x65)}  # (vertical, horizontal)
STARTS = {'^': 0x13, '>': 0x14, 'v': 0x15, '<': 0x16}
STATICS = {                                              # object code = 23 + static number
    'w': 0, 'b': 1, 't': 2, 'l': 3, 'c': 4, 'x': 7, 'y': 8, 'p': 11, 'u': 12, 'o': 14, 'r': 35,
    'k': 20, 'j': 21, 'f': 24, 'h': 25, 'a': 26, 'm': 27, 'n': 28,
    '1': 29, '2': 30, '3': 31, '4': 32, '5': 33,
}
BLOCKING = {1, 2, 3, 5, 7, 8, 10, 11, 12, 13, 16, 17, 18, 22, 35, 36, 37, 39, 40, 45, 46}
EXIT = '!'                                               # victory trigger tile
SECRETEXIT = '~'                                         # floor of a secret elevator room
ENEMIES = {'guard': (108, 36), 'officer': (116, 36), 'ss': (126, 36), 'dog': (134, 36), 'mutant': (216, 18)}
DIR4 = {'E': 0, 'N': 1, 'W': 2, 'S': 3}
DIR8 = {'E': 0, 'NE': 1, 'N': 2, 'NW': 3, 'W': 4, 'SW': 5, 'S': 6, 'SE': 7}
WRECK = 124
AMBUSHTILE = 0x6A
custom = {}                                              # char -> (object code, ambush) from @enemy/@turn
FIRSTAREA = 0x6B
NUMAREAS = 37
RLEW_TAG = 0xABCD
SECRET = 0x62
VICTORY = 0x63


def fail(msg):
    sys.exit('mkmap: ' + msg)


def define(key, args):
    if not args or len(args[0]) != 1:
        fail('@%s needs a single tile character first' % key)
    ch = args[0]
    builtin = set(WALLS) | set(PUSHWALLS) | set(DOORS) | set(STARTS) | set(STATICS) | {EXIT, SECRETEXIT, '.', ' ', ';', '@'}
    if ch in builtin or ch in custom:
        fail('tile character %r is already in use' % ch)
    if key == 'turn':
        if len(args) != 2 or args[1] not in DIR8:
            fail('usage: @turn <char> <E|NE|N|NW|W|SW|S|SE>')
        custom[ch] = (0x5A + DIR8[args[1]], False)
        return
    if len(args) == 2 and args[1] == 'wreck':
        custom[ch] = (WRECK, False)
        return
    if len(args) < 4 or args[1] not in ENEMIES or args[2] not in ('stand', 'patrol') or args[3] not in DIR4:
        fail('usage: @enemy <char> <guard|officer|ss|dog|mutant> <stand|patrol> <N|E|S|W> [medium|hard] [ambush]')
    if args[1] == 'dog' and args[2] == 'stand':
        fail('dogs can only patrol')
    extra = args[4:]
    if any(e not in ('medium', 'hard', 'ambush') for e in extra):
        fail('unknown @enemy option in %r' % ' '.join(extra))
    base, step = ENEMIES[args[1]]
    tier = 2 if 'hard' in extra else 1 if 'medium' in extra else 0
    code = base + (4 if args[2] == 'patrol' else 0) + DIR4[args[3]] + tier * step
    custom[ch] = (code, 'ambush' in extra)


def parse(path):
    header = {'name': 'Floor 1', 'music': '', 'par': '01:30', 'ceiling': (56, 56, 60), 'floor': (104, 100, 96)}
    rows = []
    with open(path, encoding='ascii') as f:
        for line in f:
            line = line.rstrip('\r\n')
            if line.startswith(';'):
                continue
            if line.startswith('@'):
                key, _, value = line[1:].partition(' ')
                if key in ('enemy', 'turn'):
                    define(key, value.split())
                    continue
                if key in ('ceiling', 'floor'):
                    header[key] = tuple(int(v) for v in value.split())
                elif key in header:
                    header[key] = value.strip()
                else:
                    fail('unknown header @' + key)
                continue
            if line.strip():
                rows.append(line)
    if len(rows) > 64 or any(len(r) > 64 for r in rows):
        fail('levels are at most 64x64 tiles')
    grid = [[(rows[y][x] if y < len(rows) and x < len(rows[y]) else '#') for x in range(64)] for y in range(64)]
    for i in range(64):
        for x, y in ((i, 0), (i, 63), (0, i), (63, i)):
            if grid[y][x] not in WALLS:
                grid[y][x] = '#'
    return header, grid


def is_wall(ch):
    return ch in WALLS or ch in PUSHWALLS


def build(grid):
    plane1 = [0] * 4096
    plane2 = [0] * 4096
    floor = [[not is_wall(grid[y][x]) and grid[y][x] not in DOORS for x in range(64)] for y in range(64)]

    # areas: each region of floor bounded by walls and doors. Area 0 is special to
    # the engine (an elevator used there leads to the secret floor), so it is kept
    # for regions containing a "~" tile and other regions are numbered from 1.
    area = [[-1] * 64 for _ in range(64)]
    count = 1
    for y in range(64):
        for x in range(64):
            if floor[y][x] and area[y][x] < 0:
                region, q = [(x, y)], deque([(x, y)])
                area[y][x] = -2
                while q:
                    cx, cy = q.popleft()
                    for nx, ny in ((cx + 1, cy), (cx - 1, cy), (cx, cy + 1), (cx, cy - 1)):
                        if floor[ny][nx] and area[ny][nx] == -1:
                            area[ny][nx] = -2
                            region.append((nx, ny))
                            q.append((nx, ny))
                if any(grid[ty][tx] == SECRETEXIT for tx, ty in region):
                    num = 0
                else:
                    if count >= NUMAREAS:
                        fail('more than %d separate areas' % (NUMAREAS - 1))
                    num = count
                    count += 1
                for tx, ty in region:
                    area[ty][tx] = num

    starts = 0
    for y in range(64):
        for x in range(64):
            ch, i = grid[y][x], y * 64 + x
            if ch in WALLS:
                plane1[i] = WALLS[ch]
            elif ch in PUSHWALLS:
                plane1[i] = PUSHWALLS[ch]
                plane2[i] = SECRET
            elif ch in DOORS:
                ew = floor[y][x - 1] and floor[y][x + 1]
                ns = floor[y - 1][x] and floor[y + 1][x]
                if ew == ns:
                    fail('door at %d,%d needs floor on exactly one axis and walls on the other' % (x, y))
                plane1[i] = DOORS[ch][0] if ew else DOORS[ch][1]
            else:
                plane1[i] = FIRSTAREA + area[y][x]
                if ch in custom:
                    plane2[i] = custom[ch][0]
                    if custom[ch][1]:
                        plane1[i] = AMBUSHTILE
                elif ch in STARTS:
                    plane2[i] = STARTS[ch]
                    starts += 1
                elif ch in STATICS:
                    plane2[i] = 23 + STATICS[ch]
                elif ch == EXIT:
                    plane2[i] = VICTORY
                elif ch not in '. ' + SECRETEXIT:
                    fail('unknown tile %r at %d,%d' % (ch, x, y))
    if starts != 1:
        fail('need exactly one player start, found %d' % starts)
    return plane1, plane2, count


def check_exit(grid):
    # the exit is a walkable tile beside an elevator switch on its east or west side, or a "!" tile
    blocked = lambda ch: is_wall(ch) or (ch in STATICS and STATICS[ch] in BLOCKING)
    start = next((x, y) for y in range(64) for x in range(64) if grid[y][x] in STARTS)
    seen, q = {start}, deque([start])
    while q:
        x, y = q.popleft()
        if grid[y][x] == EXIT or 'E' in (grid[y][x - 1], grid[y][x + 1]):
            return
        for n in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
            if n not in seen and not blocked(grid[n[1]][n[0]]):
                seen.add(n)
                q.append(n)
    fail('no reachable exit (an elevator switch "E" faced east or west, or a "!" tile)')


def rlew(words):
    out, i = [], 0
    while i < len(words):
        run = 1
        while i + run < len(words) and words[i + run] == words[i] and run < 0xFFFF:
            run += 1
        if run > 3 or words[i] == RLEW_TAG:
            out += [RLEW_TAG, run, words[i]]
            i += run
        else:
            out.append(words[i])
            i += 1
    return out


def carmack(words):
    # every word is written literally; words whose high byte is a Carmack tag are escaped
    out = bytearray()
    for w in words:
        if w >> 8 in (0xA7, 0xA8):
            out += bytes([0, w >> 8, w & 0xFF])
        else:
            out += struct.pack('<H', w)
    return bytes(out)


def plane_bytes(plane):
    stream = [len(plane) * 2] + rlew(plane)                 # RLEW data starts with its expanded size
    data = struct.pack('<H', len(stream) * 2) + carmack(stream)
    if len(data) > 0xFFFF:
        fail('plane too large')
    return data


def write(path, header, planes):
    name = header['name'].encode('ascii')
    music = header['music'].encode('ascii')
    par = header['par'].encode('ascii')[:5].ljust(5, b'\0')
    blobs = [plane_bytes(p) for p in planes]
    offset = 49 + len(name) + len(music)
    offsets = []
    for b in blobs:
        offsets.append(offset)
        offset += len(b)
    mins, secs = header['par'].split(':')
    out = bytearray()
    out += struct.pack('<IHHH', 0x21444921, RLEW_TAG, 64, 64)
    out += bytes(list(header['ceiling']) + [255]) + bytes(list(header['floor']) + [255])
    out += struct.pack('<HHH', *[len(b) for b in blobs])
    out += struct.pack('<III', *offsets)
    out += struct.pack('<HHf', len(name), len(music), int(mins) * 60 + int(secs))
    out += par + name + music
    for b in blobs:
        out += b
    with open(path, 'wb') as f:
        f.write(out)


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__ or 'usage: mkmap.py level.txt level.map')
    header, grid = parse(sys.argv[1])
    plane1, plane2, areas = build(grid)
    check_exit(grid)
    write(sys.argv[2], header, [plane1, plane2, [0] * 4096])
    print('%s: %d areas' % (sys.argv[2], areas))


if __name__ == '__main__':
    main()
