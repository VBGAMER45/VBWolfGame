#!/usr/bin/env python3
# This file is part of VBWolf, a Visual Basic 6 port of the Wolf3D Browser
# Version GPL Source Code. Licensed under the GNU GPL version 2 or later;
# see COPYING.txt.
#
# Generates an original level as a text map for tools/mkmap.py:
#   python tools/genlevel.py --seed 7 --difficulty 3 --cols 4 --rows 4 \
#       --name "Depot" --par 04:00 maps/w05.txt
#
# The floor is a grid of cells; each holds a room, a maze or nothing. Cells
# are joined by corridors (a spanning tree plus a few loops). The exit
# elevator is the farthest dead end, behind the gold door; from difficulty 2
# the gold key lies in a wing behind a silver door. Rooms may hide push-wall
# closets, and --secret-elevator turns one into the way to the secret floor.
# Enemy density and toughness rise with --difficulty (1..5). The same seed
# always produces the same level.

import argparse
import random
import sys
from collections import deque

CELL = 15
MATERIALS = ['#', 'R', 'W', 'M', 'T']
ACCENT = {'#': 'V', 'R': 'P', 'W': 'F'}
PUSHWALL = {'#': '*', 'R': '&', 'W': '%', 'M': '$'}
WEIGHTS = {
    1: [('guard', 75), ('dog', 25)],
    2: [('guard', 50), ('dog', 20), ('ss', 15), ('officer', 15)],
    3: [('guard', 40), ('dog', 15), ('ss', 25), ('officer', 15), ('mutant', 5)],
    4: [('guard', 30), ('dog', 10), ('ss', 30), ('officer', 15), ('mutant', 15)],
    5: [('guard', 20), ('dog', 10), ('ss', 30), ('officer', 20), ('mutant', 20)],
}
DENSITY = {1: 0.028, 2: 0.038, 3: 0.048, 4: 0.058, 5: 0.070}   # enemies per floor tile
CHAR_POOL = list('ABCHIJKNOQUXYZdegiqsz6789(),+-=/?:|_"') + ["'"]
TURN_CHAR = {'E': '}', 'S': '{', 'W': ']', 'N': '['}
BLOCKING = set('btlxypur')
DX = {'E': (1, 0), 'W': (-1, 0), 'S': (0, 1), 'N': (0, -1)}
OPP = {'E': 'W', 'W': 'E', 'N': 'S', 'S': 'N'}


class Room:
    def __init__(self, cx, cy, kind):
        self.cx, self.cy, self.kind = cx, cy, kind
        self.ox, self.oy = 1 + cx * CELL, 1 + cy * CELL
        self.links = {}             # side -> neighbour cell
        self.closet = None          # 'N' or 'S'
        self.mat = '#'

    def interior(self):
        return [(x, y) for y in range(self.y0, self.y1 + 1) for x in range(self.x0, self.x1 + 1)]


def build(args):
    rnd = random.Random(args.seed)
    d = args.difficulty
    cols, rows = args.cols, args.rows
    W, H = cols * CELL + 2, rows * CELL + 2
    if W > 64 or H > 64:
        sys.exit('genlevel: at most 4 x 4 cells fit in a 64 x 64 map')
    g = [['#'] * W for _ in range(H)]
    reserved = set()
    things = {}                     # (x, y) -> tile character

    # ---- cells and connections
    start = (0, rows - 1)
    cells = {(cx, cy) for cx in range(cols) for cy in range(rows)}
    for c in rnd.sample(sorted(cells - {start}), len(cells) // 6):
        trial = cells - {c}
        if connected(trial, start):
            cells = trial
    parent, order = {start: None}, [start]
    stack = [start]
    while stack:                                    # randomized depth-first spanning tree
        c = stack[-1]
        nbrs = [n for n in neighbours(c) if n in cells and n not in parent]
        if not nbrs:
            stack.pop()
            continue
        n = rnd.choice(nbrs)
        parent[n] = c
        order.append(n)
        stack.append(n)
    edges = {frozenset((c, p)) for c, p in parent.items() if p}
    depth = {start: 0}
    for c in order[1:]:
        depth[c] = depth[parent[c]] + 1
    children = {c: [n for n in parent if parent[n] == c] for c in cells}

    leaves = [c for c in cells if c != start and not children[c]]
    exit_cell = max(leaves, key=lambda c: (depth[c], rnd.random()))

    def subtree(c):
        out, q = [], [c]
        while q:
            n = q.pop()
            out.append(n)
            q.extend(children[n])
        return out

    on_exit_path = set()
    c = exit_cell
    while c:
        on_exit_path.add(c)
        c = parent[c]
    silver_root, silver_cells = None, set()
    if d >= 2:
        cands = [c for c in cells if c not in on_exit_path and depth[c] >= 2 and 1 <= len(subtree(c)) <= 4]
        if cands:
            silver_root = max(cands, key=lambda c: (len(subtree(c)) > 1, depth[c], rnd.random()))
            silver_cells = set(subtree(silver_root))

    # extra loops, never into the exit or the silver wing
    sealed = silver_cells | {exit_cell}
    loops = [frozenset((a, b)) for a in sorted(cells) for b in neighbours(a)
             if b in cells and a < b and frozenset((a, b)) not in edges and not ({a, b} & sealed)]
    rnd.shuffle(loops)
    edges |= set(loops[:max(1, len(cells) // 4)])

    rooms = {}
    for c in cells:
        kind = 'start' if c == start else 'exit' if c == exit_cell else 'room'
        if kind == 'room' and d >= 2 and rnd.random() < 0.18:
            kind = 'maze'
        rooms[c] = Room(c[0], c[1], kind)
    for e in edges:
        a, b = sorted(e)
        side = side_of(a, b)
        rooms[a].links[side] = b
        rooms[b].links[OPP[side]] = a

    # secret closets go on a north or south side with no corridor
    candidates = [r for r in rooms.values() if r.kind == 'room' and not (set('NS') <= set(r.links))]
    rnd.shuffle(candidates)
    secret_rooms = candidates[:args.secrets]
    for r in secret_rooms:
        r.closet = rnd.choice([s for s in 'NS' if s not in r.links])
    elevator_room = None
    if args.secret_elevator:
        if not secret_rooms:
            sys.exit('genlevel: no room has space for the secret elevator; try another seed')
        elevator_room = secret_rooms[0]

    # ---- room shapes
    for r in rooms.values():
        r.mat = rnd.choice(MATERIALS)
        if r.kind == 'maze':
            r.x0, r.y0, r.x1, r.y1 = r.ox + 2, r.oy + 2, r.ox + 12, r.oy + 12
        elif r.kind == 'exit':
            r.mat = 'M'
            r.x0, r.y0 = r.ox + 6, r.oy + 6
            r.x1, r.y1 = r.x0 + 2, r.y0 + 2
        else:
            lo_y = r.oy + (4 if r.closet == 'N' else 2)
            hi_y = r.oy + (10 if r.closet == 'S' else 12)
            w = rnd.randint(5, 11)
            h = rnd.randint(5, hi_y - lo_y + 1)
            r.x0 = rnd.randint(r.ox + 2, r.ox + 12 - w + 1)
            r.x1 = r.x0 + w - 1
            r.y0 = rnd.randint(lo_y, hi_y - h + 1)
            r.y1 = r.y0 + h - 1
        if r.mat in ('M', 'T') and r.closet:
            r.mat = rnd.choice(['#', 'R', 'W', 'M'])
        for y in range(r.y0 - 1, r.y1 + 2):
            for x in range(r.x0 - 1, r.x1 + 2):
                on_ring = x in (r.x0 - 1, r.x1 + 1) or y in (r.y0 - 1, r.y1 + 1)
                corner = x in (r.x0 - 1, r.x1 + 1) and y in (r.y0 - 1, r.y1 + 1)
                if on_ring:
                    g[y][x] = ACCENT[r.mat] if (r.mat in ACCENT and not corner and rnd.random() < 0.12) else r.mat
                else:
                    g[y][x] = '.'
        if r.kind == 'maze':
            carve_maze(g, r, rnd)

    # ---- corridors and doors
    for e in sorted(edges, key=sorted):
        a, b = sorted(e)
        ra, rb = rooms[a], rooms[b]
        side = side_of(a, b)
        if b in silver_cells and a not in silver_cells:
            kinds = (None, 'S')
        elif b == exit_cell or a == exit_cell:
            kinds = ('G', None) if a == exit_cell else (None, 'G')
        else:
            roll = rnd.random()
            kinds = ('D', None) if roll < 0.35 else (None, 'D') if roll < 0.7 else (None, None) if roll < 0.85 else ('D', 'D')
        if a in silver_cells and b not in silver_cells:
            kinds = ('S', None)
        carve_link(g, ra, rb, side, kinds, rnd, reserved)

    # ---- the start, the exit elevator and secret closets
    rs = rooms[start]
    sx, sy = (rs.x0 + rs.x1) // 2, (rs.y0 + rs.y1) // 2
    face = next(iter(rs.links))
    things[(sx, sy)] = {'N': '^', 'E': '>', 'S': 'v', 'W': '<'}[face]
    reserve_around(reserved, sx, sy, 1)

    rx = rooms[exit_cell]
    door_side = next(iter(rx.links))
    sw_x = rx.x0 - 1 if door_side == 'E' else rx.x1 + 1
    g[rx.y0 + 1][sw_x] = 'E'
    for p in rx.interior():
        reserved.add(p)
    things[(rx.x0 + 1, rx.y0)] = 'o'

    treasure = ['1', '2', '3', '4']
    for r in secret_rooms:
        cx = rnd.randint(r.x0 + 1, r.x1 - 1)
        if r.closet == 'N':
            wall_y, rows_ = r.y0 - 1, [r.y0 - 2, r.y0 - 3]
        else:
            wall_y, rows_ = r.y1 + 1, [r.y1 + 2, r.y1 + 3]
        ring_rows = range(min(rows_) - 1, max(rows_) + 2)
        for y in ring_rows:
            for x in range(cx - 2, cx + 3):
                if g[y][x] == '#' or y == wall_y:
                    g[y][x] = r.mat if (y in (min(ring_rows), max(ring_rows)) or x in (cx - 2, cx + 2) or y == wall_y) else g[y][x]
        floor_ch = '~' if r is elevator_room else '.'
        for y in rows_:
            for x in range(cx - 1, cx + 2):
                g[y][x] = floor_ch
        g[wall_y][cx] = PUSHWALL[r.mat]
        reserved.add((cx, r.y0 if r.closet == 'N' else r.y1))      # stand here to push
        if r is elevator_room:
            g[rows_[0]][cx + 2] = 'E'
            things[(cx - 1, rows_[1])] = 'o'
        else:
            spots = [(cx - 1, rows_[0]), (cx + 1, rows_[0]), (cx - 1, rows_[1]), (cx + 1, rows_[1])]
            for i, p in enumerate(spots[:2 + min(d, 2)]):
                things[p] = '5' if (i == 3 and d >= 3) else rnd.choice(treasure)

    # ---- patrols, decoration, items and enemies
    defs, chars = [], {}

    def enemy_char(kind, mode, facing, tier='', ambush=False):
        key = (kind, mode, facing, tier, ambush)
        if key not in chars:
            if not CHAR_POOL:
                sys.exit('genlevel: ran out of tile characters for enemies')
            chars[key] = CHAR_POOL.pop(0)
            defs.append('@enemy %s %s %s %s%s%s' % (chars[key], kind, mode, facing,
                                                   ' ' + tier if tier else '', ' ambush' if ambush else ''))
        return chars[key]

    weights = WEIGHTS[d]
    stats = {'enemies': 0, 'rooms': len(rooms)}
    near_start = {rooms[n] for n in rs.links.values()}
    for r in sorted(rooms.values(), key=lambda r: (r.cy, r.cx)):
        if r.kind in ('exit',):
            continue
        floor_tiles = [p for p in r.interior() if g[p[1]][p[0]] == '.']
        free = lambda: [p for p in floor_tiles if p not in reserved and p not in things]
        w, h = r.x1 - r.x0 + 1, r.y1 - r.y0 + 1
        patrol_slot = None
        if r.kind == 'room' and w >= 7 and h >= 7 and rnd.random() < 0.75:
            loop = ([(x, r.y0 + 1) for x in range(r.x0 + 1, r.x1)] + [(r.x1 - 1, y) for y in range(r.y0 + 1, r.y1)] +
                    [(x, r.y1 - 1) for x in range(r.x0 + 1, r.x1)] + [(r.x0 + 1, y) for y in range(r.y0 + 1, r.y1)])
            if not any(p in things for p in loop):
                reserved.update(loop)
                things[(r.x0 + 1, r.y0 + 1)] = TURN_CHAR['E']
                things[(r.x1 - 1, r.y0 + 1)] = TURN_CHAR['S']
                things[(r.x1 - 1, r.y1 - 1)] = TURN_CHAR['W']
                things[(r.x0 + 1, r.y1 - 1)] = TURN_CHAR['N']
                patrol_slot = (r.x0 + 3, r.y0 + 1)
        if r.kind == 'room':
            # pillars on an inner grid, lights, and furniture in the corners
            if w >= 7 and h >= 7:
                for y in range(r.y0 + 2, r.y1 - 1, 3):
                    for x in range(r.x0 + 2, r.x1 - 1, 3):
                        if (x, y) not in reserved and (x, y) not in things and rnd.random() < 0.6:
                            things[(x, y)] = 'x'
            for y in range(r.y0 + 1, r.y1, 4):
                for x in range(r.x0 + 1, r.x1, 4):
                    if (x, y) not in things:
                        things[(x, y)] = 'o'
            for p in [(r.x0, r.y0), (r.x1, r.y0), (r.x0, r.y1), (r.x1, r.y1)]:
                if p not in reserved and p not in things and rnd.random() < 0.5:
                    things[p] = rnd.choice('blprut' if r.mat == 'W' else 'blbrup')
        if r.kind == 'start':
            for p in rnd.sample(free(), min(2, len(free()))):
                things[p] = 'a'
            continue

        # enemies
        n = round(len(floor_tiles) * DENSITY[d] * (0.6 if r.kind == 'maze' else 1) * (0.5 if r in near_start else 1))
        extra_med = round(n * 0.3)
        extra_hard = round(n * 0.3)
        plan = [''] * n + ['medium'] * extra_med + ['hard'] * extra_hard
        for tier in plan:
            kind = pick(weights, rnd)
            if patrol_slot and (kind == 'dog' or rnd.random() < 0.4):
                things[patrol_slot] = enemy_char(kind, 'patrol', 'E', tier)
                stats['enemies'] += tier == ''
                patrol_slot = None
                continue
            if kind == 'dog':
                kind = 'guard'
            spots = [p for p in free() if abs(p[0] - sx) + abs(p[1] - sy) > 3]
            if not spots:
                break
            p = rnd.choice(spots)
            facing = rnd.choice('NESW')
            ambush = tier == '' and rnd.random() < 0.15
            things[p] = enemy_char(kind, 'stand', facing, tier, ambush)
            stats['enemies'] += tier == ''

        # supplies, scaled to the enemies in the room
        for item, count in (('a', max(1, n // 2)), ('h', 1 if n >= 3 or rnd.random() < 0.3 else 0),
                            ('f', 1 if rnd.random() < 0.3 else 0), (rnd.choice(treasure), 1 if rnd.random() < 0.5 else 0)):
            for _ in range(count):
                spots = free()
                if spots:
                    things[rnd.choice(spots)] = item

    # one heavier weapon per floor, in a room away from the start
    far = max((r for r in rooms.values() if r.kind == 'room'), key=lambda r: depth[(r.cx, r.cy)], default=None)
    if far:
        spots = [p for p in far.interior() if g[p[1]][p[0]] == '.' and p not in reserved and p not in things]
        if spots:
            things[rnd.choice(spots)] = 'm' if d <= 2 else 'n'

    # keys
    def key_spot(cellset, avoid):
        pool = sorted((rooms[c] for c in cellset if rooms[c].kind in ('room', 'maze') and c not in avoid),
                      key=lambda r: -depth[(r.cx, r.cy)])
        for r in pool:
            spots = [p for p in r.interior() if g[p[1]][p[0]] == '.' and p not in reserved and p not in things]
            if spots:
                return rnd.choice(spots)
        sys.exit('genlevel: nowhere to put a key; try another seed')

    if silver_cells:
        things[key_spot(silver_cells, set())] = 'k'
        things[key_spot(cells - silver_cells - {exit_cell, start}, set())] = 'j'
    else:
        things[key_spot(cells - {exit_cell, start}, set())] = 'k'

    for (x, y), ch in things.items():
        g[y][x] = ch
    check_progression(g, sx, sy)
    return g, defs, stats


def pick(weights, rnd):
    total = sum(w for _, w in weights)
    roll = rnd.uniform(0, total)
    for kind, w in weights:
        roll -= w
        if roll <= 0:
            return kind
    return weights[-1][0]


def neighbours(c):
    x, y = c
    return [(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)]


def connected(cells, start):
    seen, q = {start}, [start]
    while q:
        c = q.pop()
        for n in neighbours(c):
            if n in cells and n not in seen:
                seen.add(n)
                q.append(n)
    return seen == cells


def side_of(a, b):
    if b[0] > a[0]:
        return 'E'
    if b[0] < a[0]:
        return 'W'
    return 'S' if b[1] > a[1] else 'N'


def reserve_around(reserved, x, y, r):
    for yy in range(y - r, y + r + 1):
        for xx in range(x - r, x + r + 1):
            reserved.add((xx, yy))


def carve_maze(g, r, rnd):
    """Fills the room's box with a maze on a 2-tile lattice, with a few loops."""
    for p in r.interior():
        g[p[1]][p[0]] = r.mat
    n = (r.x1 - r.x0) // 2 + 1
    seen, stack = {(0, 0)}, [(0, 0)]
    g[r.y0][r.x0] = '.'
    while stack:
        i, j = stack[-1]
        opts = [(i + di, j + dj) for di, dj in ((1, 0), (-1, 0), (0, 1), (0, -1))
                if 0 <= i + di < n and 0 <= j + dj < n and (i + di, j + dj) not in seen]
        if not opts:
            stack.pop()
            continue
        ni, nj = rnd.choice(opts)
        g[r.y0 + j + nj][r.x0 + i + ni] = '.'
        g[r.y0 + 2 * nj][r.x0 + 2 * ni] = '.'
        seen.add((ni, nj))
        stack.append((ni, nj))
    for _ in range(n):                               # knock out a few walls for loops
        i, j = rnd.randrange(n - 1), rnd.randrange(n)
        g[r.y0 + 2 * j][r.x0 + 2 * i + 1] = '.'


def carve_link(g, ra, rb, side, kinds, rnd, reserved):
    """Joins room a to room b (b lies to the east or south) with a bent corridor."""
    def spot(r, horizontal):
        if r.kind == 'maze':
            vals = range(r.y0, r.y1 + 1, 2) if horizontal else range(r.x0, r.x1 + 1, 2)
        else:
            vals = range(r.y0, r.y1 + 1) if horizontal else range(r.x0, r.x1 + 1)
        return rnd.choice(list(vals))

    if side == 'E':
        ya, yb = spot(ra, True), spot(rb, True)
        bx = ra.ox + CELL - 1
        path = [(x, ya) for x in range(ra.x1 + 2, bx + 1)] + [(bx, y) for y in range(min(ya, yb), max(ya, yb) + 1)] + \
               [(x, yb) for x in range(bx, rb.x0 - 1)]
        ends = [((ra.x1 + 1, ya), (ra.x1, ya)), ((rb.x0 - 1, yb), (rb.x0, yb))]
    else:
        xa, xb = spot(ra, False), spot(rb, False)
        by = ra.oy + CELL - 1
        path = [(xa, y) for y in range(ra.y1 + 2, by + 1)] + [(x, by) for x in range(min(xa, xb), max(xa, xb) + 1)] + \
               [(xb, y) for y in range(by, rb.y0 - 1)]
        ends = [((xa, ra.y1 + 1), (xa, ra.y1)), ((xb, rb.y0 - 1), (xb, rb.y0))]
    for x, y in path:
        g[y][x] = '.'
    for (wall, inside), kind in zip(ends, kinds):
        g[wall[1]][wall[0]] = kind or '.'
        reserved.add(inside)
        reserved.add(wall)


def check_progression(g, sx, sy):
    """Silver key, then gold key, then the exit must each be reachable in turn."""
    H, W = len(g), len(g[0])
    solid = set('#VRPWFTMEDGSL*&%$') | BLOCKING

    def reach(passable_doors):
        seen, q = {(sx, sy)}, deque([(sx, sy)])
        while q:
            x, y = q.popleft()
            for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                ch = g[ny][nx]
                ok = ch not in solid or ch in passable_doors
                if ok and (nx, ny) not in seen:
                    seen.add((nx, ny))
                    q.append((nx, ny))
        return seen

    def find(ch):
        return [(x, y) for y in range(H) for x in range(W) if g[y][x] == ch]

    open_ = reach('D')
    silver = find('j')
    if silver and not all(p in open_ for p in silver):
        sys.exit('genlevel: the silver key is locked away; try another seed')
    with_silver = reach('DS')
    if not all(p in with_silver for p in find('k')):
        sys.exit('genlevel: the gold key cannot be reached; try another seed')
    with_gold = reach('DSG')
    exits = [(x, y) for y in range(H) for x in range(W) if g[y][x] == 'E']
    if not any((x - 1, y) in with_gold or (x + 1, y) in with_gold for x, y in exits):
        sys.exit('genlevel: the exit cannot be reached; try another seed')


LEGEND = """; Walls: # stone  V stone vent  R brick  P brick plaque  W wood  F wood painting  T tile  M metal
;        E elevator switch (use it facing east or west)
; Push walls (secret): * stone  & brick  % wood  $ metal
; Doors: D normal  G gold lock  S silver lock  L elevator      Start: ^ > v <  (facing N E S W)
; Objects: w puddle  b barrel  r wooden barrel  t table  l floor lamp  c chandelier  o ceiling light
;          x pillar  y tree  p plant  u urn
; Items:   k gold key  j silver key  f food  h first aid  a ammo  m machine gun  n rotary gun
;          1-4 treasure  5 extra life  ! exit tile  ~ secret elevator floor      Floor: .
"""


def main():
    ap = argparse.ArgumentParser(description='Generate an original VBWolf level as a text map.')
    ap.add_argument('out')
    ap.add_argument('--seed', type=int, required=True)
    ap.add_argument('--difficulty', type=int, choices=range(1, 6), required=True)
    ap.add_argument('--cols', type=int, default=4)
    ap.add_argument('--rows', type=int, default=4)
    ap.add_argument('--secrets', type=int, default=2)
    ap.add_argument('--secret-elevator', action='store_true')
    ap.add_argument('--name', required=True)
    ap.add_argument('--par', default='04:00')
    ap.add_argument('--ceiling', default='48 48 56')
    ap.add_argument('--floor', default='96 92 90')
    args = ap.parse_args()
    g, defs, stats = build(args)
    with open(args.out, 'w', newline='\r\n') as f:
        f.write('; VBWolf level "%s" (original), generated by:\n' % args.name)
        f.write(';   python tools/genlevel.py %s --seed %d --difficulty %d --cols %d --rows %d --secrets %d%s --name "%s" --par %s\n'
                % (args.out, args.seed, args.difficulty, args.cols, args.rows, args.secrets,
                   ' --secret-elevator' if args.secret_elevator else '', args.name, args.par))
        f.write('@name %s\n@par %s\n@ceiling %s\n@floor %s\n' % (args.name, args.par, args.ceiling, args.floor))
        f.write(LEGEND)
        f.write('\n'.join(defs + ['@turn %s %s' % (c, d) for d, c in TURN_CHAR.items()]) + '\n')
        f.write('\n'.join(''.join(r) for r in g) + '\n')
    print('%s: %d rooms, %d enemies (plus medium/hard extras)' % (args.out, stats['rooms'], stats['enemies']))


if __name__ == '__main__':
    main()
