# Content packs

A content pack is a folder of your own art, sounds and levels that VBWolf lays
over its built-in content. Every file replaces one slot; anything the pack
does not provide stays built-in. Choose a pack from the title screen
(**Content Pack > Choose folder...**); the choice is remembered in
`VBWolf.ini`. **Use built-in content** switches it off again.

`packs/example` is a small working pack, built by
`python tools/make_example_pack.py packs/example`.

Only use content you have the rights to share. VBWolf's own content is GPL;
a pack is your work and you choose its license.

## Layout

```
mypack\
  pack.txt                 optional:  @name My Pack
  walls\<slot>.png|bmp     wall and door textures
  sprites\<slot>.png|bmp   objects, pickups and enemies
  weapons\<w>_<f>.png|bmp  first-person weapon frames
  sounds\<name>.wav        sound effects
  maps\episode.txt         optional: your own level set (with its .map files)
```

Slot numbers can be written plainly (`14.png`) or with three digits
(`014.png`).

## Images

- PNG or BMP, any size. Images are scaled to 64 x 64; square images look best.
- Walls and doors are opaque.
- Sprites and weapon frames are see-through where the alpha channel is below
  half, or where the colour is pure magenta (255, 0, 255), which is handy for
  BMPs.
- Sprites stand on the floor and reach the ceiling: the top row of the image
  touches the ceiling, the bottom row touches the floor. Small objects belong
  at the bottom of the image.
- Weapon frames are drawn at double size from the bottom of the view; keep the
  weapon in the lower half of the image so it doesn't hide the centre.

## Wall and door slots

Map wall code *k* (as written by `tools/mkmap.py`) uses slot `2(k-1)` for faces
seen east-west and the darker `2(k-1)+1` for faces seen north-south. If a
pack gives only the even slot, the odd one is made automatically by
darkening it.

| Slots | Built-in use | `mkmap` tile |
| --- | --- | --- |
| 0, 1 | stone | `#` |
| 2, 3 | stone with vent | `V` |
| 4, 5 | brick | `R` |
| 6, 7 | brick with plaque | `P` |
| 8, 9 | wood | `W` |
| 10, 11 | wood with painting | `F` |
| 12, 13 | blue tiles | `T` |
| 14, 15 | metal panels (also the title-screen backdrop) | `M` |
| 16 - 39 | unused wall codes 9 - 20 | |
| 40, 42 | elevator rails (north/south faces of the switch wall) | `E` |
| 41 | elevator switch, ready | `E` |
| 43 | elevator switch, pressed | |
| 44 - 97 | unused wall codes 23 - 49 | |
| 98, 99 | door | `D` |
| 100, 101 | door frame (the wall beside a door) | |
| 102, 103 | elevator door | `L` |
| 104 | gold-locked door | `G` |
| 105 | silver-locked door | `S` |

## Sprite slots

### Objects and pickups

Static object *n* (map object code 23 + *n*) uses sprite slot `2 + n`.
Objects marked *blocks* stop the player.

| Slot | Built-in use | `mkmap` tile |
| --- | --- | --- |
| 2 | puddle | `w` |
| 3 | barrel, blocks | `b` |
| 4 | table, blocks | `t` |
| 5 | floor lamp, blocks | `l` |
| 6 | chandelier | `c` |
| 9 | pillar, blocks | `x` |
| 10 | potted tree, blocks | `y` |
| 13 | plant, blocks | `p` |
| 14 | urn, blocks | `u` |
| 16 | ceiling light | `o` |
| 22 | gold key | `k` |
| 23 | silver key | `j` |
| 26 | food (+10 health) | `f` |
| 27 | first aid (+25 health) | `h` |
| 28 | ammo clip (also dropped by enemies) | `a` |
| 29 | machine gun | `m` |
| 30 | rotary gun | `n` |
| 31 - 34 | treasure: 100, 500, 1000, 5000 points | `1` - `4` |
| 35 | extra life | `5` |
| 37 | wooden barrel, blocks | `r` |
| 51 | ammo box (dropped) | |

Other slots from 2 to 53 have no built-in art and show a magenta box; they
belong to object codes the built-in levels don't use.

### Enemies

Each enemy has rotated frames for standing and walking, and single frames
for the rest. A group of 8 slots holds one frame seen from 8 directions: the
first faces the viewer, the third faces screen-left, the fifth faces away
and the seventh faces screen-right.

| Enemy (`@enemy` type) | Stand (8) | Walk 1-4 (8 each) | Pain | Dying | Dead | Attack |
| --- | --- | --- | --- | --- | --- | --- |
| Sentry (`guard`) | 54-61 | 62-93 | 94, 98 | 95-97 | 99 | 100-102 |
| Hound (`dog`) | - | 103-134 | - | 135-137 | 138 | 139-141 |
| Enforcer (`ss`) | 142-149 | 150-181 | 182, 186 | 183-185 | 187 | 188-190 |
| Brute (`mutant`) | 191-198 | 199-230 | 231, 235 | 232-234, 236 | 237 | 238-241 |
| Scout (`officer`) | 242-249 | 250-281 | 282, 286 | 283-285, 287 | 288 | 289-291 |

## Weapon frames

`weapons\<w>_<f>` where *w* is 0 knife, 1 pistol, 2 machine gun, 3 rotary gun
and *f* is the frame: 0 idle, 1 raised, 2 firing, 3 recoil.

## Sounds

`sounds\<name>.wav`, PCM, 8 or 16 bit, mono or stereo, any sample rate (up to
5 seconds are used):

`pistol`, `auto`, `rotary`, `knife`, `enemy_shot`, `heavy_shot`, `bite`,
`door_open`, `door_close`, `pushwall`, `ammo`, `health`, `treasure`, `key`,
`weapon`, `robot_death`, `hound_death`, `player_death`, `elevator`, `alert`,
`ricochet`, `gibs`.

## Levels

Put compiled maps and an `episode.txt` in the pack's `maps` folder (see the
README for the level format, `tools/mkmap.py` and `tools/genlevel.py`). When
present, the pack's episode replaces the built-in one.
