# VBWolf

A Visual Basic 6 port of the **Wolf3D Browser Version GPL Source Code**
(<https://github.com/id-Software/wolf3d-browser>),
Copyright (C) 2012 id Software LLC, a ZeniMax Media company.

## License

VBWolf is free software under the **GNU General Public License, version 2 or
later**, the same license as the source it is ported from. See `COPYING.txt`.
If you distribute `VBWolf.exe`, you must also make this source code available
under the GPL.

## Game data

id's GPL release contains no game data ("the game data remains subject to the
original EULA"), and VBWolf does not use any. The textures, objects, weapon
views and enemies are original art drawn in code at startup (`modArt.bas`);
the enemies are robots (sentry, scout, enforcer, brute and a robot hound) that
fill the generic enemy types' sprite slots. The sound effects are original
too, synthesized at startup (`modSound.bas`), and the levels in `maps/` are
original.

VBWolf does not read the original game's data files. If you own Wolfenstein
3D, source ports such as ECWolf or Wolf4SDL play it from your installed copy.

## Content packs

You can replace any of the built-in art, sounds and levels with your own:
put PNG/BMP images, WAV sounds and optionally a level set in a folder, then
choose it from the title screen under **Content Pack**. Each file replaces one
slot and everything else stays built-in. `PACKS.md` describes the folder
layout and every slot; `packs/example` is a small working pack.

## Running

`VBWolf.exe` opens the title screen. New Game asks for a skill level, then
plays the episode listed in `maps\episode.txt`: seven floors (Intake,
Assembly, Core, Foundry, Cooling Towers, Depot and Reactor Spire) with enemies
growing denser and tougher as you go, plus a secret floor, the Vault, behind a
secret elevator on Cooling Towers. Esc pauses; from the pause menu you can resume or quit to the title.

For testing, the command line can name another episode file
(`VBWolf.exe maps\other.txt`) or a single map to play straight away
(`VBWolf.exe maps\mylevel.map`).

| Key | Action |
| --- | --- |
| Up / Down (or W / S) | Move |
| Left / Right | Turn |
| Alt + Left / Right, or A / D | Strafe |
| Shift | Run |
| Ctrl | Fire (hold for automatic fire) |
| Space | Open doors, push secret walls, press the elevator switch |
| 1 - 4 | Knife, pistol, machine gun, rotary gun |
| Up / Down, Enter | Choose in menus |
| Enter or Space | Continue after a floor's results |
| Esc | Pause menu (on the title screen: quit) |

## Building

Open `VBWolf.vbp` in the VB6 IDE, or run:

```
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make VBWolf.vbp /out build.log
```

VB6 needs CRLF line endings in `.bas`, `.frm` and `.vbp` files.

## Making levels

Levels are written as text, one character per tile with north at the top; the
legend is at the top of `maps/w00.txt`. Compile a level to the engine's `.map`
format (the format of the browser version's map files) with Python 3:

```
python tools/mkmap.py maps/w00.txt maps/w00.map
```

The compiler checks for one player start, doors placed between two walls, at
most 37 separate areas, and an exit the player can reach (an elevator switch
`E` to use facing east or west, or a `!` tile).

Enemies and patrol routes use tile characters you define at the top of the
level file:

```
@enemy g guard stand S               ; type, stand or patrol, facing N/E/S/W
@enemy z mutant stand S medium       ; only on medium skill and above (or hard)
@enemy q guard stand W ambush        ; ignores noise until it sees you
@enemy d dog patrol W                ; dogs only patrol, and cannot open doors
@turn } E                            ; patrols turn east on this tile
```

Types are `guard`, `officer`, `ss`, `dog` and `mutant`; `@enemy <char> wreck`
places a destroyed robot.

`maps\episode.txt` lists the compiled maps in play order. Write `secret` after
a map to make it a secret floor: it is only reached by using an elevator switch
while standing on a `~` floor tile (put them in a small room hidden behind a
push wall), and play then continues with the floor after the one it was
entered from.

### Generating levels

`tools/genlevel.py` builds larger original floors: rooms and mazes on a grid,
joined by bent corridors with loops, a silver-key wing holding the gold key
(from difficulty 2), the exit elevator behind the gold door, push-wall
closets, patrol loops, and enemies whose number and toughness rise with
`--difficulty` (1 to 5). It checks that each key and the exit can be reached
in order before writing the map. The same seed always gives the same floor,
and the command used is written at the top of the output, so floors 4 to 8 in
`maps/` can be regenerated or tweaked by hand afterwards:

```
python tools/genlevel.py maps/w05.txt --seed 300 --difficulty 4 --cols 4 --rows 4 --name Depot --par 05:00
python tools/mkmap.py maps/w05.txt maps/w05.map
```

Options: `--cols`/`--rows` (up to 4 x 4 cells), `--secrets N` (push-wall
closets), `--secret-elevator` (one closet becomes the way to the secret floor),
`--ceiling`/`--floor` colours.

## Status

The engine core, enemies and combat run: enemies stand guard or patrol, spot
you, open doors, chase, dodge and shoot; you can fight with all four weapons,
take damage, die, respawn with one life fewer, and reach game over, with sound
effects. A title screen, skill choice, pause menu, end-of-floor tally with the
original bonus rules and a three-floor episode tie it together. `PORTING.md`
lists every source file and its state.
