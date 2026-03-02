# VimbreWriter / Vim Rewriter

**Vi‑like modal editing for LibreOffice / OpenOffice Writer**

This macro brings the power of Vim’s modal editing to LibreOffice and OpenOffice Writer.  
It provides normal, insert, visual, and visual‑line modes, movements with multipliers, operators (`d`, `c`, `y`, etc.), registers, text objects, a command mode for formatting, and a rich status line – all inside your favourite word processor.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
  - [Basic Workflow](#basic-workflow)
  - [Modes](#modes)
- [Key Reference](#key-reference)
  - [Global Keys](#global-keys)
  - [Mode Switching](#mode-switching)
  - [Movement](#movement)
  - [Editing](#editing)
  - [Visual Mode](#visual-mode)
  - [Command Mode](#command-mode)
  - [Registers](#registers)
  - [Text Objects](#text-objects)
  - [Special Commands](#special-commands)
- [Status Line](#status-line)
- [Development](#development)
  - [Building from Source](#building-from-source)
  - [Testing with Makefile](#testing-with-makefile)
  - [Code Style](#code-style)
- [Known Differences from Vim](#known-differences-from-vim)
- [Credits & License](#credits--license)

---

## Features

- **Modal editing**: Normal, Insert, Visual (characterwise), Visual‑Line.
- **Operators**: `d` (delete), `c` (change), `y` (yank), `p`/`P` (paste) – all work with system clipboard.
- **Registers**: Use `"` followed by a letter to yank into or paste from named registers.
- **Multipliers**: e.g. `5j` moves down five lines, `3dd` deletes three lines.
- **Movements**: character (`h`/`l`), line (`j`/`k`), word (`w`/`b`/`e`), sentence (`)`/`(`), paragraph (`}`/`{`), start of line (`0`), first non‑blank (`^`), end of line (`$`), document start (`gg`) and end (`G`), page jumps (`<count>g`), half‑page scroll (`Ctrl+d`/`Ctrl+u`).
- **Find character**: `f`/`F` (to next/previous occurrence), `t`/`T` (up to next/previous).
- **Text objects**: `i` and `a` with `(`, `{`, `[`, `<`, `'`, `"` to select inside or around delimiters (experimental).
- **Undo/Redo**: `u` and `Ctrl+r` (with multiplier).
- **Find & Replace**: `/` focuses the find bar, `\` opens the search dialog.
- **Command mode**: `:` followed by a letter applies formatting (bold, italic, underline, align, etc.).
- **Rich status line**: shows mode, multiplier, active operator, register, page, word count, file name.

---

## Installation

1. Download the [latest extension file](https://raw.github.com/jonnieey/vimbrewriter/master/dist/vimbrewriter-{latest}oxt).
2. Open it with LibreOffice/OpenOffice, or go to **Tools → Extension Manager → Add** and select the `.oxt` file.
3. Restart LibreOffice/OpenOffice if prompted.
4. Enable VimbreWriter via **Tools → Add-Ons → VimbreWriter**.


### Installation manual
1. Clone the repo
2. cd vimbrewriter
3. run `VIMBREWRITER_VERSION="VERSION_NUMBER" make extension` (produces oxt file in ./dist/)
4. Open it with LibreOffice/OpenOffice, or go to **Tools → Extension Manager → Add** and select the `.oxt` file.
5. Enable VimbreWriter via **Tools → Add-Ons → VimbreWriter**.
 
or
**requires unopkg**
1. Clone the repo
2. cd vimbrewriter
3. run `VIMBREWRITER_VERSION="VERSION_NUMBER" make install` (installs using unopkg)
5. Enable VimbreWriter via **Tools → Add-Ons → VimbreWriter**.

---

## Usage

### Basic Workflow

- After enabling, you start in **Normal Mode** (cursor becomes a block).
- Navigate with `h` (left), `j` (down), `k` (up), `l` (right).
- Press `i` to enter **Insert Mode** and type normally. `Esc` returns to Normal.
- Delete a line with `dd`, copy a line with `yy`, paste with `p` (after cursor) or `P` (before).
- Undo with `u`, redo with `Ctrl+r`.
- Search: press `/`, type your query, and Enter focuses the find bar.
- Replace: press `\` to open the Find & Replace dialog.
- Use multipliers: `3j` moves down three lines, `2dd` deletes two lines.

### Modes

| Mode          | Description                                                                 |
|---------------|-----------------------------------------------------------------------------|
| **NORMAL**    | Default mode. Keys trigger movements and commands. The cursor is a block.  |
| **INSERT**    | Type text normally. All keys insert their characters. `Esc` returns to NORMAL. |
| **VISUAL**    | Character‑wise selection. Movements expand the selection. Operators act on the selected text. |
| **VISUAL_LINE** | Line‑wise selection. The whole line is selected. `j`/`k` expand linewise. |
| **CMD**       | Temporary mode after pressing `:` – next key applies a formatting command.  |

---

## Key Reference

### Global Keys

| Key          | Action                              |
|--------------|-------------------------------------|
| `Esc` or `Ctrl+[` | Return to NORMAL mode.              |

### Mode Switching

| Key | Action                                                                                      |
|-----|---------------------------------------------------------------------------------------------|
| `i` | Enter INSERT mode before the cursor.                                                        |
| `I` | Move cursor to start of line and enter INSERT mode.                                         |
| `a` | Enter INSERT mode after the cursor.                                                         |
| `A` | Move cursor to end of line and enter INSERT mode.                                           |
| `o` | Open a new line below the current line and enter INSERT mode.                               |
| `O` | Open a new line above the current line and enter INSERT mode.                               |
| `v` | Enter VISUAL mode (characterwise).                                                          |
| `V` | Enter VISUAL_LINE mode.                                                                     |

### Movement

Multipliers can be used with all movements, e.g. `5j` moves down five lines.

| Key            | Action                                                                          |
|----------------|---------------------------------------------------------------------------------|
| `h`            | Move left one character.                                                        |
| `l`            | Move right one character.                                                       |
| `j`            | Move down one line.                                                             |
| `k`            | Move up one line.                                                               |
| `0`            | Move to **column 0** (start of line).                                           |
| `^`            | Move to the **first non‑blank** character of the line.                          |
| `$`            | Move to the **end** of the line.                                                |
| `w` / `W`      | Move to the start of the next word. (Both behave the same – next word.)         |
| `b` / `B`      | Move to the start of the previous word. (Both behave the same.)                 |
| `e`            | Move to the end of the current/next word.                                       |
| `)`            | Move to the next sentence.                                                      |
| `(`            | Move to the previous sentence.                                                  |
| `}`            | Move to the next paragraph.                                                     |
| `{`            | Move to the previous paragraph.                                                 |
| `gg`           | Move to the very start of the document.                                         |
| `G`            | Move to the very end of the document.                                           |
| `<count>g`     | Jump to page `<count>` (e.g. `5g` goes to page 5). **Not** a line number jump.  |
| `pagedown`     | Move down half a screen (page down).                                            |
| `Ctrl+shift+>` | Move down half a screen (page down).                                            | 
| `pageup`       | Move up half a screen (page up).                                                |
| `ctrl+shift+u` | Move up half a screen (page up).                                                |
| Arrow keys     | In VISUAL / VISUAL_LINE modes, the arrow keys work like `h`/`j`/`k`/`l`.        |
----------------------------------------------------------------------------------------------------

### Editing

| Key                | Action                                                                                           |
|--------------------|--------------------------------------------------------------------------------------------------|
| `x`                | Delete character under cursor (like `dl`).                                                       |
| `d` + movement     | Delete from cursor to end of movement. Yanked text is saved to clipboard or register.            |
| `dd`               | Delete the current line. Preserves horizontal cursor position if possible.                       |
| `c` + movement     | Delete to end of movement and enter INSERT mode.                                                 |
| `cc`               | Delete the current line and enter INSERT mode.                                                   |
| `y` + movement     | Yank (copy) text to clipboard or register.                                                       |
| `yy`               | Yank the current line.                                                                           |
| `p`                | Paste **after** the cursor.                                                                      |
| `P`                | Paste **before** the cursor.                                                                     |
| `u`                | Undo. Can be prefixed with a multiplier (e.g. `3u` undoes three steps).                          |
| `z`                | Redo.                                                                                             |
| `r`<char>          | Replace the character under the cursor with `<char>`.                                             |
| `D`                | Delete from cursor to end of line.                                                               |
| `C`                | Change from cursor to end of line (delete then INSERT).                                          |
| `S`                | Substitute whole line – same as `cc`.                                                            |
| `s`                | Substitute character – same as `cl` (delete char and enter INSERT).                              |

### Visual Mode

| Key                | Action                                                                                     |
|--------------------|--------------------------------------------------------------------------------------------|
| Movements (`h`,`j`,`k`,`l`, `^`, `$`, `w`, etc.) | Expand the selection.                                        |
| `d`                | Delete the selected text and return to NORMAL.                                             |
| `c`                | Delete the selected text and enter INSERT.                                                 |
| `y`                | Yank the selected text and return to NORMAL.                                               |
| `x`                | Same as `d`.                                                                               |
| `D` / `C`          | Delete/change to end of line from selection start (same as `$` then `d`/`c`).              |
| `s`                | Substitute whole line (same as `cc`).                                                      |
| `Esc`              | Return to NORMAL without changing text.                                                    |

In **VISUAL_LINE** mode, the selection always spans whole lines. The base line is remembered so that `j`/`k` expand the selection correctly.

### Command Mode

Press `:` to enter a temporary command mode. The next key pressed executes a formatting action on the current word (or selection, if one exists). Supported commands:

| Key   | Action                        |
|-------|-------------------------------|
| `b`   | Bold                          |
| `e`   | Align center                  |
| `h`   | Highlight (yellow background) |
| `i`   | Italic                        |
| `j`   | Justify                       |
| `l`   | Align left                    |
| `p`   | Paste (formatted)             |
| `P`   | Paste unformatted             |
| `r`   | Align right                   |
| `s`   | Subscript                     |
| `S`   | Superscript                   |
| `t`   | Strike through                |
| `u`   | Underline                     |
| `w`   | Save document                 |
| `]`   | Increase indent               |
| `[`   | Decrease indent               |
-----------------------------------------

After the command, you return to Normal mode.

### Registers

VimbreWriter supports named registers (letters a–z). Use `"` followed by a register letter before a yank or delete operation to store text in that register. Use the same register with `p`/`P` to paste from it.

- Example: `"ayy` – yank current line into register `a`.
- Example: `"ap` – paste contents of register `a` after cursor.

The unnamed register (`"`) and system clipboard (`*` or `+`) are also supported.

### Text Objects

Experimental support for selecting inside (`i`) or around (`a`) delimiters. After pressing `i` or `a` (in Normal mode), the next key chooses the delimiter:

| Key | Delimiters tried (smart quotes first, then straight) |
|-----|-------------------------------------------------------|
| `(` or `)` | `(` and `)`                                           |
| `{` or `}` | `{` and `}`                                           |
| `[` or `]` | `[` and `]`                                           |
| `<` or `>` | `<` and `>`                                           |
| `.`        | `.` and `.`                                           |
| `,`        | `,` and `,`                                           |
| `'`        | ‘ ’ then `'`                                           |
| `"`        | “ ” then `"`                                           |

Example: `ci(` – change inside parentheses (delete text between `(` and `)` and enter Insert mode).

### Special Commands

| Key | Action                                                                                       |
|-----|----------------------------------------------------------------------------------------------|
| `/` | Focus the find bar (like incremental search). Type your query and press Enter.               |
| `\` | Open the Find & Replace dialog.                                                              |

---

## Status Line

The default status bar is replaced with a custom one showing:
`MODE | Multiplier | special: … | reg: … | modifier: … | page: X/Y | paragraphs: N | Words: M | File: name`
