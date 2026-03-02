# VimbreWriter

*(Unfortunately, vimbrewriter, formerly vibreoffice, was in an experimental stage, and the creator [seanyeh](https://github.com/seanyeh) no longer has much time to work on it. So I guess I'm learning Visual basic to pick up the mantel)*

VimbreWriter is an extension for LibreOffice and OpenOffice that brings key bindings from vi/vim to your office suite. It is a continuation of the experimental vibreoffice project, written in StarBasic.

## Features

- **Modes**: Normal, Insert, Visual (character/line), Command.
- **Movements**: `hjkl`, `w`/`b`/`e`, `^`/`$`, `gg`/`G`, `}`/`{`, `)`/`(`.
- **Editing**: `d`/`c`/`y` + motion, `dd`/`cc`/`yy`, `p`/`P`, `u`/`Ctrl+r`.
- **Search**: `/` (find bar), `\` (dialog), `f`/`F`/`t`/`T`.
- **Registers**: `"` + letter to store/recall text.
- **Text objects**: `i`/`a` + `( { [ < ' "` (experimental).
- **Command mode**: `:` + letter for formatting (bold, italic, align, etc.).
- **Status line**: Shows mode, multiplier, register, page, word count.

---

## Installation

### From Release
Download the latest `.oxt` extension file and open it with LibreOffice/OpenOffice, or go to **Tools -> Extension Manager** to add it.

### From Source
You can build and install the extension yourself using the provided Makefile.

**VIMBREWRITER_VERSION="0.0.0" is used for testing, do not use it**

```bash
# Makes and install directly (requires unopkg)
VIMBREWRITER_VERSION="0.0.1" make install

# Build the extension (creates .oxt in dist/)
VIMBREWRITER_VERSION="0.0.1" make extension
```

To enable/disable VimbreWriter, go to **Tools -> Add-Ons -> VimbreWriter**.

### Bindings (Compact)

| Key | Action | Key | Action |
| :--- | :--- | :--- | :--- |
| `h` `j` `k` `l` | Left, Down, Up, Right | `i` `I` `a` `A` | Insert (before/line start/after/line end) |
| `w` `W` `b` `B` | Next/Prev Word | `o` `O` | Open line below/above |
| `e` | End of word | `x` | Delete char |
| `^` `$` | Start/End of line | `r` | Replace char |
| `gg` `G` | Top/Bottom of doc | `u` `C-r` | Undo / Redo |
| `f` `t` | Find char inline | `v` `V` | Visual / Visual Line Mode |
| `{` `}` | Prev/Next Paragraph | `/` | Find |
| `(` `)` | Prev/Next Sentence | `\` | Find & Replace Dialog |
| `d` `c` `y` | Delete, Change, Yank | `p` `P` | Paste after/before |

**Note**: Operators like `d`, `c`, `y` work with motions (e.g., `dw`, `cw`).

## Full Documentation

For a complete key reference, modes, text objects, registers, and development notes, see **[README-COMP.md](README-COMP.md)**.

## Developer Guidelines

### Changing Code
The core logic resides in `@src/vimbrewriter.vbs`. This is a Visual Basic script.
- **Logic**: All key handling, state management (Modes), and document manipulation happen here.
- **Conventions**: Maintain the existing indentation and variable naming style

### Testing
Use the Makefile to streamline the testing loop.

```bash
# Builds the testing extension, kills running LibreOffice instances,
# installs the extension, and opens the test document.
# uninstalls after quitting soffice
make testing
```
**Warning**: `make testing` will kill `soffice.bin`. Save your work in other documents before running.

### Formatting
- run `make lint`
- Keep indentation consistent (spaces/tabs as per existing file).

#### Example
```
VIMBREWRITER_VERSION="1.2.3" make lint ; make install
```

### Known differences/issues
- tag block deletion doesn't work, but I don't believe it's necessary,
  especially given the `iu` and `au` functionality.
- Movement will differ between Vi and LibreOffice. The goal is close enough.
- Unlike vi/vim, movement keys will wrap to the next line
- Due to line wrapping, you may find your cursor move a bit for commands that
  that would otherwise leave you in the same position (such as `dd`) this has
  mostly been mitigated but a side effect is `dd` will not deleted entire
  paragraphs

### I found a bug/How do I help
- If you find a bug or have a suggestion follow these simple steps:
    - check if there's a pre-existing issue
        - If there is and the issue is closed, I didn't want to deal with it
            - either create a pull request or create an issue
        - If there is and it's open, give information on that issue!
    - No issue? Create one!
    - If you know how to program, create a pull request
        - It is best to create an issue first, or explain what the code is doing
    - profit   

## License
MIT © 2014 Sean Yeh, maintained by contributors.
