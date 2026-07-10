# STIX Two Math Character Glyphs to SVG (Python Version)

This tool is a Python port of the JavaScript glyph extraction tool. It loads the `STIXTwoMath-Regular.otf` font and extracts vector SVG outlines for mathematical Unicode characters, matching the output structure and coordinates of the JavaScript version character-for-character.

## Prerequisites

- **Python 3.10+**
- **pip** package manager

## Setup

1. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Make sure the font file `STIXTwoMath-Regular.otf` is present (either in this folder or in the sibling JS folder `../stix_two_math_chars_to_svg/`).

## Running the Script

To generate the SVG glyphs in the current folder:
```bash
python stix_two_math_to_svg.py
```

To specify custom font or output paths, use the optional arguments:
```bash
python stix_two_math_to_svg.py --font-path /path/to/STIXTwoMath-Regular.otf --out-dir /path/to/output
```
