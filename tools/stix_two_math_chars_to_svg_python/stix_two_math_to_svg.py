# Copyright (c) 2026 MyCompany LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import sys
import math
import argparse
from fontTools.ttLib import TTFont
from fontTools.pens.basePen import BasePen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.boundsPen import BoundsPen
from fontTools.misc.transform import Transform


class ExplicitSVGPen(BasePen):
    def __init__(self, glyphSet, decimal_places=2):
        super().__init__(glyphSet)
        self.decimal_places = decimal_places
        self.commands = []

    def js_round(self, x):
        return math.floor(x + 0.5)

    def float_to_string(self, v):
        factor = 10.0 ** self.decimal_places
        rounded = self.js_round(v * factor) / factor
        if self.js_round(v) == rounded:
            return str(int(self.js_round(v)))
        else:
            return f"{rounded:.{self.decimal_places}f}"

    def pack_values(self, *args):
        s = ""
        for i, v in enumerate(args):
            if v >= 0 and i > 0:
                s += " "
            s += self.float_to_string(v)
        return s

    def _moveTo(self, p):
        self.commands.append("M" + self.pack_values(p[0], p[1]))

    def _lineTo(self, p):
        self.commands.append("L" + self.pack_values(p[0], p[1]))

    def _curveToOne(self, p1, p2, p3):
        self.commands.append("C" + self.pack_values(p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]))

    def _qCurveToOne(self, p1, p2):
        self.commands.append("Q" + self.pack_values(p1[0], p1[1], p2[0], p2[1]))

    def _closePath(self):
        self.commands.append("Z")

    def getCommands(self):
        return "".join(self.commands)

def hex_from_grapheme(character):
    codes = [f"{ord(c):04x}" for c in character]
    return "+".join(codes)

def generate_svg_for_text(text, x, y, font_size, cmap, glypset, units_per_em, auto_center=True):
    if auto_center:
        # Calculate cumulative bounds of the entire text string
        total_width = 0.0
        y_min_total = float('inf')
        y_max_total = float('-inf')
        x_min_total = float('inf')
        x_max_total = float('-inf')
        
        glyphs_to_draw = []
        current_x = 0.0
        
        for char in text:
            codepoint = ord(char)
            glyph_name = cmap.get(codepoint) or '.notdef'
            glyph = glypset[glyph_name]
            
            bp = BoundsPen(glypset)
            glyph.draw(bp)
            if bp.bounds:
                gx_min, gy_min, gx_max, gy_max = bp.bounds
            else:
                gx_min, gy_min, gx_max, gy_max = 0.0, 0.0, 0.0, 0.0
                
            glyphs_to_draw.append((glyph, current_x, gx_min, gy_min, gx_max, gy_max))
            
            x_min_total = min(x_min_total, current_x + gx_min)
            x_max_total = max(x_max_total, current_x + gx_max)
            y_min_total = min(y_min_total, gy_min)
            y_max_total = max(y_max_total, gy_max)
            
            current_x += glyph.width
            
        if not glyphs_to_draw:
            return ""
            
        # Total text dimensions in font design units
        width = x_max_total - x_min_total
        height = y_max_total - y_min_total
        
        # Fit inside 80x80 box
        box_size = 80.0
        if width > 0 and height > 0:
            scale = min(box_size / width, box_size / height)
        elif width > 0:
            scale = box_size / width
        elif height > 0:
            scale = box_size / height
        else:
            scale = box_size / units_per_em
            
        # Avoid scaling tiny elements (like dots, primes, marks) to giant size
        max_scale = 65.0 / units_per_em
        if scale > max_scale:
            scale = max_scale
            
        # Center the cumulative bounding box at (50, 50)
        center_x = (x_min_total + x_max_total) / 2.0
        center_y = (y_min_total + y_max_total) / 2.0
        
        tx = 50.0 - center_x * scale
        ty = 50.0 + center_y * scale
        
        # Draw glyphs
        all_commands = []
        for glyph, offset_x, gx_min, gy_min, gx_max, gy_max in glyphs_to_draw:
            pen = ExplicitSVGPen(glypset)
            # Apply scale, vertical flip, and centering translation
            t = Transform(scale, 0, 0, -scale, tx + offset_x * scale, ty)
            tpen = TransformPen(pen, t)
            glyph.draw(tpen)
            all_commands.append(pen.getCommands())
            
        path_d = "".join(all_commands)
    else:
        # Original manual positioning logic
        scale = font_size / units_per_em
        current_x = x
        all_commands = []
        for char in text:
            codepoint = ord(char)
            glyph_name = cmap.get(codepoint) or '.notdef'
            glyph = glypset[glyph_name]
            pen = ExplicitSVGPen(glypset)
            t = Transform(scale, 0, 0, -scale, current_x, y)
            tpen = TransformPen(pen, t)
            glyph.draw(tpen)
            all_commands.append(pen.getCommands())
            current_x += glyph.width * scale
        path_d = "".join(all_commands)
        
    svg_content = (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">\n'
        f'        <path d="{path_d}" />\n'
        f'    </svg>'
    )
    return svg_content

def main():
    parser = argparse.ArgumentParser(description="STIX Two Math Character Glyphs to SVG Extractor")
    parser.add_argument("--font-path", help="Path to STIXTwoMath-Regular.otf font file")
    parser.add_argument("--out-dir", help="Output directory for generated SVG files")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Resolve font path
    font_path = args.font_path
    if not font_path:
        # Check current/script dir
        font_path = os.path.join(script_dir, "STIXTwoMath-Regular.otf")
        if not os.path.exists(font_path):
            # Check relative path to JS tool setup
            font_path = os.path.abspath(os.path.join(script_dir, "..", "stix_two_math_chars_to_svg", "STIXTwoMath-Regular.otf"))
            
    if not os.path.exists(font_path):
        print(f"Error: STIXTwoMath-Regular.otf font file not found at '{font_path}'.", file=sys.stderr)
        sys.exit(1)

    # 2. Resolve output directory
    out_dir = args.out_dir
    if not out_dir:
        out_dir = script_dir
    os.makedirs(out_dir, exist_ok=True)

    print(f"Loading font from: {font_path}")
    print(f"Writing SVGs to: {out_dir}")

    # 3. Load font
    font = TTFont(font_path)
    cmap = font.getBestCmap()
    glypset = font.getGlyphSet()
    units_per_em = font['head'].unitsPerEm

    # 4. Characters blocks to extract (in original order to preserve overwriting behavior)
    blocks = [
        # List 1
        (["↔","+", "−", "⋅", "×", "÷", "‾", "±", "∓", "√", "∛", "∜", "¢", "∞", "π", "!", "∘", "=", "≠", "<", "≮", ">", "≯", "≤", "≰", "≥", "≱", "≈", "∝", "≅", "≇", "∼", "≁", "∈", "∉", "∋", "⊂", "⊄", "⊆", "⊈", "⊃", "⊅", "⊇", "⊉", "∪", "∩", "↔︎", "→", "∧", "∨", "¬", "⊼", "⊽", "⊕", "⊙", "∀", "∃", "∅", "ℂ", "ℤ", "ℕ", "ℚ", "ℝ", "α", "β", "χ", "δ", "Δ", "γ", "λ", "μ", "ω", "π", "φ", "ρ", "Σ", "τ", "θ", "lim", "→", "∞", "'", "″", "‴", "⁗", "∫", "∬", "∭", "∂", "∑", "∏", "Δ", "∇", "⃗", "∠", "∟", "⊥", "⊥̸", "∥", "∦", "≅", "≇", "∼", "≁", "π", "°", "△", "▱", "◯", "⊙", "⌢", "→", "↔︎", "≈", "±", "℃", "℉", "≪", "≫", "■", "ⓢ", "Ⓢ", "⒨", "⒱", "⒩", "|", "&", "@", "∴", "∵", "⋯", "⋮", "⋱", "…"], 25, 70, 60),
        # List 2
        (["⒨","⋅"], 7, 68, 60),
        # List 3
        (["⋅","‾"], 40, 68, 60),
        # List 4
        (["‾"], 35, 68, 60),
        # List 5
        (["∫","∬","∭","◯","⒱","⒩","△","▱"], 12, 65, 60),
        # List 6
        (["⃗"], 65, 65, 60)
    ]

    total_generated = 0
    for chars, x, y, font_size in blocks:
        for char in chars:
            svg_content = generate_svg_for_text(char, x, y, font_size, cmap, glypset, units_per_em)
            filename = f"{hex_from_grapheme(char)}.svg"
            filepath = os.path.join(out_dir, filename)
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(svg_content)
            total_generated += 1

    print(f"Successfully generated {total_generated} SVG files.")

if __name__ == "__main__":
    main()
