# Plan: Porting SVG Glyph Extraction and Word Automation to Python + VBA

This plan outlines how to transition the repository's tools from Node.js and PowerShell to a unified **Python + VBA** stack, avoiding the need for JavaScript or PowerShell runtimes.

---

## 1. Scope & Architecture

The new stack simplifies dependencies by using:
* **Python**: Handles all offline, batch-processing tasks (SVG character glyph extraction) and setup tasks (Microsoft Word AutoCorrect shortcuts insertion).
* **VBA (Visual Basic for Applications)**: Handles all interactive, inside-app tasks inside MS Word (keyboard shortcuts navigation, audio triggers, user form editing, and paste intercepts).

### Project Structure Alignment
The Python scripts and configuration files will reside in:
`tools/stix_two_math_chars_to_svg_python/` (mirroring the JavaScript setup).

---

## 2. Python Tasks (To Be Implemented Manually)

### Task A: SVG Glyph Extraction (Replacing `stix_two_math_to_svg.js`)
* **Objective**: Load the `STIXTwoMath-Regular.otf` font and extract vector SVG outlines for mathematical Unicode characters.
* **Implementation Strategy**:
  1. Load the OpenType font using `fontTools.ttLib.TTFont`.
  2. Map character strings to glyph names using the cmap (character map).
  3. Instantiate `fontTools.pens.svgPen.SVGPen` to draw the glyph outlines.
  4. Perform the coordinate flip (transform matrix) to convert from font units (Y-up) to SVG coordinates (Y-down) and scale according to desired font size.
  5. Save the output files named as `<hex_unicode>.svg`.

### Task B: Word AutoCorrect Insertion (Replacing `addmathcodes.ps1`)
* **Objective**: Add math AutoCorrect pairings to MS Word.
* **Implementation Strategy**:
  1. Use the `win32com.client` library (from `pywin32`) to open a connection to Microsoft Word.
  2. Access the `OMathAutoCorrect.Entries` collection.
  3. Loop through the desired shortcode-to-symbol mapping and call `.Add(shortcode, symbol)`.
  4. Implement a backup and restore function for the MS Office AutoCorrect ACL file (`mso0127.acl`).

---

## 3. VBA Tasks (Retained & Integrated)

The current VBA modules inside `OMathNavEnhancements/src/` will remain the core mechanism for Microsoft Word integrations:
1. **[OMathNav.bas](file:///c:/salo/daisy/project/math-a11y/OMathNavEnhancements/src/OMathNav.bas)**: Handles cursor-event polling, shortcut bindings (Alt+[, Alt+], Alt+Shift+[, Alt+Shift+]), sound plays, and forms launch.
2. **[frmMathEdit.frm](file:///c:/salo/daisy/project/math-a11y/OMathNavEnhancements/src/frmMathEdit.frm)**: The dialog form for editing equations.
3. **[fixThenPasteClipboard.bas](file:///c:/salo/daisy/project/math-a11y/tools/fixBadMathML/fixThenPasteClipboard.bas)**: Intercepts pastes and corrects MathML namespaces.
