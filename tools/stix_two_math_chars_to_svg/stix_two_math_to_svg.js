const opentype = require('opentype.js');
const fs = require('fs');

const font = opentype.loadSync('STIXTwoMath-Regular.otf');
let chars = ["↔","+", "−", "⋅", "×", "÷", "‾", "±", "∓", "√", "∛", "∜", "¢", "∞", "π", "!", "∘", "=", "≠", "<", "≮", ">", "≯", "≤", "≰", "≥", "≱", "≈", "∝", "≅", "≇", "∼", "≁", "∈", "∉", "∋", "⊂", "⊄", "⊆", "⊈", "⊃", "⊅", "⊇", "⊉", "∪", "∩", "↔︎", "→", "∧", "∨", "¬", "⊼", "⊽", "⊕", "⊙", "∀", "∃", "∅", "ℂ", "ℤ", "ℕ", "ℚ", "ℝ", "α", "β", "χ", "δ", "Δ", "γ", "λ", "μ", "ω", "π", "φ", "ρ", "Σ", "τ", "θ", "lim", "→", "∞", "'", "″", "‴", "⁗", "∫", "∬", "∭", "∂", "∑", "∏", "Δ", "∇", "⃗", "∠", "∟", "⊥", "⊥̸", "∥", "∦", "≅", "≇", "∼", "≁", "π", "°", "△", "▱", "◯", "⊙", "⌢", "→", "↔︎", "≈", "±", "℃", "℉", "≪", "≫", "■", "ⓢ", "Ⓢ", "⒨", "⒱", "⒩", "|", "&", "@", "∴", "∵", "⋯", "⋮", "⋱", "…"]

function hexFromGrapheme(character) {
    const codes = Array.from(character, char => char.codePointAt(0).toString(16).padStart(4, '0'));
    return codes.join('+'); 
}

chars.forEach(char => {
    const path = font.getPath(char, 25, 70, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

chars = ["⒨","⋅"]

chars.forEach(char => {
    const path = font.getPath(char, 7, 68, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

chars = ["⋅","‾"]

chars.forEach(char => {
    const path = font.getPath(char, 40, 68, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

chars = ["‾"]

chars.forEach(char => {
    const path = font.getPath(char, 35, 68, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

chars = ["∫","∬","∭","◯","⒱","⒩","△","▱"]

chars.forEach(char => {
    const path = font.getPath(char, 12, 65, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

chars = ["⃗"]

chars.forEach(char => {
    const path = font.getPath(char, 65, 65, 60);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <path d="${path.toPathData()}" />
    </svg>`;
    fs.writeFileSync(`${hexFromGrapheme(char)}.svg`, svg);
});

