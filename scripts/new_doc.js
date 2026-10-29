/**
 * new_doc.js — 《岩土工程学报》(CJGE) paper template
 *
 * All formatting (except MathType equations) follows the CJGE layout:
 *   - A4, margins left 1.65 cm / right 1.53 cm / top and bottom 2.15 cm, header 1.5 cm
 *   - Front matter (题名 … Key words) single column; body two columns (8.53 cm, gap 0.74 cm)
 *   - 宋体 / Times New Roman 五号 body, fixed 15.6 pt line spacing, 2-character first-line indent
 *   - 题名 黑体二号, 作者 仿宋五号, 单位 六号, 摘要/关键词 小五 (labels 黑体), English block in TNR
 *   - Headings: 0  引    言 (黑体四号) / 1.1  标题 (黑体五号) / （1）标题 (宋体五号)
 *   - Captions 黑体小五 + English TNR 小五; 三线表 0.75 / 0.5 pt, 小五 table text
 *   - References 小五, fixed 16 pt, 2-character hanging indent, [n] numbering
 *   - Page number in the header (odd / even / first page), no footer
 *
 * Usage:
 *   node scripts/new_doc.js                    # sample content below
 *   node scripts/new_doc.js paper.md out.docx  # convert Markdown
 *
 * Dependencies:
 *   npm install docx temml fast-xml-parser
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const {
  Document, Packer,
  Paragraph, TextRun, Math, MathRun,
  Table, TableRow, TableCell,
  Header, Footer,
  PageNumber, AlignmentType, LineRuleType, HeadingLevel,
  LevelFormat, LevelSuffix, BorderStyle, WidthType, ShadingType, VerticalAlign,
  SectionType, TabStopType,
  TableOfContents, PageBreak, ImageRun,
} = require('docx');

// MathML to docx Math converter
const { mathmlToDocxChildren } = require('./mathml-to-docx');
const temml = require('temml');

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const INPUT_MARKDOWN = process.argv[2] ? path.resolve(process.argv[2]) : null;
const OUTPUT_PATH = process.argv[3]
  ? path.resolve(process.argv[3])
  : (INPUT_MARKDOWN
      ? path.join(path.dirname(INPUT_MARKDOWN), `${path.parse(INPUT_MARKDOWN).name}.docx`)
      : 'output.docx');

// Page / margin (DXA: 1440 = 1 inch, 567 ≈ 1 cm) — 《岩土工程学报》
const PAGE_W        = 11906;   // A4
const PAGE_H        = 16838;
const MARGIN_L      = 935;     // 1.65 cm
const MARGIN_R      = 867;     // 1.53 cm
const MARGIN_TB     = 1219;    // 2.15 cm (版心高 25.4 cm)
const HEADER_DIST   = 850;     // 1.5 cm
const COLUMN_SPACE  = 420;     // 0.74 cm
const CONTENT_W     = PAGE_W - MARGIN_L - MARGIN_R;          // 10104 DXA = 17.82 cm (版心宽)
const COLUMN_W      = (CONTENT_W - COLUMN_SPACE) / 2;        // 4842 DXA = 8.54 cm (栏宽)
const LATIN         = 'Times New Roman';

// Three-line table rules: top/bottom 0.75 pt, header rule 0.5 pt
const THICK = { style: BorderStyle.SINGLE, size: 6, color: '000000' }; // 0.75 pt
const THIN  = { style: BorderStyle.SINGLE, size: 4, color: '000000' }; // 0.5 pt
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// ISSUE 3 FIX: Manual heading counters for synchronized numbering
let currentChapter = 0;
let currentSection = 0;
let currentSubsection = 0;

// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 5/7 FIX: Inline Math Detection and Conversion
// ─────────────────────────────────────────────────────────────────────────────

// Unicode math symbols to LaTeX mapping
const UNICODE_TO_LATEX = {
  // Greek letters
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ε': '\\varepsilon',
  'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta', 'ι': '\\iota', 'κ': '\\kappa',
  'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi',
  'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon', 'φ': '\\phi',
  'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
  'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi',
  'Π': '\\Pi', 'Σ': '\\Sigma', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
  // Subscript digits
  '₀': '_0', '₁': '_1', '₂': '_2', '₃': '_3', '₄': '_4',
  '₅': '_5', '₆': '_6', '₇': '_7', '₈': '_8', '₉': '_9',
  'ₙ': '_n', 'ₓ': '_x', 'ᵢ': '_i', 'ₜ': '_t', 'ₛ': '_s',
  // Superscript
  '⁰': '^0', '¹': '^1', '²': '^2', '³': '^3', '⁴': '^4',
  '⁵': '^5', '⁶': '^6', '⁷': '^7', '⁸': '^8', '⁹': '^9',
  'ⁿ': '^n', 'ⁱ': '^i',
  // Special symbols
  '∞': '\\infty', '∑': '\\sum', '∏': '\\prod', '∫': '\\int',
  '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '≈': '\\approx',
  '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow',
  '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊃': '\\supset',
  '∀': '\\forall', '∃': '\\exists', '∧': '\\land', '∨': '\\lor',
  '×': '\\times', '÷': '\\div', '±': '\\pm', '∓': '\\mp',
  '·': '\\cdot', '…': '\\ldots', '⋯': '\\cdots',
  '′': "'", '″': "''",
  '⟨': '\\langle', '⟩': '\\rangle',
  // Superscript letter (for π*, Q*, etc.)
  '*': '^*',
};

/**
 * Convert text with Unicode math symbols to LaTeX format
 * @param {string} text - Text with Unicode math symbols
 * @returns {string} LaTeX formatted text
 */
function unicodeToLatex(text) {
  let result = text;
  for (const [unicode, latex] of Object.entries(UNICODE_TO_LATEX)) {
    result = result.split(unicode).join(latex);
  }
  return result;
}

/**
 * Detect if text contains math content (needs formula editor rendering)
 * ISSUE 7 FIX: Strict detection - don't match plain numbers or English words
 * @param {string} text - Text to detect
 * @returns {boolean}
 */
function containsMath(text) {
  // Detect Greek letters
  if (/[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text)) return true;
  // Detect Unicode subscript/superscript characters
  if (/[₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]/.test(text)) return true;
  // Detect math operators and special symbols
  if (/[∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩]/.test(text)) return true;
  // Detect starred symbols like π*, Q* (but not plain words like Agent)
  if (/[A-Z]\*/.test(text)) return true;
  // Detect $...$ LaTeX delimiters
  if (/\$[^$]+\$/.test(text)) return true;
  return false;
}

/**
 * ISSUE 8 FIX: Detect if text contains citation [n] format
 */
function containsCitation(text) {
  return /\[\d+(?:\s*[-–,，]\s*\d+)*\]/.test(text);
}

/**
 * ISSUE 7 FIX: Parse text into TextRun and Math mixed array (for inline formulas)
 * Strict regex - only matches actual math content, not plain numbers or words
 * @param {string} text - Input text
 * @returns {Array} Array of TextRun and Math objects
 */
function parseInlineContent(text) {
  const children = [];
  
  // Regex matching math content blocks (in priority order):
  // 1. $...$  explicit LaTeX (highest priority)
  // 2. Function form Q(s,a) V(s) Rₓ(a) etc. (only with subscripts or specific single letters)
  // 3. Greek letters (alone or with subscripts)
  // 4. Variables with subscripts like Qₙ xₙ αₙ etc.
  // 5. Starred symbols like π* Q* (single letter only)
  // 
  // NOTE: Does NOT match plain English words like Agent, Watkins, Dayan
  // Does NOT match plain numbers like 1992, 500
  // Does NOT match plain parentheses expressions like (1992)
  
  const mathPattern = /\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;
  
  let lastIndex = 0;
  let match;
  
  while ((match = mathPattern.exec(text)) !== null) {
    // Add plain text before match
    if (match.index > lastIndex) {
      const plainText = text.slice(lastIndex, match.index);
      if (plainText) {
        children.push(new TextRun(plainText));
      }
    }
    
    // Get matched math content
    const mathContent = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
    if (mathContent) {
      // Convert Unicode to LaTeX and create Math object
      const latex = unicodeToLatex(mathContent);
      try {
        const mathml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
        const mathChildren = mathmlToDocxChildren(mathml);
        if (mathChildren && mathChildren.length) {
          children.push(new Math({ children: mathChildren }));
        } else {
          // fallback
          children.push(new Math({ children: [new MathRun(mathContent)] }));
        }
      } catch (e) {
        // Parse failed, use MathRun to display original text
        children.push(new Math({ children: [new MathRun(mathContent)] }));
      }
    }
    
    lastIndex = match.index + match[0].length;
  }
  
  // Add remaining plain text
  if (lastIndex < text.length) {
    children.push(new TextRun(text.slice(lastIndex)));
  }
  
  // If no math content matched, return plain text
  if (children.length === 0) {
    children.push(new TextRun(text));
  }
  
  return children;
}

/**
 * ISSUE 8 FIX: Parse text with math content and citations
 * Citations [n] are converted to superscript format
 * @param {string} text - Input text
 * @returns {Array} Array of TextRun and Math objects
 */
function parseInlineContentWithCitations(text) {
  const children = [];
  
  // Combined regex: match math content or citations
  // Citations [n] become superscript
  // Math content becomes Math objects
  const combinedPattern = /(\[\d+(?:\s*[-–,，]\s*\d+)*\])|\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;
  
  let lastIndex = 0;
  let match;
  
  while ((match = combinedPattern.exec(text)) !== null) {
    // Add plain text before match
    if (match.index > lastIndex) {
      const plainText = text.slice(lastIndex, match.index);
      if (plainText) {
        children.push(new TextRun(plainText));
      }
    }
    
    if (match[1]) {
      // Citation [n] - convert to superscript
      children.push(new TextRun({
        text: match[1],
        superScript: true,
      }));
    } else {
      // Math content
      const mathContent = match[2] || match[3] || match[4] || match[5] || match[6] || match[7];
      if (mathContent) {
        const latex = unicodeToLatex(mathContent);
        try {
          const mathml = temml.renderToString(latex, { displayMode: false, throwOnError: false });
          const mathChildren = mathmlToDocxChildren(mathml);
          if (mathChildren && mathChildren.length) {
            children.push(new Math({ children: mathChildren }));
          } else {
            children.push(new Math({ children: [new MathRun(mathContent)] }));
          }
        } catch (e) {
          children.push(new Math({ children: [new MathRun(mathContent)] }));
        }
      }
    }
    
    lastIndex = match.index + match[0].length;
  }
  
  // Add remaining plain text
  if (lastIndex < text.length) {
    children.push(new TextRun(text.slice(lastIndex)));
  }
  
  if (children.length === 0) {
    children.push(new TextRun(text));
  }
  
  return children;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: body paragraph (首行缩进 2 字符, 单倍行距) - supports inline formulas
// ─────────────────────────────────────────────────────────────────────────────

function makeBodyParagraph(text, overrides = {}) {
  const children = containsMath(text) || containsCitation(text)
    ? parseInlineContentWithCitations(text)
    : [new TextRun(text)];

  // Inline equations must not be clipped by the fixed line spacing: use "at least" instead.
  const spacing = children.some((c) => c instanceof Math) ? { spacing: { line: 312, lineRule: LineRuleType.AT_LEAST } } : {};
  return new Paragraph({
    ...spacing,
    ...overrides,
    children,
  });
}

function body(text) {
  return makeBodyParagraph(text);
}

function bodyIndented(text, left = 0) {
  return makeBodyParagraph(text, {
    indent: { left, firstLine: 0 },
  });
}

function bodyNoIndent(text, overrides = {}) {
  return makeBodyParagraph(text, {
    indent: { firstLine: 0 },
    ...overrides,
  });
}

// Paragraph with multiple TextRuns
function bodyMulti(runs) {
  return new Paragraph({
    children: runs,
  });
}

const INLINE_LATEX_TO_UNICODE = {
  '\\alpha': 'α',
  '\\beta': 'β',
  '\\gamma': 'γ',
  '\\delta': 'δ',
  '\\epsilon': 'ε',
  '\\varepsilon': 'ε',
  '\\zeta': 'ζ',
  '\\eta': 'η',
  '\\theta': 'θ',
  '\\vartheta': 'ϑ',
  '\\iota': 'ι',
  '\\kappa': 'κ',
  '\\lambda': 'λ',
  '\\mu': 'μ',
  '\\nu': 'ν',
  '\\xi': 'ξ',
  '\\pi': 'π',
  '\\rho': 'ρ',
  '\\sigma': 'σ',
  '\\tau': 'τ',
  '\\upsilon': 'υ',
  '\\phi': 'φ',
  '\\varphi': 'φ',
  '\\chi': 'χ',
  '\\psi': 'ψ',
  '\\omega': 'ω',
  '\\Gamma': 'Γ',
  '\\Delta': 'Δ',
  '\\Theta': 'Θ',
  '\\Lambda': 'Λ',
  '\\Xi': 'Ξ',
  '\\Pi': 'Π',
  '\\Sigma': 'Σ',
  '\\Phi': 'Φ',
  '\\Psi': 'Ψ',
  '\\Omega': 'Ω',
  '\\langle': '⟨',
  '\\rangle': '⟩',
  '\\leq': '≤',
  '\\geq': '≥',
  '\\neq': '≠',
  '\\to': '→',
  '\\rightarrow': '→',
  '\\leftarrow': '←',
  '\\cdot': '·',
  '\\times': '×',
  '\\infty': '∞',
};

const SUBSCRIPT_CHARS = {
  '0': '₀',
  '1': '₁',
  '2': '₂',
  '3': '₃',
  '4': '₄',
  '5': '₅',
  '6': '₆',
  '7': '₇',
  '8': '₈',
  '9': '₉',
  '+': '₊',
  '-': '₋',
  '=': '₌',
  '(': '₍',
  ')': '₎',
  'a': 'ₐ',
  'e': 'ₑ',
  'h': 'ₕ',
  'i': 'ᵢ',
  'j': 'ⱼ',
  'k': 'ₖ',
  'l': 'ₗ',
  'm': 'ₘ',
  'n': 'ₙ',
  'o': 'ₒ',
  'p': 'ₚ',
  'r': 'ᵣ',
  's': 'ₛ',
  't': 'ₜ',
  'u': 'ᵤ',
  'v': 'ᵥ',
  'x': 'ₓ',
  'β': 'ᵦ',
  'γ': 'ᵧ',
  'ρ': 'ᵨ',
  'φ': 'ᵩ',
  'χ': 'ᵪ',
};

const SUPERSCRIPT_CHARS = {
  '0': '⁰',
  '1': '¹',
  '2': '²',
  '3': '³',
  '4': '⁴',
  '5': '⁵',
  '6': '⁶',
  '7': '⁷',
  '8': '⁸',
  '9': '⁹',
  '+': '⁺',
  '-': '⁻',
  '=': '⁼',
  '(': '⁽',
  ')': '⁾',
  'i': 'ⁱ',
  'n': 'ⁿ',
};

function convertCharsWithMap(text, mapping) {
  return [...text].map((char) => mapping[char] || char).join('');
}

function unwrapLatexTextCommands(text) {
  let result = text;
  let previous = '';

  while (result !== previous) {
    previous = result;
    result = result.replace(
      /\\(?:mathrm|mathbf|mathit|text|operatorname)\{([^{}]*)\}/g,
      '$1'
    );
  }

  return result;
}

function simplifyInlineLatex(latex) {
  let text = latex.trim();
  text = unwrapLatexTextCommands(text);

  for (const [command, replacement] of Object.entries(INLINE_LATEX_TO_UNICODE)) {
    text = text.split(command).join(replacement);
  }

  text = text
    .replace(/\\left/g, '')
    .replace(/\\right/g, '')
    .replace(/\\,/g, ' ')
    .replace(/\\;/g, ' ')
    .replace(/\\!/g, '')
    .replace(/\\ /g, ' ')
    .replace(/\\\{/g, '{')
    .replace(/\\\}/g, '}');

  text = text
    .replace(/_\{([^{}]+)\}/g, (_, group) => convertCharsWithMap(group, SUBSCRIPT_CHARS))
    .replace(/_([A-Za-z0-9+\-=()α-ωΑ-Ω])/g, (_, group) => convertCharsWithMap(group, SUBSCRIPT_CHARS))
    .replace(/\^\{([^{}]+)\}/g, (_, group) => group === '*' ? '*' : convertCharsWithMap(group, SUPERSCRIPT_CHARS))
    .replace(/\^([A-Za-z0-9+\-=()α-ωΑ-Ω*])/g, (_, group) => group === '*' ? '*' : convertCharsWithMap(group, SUPERSCRIPT_CHARS));

  text = text
    .replace(/[{}]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  return text;
}

function buildInlineMathRuns(text, options = {}) {
  const {
    bold = false,
    normalEastAsia = 'SimHei',
    normalItalic = false,
    mathItalic = true,
  } = options;

  const runs = [];
  const pattern = /\$([^$]+)\$/g;
  let lastIndex = 0;
  let match;

  const normalRun = (value) => new TextRun({
    text: value,
    bold,
    italics: normalItalic,
    font: { ascii: LATIN, eastAsia: normalEastAsia, hAnsi: LATIN },
  });

  const mathRun = (value) => new TextRun({
    text: simplifyInlineLatex(value),
    bold,
    italics: mathItalic,
    font: { ascii: LATIN, eastAsia: normalEastAsia, hAnsi: LATIN },
  });

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(normalRun(text.slice(lastIndex, match.index)));
    }

    if (match[1].trim()) {
      runs.push(mathRun(match[1]));
    }

    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    runs.push(normalRun(text.slice(lastIndex)));
  }

  if (runs.length === 0) {
    runs.push(normalRun(text));
  }

  return runs;
}

function plainHeading(level, text) {
  return new Paragraph({
    heading: level,
    indent: { firstLine: 0 },
    children: buildInlineMathRuns(text, { bold: false, normalEastAsia: 'SimHei' }),
  });
}

function centeredText(text, options = {}) {
  const {
    size = 24,
    bold = false,
    spacing = { before: 0, after: 120 },
  } = options;

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing,
    indent: { firstLine: 0 },
    children: [new TextRun({
      text,
      bold,
      size,
      font: { ascii: LATIN, eastAsia: 'SimHei', hAnsi: LATIN },
    })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: containsMath(text) || containsCitation(text)
      ? parseInlineContentWithCitations(text)
      : [new TextRun(text)],
  });
}

function parseHtmlTable(html) {
  const rowMatches = [...html.matchAll(/<tr>(.*?)<\/tr>/gsi)];
  const rows = rowMatches.map((rowMatch) => {
    const cellMatches = [...rowMatch[1].matchAll(/<t[dh]>(.*?)<\/t[dh]>/gsi)];
    return cellMatches.map((cellMatch) =>
      cellMatch[1]
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .trim()
    );
  }).filter((row) => row.length);

  if (rows.length < 2) {
    throw new Error('HTML table requires at least one header row and one body row');
  }

  return {
    headers: rows[0],
    rows: rows.slice(1),
  };
}

function normalizeCaptionText(text) {
  return text
    .replace(/\s+/g, ' ')
    .replace(/^([\u56fe\u8868])\s*(\d+)\s*/, '$1$2 ')   // CJGE: "\u56fe1 \u9898\u540d" / "\u88681 \u9898\u540d"
    .trim();
}

function getImageType(imagePath) {
  const ext = path.extname(imagePath).toLowerCase();
  if (ext === '.jpg') return 'jpg';
  if (ext === '.jpeg') return 'jpeg';
  if (ext === '.png') return 'png';
  if (ext === '.gif') return 'gif';
  if (ext === '.bmp') return 'bmp';
  throw new Error(`Unsupported image format: ${imagePath}`);
}

function getImageDimensions(buffer, imagePath) {
  const ext = path.extname(imagePath).toLowerCase();

  if (ext === '.png') {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20),
    };
  }

  if (ext === '.jpg' || ext === '.jpeg') {
    let offset = 2;
    while (offset < buffer.length) {
      if (buffer[offset] !== 0xFF) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      const length = buffer.readUInt16BE(offset + 2);
      if (marker >= 0xC0 && marker <= 0xC3) {
        return {
          height: buffer.readUInt16BE(offset + 5),
          width: buffer.readUInt16BE(offset + 7),
        };
      }
      offset += 2 + length;
    }
  }

  throw new Error(`Unable to read image dimensions: ${imagePath}`);
}

function imageBlock(imagePath) {
  const data = fs.readFileSync(imagePath);
  const { width, height } = getImageDimensions(data, imagePath);
  const maxWidth = 302;   // 8.0 cm single-column figure (px at 96 dpi); use 643 px for 17.0 cm
  const maxHeight = 360;
  const scale = globalThis.Math.min(maxWidth / width, maxHeight / height, 1);

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    spacing: { before: 60, after: 0, line: 240, lineRule: LineRuleType.AUTO },  // never clip pictures
    children: [new ImageRun({
      type: getImageType(imagePath),
      data,
      transformation: {
        width: globalThis.Math.round(width * scale),
        height: globalThis.Math.round(height * scale),
      },
      altText: {
        title: path.basename(imagePath),
        description: path.basename(imagePath),
        name: path.basename(imagePath),
      },
    })],
  });
}

function parseFormulaBlock(lines, startIndex) {
  const formulaLines = [];
  let i = startIndex + 1;
  while (i < lines.length && lines[i].trim() !== '$$') {
    formulaLines.push(lines[i]);
    i += 1;
  }

  const rawLatex = formulaLines.join(' ').trim();
  const tagMatch = rawLatex.match(/\\tag\{([^}]+)\}\s*$/);
  const number = tagMatch ? tagMatch[1] : '';
  const latex = tagMatch ? rawLatex.slice(0, tagMatch.index).trim() : rawLatex;

  return {
    block: formula(latex, number),
    nextIndex: i,
  };
}

function buildContentFromMarkdown(markdownPath) {
  // Returns { front, body }: front matter (通栏, single column) and body (双栏).
  // Markdown conventions: "# 题名", then 作者 / 单位 lines, "摘  要：…", "关键词：…",
  // "中图分类号：…", "作者简介：…", "## English title" block with "Abstract: …" / "Key words: …",
  // then "## 0 引言"-style headings (## / ### / ####), "参考文献" and [n] entries.
  const raw = fs.readFileSync(markdownPath, 'utf8');
  const lines = raw.split(/\r?\n/);
  const front = [];
  const content = [];
  const baseDir = path.dirname(markdownPath);
  let i = 0;
  let inReferences = false;
  let inFront = true;
  let frontCn = 0;   // plain lines seen before 摘要: 0 title, 1 authors, 2+ affiliation
  let frontEn = -1;  // plain lines after 作者简介: 0 title, 1 authors, 2+ affiliation
  const out = () => (inFront ? front : content);

  const flushParagraph = (paragraphLines) => {
    if (!paragraphLines.length) return;
    const rawText = paragraphLines.join(' ').replace(/\s+/g, ' ').trim();
    if (!rawText) return;
    if (inFront) {
      if (/^摘\s*要\s*[:：]/.test(rawText)) { front.push(abstractCn(rawText.replace(/^摘\s*要\s*[:：]\s*/, ''))); return; }
      if (/^关\s*键\s*词\s*[:：]/.test(rawText)) { front.push(keywordsCn(rawText.replace(/^关\s*键\s*词\s*[:：]\s*/, ''))); return; }
      if (/^(中图分类号|文献标识码|文章编号)\s*[:：]/.test(rawText)) { front.push(clcLine(rawText)); return; }
      if (/^作者简介\s*[:：]/.test(rawText)) { front.push(bioCn(rawText.replace(/^作者简介\s*[:：]\s*/, ''))); frontEn = 0; return; }
      if (/^Abstract\s*[:：]/i.test(rawText)) { front.push(abstractEn(rawText.replace(/^Abstract\s*[:：]\s*/i, ''))); return; }
      if (/^Key\s*words\s*[:：]/i.test(rawText)) { front.push(keywordsEn(rawText.replace(/^Key\s*words\s*[:：]\s*/i, ''))); return; }
      if (frontEn >= 0) {
        front.push(frontEn === 0 ? titleEn(rawText) : frontEn === 1 ? authorsEn(rawText) : affiliationEn(rawText));
        frontEn += 1;
        return;
      }
      if (frontCn >= 1) {
        front.push(frontCn === 1 ? authorsCn(rawText) : affiliationCn(rawText));
        frontCn += 1;
        return;
      }
    }
    if (/^(利益冲突声明|作者贡献)/.test(rawText)) { content.push(statement(rawText)); return; }
    content.push(body(rawText));
  };

  while (i < lines.length) {
    const rawLine = lines[i];
    const line = rawLine.trim();

    if (!line || line === '---') {
      i += 1;
      continue;
    }

    if (line.startsWith('# ')) {
      front.push(titleCn(line.slice(2).trim()));
      frontCn = 1;
      i += 1;
      continue;
    }

    if (inFront && /^##\s+/.test(line) && frontEn >= 0 && !/^##\s+\d/.test(line)) {
      front.push(titleEn(line.replace(/^##\s+/, '').trim()));
      frontEn = 1;
      i += 1;
      continue;
    }

    if (line === '$$') {
      const { block, nextIndex } = parseFormulaBlock(lines, i);
      out().push(block);
      i = nextIndex + 1;
      continue;
    }

    const imageMatch = line.match(/^!\[[^\]]*]\(([^)]+)\)$/);
    if (imageMatch) {
      const imagePath = path.resolve(baseDir, imageMatch[1]);
      out().push(imageBlock(imagePath));
      let j = i + 1;
      while (j < lines.length && !lines[j].trim()) j += 1;
      if (j < lines.length && /^图/.test(lines[j].trim())) {
        let en = '';
        let k = j + 1;
        while (k < lines.length && !lines[k].trim()) k += 1;
        if (k < lines.length && /^Fig/i.test(lines[k].trim())) en = lines[k].trim();
        out().push(...figCaption(normalizeCaptionText(lines[j].trim()), en));
        i = (en ? k : j) + 1;
      } else {
        i += 1;
      }
      continue;
    }

    if (line.startsWith('<table>')) {
      const { headers, rows } = parseHtmlTable(line);
      const target = out();
      // Caption lines just above the table: "表1 …" and optionally "Table 1 …".
      let cn = '';
      let en = '';
      if (target.length && i > 0) {
        let k = i - 1;
        while (k > 0 && !lines[k].trim()) k -= 1;
        if (/^Table/i.test(lines[k].trim())) {
          en = lines[k].trim();
          k -= 1;
          while (k > 0 && !lines[k].trim()) k -= 1;
        }
        if (/^表/.test(lines[k].trim())) cn = lines[k].trim();
        if (cn) target.splice(target.length - (en ? 2 : 1), en ? 2 : 1);
      }
      if (cn) target.push(...tableCaption(normalizeCaptionText(cn), en));
      target.push(threeLineTable(headers, rows));
      i += 1;
      continue;
    }

    if (/^#{0,2}\s*参考文献\s*[:：]?\s*$/.test(line)) {
      inReferences = true;
      inFront = false;
      content.push(refHeading());
      i += 1;
      continue;
    }

    if (/^##\s+/.test(line)) {
      inFront = false;
      content.push(h1Manual(line.replace(/^##\s+/, '').trim()));
      inReferences = false;
      i += 1;
      continue;
    }

    if (/^###\s+/.test(line)) {
      content.push(h2Manual(line.replace(/^###\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^####\s+/.test(line)) {
      content.push(h3Manual(line.replace(/^####\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^\[\d+\]\s+/.test(line) && inReferences) {
      content.push(ref(line.replace(/^\[\d+\]\s+/, '').trim()));
      i += 1;
      continue;
    }

    if (/^- /.test(line)) {
      out().push(bullet(line.replace(/^- /, '').trim()));
      i += 1;
      continue;
    }

    const paragraphLines = [rawLine];
    let j = i + 1;
    while (j < lines.length) {
      const nextTrimmed = lines[j].trim();
      if (
        !nextTrimmed ||
        nextTrimmed === '$$' ||
        nextTrimmed === '---' ||
        /^#/.test(nextTrimmed) ||
        /^!\[/.test(nextTrimmed) ||
        nextTrimmed.startsWith('<table>') ||
        /^- /.test(nextTrimmed) ||
        (/^\[\d+\]\s+/.test(nextTrimmed) && inReferences)
      ) {
        break;
      }
      paragraphLines.push(lines[j]);
      j += 1;
    }
    flushParagraph(paragraphLines);
    i = j;
  }

  return { front, body: content };
}

// ─────────────────────────────────────────────────────────────────────────────
// Headings (《岩土工程学报》): 一级 "0  引    言" 黑体四号 · 二级 "1.1  标题" 黑体五号 ·
// 三级 "（1）标题" 宋体五号缩进 2 字符 · 四级 "a）标题。" 与正文接排 (use body()).
// The number and the title are separated by two spaces; the introduction is numbered 0.
// ─────────────────────────────────────────────────────────────────────────────

/** Reset heading counters (call at start of document); the first h1() is "0" (引言). */
function resetHeadingCounters() {
  currentChapter = -1;
  currentSection = 0;
  currentSubsection = 0;
}

function headingRuns(text, eastAsia = 'SimHei') {
  return buildInlineMathRuns(text, { bold: false, normalEastAsia: eastAsia });
}

/** H1 with the number already in the text, e.g. "0  引    言" */
function h1Manual(text) {
  const m = text.match(/^(\d+)\s/);
  if (m) currentChapter = Number(m[1]);
  else currentChapter++;
  currentSection = 0;
  currentSubsection = 0;
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: headingRuns(text) });
}

/** H1 with automatic numbering: 0, 1, 2 … → "1  试验设计" */
function h1(text) {
  currentChapter++;
  currentSection = 0;
  currentSubsection = 0;
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: headingRuns(`${currentChapter}  ${text}`) });
}

/** H2 with the number already in the text, e.g. "1.1  试验装置" */
function h2Manual(text) {
  currentSection++;
  currentSubsection = 0;
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: headingRuns(text) });
}

/** H2 with automatic numbering → "1.1  试验装置" */
function h2(text) {
  currentSection++;
  currentSubsection = 0;
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: headingRuns(`${currentChapter}.${currentSection}  ${text}`),
  });
}

/** H3 with the number already in the text, e.g. "（1）级配影响" */
function h3Manual(text) {
  currentSubsection++;
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: headingRuns(text, 'SimSun') });
}

/** H3 with automatic numbering → "（1）级配影响" (own line, no end punctuation) */
function h3(text) {
  currentSubsection++;
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: headingRuns(`（${currentSubsection}）${text}`, 'SimSun') });
}

/** "参考文献：" 黑体五号, flush left */
function refHeading(text = '参考文献：') {
  return new Paragraph({ style: 'RefHeading', children: [new TextRun(text)] });
}

// ─────────────────────────────────────────────────────────────────────────────
// Front matter (通栏). Labels: 中文黑体, English bold.
// ─────────────────────────────────────────────────────────────────────────────

const labelRun = (text, eastAsia = 'SimHei', bold = false) =>
  new TextRun({ text, bold, font: { ascii: LATIN, hAnsi: LATIN, eastAsia } });

function titleCn(text) {
  return new Paragraph({ style: 'Title', children: [new TextRun(text)] });
}

/** e.g. authorsCn('简富献', ['1', '2'], '蔡正银', ['*1']) or authorsCn('简富献1, 2，蔡正银*1') */
function authorsCn(...args) {
  return new Paragraph({ style: 'AuthorsCn', children: authorRuns(args, '，') });
}

function authorsEn(...args) {
  return new Paragraph({ style: 'AuthorsEn', children: authorRuns(args, ', ') });
}

function authorRuns(args, sep) {
  if (args.length === 1) return [new TextRun(args[0])];
  const runs = [];
  for (let k = 0; k < args.length; k += 2) {
    if (k) runs.push(new TextRun(sep));
    runs.push(new TextRun(args[k]));
    const marks = args[k + 1] || [];
    if (marks.length) runs.push(new TextRun({ text: marks.join(', '), superScript: true }));
  }
  return runs;
}

function affiliationCn(text) {
  return new Paragraph({ style: 'AffiliationCn', children: [new TextRun(text)] });
}

function affiliationEn(text) {
  return new Paragraph({ style: 'AffiliationEn', children: [new TextRun(text)] });
}

function abstractCn(text) {
  return new Paragraph({ style: 'AbstractCn', children: [labelRun('摘  要：'), ...parseInlineContentWithCitations(text)] });
}

function keywordsCn(text) {
  return new Paragraph({ style: 'AbstractCn', children: [labelRun('关键词：'), new TextRun(text)] });
}

/** "中图分类号：TU42　文献标识码：A　文章编号：…" — labels in 黑体 */
function clcLine(text) {
  const parts = text.split(/(中图分类号\s*[:：]|文献标识码\s*[:：]|文章编号\s*[:：])/).filter(Boolean);
  return new Paragraph({
    style: 'FrontNote',
    children: parts.map((p) => (/^(中图分类号|文献标识码|文章编号)/.test(p) ? labelRun(p) : new TextRun(p))),
  });
}

function bioCn(text) {
  return new Paragraph({ style: 'FrontNote', children: [labelRun('作者简介：'), new TextRun(text)] });
}

function titleEn(text) {
  return new Paragraph({ style: 'TitleEn', children: [new TextRun(text)] });
}

function abstractEn(text) {
  return new Paragraph({ style: 'AbstractEn', children: [labelRun('Abstract', LATIN, true), new TextRun(': '), new TextRun(text)] });
}

function keywordsEn(text) {
  return new Paragraph({ style: 'AbstractEn', children: [labelRun('Key words', LATIN, true), new TextRun(': '), new TextRun(text)] });
}

/** 文末声明: statement('利益冲突声明/Conflict of Interests：', '所有作者声明…All authors …') */
function statement(label, text) {
  if (text === undefined) {
    const m = label.match(/^((利益冲突声明|作者贡献)\s*\/[^：:]*[：:])(.*)$/);
    if (!m) return new Paragraph({ style: 'Statement', children: [new TextRun(label)] });
    [label, text] = [m[1], m[3]];
  }
  const cn = label.match(/^[^/]+/)[0];
  return new Paragraph({
    style: 'Statement',
    children: [labelRun(cn, 'SimHei', true), labelRun(label.slice(cn.length), 'SimSun', true), new TextRun(text)],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Captions: Chinese 黑体小五 + English TNR 小五, centred; figure caption below the figure,
// table caption above the table. Pass "图1 题名" / "Fig. 1 Title" (numbers are kept as written).
// ─────────────────────────────────────────────────────────────────────────────

function figCaption(label, labelEn = '') {
  const out = [new Paragraph({ style: 'FigureCaption', children: headingRuns(label) })];
  if (labelEn) out.push(new Paragraph({ style: 'FigureCaptionEn', children: [new TextRun(labelEn)] }));
  return out;
}

function tableCaption(label, labelEn = '') {
  const out = [new Paragraph({ style: 'TableCaption', children: headingRuns(label) })];
  if (labelEn) out.push(new Paragraph({ style: 'TableCaptionEn', children: [new TextRun(labelEn)] }));
  return out;
}

/** 分图编号 "(a) 说明", centred below the sub-figure */
function subCaption(text) {
  return new Paragraph({ style: 'SubCaption', children: [new TextRun(text)] });
}

// ─────────────────────────────────────────────────────────────────────────────
// 三线表: top and bottom rules 0.75 pt, header rule 0.5 pt, no vertical rules, no fill,
// 小五 text. Default width = one column (8.5 cm); pass colWidths summing to CONTENT_W for a
// full-width (双栏) table.
// ─────────────────────────────────────────────────────────────────────────────

function threeLineTable(headers, rows, colWidths) {
  const n = headers.length;

  if (!colWidths) {
    const w = globalThis.Math.floor(COLUMN_W / n);
    colWidths = Array(n).fill(w);
    colWidths[n - 1] = COLUMN_W - w * (n - 1);
  }
  if (colWidths.length !== n) throw new Error('colWidths length must match headers length');
  const tableW = colWidths.reduce((sum, w) => sum + w, 0);

  const cellOf = (text, w, borders) => {
    const runs = containsMath(text) ? parseInlineContent(text) : [new TextRun(text)];
    return new TableCell({
    width:   { size: w, type: WidthType.DXA },
    borders,
    margins: { top: 0, bottom: 0, left: 57, right: 57 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      style: 'TableText',
      ...(runs.some((c) => c instanceof Math) ? { spacing: { line: 280, lineRule: LineRuleType.AT_LEAST } } : {}),
      children: runs,
    })],
    });
  };

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE })
    ),
  });

  const bodyRows = rows.map((row, ri) => {
    const isLast = ri === rows.length - 1;
    return new TableRow({
      children: row.map((cell, i) =>
        cellOf(String(cell), colWidths[i], { top: NONE, bottom: isLast ? THICK : NONE, left: NONE, right: NONE })
      ),
    });
  });

  return new Table({
    width:        { size: tableW, type: WidthType.DXA },
    alignment:    AlignmentType.CENTER,
    columnWidths: colWidths,
    borders: { top: THICK, bottom: THICK, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    rows:         [headerRow, ...bodyRows],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Reference entry: 小五, 行距固定值 16 磅, 悬挂缩进 2 字符, "[n] " numbering
// ─────────────────────────────────────────────────────────────────────────────

function ref(text) {
  return new Paragraph({
    style: 'Reference',
    numbering: { reference: 'references', level: 0 },
    children: [new TextRun(text)],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: empty paragraph (spacing)
// ─────────────────────────────────────────────────────────────────────────────

function blank() {
  return new Paragraph({ children: [] });
}

// ─────────────────────────────────────────────────────────────────────────────
// Page break helper
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create a page break paragraph
 * @returns {Paragraph} Paragraph containing a page break
 */
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}



// ─────────────────────────────────────────────────────────────────────────────
// ISSUE 5 FIX: Block formula using temml + mathmlToDocxChildren (Word native math)
// ISSUE 2 FIX: Formula table borders all set to NONE including insideHorizontal/insideVertical
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Convert LaTeX formula to docx Math component
 * @param {string} latex - LaTeX formula string
 * @returns {Math} docx Math object
 */
function latexToMath(latex) {
  try {
    const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
    const children = mathmlToDocxChildren(mathml);
    if (children && children.length) {
      return new Math({ children });
    }
  } catch (e) {
    console.warn(`[formula] LaTeX parse error: ${latex}`, e.message);
  }
  // Fallback: return plain text
  return new Math({ children: [new MathRun(latex)] });
}

/**
 * Block formula layout using 3-column borderless table
 * ISSUE 2 FIX: All borders including insideHorizontal/insideVertical set to NONE
 * @param {string} latex - LaTeX formula string
 * @param {number|string} number - Equation number
 * @returns {Table} Formula table
 */
function formula(latex, number) {
  // Use 3-column borderless table layout: left margin | centered formula | right-aligned number
  const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  
  const leftCell = new TableCell({
    width: { size: 454, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
    children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
  });
  
  // Use temml + mathmlToDocxChildren to create Word native formula
  const mathObj = latexToMath(latex);
  const formulaCell = new TableCell({
    width: { size: COLUMN_W - 908, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      indent: { firstLine: 0 },
      spacing: { line: 240, lineRule: LineRuleType.AUTO },
      children: [mathObj],
    })],
  });
  
  const numberCell = new TableCell({
    width: { size: 454, type: WidthType.DXA },
    borders: noBorders,
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,  // ISSUE 11 FIX: Center content vertically
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      indent: { firstLine: 0 },
      spacing: { line: 240, lineRule: LineRuleType.AUTO },
      children: [new TextRun(`(${number})`)],
    })],
  });
  
  return new Table({
    width: { size: COLUMN_W, type: WidthType.DXA },
    columnWidths: [454, COLUMN_W - 908, 454],
    // ISSUE 2 FIX: Include insideHorizontal and insideVertical as NONE
    borders: {
      top: NONE,
      bottom: NONE,
      left: NONE,
      right: NONE,
      insideHorizontal: NONE,
      insideVertical: NONE,
    },
    rows: [new TableRow({
      children: [leftCell, formulaCell, numberCell],
    })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Document structure: styles + numbering (《岩土工程学报》 CJGE)
// Sizes in half-points (五号 21, 小五 18, 六号 15, 四号 28, 二号 44); spacing in twips (15.6 pt = 312).
// Chinese text 宋体 / 黑体 / 仿宋, Latin letters and digits Times New Roman.
// ─────────────────────────────────────────────────────────────────────────────

const fontOf = (eastAsia, latin = LATIN) => ({ ascii: latin, hAnsi: latin, cs: latin, eastAsia });
const exact = (pt) => ({ line: globalThis.Math.round(pt * 20), lineRule: LineRuleType.EXACT });
const pStyle = (id, name, run, paragraph, extra = {}) => ({
  id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true, run, paragraph, ...extra,
});

const STYLES = {
  default: {
    document: {
      run: { font: fontOf('SimSun'), size: 21 },                       // 五号
      paragraph: {
        alignment: AlignmentType.JUSTIFIED,
        spacing: exact(15.6),                                          // 固定值 15.6 磅
        indent:  { firstLine: 420 },                                   // 首行缩进 2 字符
      },
    },
  },
  paragraphStyles: [
    pStyle('Title', 'Title', { font: fontOf('SimHei'), size: 44, bold: false, color: '000000' },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(30), before: 0, after: 160 } }),
    pStyle('AuthorsCn', 'Authors CN', { font: fontOf('FangSong'), size: 21 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(15.6), after: 120 } }),
    pStyle('AffiliationCn', 'Affiliation CN', { font: fontOf('SimSun'), size: 15 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(12), after: 200 } }),
    pStyle('AbstractCn', 'Abstract CN', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.JUSTIFIED, indent: { left: 426, right: 426, firstLine: 0 }, spacing: exact(14) }),
    pStyle('FrontNote', 'Front Note', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.JUSTIFIED, indent: { left: 426, right: 426, firstLine: 0 }, spacing: exact(15.6) }),
    pStyle('TitleEn', 'Title EN', { font: fontOf(LATIN), size: 28, bold: true, color: '000000' },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(20), before: 270, after: 236 } }),
    pStyle('AuthorsEn', 'Authors EN', { font: fontOf(LATIN), size: 21 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: exact(15.6) }),
    pStyle('AffiliationEn', 'Affiliation EN', { font: fontOf(LATIN), size: 15 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(15.6), after: 40 } }),
    pStyle('AbstractEn', 'Abstract EN', { font: fontOf(LATIN), size: 18 },
      { alignment: AlignmentType.JUSTIFIED, indent: { left: 426, right: 426, firstLine: 0 }, spacing: exact(15.6) }),
    pStyle('PageNote', 'First Page Note', { font: fontOf('SimSun'), size: 15 },
      { alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 0 }, spacing: exact(12) }),
    pStyle('Heading1', 'Heading 1', { font: fontOf('SimHei'), size: 28, bold: false, color: '000000' },
      { alignment: AlignmentType.LEFT, indent: { firstLine: 0 }, spacing: { ...exact(18), before: 300, after: 0 }, outlineLevel: 0 }),
    pStyle('Heading2', 'Heading 2', { font: fontOf('SimHei'), size: 21, bold: false, color: '000000' },
      { alignment: AlignmentType.LEFT, indent: { firstLine: 0 }, spacing: { ...exact(15.6), before: 0, after: 0 }, outlineLevel: 1 }),
    pStyle('Heading3', 'Heading 3', { font: fontOf('SimSun'), size: 21, bold: false, color: '000000' },
      { alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 420 }, spacing: { ...exact(15.6), before: 0, after: 0 }, outlineLevel: 2 }),
    pStyle('RefHeading', 'Reference Heading', { font: fontOf('SimHei'), size: 21 },
      { alignment: AlignmentType.LEFT, indent: { firstLine: 0 }, spacing: exact(15.6) }),
    pStyle('FigureCaption', 'Figure Caption', { font: fontOf('SimHei'), size: 18, bold: false },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(15.6), before: 60 } }),
    pStyle('FigureCaptionEn', 'Figure Caption EN', { font: fontOf(LATIN), size: 18, bold: false },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(15.6), after: 60 } }),
    pStyle('TableCaption', 'Table Caption', { font: fontOf('SimHei'), size: 18, bold: false },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: { ...exact(14), before: 60 } }),
    pStyle('TableCaptionEn', 'Table Caption EN', { font: fontOf(LATIN), size: 18, bold: false },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: exact(14) }),
    pStyle('SubCaption', 'Sub Caption', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: exact(14) }),
    pStyle('TableText', 'Table Text', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.CENTER, indent: { firstLine: 0 }, spacing: exact(14) }),
    pStyle('Statement', 'Statement', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 0 }, spacing: exact(15.6) }),
    pStyle('Reference', 'Reference', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.JUSTIFIED, indent: { left: 360, hanging: 360 }, spacing: exact(16) }),
    pStyle('HeaderText', 'Header Text', { font: fontOf('SimSun'), size: 18 },
      { alignment: AlignmentType.LEFT, indent: { firstLine: 0 }, spacing: { line: 240, lineRule: LineRuleType.AUTO } }),
  ],
};

const NUMBERING = {
  config: [
    {
      reference: 'references',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '[%1]', suffix: LevelSuffix.SPACE,
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 360 } } } },
      ],
    },
    {
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
    {
      reference: 'numbers',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ],
    },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Running headers (CJGE): page number in the header, no footer.
//   odd  : 第N期 | 第一作者，等. 题名 | 页码        even: 页码 | 岩 土 工 程 学 报 | YYYY年
//   first: 第V卷 第N期 / YYYY年 M月 | 岩 土 工 程 学 报 / Chinese Journal of … | Vol.V No.N / Mon. YYYY
// Header rules: 0.5 pt single; first page: thick-thin double.
// ─────────────────────────────────────────────────────────────────────────────

const JOURNAL = {
  year: new Date().getFullYear(),
  volume: '',          // e.g. '48' — first-page header is printed only when volume and issue are set
  issue: '',           // e.g. '9'
  month: '',           // e.g. 9
  runningHeader: '',   // e.g. '简富献，等. 滑坡堰塞体河谷堆积形态特征研究'
};

const TABS = [
  { type: TabStopType.CENTER, position: globalThis.Math.round(CONTENT_W / 2) },
  { type: TabStopType.RIGHT, position: CONTENT_W },
];
const HEADER_RULE = { bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000', space: 1 } };
const FIRST_RULE = { bottom: { style: BorderStyle.THICK_THIN_SMALL_GAP, size: 18, color: '000000', space: 1 } };

function headerLine(children, border) {
  return new Paragraph({ style: 'HeaderText', tabStops: TABS, border, children });
}

function buildHeaders() {
  const months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.'];
  const m = Number(JOURNAL.month) || 0;
  const first = JOURNAL.volume && JOURNAL.issue
    ? [
        headerLine([
          new TextRun(`第${JOURNAL.volume}卷 第${JOURNAL.issue}期\t`),
          new TextRun({ text: '岩 土 工 程 学 报', size: 24 }),
          new TextRun(`\tVol.${JOURNAL.volume} No.${JOURNAL.issue}`),
        ]),
        headerLine([new TextRun(`${JOURNAL.year}年${m ? ` ${m}月` : ''}\tChinese Journal of Geotechnical Engineering\t${m ? months[m - 1] + ' ' : ''}${JOURNAL.year}`)], FIRST_RULE),
      ]
    : [new Paragraph({ style: 'HeaderText', children: [] })];
  return {
    first: new Header({ children: first }),
    default: new Header({
      children: [headerLine([
        new TextRun(`${JOURNAL.issue ? `第${JOURNAL.issue}期` : ''}\t${JOURNAL.runningHeader}\t`),
        new TextRun({ children: [PageNumber.CURRENT] }),
      ], HEADER_RULE)],
    }),
    even: new Header({
      children: [headerLine([
        new TextRun({ children: [PageNumber.CURRENT] }),
        new TextRun(`\t岩 土 工 程 学 报\t${JOURNAL.year}年`),
      ], HEADER_RULE)],
    }),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// ██  CONTENT SECTION — Edit below this line  ████████████████████████████████
// FRONT = 通栏 front matter (single column); CONTENT = 双栏 body.
// ─────────────────────────────────────────────────────────────────────────────

resetHeadingCounters();

let FRONT;
let CONTENT;
if (INPUT_MARKDOWN) {
  ({ front: FRONT, body: CONTENT } = buildContentFromMarkdown(INPUT_MARKDOWN));
} else {
  JOURNAL.runningHeader = '张三，等. 论文题名示例';
  FRONT = [
    titleCn('论文题名示例'),
    authorsCn('张三', ['1', '2'], '李四', ['*1']),
    affiliationCn('（1. 某某大学土木工程学院，江苏 南京 210000；2. 某某研究院，江苏 南京 210024）'),
    abstractCn('为研究……，开展了……试验，得到了……。结果表明：……。'),
    keywordsCn('关键词一；关键词二；关键词三'),
    clcLine('中图分类号：TU42　　　　文献标识码：A　　　　文章编号：1000-4548(2026)00-0000-00'),
    bioCn('张三(1990— )，男，博士，主要从事××方面的研究工作。E-mail: zhangsan@example.com。'),
    titleEn('Title of the paper'),
    authorsEn('ZHANG San', ['1', '2'], 'LI Si', ['1']),
    affiliationEn('(1. School of Civil Engineering, XX University, Nanjing 210000, China; 2. XX Research Institute, Nanjing 210024, China)'),
    abstractEn('To study ……, tests are carried out. The results show that ……'),
    keywordsEn('keyword one; keyword two; keyword three'),
  ];
  CONTENT = [
    h1('引    言'),
    body('正文示例，引用写作上标[1]。数值与单位之间空一格，如 124 m、0.5～1.0 mm。'),
    h1('试验设计'),
    h2('试验装置'),
    h3('级配影响'),
    body('a）相似性。四级标题与正文接排。'),
    ...tableCaption('表1 试样参数', 'Table 1 Parameters of samples'),
    threeLineTable(['参数', '数值'], [['跨度/m', '1.2'], ['θ/(°)', '36']]),
    h1('结    论'),
    body('为……，主要得到以下 2 点结论。'),
    body('（1）……。'),
    body('（2）……。'),
    statement('利益冲突声明/Conflict of Interests：', '所有作者声明不存在利益冲突。All authors disclose no relevant conflict of interest.'),
    statement('作者贡献/Authors\' Contributions：', '张三设计研究内容，李四参与论文的写作和修改。所有作者均阅读并同意最终稿件的提交。All the authors have read the final version of manuscript and consented for submission.'),
    refHeading(),
    ref('DONG J Y, TUNG Y S, CHEN C C, et al. Logistic regression model for predicting the failure probability of a landslide dam[J]. Engineering Geology, 2011, 117(1/2): 52-61.'),
  ];
}

// ─────────────────────────────────────────────────────────────────────────────
// Build & write document: A4, CJGE margins, front matter single column, body two columns
// ─────────────────────────────────────────────────────────────────────────────

const PAGE = {
  size:   { width: PAGE_W, height: PAGE_H },
  margin: { top: MARGIN_TB, bottom: MARGIN_TB, left: MARGIN_L, right: MARGIN_R, header: HEADER_DIST, footer: 567 },
};

const doc = new Document({
  styles:    STYLES,
  numbering: NUMBERING,
  evenAndOddHeaderAndFooters: true,
  sections: [
    {
      properties: { page: PAGE, titlePage: true, column: { count: 1 } },
      headers: buildHeaders(),
      children: FRONT,
    },
    {
      properties: { type: SectionType.CONTINUOUS, page: PAGE, column: { count: 2, space: COLUMN_SPACE } },
      children: CONTENT,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  const out = path.resolve(OUTPUT_PATH);
  fs.writeFileSync(out, buf);
  console.log(`✓  Written: ${out}`);
}).catch(err => {
  console.error('Error building document:', err.message);
  process.exit(1);
});
