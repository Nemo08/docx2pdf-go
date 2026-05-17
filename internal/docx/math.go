package docx

import (
	"encoding/xml"
	"strings"
)

// Math support: we walk the OMML (Office Math Markup Language) subtree
// and emit a structurally-aware *string* approximation rather than a true
// math-typesetting layout. The output is meant to be readable inline:
//
//	m:f  (fraction)        → "a/b"
//	m:rad (radical)        → "√(x)" or "ⁿ√(x)" for nth-root
//	m:sSup (superscript)   → "x^(2)" with Unicode superscripts for 0-9 + - = ( )
//	m:sSub (subscript)     → "x_(2)" with Unicode subscripts
//	m:sSubSup              → "x_(i)^(2)"
//	m:nary (∑/∫/∏ etc.)    → "∑_(i=1)^(n) f(i)"
//	m:d   (delimited)      → "(a, b)" or "{a; b}" depending on declared chars
//	m:func (function-apply)→ "sin(x)"
//	m:limLow / m:limUpp    → "lim_(x→0) f(x)" / "lim^∞ f"
//	m:acc  (accent)        → "x̂" by stacking the accent char after the base
//	m:bar / m:box          → "‾x" or "⌜x⌝"
//	m:groupChr             → "⟨x⟩" using its declared bracket char
//	m:matrix / m:eqArr     → "[a b; c d]" / "{a; b}"
//	m:r/m:t (run / text)   → the visible text content
//
// This is far better than the previous flat-extract: variables stay
// grouped, exponents and limits read correctly, and matrices keep their
// row structure. It is NOT a substitute for proper glyph positioning —
// that requires a math engine which is out of scope.

// extractMathText walks an m:oMath / m:oMathPara subtree starting at start
// and returns a single readable string. The caller emits this as an italic
// run.
func extractMathText(dec *xml.Decoder, start xml.StartElement) (string, error) {
	mr, err := decodeMathNode(dec, start)
	if err != nil {
		return "", err
	}
	return mr.render(), nil
}

// ExtractMathTree walks an m:oMath / m:oMathPara subtree and returns the
// structural MathNode tree plus the flat string approximation (suitable
// as a textual fallback). The renderer prefers the tree when it can paint
// 2D math; the string keeps text-search of the PDF working.
func ExtractMathTree(dec *xml.Decoder, start xml.StartElement) (*MathNode, string, error) {
	mr, err := decodeMathNode(dec, start)
	if err != nil {
		return nil, "", err
	}
	if mr == nil {
		return nil, "", nil
	}
	return mr, mr.render(), nil
}

// mathNode is an OMML subtree we know how to render to a string. Each
// node has a kind plus a small set of slots — argument lists for things
// like sub/sup, delimiters, n-ary operators, accents.
type mathNode = MathNode

// MathNode is an OMML subtree the renderer can either flatten to a string
// (via render()) or lay out structurally on the PDF canvas (the render
// package walks the tree directly). Each node has a kind plus a small set
// of slots for sub/sup, delimiters, n-ary operators, accents, etc. All
// fields are exported so the render package can read them without going
// through accessor methods.
type MathNode struct {
	Kind     string
	Text     string      // raw text for "r" / "t" / accentChar / numerator-tex etc.
	Children []*MathNode // generic ordered children (e.g. inside m:e, m:oMath body)
	// Named slots for structured elements. Empty when not applicable.
	Num   *MathNode // m:f numerator
	Den   *MathNode // m:f denominator
	Base  *MathNode // m:sSup / m:sSub / m:sSubSup / m:rad / m:nary / m:limLow / m:limUpp / m:acc / m:groupChr / m:bar / m:box / m:func base
	Sup   *MathNode // superscript
	Sub   *MathNode // subscript
	Deg   *MathNode // m:rad degree
	LimLo *MathNode // m:limLow / m:nary lower limit
	LimUp *MathNode // m:limUpp / m:nary upper limit
	Arg   *MathNode // m:func argument
	// matrix rows; each row is a list of "e" cells.
	Rows [][]*MathNode
	// Per-element formatting hints pulled from props.
	BegChar  string // m:dPr begChr
	EndChar  string // m:dPr endChr
	SepChar  string // m:dPr sepChar (defaults to ",")
	NaryChar string // m:naryPr chr (∑, ∫, ∏ ...)
	AccChar  string // m:accPr chr
}

func newMathNode(kind string) *mathNode { return &mathNode{Kind: kind} }

// decodeMathNode walks one element and returns its mathNode. Recognizes
// OMML structural elements; treats unknown ones as opaque text wrappers
// (their CharData content is concatenated into the node's text).
func decodeMathNode(dec *xml.Decoder, start xml.StartElement) (*mathNode, error) {
	n := newMathNode(start.Name.Local)
	// Pull display attributes off props elements at decode time.
	for {
		tok, err := dec.Token()
		if err != nil {
			return n, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			child, err := decodeMathNode(dec, t)
			if err != nil {
				return n, err
			}
			// Route the child into the appropriate slot based on its
			// element name (m:num / m:den / m:e / m:sup / m:sub / etc.).
			switch t.Name.Local {
			case "num":
				n.Num = child
			case "den":
				n.Den = child
			case "e":
				if n.Kind == "mr" {
					// matrix row child cells appear as m:e directly inside m:mr.
					if len(n.Rows) == 0 {
						n.Rows = append(n.Rows, nil)
					}
					n.Rows[0] = append(n.Rows[0], child)
				} else {
					n.Base = child
				}
			case "sup":
				n.Sup = child
			case "sub":
				n.Sub = child
			case "deg":
				n.Deg = child
			case "lim":
				// m:limLow and m:limUpp wrap a m:lim element holding the
				// limit expression.
				if n.Kind == "limLow" || n.Kind == "naryLimLow" {
					n.LimLo = child
				} else if n.Kind == "limUpp" || n.Kind == "naryLimUpp" {
					n.LimUp = child
				} else {
					n.Children = append(n.Children, child)
				}
			case "fName":
				n.Arg = child
			case "mr":
				n.Rows = append(n.Rows, child.Rows[0])
			case "dPr":
				n.BegChar = child.BegChar
				n.EndChar = child.EndChar
				n.SepChar = child.SepChar
			case "naryPr":
				n.NaryChar = child.NaryChar
				// nary props also carry sub-position / lim-loc info; we
				// ignore those (formatting only).
			case "accPr":
				n.AccChar = child.AccChar
			case "begChr":
				n.BegChar = child.Text
			case "endChr":
				n.EndChar = child.Text
			case "sepChr":
				n.SepChar = child.Text
			case "chr":
				// chr is reused across many props elements; n.Kind tells
				// us which to assign to. naryPr.chr → naryChar;
				// accPr.chr → accChar; groupChrPr.chr → base text.
				if n.Kind == "naryPr" {
					n.NaryChar = child.Text
				} else if n.Kind == "accPr" {
					n.AccChar = child.Text
				} else if n.Kind == "groupChrPr" {
					n.AccChar = child.Text
				}
			default:
				n.Children = append(n.Children, child)
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return n, nil
			}
		case xml.CharData:
			// m:t holds the literal glyph text; everything else is
			// structural. Accumulate CharData into n.Text.
			n.Text += string(t)
		}
		// For attribute-bearing elements (begChr / endChr / sepChr / chr),
		// we also want the "val" attribute (some writers put the literal
		// glyph in val instead of in CharData).
		if start.Name.Local == "begChr" || start.Name.Local == "endChr" || start.Name.Local == "sepChr" || start.Name.Local == "chr" {
			if n.Text == "" {
				for _, a := range start.Attr {
					if a.Name.Local == "val" && a.Value != "" {
						n.Text = a.Value
					}
				}
			}
		}
	}
}

// render returns the readable approximation for this node's subtree.
func (n *mathNode) render() string {
	if n == nil {
		return ""
	}
	switch n.Kind {
	case "t":
		return n.Text
	case "r", "e", "num", "den", "deg", "sup", "sub", "lim", "fName", "oMath", "oMathPara":
		return n.joinChildren()
	case "f":
		// Fraction: render as "num/den" with parens around multi-token
		// halves so the precedence reads sensibly.
		num, den := n.Num.render(), n.Den.render()
		return wrapIfComplex(num) + "/" + wrapIfComplex(den)
	case "rad":
		base := n.Base.render()
		deg := n.Deg.render()
		if deg != "" {
			return supScript(deg) + "√(" + base + ")"
		}
		return "√(" + base + ")"
	case "sSup":
		return n.Base.render() + supScript(n.Sup.render())
	case "sSub":
		return n.Base.render() + subScript(n.Sub.render())
	case "sSubSup":
		return n.Base.render() + subScript(n.Sub.render()) + supScript(n.Sup.render())
	case "nary":
		op := n.NaryChar
		if op == "" {
			op = "∑"
		}
		// nary's lower/upper limits live in m:sub/m:sup in the wire
		// format — the decoder routes them to n.Sub / n.Sup. Older
		// docs use m:limLow/m:limUpp inside m:nary; respect both.
		lo := n.LimLo.render()
		if lo == "" {
			lo = n.Sub.render()
		}
		up := n.LimUp.render()
		if up == "" {
			up = n.Sup.render()
		}
		body := n.Base.render()
		s := op
		if lo != "" {
			s += subScript(lo)
		}
		if up != "" {
			s += supScript(up)
		}
		if body != "" {
			s += " " + body
		}
		return s
	case "d":
		// Delimited group: render with the begChar/endChar pair, defaulting
		// to round brackets. If multiple m:e children exist they are
		// separated by sepChar (default ",").
		open, close := n.BegChar, n.EndChar
		if open == "" {
			open = "("
		}
		if close == "" {
			close = ")"
		}
		sep := n.SepChar
		if sep == "" {
			sep = ", "
		}
		parts := make([]string, 0, len(n.Children))
		for _, c := range n.Children {
			if c.Kind == "e" {
				parts = append(parts, c.render())
			}
		}
		if len(parts) == 0 {
			parts = append(parts, n.joinChildren())
		}
		return open + strings.Join(parts, sep) + close
	case "func":
		return n.Arg.render() + "(" + n.Base.render() + ")"
	case "limLow":
		return n.Base.render() + subScript(n.LimLo.render())
	case "limUpp":
		return n.Base.render() + supScript(n.LimUp.render())
	case "acc":
		// Accent: stack the accent char after the base (Unicode combining
		// behaviour). If the accent char is empty fall back to U+0302
		// (combining circumflex).
		ch := n.AccChar
		if ch == "" {
			ch = "̂"
		}
		return n.Base.render() + ch
	case "bar":
		return "‾" + n.Base.render()
	case "box":
		return "⌜" + n.Base.render() + "⌝"
	case "borderBox":
		return "[" + n.Base.render() + "]"
	case "groupChr":
		ch := n.AccChar
		if ch == "" {
			ch = "⏞"
		}
		return ch + "{" + n.Base.render() + "}"
	case "m", "matrix":
		// Matrix: rows separated by "; ", cells in a row by " ".
		out := make([]string, 0, len(n.Rows))
		for _, row := range n.Rows {
			cells := make([]string, len(row))
			for i, c := range row {
				cells[i] = c.render()
			}
			out = append(out, strings.Join(cells, " "))
		}
		return "[" + strings.Join(out, "; ") + "]"
	case "eqArr":
		// Equation array: stack of equations separated by "; ".
		parts := make([]string, 0, len(n.Children))
		for _, c := range n.Children {
			if c.Kind == "e" {
				parts = append(parts, c.render())
			}
		}
		return "{" + strings.Join(parts, "; ") + "}"
	case "phant":
		// Phantom: invisible — drop the content (matches Word's intent).
		return ""
	default:
		// Unknown structural wrapper: render its children inline.
		return n.joinChildren()
	}
}

func (n *mathNode) joinChildren() string {
	if n == nil {
		return ""
	}
	var b strings.Builder
	b.WriteString(n.Text)
	for _, c := range n.Children {
		b.WriteString(c.render())
	}
	return b.String()
}

// wrapIfComplex wraps a string in parentheses if it contains any
// "operator-like" character (space, +, -, *, /, =). Single tokens like
// "x" or "23" are returned unchanged so the fraction reads naturally.
func wrapIfComplex(s string) string {
	for _, r := range s {
		switch r {
		case ' ', '+', '-', '*', '/', '=':
			return "(" + s + ")"
		}
	}
	return s
}

// supScript / subScript convert a string into Unicode superscript /
// subscript glyphs when possible, otherwise wraps with "^( )" / "_( )"
// so the structure is at least visible.
var supTable = map[rune]rune{
	'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵',
	'6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
	'+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
	'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ', 'f': 'ᶠ',
	'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ', 'k': 'ᵏ', 'l': 'ˡ',
	'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ',
	't': 'ᵗ', 'u': 'ᵘ', 'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ',
	'z': 'ᶻ',
}

var subTable = map[rune]rune{
	'0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅',
	'6': '₆', '7': '₇', '8': '₈', '9': '₉',
	'+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
	'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ', 'k': 'ₖ',
	'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ', 'p': 'ₚ', 'r': 'ᵣ',
	's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ', 'v': 'ᵥ', 'x': 'ₓ',
}

func supScript(s string) string {
	if s == "" {
		return ""
	}
	out := make([]rune, 0, len(s))
	allMapped := true
	for _, r := range s {
		if m, ok := supTable[r]; ok {
			out = append(out, m)
		} else {
			allMapped = false
			break
		}
	}
	if allMapped {
		return string(out)
	}
	return "^(" + s + ")"
}

func subScript(s string) string {
	if s == "" {
		return ""
	}
	out := make([]rune, 0, len(s))
	allMapped := true
	for _, r := range s {
		if m, ok := subTable[r]; ok {
			out = append(out, m)
		} else {
			allMapped = false
			break
		}
	}
	if allMapped {
		return string(out)
	}
	return "_(" + s + ")"
}

// mathRun returns a Run carrying the extracted math text styled in italic,
// inheriting the surrounding paragraph's run properties.
func mathRun(text string, paraRPr RunProps) Run {
	rp := paraRPr
	rp.Italic = true
	return Run{Text: text, Props: rp}
}
