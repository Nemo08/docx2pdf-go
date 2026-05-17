package render

import (
	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

// math2d lays out a MathNode tree on the PDF canvas as a 2D expression.
// We use a simple box model:
//
//   - Each box reports its width, ascent (above baseline), and descent
//     (below baseline). All values in PostScript points.
//   - Composite layouts (fractions, radicals, n-ary) stack their child
//     boxes around a baseline using these measurements.
//   - The atomic case is a textual leaf rendered at the active font size.
//
// Limitations: we paint fraction bars and radical vinculums as gopdf
// lines, but variable-height brackets are approximated by stretching the
// glyph baseline rather than drawing a custom path. Matrices use a
// uniform column gap. Good enough to make most algebra readable.

// mathBox represents one laid-out subtree.
type mathBox struct {
	w       float64
	ascent  float64
	descent float64
	draw    func(r *renderer, x, baseline float64)
}

// height returns ascent + descent.
func (b mathBox) height() float64 { return b.ascent + b.descent }

// buildMathBox recursively measures and lays out one MathNode at the
// given font size. Returns the resulting box; the caller invokes
// box.draw(x, baseline) to actually paint at a target location.
func (r *renderer) buildMathBox(n *docx.MathNode, fontSize float64) mathBox {
	if n == nil {
		return mathBox{}
	}
	switch n.Kind {
	case "t":
		return r.mathTextBox(n.Text, fontSize)
	case "r":
		// w:r in OMML carries a w:t leaf — n.Text was populated by the
		// decoder via CharData. Render directly. Children, when present,
		// are non-text formatting wrappers we treat as a sequence.
		if n.Text != "" {
			return r.mathTextBox(n.Text, fontSize)
		}
		return r.mathSequence(n.Children, fontSize)
	case "e", "num", "den", "deg", "sup", "sub", "lim", "fName", "oMath", "oMathPara":
		if n.Text != "" && len(n.Children) == 0 {
			return r.mathTextBox(n.Text, fontSize)
		}
		return r.mathSequence(n.Children, fontSize)
	case "f":
		return r.mathFractionBox(n, fontSize)
	case "rad":
		return r.mathRadicalBox(n, fontSize)
	case "sSup":
		return r.mathSupBox(n, fontSize)
	case "sSub":
		return r.mathSubBox(n, fontSize)
	case "sSubSup":
		return r.mathSubSupBox(n, fontSize)
	case "nary":
		return r.mathNaryBox(n, fontSize)
	case "d":
		return r.mathDelimBox(n, fontSize)
	case "func":
		return r.mathFuncBox(n, fontSize)
	case "acc":
		return r.mathAccBox(n, fontSize)
	case "bar":
		return r.mathBarBox(n, fontSize)
	case "limLow":
		return r.mathLimBox(n, n.LimLo, true, fontSize)
	case "limUpp":
		return r.mathLimBox(n, n.LimUp, false, fontSize)
	case "m", "matrix":
		return r.mathMatrixBox(n, fontSize)
	}
	// Unknown kind: fall back to the textual approximation.
	if n.Text != "" {
		return r.mathTextBox(n.Text, fontSize)
	}
	return r.mathSequence(n.Children, fontSize)
}

// mathTextBox renders a single string at the active font size.
func (r *renderer) mathTextBox(s string, fontSize float64) mathBox {
	if s == "" {
		return mathBox{}
	}
	w := mustMeasureMath(r, s, fontSize)
	return mathBox{
		w:       w,
		ascent:  fontSize * 0.75,
		descent: fontSize * 0.25,
		draw: func(r *renderer, x, baseline float64) {
			r.pdf.SetFontSize(fontSize)
			r.pdf.SetX(x)
			r.pdf.SetY(baseline - fontSize*0.75)
			_ = r.pdf.Cell(nil, s)
		},
	}
}

func mustMeasureMath(r *renderer, s string, fontSize float64) float64 {
	old := r.opts.DefaultFontSize
	// Ensure a font is selected — MeasureTextWidth panics without one. The
	// renderer's default family is reliably registered, so SetFont is a
	// safe no-op when an active face is already set.
	defer func() { _ = recover() }()
	_ = r.pdf.SetFont(defaultFamily, "", fontSize)
	r.pdf.SetFontSize(fontSize)
	w, _ := r.pdf.MeasureTextWidth(s)
	r.pdf.SetFontSize(old)
	return w
}

// mathSequence lays out a list of children in a single horizontal row.
func (r *renderer) mathSequence(kids []*docx.MathNode, fontSize float64) mathBox {
	if len(kids) == 0 {
		return mathBox{ascent: fontSize * 0.75, descent: fontSize * 0.25}
	}
	if len(kids) == 1 {
		return r.buildMathBox(kids[0], fontSize)
	}
	boxes := make([]mathBox, len(kids))
	totalW := 0.0
	maxA, maxD := 0.0, 0.0
	for i, c := range kids {
		boxes[i] = r.buildMathBox(c, fontSize)
		totalW += boxes[i].w
		if boxes[i].ascent > maxA {
			maxA = boxes[i].ascent
		}
		if boxes[i].descent > maxD {
			maxD = boxes[i].descent
		}
	}
	return mathBox{
		w:       totalW,
		ascent:  maxA,
		descent: maxD,
		draw: func(r *renderer, x, baseline float64) {
			cx := x
			for _, b := range boxes {
				if b.draw != nil {
					b.draw(r, cx, baseline)
				}
				cx += b.w
			}
		},
	}
}

// mathFractionBox stacks numerator over denominator with a horizontal bar.
func (r *renderer) mathFractionBox(n *docx.MathNode, fontSize float64) mathBox {
	num := r.buildMathBox(n.Num, fontSize*0.9)
	den := r.buildMathBox(n.Den, fontSize*0.9)
	w := num.w
	if den.w > w {
		w = den.w
	}
	const barGap = 1.5
	// The fraction's box ascent reaches the top of the numerator;
	// descent reaches the bottom of the denominator. Baseline sits on
	// the fraction bar.
	asc := num.height() + barGap
	desc := den.height() + barGap
	return mathBox{
		w:       w + 2, // a small margin on each side for the bar
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			barY := baseline
			if num.draw != nil {
				num.draw(r, x+(w-num.w)/2+1, barY-barGap-num.descent)
			}
			if den.draw != nil {
				den.draw(r, x+(w-den.w)/2+1, barY+barGap+den.ascent)
			}
			r.pdf.SetLineWidth(0.5)
			r.pdf.SetStrokeColor(0, 0, 0)
			r.pdf.Line(x+1, barY, x+1+w, barY)
		},
	}
}

// mathRadicalBox draws a √ symbol with a horizontal vinculum over the
// base; an optional degree sits as a small superscript on the radical's
// upper-left.
func (r *renderer) mathRadicalBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	const symW = 0.4 // width of the radical sign as a fraction of fontSize
	rsW := fontSize * symW
	w := rsW + base.w + 2
	asc := base.ascent + 2
	desc := base.descent
	deg := r.buildMathBox(n.Deg, fontSize*0.6)
	if deg.w > 0 {
		w += deg.w * 0.7
	}
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			// Paint √ symbol as two strokes plus the vinculum.
			midY := baseline + fontSize*0.2
			topY := baseline - asc + 1
			leftX := x + rsW*0.2
			if deg.w > 0 {
				deg.draw(r, x, baseline-asc*0.7)
				leftX += deg.w * 0.7
			}
			r.pdf.SetLineWidth(0.8)
			r.pdf.SetStrokeColor(0, 0, 0)
			r.pdf.Line(leftX, midY, leftX+rsW*0.4, baseline+desc) // down-stroke
			r.pdf.Line(leftX+rsW*0.4, baseline+desc, leftX+rsW*0.8, topY)
			r.pdf.Line(leftX+rsW*0.8, topY, leftX+rsW+base.w+1, topY) // vinculum
			if base.draw != nil {
				base.draw(r, leftX+rsW, baseline)
			}
		},
	}
}

// mathSupBox stacks a superscript on the base's upper-right.
func (r *renderer) mathSupBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	sup := r.buildMathBox(n.Sup, fontSize*0.75)
	supRise := fontSize * 0.45
	w := base.w + sup.w
	asc := base.ascent + supRise*0.5
	if base.ascent < supRise+sup.ascent {
		asc = supRise + sup.ascent
	}
	desc := base.descent
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x, baseline)
			}
			if sup.draw != nil {
				sup.draw(r, x+base.w, baseline-supRise)
			}
		},
	}
}

// mathSubBox stacks a subscript on the base's lower-right.
func (r *renderer) mathSubBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	sub := r.buildMathBox(n.Sub, fontSize*0.75)
	subDrop := fontSize * 0.25
	w := base.w + sub.w
	asc := base.ascent
	desc := base.descent + subDrop
	if base.descent < subDrop+sub.descent {
		desc = subDrop + sub.descent
	}
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x, baseline)
			}
			if sub.draw != nil {
				sub.draw(r, x+base.w, baseline+subDrop)
			}
		},
	}
}

// mathSubSupBox stacks both subscript and superscript on the base.
func (r *renderer) mathSubSupBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	sub := r.buildMathBox(n.Sub, fontSize*0.75)
	sup := r.buildMathBox(n.Sup, fontSize*0.75)
	supRise := fontSize * 0.45
	subDrop := fontSize * 0.25
	wExt := sub.w
	if sup.w > wExt {
		wExt = sup.w
	}
	w := base.w + wExt
	asc := supRise + sup.ascent
	if base.ascent > asc {
		asc = base.ascent
	}
	desc := subDrop + sub.descent
	if base.descent > desc {
		desc = base.descent
	}
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x, baseline)
			}
			if sup.draw != nil {
				sup.draw(r, x+base.w, baseline-supRise)
			}
			if sub.draw != nil {
				sub.draw(r, x+base.w, baseline+subDrop)
			}
		},
	}
}

// mathNaryBox renders an n-ary operator (∑ / ∫ / ∏) with stacked
// limits when present. The operator glyph defaults to ∑.
func (r *renderer) mathNaryBox(n *docx.MathNode, fontSize float64) mathBox {
	glyph := n.NaryChar
	if glyph == "" {
		glyph = "∑"
	}
	op := r.mathTextBox(glyph, fontSize*1.4)
	base := r.buildMathBox(n.Base, fontSize)
	lo := r.buildMathBox(n.LimLo, fontSize*0.7)
	hi := r.buildMathBox(n.LimUp, fontSize*0.7)
	limW := lo.w
	if hi.w > limW {
		limW = hi.w
	}
	opSpan := op.w
	if limW > opSpan {
		opSpan = limW
	}
	w := opSpan + base.w + 2
	asc := op.ascent
	if hi.height() > 0 {
		asc += hi.height() + 1
	}
	desc := op.descent
	if lo.height() > 0 {
		desc += lo.height() + 1
	}
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			cx := x + (opSpan-op.w)/2
			if hi.draw != nil {
				hi.draw(r, x+(opSpan-hi.w)/2, baseline-op.ascent-1-hi.descent)
			}
			if op.draw != nil {
				op.draw(r, cx, baseline)
			}
			if lo.draw != nil {
				lo.draw(r, x+(opSpan-lo.w)/2, baseline+op.descent+1+lo.ascent)
			}
			if base.draw != nil {
				base.draw(r, x+opSpan+2, baseline)
			}
		},
	}
}

// mathDelimBox surrounds a body with paired delimiters (paren / bracket /
// brace / pipe). Delimiters stretch by simply scaling their font size to
// match the body height.
func (r *renderer) mathDelimBox(n *docx.MathNode, fontSize float64) mathBox {
	beg := n.BegChar
	if beg == "" {
		beg = "("
	}
	end := n.EndChar
	if end == "" {
		end = ")"
	}
	sep := n.SepChar
	if sep == "" {
		sep = ","
	}
	// Body: all child slots joined by sep.
	parts := []mathBox{}
	if n.Base != nil {
		parts = append(parts, r.buildMathBox(n.Base, fontSize))
	}
	for _, c := range n.Children {
		if c.Kind == "e" {
			parts = append(parts, r.buildMathBox(c, fontSize))
		}
	}
	if len(parts) == 0 {
		parts = []mathBox{r.mathSequence(n.Children, fontSize)}
	}
	bodyW := 0.0
	maxA, maxD := 0.0, 0.0
	for i, p := range parts {
		bodyW += p.w
		if i > 0 {
			bodyW += r.mathTextBox(sep, fontSize).w + 2
		}
		if p.ascent > maxA {
			maxA = p.ascent
		}
		if p.descent > maxD {
			maxD = p.descent
		}
	}
	delimScale := 1.0
	bodyH := maxA + maxD
	if bodyH > fontSize*1.4 {
		delimScale = bodyH / (fontSize * 1.1)
	}
	begBox := r.mathTextBox(beg, fontSize*delimScale)
	endBox := r.mathTextBox(end, fontSize*delimScale)
	asc := begBox.ascent
	if maxA > asc {
		asc = maxA
	}
	desc := begBox.descent
	if maxD > desc {
		desc = maxD
	}
	return mathBox{
		w:       begBox.w + bodyW + endBox.w + 2,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			cx := x
			begBox.draw(r, cx, baseline)
			cx += begBox.w
			for i, p := range parts {
				if i > 0 {
					sepBox := r.mathTextBox(sep, fontSize)
					sepBox.draw(r, cx, baseline)
					cx += sepBox.w + 2
				}
				if p.draw != nil {
					p.draw(r, cx, baseline)
				}
				cx += p.w
			}
			endBox.draw(r, cx, baseline)
		},
	}
}

// mathFuncBox renders fName(arg) — e.g. sin(x).
func (r *renderer) mathFuncBox(n *docx.MathNode, fontSize float64) mathBox {
	name := r.buildMathBox(n.Arg, fontSize)
	if name.w == 0 {
		// Empty name field — children carry the function name.
		name = r.mathSequence(n.Children, fontSize)
	}
	body := r.buildMathBox(n.Base, fontSize)
	gap := fontSize * 0.15
	w := name.w + gap + body.w
	asc := name.ascent
	if body.ascent > asc {
		asc = body.ascent
	}
	desc := name.descent
	if body.descent > desc {
		desc = body.descent
	}
	return mathBox{
		w:       w,
		ascent:  asc,
		descent: desc,
		draw: func(r *renderer, x, baseline float64) {
			if name.draw != nil {
				name.draw(r, x, baseline)
			}
			if body.draw != nil {
				body.draw(r, x+name.w+gap, baseline)
			}
		},
	}
}

// mathAccBox layers an accent character above the base.
func (r *renderer) mathAccBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	accChar := n.AccChar
	if accChar == "" {
		accChar = "̂"
	}
	acc := r.mathTextBox(accChar, fontSize*0.6)
	asc := base.ascent + acc.height()
	return mathBox{
		w:       base.w,
		ascent:  asc,
		descent: base.descent,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x, baseline)
			}
			if acc.draw != nil {
				acc.draw(r, x+(base.w-acc.w)/2, baseline-base.ascent)
			}
		},
	}
}

// mathBarBox draws a horizontal overline over the base.
func (r *renderer) mathBarBox(n *docx.MathNode, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	asc := base.ascent + 2
	return mathBox{
		w:       base.w,
		ascent:  asc,
		descent: base.descent,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x, baseline)
			}
			y := baseline - base.ascent - 1
			r.pdf.SetLineWidth(0.5)
			r.pdf.SetStrokeColor(0, 0, 0)
			r.pdf.Line(x, y, x+base.w, y)
		},
	}
}

// mathLimBox renders limLow / limUpp: a base with a low (or high) limit
// underneath.
func (r *renderer) mathLimBox(n *docx.MathNode, lim *docx.MathNode, low bool, fontSize float64) mathBox {
	base := r.buildMathBox(n.Base, fontSize)
	limBox := r.buildMathBox(lim, fontSize*0.7)
	w := base.w
	if limBox.w > w {
		w = limBox.w
	}
	if low {
		return mathBox{
			w:       w,
			ascent:  base.ascent,
			descent: base.descent + limBox.height() + 1,
			draw: func(r *renderer, x, baseline float64) {
				if base.draw != nil {
					base.draw(r, x+(w-base.w)/2, baseline)
				}
				if limBox.draw != nil {
					limBox.draw(r, x+(w-limBox.w)/2, baseline+base.descent+1+limBox.ascent)
				}
			},
		}
	}
	return mathBox{
		w:       w,
		ascent:  base.ascent + limBox.height() + 1,
		descent: base.descent,
		draw: func(r *renderer, x, baseline float64) {
			if base.draw != nil {
				base.draw(r, x+(w-base.w)/2, baseline)
			}
			if limBox.draw != nil {
				limBox.draw(r, x+(w-limBox.w)/2, baseline-base.ascent-1-limBox.descent)
			}
		},
	}
}

// mathMatrixBox lays the rows out in a grid with uniform column spacing.
func (r *renderer) mathMatrixBox(n *docx.MathNode, fontSize float64) mathBox {
	if len(n.Rows) == 0 {
		return mathBox{}
	}
	cols := 0
	for _, r := range n.Rows {
		if len(r) > cols {
			cols = len(r)
		}
	}
	cellBoxes := make([][]mathBox, len(n.Rows))
	colW := make([]float64, cols)
	rowH := make([]float64, len(n.Rows))
	for i, row := range n.Rows {
		cellBoxes[i] = make([]mathBox, cols)
		for j := 0; j < cols; j++ {
			if j < len(row) {
				cellBoxes[i][j] = r.buildMathBox(row[j], fontSize)
			}
			if cellBoxes[i][j].w > colW[j] {
				colW[j] = cellBoxes[i][j].w
			}
			h := cellBoxes[i][j].height()
			if h > rowH[i] {
				rowH[i] = h
			}
		}
	}
	const colGap = 6.0
	const rowGap = 2.0
	totalW := 0.0
	for _, w := range colW {
		totalW += w + colGap
	}
	if totalW > 0 {
		totalW -= colGap
	}
	totalH := 0.0
	for _, h := range rowH {
		totalH += h + rowGap
	}
	if totalH > 0 {
		totalH -= rowGap
	}
	return mathBox{
		w:       totalW + 4,
		ascent:  totalH / 2,
		descent: totalH / 2,
		draw: func(r *renderer, x, baseline float64) {
			y := baseline - totalH/2
			for i, row := range cellBoxes {
				cx := x + 2
				for j := 0; j < cols; j++ {
					if row[j].draw != nil {
						// Center each cell in its column; baseline at row mid-line.
						cellAsc := row[j].ascent
						row[j].draw(r, cx+(colW[j]-row[j].w)/2, y+cellAsc)
					}
					cx += colW[j] + colGap
				}
				y += rowH[i] + rowGap
			}
		},
	}
}
