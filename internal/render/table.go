package render

import (
	"github.com/bobyeoh/docx2pdf-go/internal/docx"
	"github.com/signintech/gopdf"
)

// applyTableStyleToCells merges the named tblStyle's run defaults into
// every cell text run that doesn't already specify its own value, then
// layers in w:tblLook conditional emphasis: firstRow / lastRow /
// firstColumn / lastColumn / band1Horz / band2Horz / band1Vert /
// band2Vert / nwCell / neCell / swCell / seCell.
//
// The CONDITIONAL blocks come from the named tblStyle's <w:tblStylePr>
// children, parsed into ts.Conditional[<type>]. When a cell matches more
// than one condition (e.g. firstRow + firstColumn = nwCell) the more
// specific corner condition wins.
func (r *renderer) applyTableStyleToCells(t *docx.Table) {
	if t.StyleID == "" {
		return
	}
	ts, ok := r.doc.TableStyles[t.StyleID]
	if !ok {
		return
	}
	nRows := len(t.Rows)
	for ri := range t.Rows {
		nCols := len(t.Rows[ri].Cells)
		for ci := range t.Rows[ri].Cells {
			cell := &t.Rows[ri].Cells[ci]
			conds := matchingConditions(t.Look, ri, ci, nRows, nCols)

			// Build the merged extra props/shading/borders from the
			// matching condition blocks. Later entries in `conds`
			// override earlier ones (specificity order: bands → first/
			// last → corners).
			extra := docx.RunProps{}
			condShading := ""
			var condBorders docx.CellBorders
			for _, key := range conds {
				cb, ok := ts.Conditional[key]
				if !ok {
					continue
				}
				extra = docx.MergeRunProps(extra, cb.Run)
				if cb.CellShading != "" {
					condShading = cb.CellShading
				}
				if cb.Borders.Top.Has() {
					condBorders.Top = cb.Borders.Top
				}
				if cb.Borders.Bottom.Has() {
					condBorders.Bottom = cb.Borders.Bottom
				}
				if cb.Borders.Left.Has() {
					condBorders.Left = cb.Borders.Left
				}
				if cb.Borders.Right.Has() {
					condBorders.Right = cb.Borders.Right
				}
			}
			// Backstop: keep the legacy "bold the first row/col" fallback
			// when the named style didn't ship a conditional block (e.g.
			// a one-off style that only declared firstRow=on in tblLook).
			if t.Look.FirstRow && ri == 0 && !hasCondition(ts, "firstRow") {
				extra.Bold = true
			}
			if t.Look.LastRow && ri == nRows-1 && nRows > 1 && !hasCondition(ts, "lastRow") {
				extra.Bold = true
			}
			if t.Look.FirstColumn && ci == 0 && !hasCondition(ts, "firstCol") {
				extra.Bold = true
			}
			if t.Look.LastColumn && ci == nCols-1 && nCols > 1 && !hasCondition(ts, "lastCol") {
				extra.Bold = true
			}

			// Cell-level overrides: shading only when the cell didn't set
			// its own; same for borders (a missing edge inherits the
			// condition's edge if any).
			if cell.Shading == "" && condShading != "" {
				cell.Shading = condShading
			}
			if condBorders.Top.Has() && !cell.Borders.Top.Has() {
				cell.Borders.Top = condBorders.Top
			}
			if condBorders.Bottom.Has() && !cell.Borders.Bottom.Has() {
				cell.Borders.Bottom = condBorders.Bottom
			}
			if condBorders.Left.Has() && !cell.Borders.Left.Has() {
				cell.Borders.Left = condBorders.Left
			}
			if condBorders.Right.Has() && !cell.Borders.Right.Has() {
				cell.Borders.Right = condBorders.Right
			}

			for bi := range cell.Blocks {
				p, ok := cell.Blocks[bi].(docx.Paragraph)
				if !ok {
					continue
				}
				for k := range p.Runs {
					base := docx.MergeRunProps(ts.Run, extra)
					p.Runs[k].Props = docx.MergeRunProps(base, p.Runs[k].Props)
				}
				cell.Blocks[bi] = p
			}
		}
	}
}

func hasCondition(ts docx.TableStyle, name string) bool {
	_, ok := ts.Conditional[name]
	return ok
}

// matchingConditions returns the list of condition keys that apply to
// the cell at (ri, ci) in row order: bands first (lowest priority), then
// firstRow/lastRow/firstCol/lastCol, finally corner cells (highest
// priority). The caller merges in that order so specific entries win.
func matchingConditions(look docx.TableLook, ri, ci, nRows, nCols int) []string {
	var out []string
	// Banding (zero-indexed). Banding fires only when the relevant
	// "NoHBand"/"NoVBand" flag is clear.
	if !look.NoHBand && nRows > 1 {
		if ri%2 == 0 {
			out = append(out, "band1Horz")
		} else {
			out = append(out, "band2Horz")
		}
	}
	if !look.NoVBand && nCols > 1 {
		if ci%2 == 0 {
			out = append(out, "band1Vert")
		} else {
			out = append(out, "band2Vert")
		}
	}
	if look.FirstRow && ri == 0 {
		out = append(out, "firstRow")
	}
	if look.LastRow && ri == nRows-1 && nRows > 1 {
		out = append(out, "lastRow")
	}
	if look.FirstColumn && ci == 0 {
		out = append(out, "firstCol")
	}
	if look.LastColumn && ci == nCols-1 && nCols > 1 {
		out = append(out, "lastCol")
	}
	// Corner cells override the row+col combination.
	if look.FirstRow && look.FirstColumn && ri == 0 && ci == 0 {
		out = append(out, "nwCell")
	}
	if look.FirstRow && look.LastColumn && ri == 0 && ci == nCols-1 && nCols > 1 {
		out = append(out, "neCell")
	}
	if look.LastRow && look.FirstColumn && ri == nRows-1 && ci == 0 && nRows > 1 {
		out = append(out, "swCell")
	}
	if look.LastRow && look.LastColumn && ri == nRows-1 && ci == nCols-1 && nRows > 1 && nCols > 1 {
		out = append(out, "seCell")
	}
	return out
}

func (r *renderer) drawTable(t docx.Table) error {
	r.applyTableStyleToCells(&t)
	r.resolveHMerge(&t)
	r.resolveAdjacentBorders(&t)
	r.resolveVMerge(&t)
	var tblBarTop float64
	if r.opts.ShowRevisions && tableHasRevision(t) {
		tblBarTop = r.cursorY
		defer func() {
			r.drawRevisionChangeBar(tblBarTop, r.cursorY)
		}()
	}
	cols := 0
	for _, row := range t.Rows {
		if len(row.Cells) > cols {
			cols = len(row.Cells)
		}
	}
	if cols == 0 {
		return nil
	}
	widths := r.resolveColumnWidths(t, cols)
	r.activeTableSpacing = t.CellSpacingTwips
	defer func() { r.activeTableSpacing = 0 }()

	// w:tblInd shifts the table's starting X by the indent amount. We
	// implement it by temporarily widening marL and shrinking contentW for
	// the duration of the table draw, so all per-row x math just works.
	tblIndent := 0.0
	if t.IndentTwips != 0 {
		tblIndent = float64(t.IndentTwips) / 20.0
	}

	// Floating table (tblpPr): full text-wrap isn't implemented but we
	// at least honor the requested horizontal alignment and absolute
	// X offset so the table lands roughly where the source asked.
	if t.FloatPos != nil {
		fp := t.FloatPos
		sum := 0.0
		for _, w := range widths {
			sum += w
		}
		// Reference frame for X position depends on hAnchor: "margin"
		// (default) anchors at marL, "page" anchors at page edge.
		anchorX := r.marL
		anchorW := r.contentW
		if fp.HAnchor == "page" {
			anchorX = 0
			anchorW = r.pageW
		}
		off := 0.0
		switch fp.XAlign {
		case "center":
			off = (anchorW - sum) / 2
		case "right":
			off = anchorW - sum
		case "outside":
			off = anchorW - sum
		case "left", "inside":
			off = 0
		default:
			if fp.XTwips != 0 {
				off = float64(fp.XTwips) / 20.0
			}
		}
		newMarL := anchorX + off
		if newMarL < 0 {
			newMarL = 0
		}
		savedMarL, savedContentW := r.marL, r.contentW
		r.marL = newMarL
		r.contentW = sum
		defer func() {
			r.marL = savedMarL
			r.contentW = savedContentW
		}()
		// Disable tblIndent on floating tables — XAlign already places
		// the table's left edge.
		tblIndent = 0
	}

	if tblIndent != 0 {
		savedMarL, savedContentW := r.marL, r.contentW
		r.marL += tblIndent
		r.contentW -= tblIndent
		defer func() {
			r.marL = savedMarL
			r.contentW = savedContentW
		}()
	}

	// Header rows repeat after each page break (leading consecutive
	// header-flagged rows per ECMA-376).
	var headerRows []docx.TableRow
	for _, row := range t.Rows {
		if !row.IsHeader {
			break
		}
		headerRows = append(headerRows, row)
	}

	savedCtx := r.fields.tableCtx
	defer func() { r.fields.tableCtx = savedCtx }()
	r.fields.tableCtx = &tableContext{table: &t}

	// Pre-measure all row heights so each restart cell in a vMerge group
	// can stash the cumulative height onto its MergedHeightPt for vAlign.
	r.populateMergedHeights(&t, widths)

	for i, row := range t.Rows {
		r.fields.tableCtx.row = i
		// Pre-flight: if this is a body row and it won't fit on the
		// current page, force a page break and re-draw the header rows
		// BEFORE the row that triggered the break. Without this the
		// header lands after the row, mid-page, on every page where
		// the table continues.
		if len(headerRows) > 0 && i >= len(headerRows) && !r.noPageBreak {
			rowH := r.predictRowHeight(row, widths)
			if r.cursorY+rowH > r.pageH-r.marB {
				r.drawFootnotesAtBottom()
				r.newPage()
				for _, hr := range headerRows {
					if err := r.drawRow(hr, widths); err != nil {
						return err
					}
				}
			}
		}
		if err := r.drawRow(row, widths); err != nil {
			return err
		}
	}
	return nil
}

// resolveColumnWidths returns the per-column widths in points. It honors,
// in priority order:
//  1. The table's w:tblGrid (most reliable; one entry per column).
//  2. Per-cell w:tcW in the first non-merged row (used when tblGrid is
//     missing or doesn't match column count).
//  3. Equal division of contentW (fallback).
//
// When the table's w:tblW declares a pct (50% of contentW), the column
// widths are scaled to match. When w:tblLayout is "autofit" and the table
// is wider than contentW, columns are proportionally shrunk so the table
// fits — the previous behavior was to overflow the page margin.
func (r *renderer) resolveColumnWidths(t docx.Table, cols int) []float64 {
	widths := make([]float64, cols)

	// First, derive a starting set of widths.
	gridUsable := len(t.ColumnWidthsTwips) == cols
	if gridUsable {
		gridTotal := 0
		for _, w := range t.ColumnWidthsTwips {
			gridTotal += w
		}
		if gridTotal > 0 {
			for i, w := range t.ColumnWidthsTwips {
				widths[i] = float64(w) / 20.0
			}
		} else {
			gridUsable = false
		}
	}
	if !gridUsable {
		// Try to recover column widths from the first row's tcW values
		// (covers tables that ship without a w:tblGrid).
		if len(t.Rows) > 0 {
			row := t.Rows[0]
			col := 0
			anyTcW := false
			for _, cell := range row.Cells {
				span := cell.GridSpan
				if span < 1 {
					span = 1
				}
				switch cell.CellWidthType {
				case "dxa":
					if cell.CellWidthTwips > 0 {
						per := float64(cell.CellWidthTwips) / 20.0 / float64(span)
						for k := 0; k < span && col+k < cols; k++ {
							widths[col+k] = per
						}
						anyTcW = true
					}
				case "pct":
					// pct stored in 50ths of a percent (5000 = 100%).
					if cell.CellWidthTwips > 0 {
						frac := float64(cell.CellWidthTwips) / 5000.0
						per := frac * r.contentW / float64(span)
						for k := 0; k < span && col+k < cols; k++ {
							widths[col+k] = per
						}
						anyTcW = true
					}
				}
				col += span
			}
			if !anyTcW {
				for i := range widths {
					widths[i] = r.contentW / float64(cols)
				}
			} else {
				// Fill unset columns with the mean of the known widths.
				known, n := 0.0, 0
				for _, w := range widths {
					if w > 0 {
						known += w
						n++
					}
				}
				if n > 0 {
					mean := known / float64(n)
					for i := range widths {
						if widths[i] == 0 {
							widths[i] = mean
						}
					}
				}
			}
		} else {
			for i := range widths {
				widths[i] = r.contentW / float64(cols)
			}
		}
	}

	// Apply w:tblW pct override: column widths are scaled to occupy
	// (pct / 5000) of contentW. Word stores pct in 50ths of a percent
	// (5000 = 100%).
	if t.TableWidthType == "pct" && t.TableWidthTwips > 0 {
		target := float64(t.TableWidthTwips) / 5000.0 * r.contentW
		sum := 0.0
		for _, w := range widths {
			sum += w
		}
		if sum > 0 {
			scale := target / sum
			for i := range widths {
				widths[i] *= scale
			}
		}
	}

	// w:tblLayout="autofit" (default in Word for non-fixed tables) plus
	// total > contentW: scale columns down so the table fits. With
	// Layout="fixed" we respect the absolute widths even if they
	// overflow, since the source explicitly asked for that.
	if t.Layout != "fixed" {
		sum := 0.0
		for _, w := range widths {
			sum += w
		}
		if sum > r.contentW*1.05 {
			scale := r.contentW / sum
			for i := range widths {
				widths[i] *= scale
			}
		}
		// When the table has total width slack AND any column is too
		// narrow to hold its widest unbreakable atom (e.g. an English
		// word), redistribute slack from wide columns toward those
		// narrow ones. This is the most visible piece of "true autofit"
		// the user can ask for without breaking fixed-grid templates.
		r.distributeAutofitSlack(t, widths)
	}
	return widths
}

// distributeAutofitSlack walks the table's first row, measures each
// cell's intrinsic minimum width (widest unbreakable atom), and shifts
// slack from wider-than-needed columns toward the columns that are
// narrower than their minimum. Operates in-place on widths.
func (r *renderer) distributeAutofitSlack(t docx.Table, widths []float64) {
	if len(t.Rows) == 0 {
		return
	}
	// Compute per-column minimum content width across every row.
	mins := make([]float64, len(widths))
	for _, row := range t.Rows {
		col := 0
		for _, cell := range row.Cells {
			if col >= len(widths) {
				break
			}
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			if span == 1 && cell.VMerge != "continue" {
				if m := r.cellMinWidth(cell); m > mins[col] {
					mins[col] = m
				}
			}
			col += span
		}
	}
	// Find under-allocated columns vs over-allocated ones.
	deficit := 0.0
	surplus := 0.0
	type entry struct {
		idx  int
		diff float64
	}
	var needers, donors []entry
	for i, w := range widths {
		if mins[i] > w+0.5 {
			needers = append(needers, entry{i, mins[i] - w})
			deficit += mins[i] - w
		} else if w > mins[i]+8 {
			donors = append(donors, entry{i, w - mins[i] - 4})
			surplus += w - mins[i] - 4
		}
	}
	if len(needers) == 0 || surplus <= 0 {
		return
	}
	take := deficit
	if take > surplus {
		take = surplus
	}
	// Pro-rata donate from donors, fill needers.
	for _, d := range donors {
		give := take * (d.diff / surplus)
		widths[d.idx] -= give
	}
	givenSum := 0.0
	for _, n := range needers {
		givenSum += n.diff
	}
	for _, n := range needers {
		got := take * (n.diff / givenSum)
		widths[n.idx] += got
	}
}

// cellMinWidth measures the widest unbreakable atom inside a cell so
// distributeAutofitSlack can guarantee that atom fits. Side-effect free:
// pendingFootnotes is saved/restored.
func (r *renderer) cellMinWidth(cell docx.TableCell) float64 {
	savedFootnotes := r.pendingFootnotes
	defer func() { r.pendingFootnotes = savedFootnotes }()
	maxW := 0.0
	for _, p := range cell.Paragraphs() {
		for _, a := range r.runsToAtoms(p.Runs) {
			if a.kind == atomWord && a.width > maxW {
				maxW = a.width
			}
		}
	}
	// Pad for cell margins (4pt each side default).
	return maxW + 8
}

// resolveAdjacentBorders applies the OOXML §17.4.66 border-conflict
// resolution between adjacent cells: when two neighboring cells declare
// a shared edge, the visually heavier edge wins. We resolve each pair
// once (the winner is written to BOTH cells so drawCellEdge — which has
// no global view — paints consistently when it sees either side).
func (r *renderer) resolveAdjacentBorders(t *docx.Table) {
	// Horizontal adjacency: cell.Right vs next-cell.Left in each row.
	for ri := range t.Rows {
		row := &t.Rows[ri]
		for ci := 0; ci+1 < len(row.Cells); ci++ {
			a := &row.Cells[ci]
			b := &row.Cells[ci+1]
			win := pickBorder(a.Borders.Right, b.Borders.Left)
			a.Borders.Right = win
			b.Borders.Left = win
		}
	}
	// Vertical adjacency: cell.Bottom vs next-row-cell.Top, column-aligned.
	if len(t.Rows) < 2 {
		return
	}
	// Build a logical-column index per row (account for gridSpan).
	rowColMap := make([]map[int]int, len(t.Rows))
	for ri := range t.Rows {
		rowColMap[ri] = map[int]int{}
		col := 0
		for ci, cell := range t.Rows[ri].Cells {
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			rowColMap[ri][col] = ci
			col += span
		}
	}
	for ri := 0; ri+1 < len(t.Rows); ri++ {
		rowA := &t.Rows[ri]
		rowB := &t.Rows[ri+1]
		for col, ciA := range rowColMap[ri] {
			ciB, ok := rowColMap[ri+1][col]
			if !ok {
				continue
			}
			a := &rowA.Cells[ciA]
			b := &rowB.Cells[ciB]
			win := pickBorder(a.Borders.Bottom, b.Borders.Top)
			a.Borders.Bottom = win
			b.Borders.Top = win
		}
	}
}

// pickBorder applies §17.4.66 conflict resolution between two border
// specs declared on a shared edge: heavier wins; on a tie, darker color
// wins; ties beyond that favor the first argument.
func pickBorder(a, b docx.BorderEdge) docx.BorderEdge {
	// "none"/"nil" means the cell explicitly declined this edge — let the
	// other side win.
	if a.Style == "none" || a.Style == "nil" {
		return b
	}
	if b.Style == "none" || b.Style == "nil" {
		return a
	}
	if !a.Has() {
		return b
	}
	if !b.Has() {
		return a
	}
	if a.Sz > b.Sz {
		return a
	}
	if b.Sz > a.Sz {
		return b
	}
	if borderColorWeight(a.Color) > borderColorWeight(b.Color) {
		return b
	}
	return a
}

// borderColorWeight returns a non-negative "lightness" score so darker
// colors win the tie-break. Empty / "auto" treated as black (weight 0).
func borderColorWeight(hex string) int {
	if hex == "" || hex == "auto" {
		return 0
	}
	if len(hex) != 6 {
		return 0
	}
	parse := func(s string) int {
		var n int
		for _, c := range s {
			n *= 16
			switch {
			case c >= '0' && c <= '9':
				n += int(c - '0')
			case c >= 'A' && c <= 'F':
				n += int(c-'A') + 10
			case c >= 'a' && c <= 'f':
				n += int(c-'a') + 10
			}
		}
		return n
	}
	return parse(hex[:2]) + parse(hex[2:4]) + parse(hex[4:])
}

// populateMergedHeights walks every vMerge="restart" cell and sums the
// pre-measured row heights of itself plus every continue cell below in
// the same logical column. The result is stashed on cell.MergedHeightPt
// so drawRow can use it for vAlign math. We accept that this pre-flight
// re-runs the layout dry-pass; the cost is bounded by table size and
// not by content beyond the cells.
func (r *renderer) populateMergedHeights(t *docx.Table, widths []float64) {
	if len(t.Rows) == 0 {
		return
	}
	// Pre-measure every row once.
	rowHeights := make([]float64, len(t.Rows))
	for ri := range t.Rows {
		rowHeights[ri] = r.predictRowHeight(t.Rows[ri], widths)
	}
	// Logical column → cell-index map per row.
	colMap := make([]map[int]int, len(t.Rows))
	for ri := range t.Rows {
		colMap[ri] = map[int]int{}
		col := 0
		for ci, cell := range t.Rows[ri].Cells {
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			colMap[ri][col] = ci
			col += span
		}
	}
	for ri := range t.Rows {
		col := 0
		for ci := range t.Rows[ri].Cells {
			cell := &t.Rows[ri].Cells[ci]
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			if cell.VMerge == "restart" {
				total := rowHeights[ri]
				for rj := ri + 1; rj < len(t.Rows); rj++ {
					nextCi, ok := colMap[rj][col]
					if !ok {
						break
					}
					if t.Rows[rj].Cells[nextCi].VMerge != "continue" {
						break
					}
					total += rowHeights[rj]
				}
				cell.MergedHeightPt = total
			}
			col += span
		}
	}
}

// resolveVMerge marks each cell that participates in a vertical merge
// group (other than the LAST cell in the group) so its bottom edge will
// be suppressed at draw time. Without this, a vMerge="restart" cell
// would draw a bottom border at the end of its own row, producing a
// horizontal divider inside what should be one merged region.
//
// The same column position in the next row holds either a continue
// cell (still inside the merge) or anything else (group ends here, so
// keep the bottom border).
func (r *renderer) resolveVMerge(t *docx.Table) {
	if len(t.Rows) < 2 {
		return
	}
	// Logical column → cell-index map per row (account for gridSpan).
	colMap := make([]map[int]int, len(t.Rows))
	for ri := range t.Rows {
		colMap[ri] = map[int]int{}
		col := 0
		for ci, cell := range t.Rows[ri].Cells {
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			colMap[ri][col] = ci
			col += span
		}
	}
	for ri := 0; ri+1 < len(t.Rows); ri++ {
		col := 0
		for ci := range t.Rows[ri].Cells {
			cell := &t.Rows[ri].Cells[ci]
			span := cell.GridSpan
			if span < 1 {
				span = 1
			}
			// Only restart/continue cells live in a merge group.
			if cell.VMerge == "restart" || cell.VMerge == "continue" {
				if nextCi, ok := colMap[ri+1][col]; ok {
					next := t.Rows[ri+1].Cells[nextCi]
					if next.VMerge == "continue" {
						cell.SuppressBottomBorder = true
					}
				}
			}
			col += span
		}
	}
}

// resolveHMerge folds w:hMerge="continue" cells into their preceding
// "restart" cell by widening that cell's GridSpan. We then drop the
// continuation cells from the row. This matches what GridSpan would have
// produced if the doc had used it instead.
func (r *renderer) resolveHMerge(t *docx.Table) {
	for ri := range t.Rows {
		row := &t.Rows[ri]
		// Walk left-to-right; collapse runs of "continue" into the
		// nearest preceding non-continue cell.
		out := make([]docx.TableCell, 0, len(row.Cells))
		for _, cell := range row.Cells {
			if cell.HMerge == "continue" && len(out) > 0 {
				last := &out[len(out)-1]
				span := cell.GridSpan
				if span < 1 {
					span = 1
				}
				if last.GridSpan < 1 {
					last.GridSpan = 1
				}
				last.GridSpan += span
				continue
			}
			out = append(out, cell)
		}
		row.Cells = out
	}
}

// predictRowHeight computes the row's rendered height without drawing
// anything. Used by drawTable for pre-flight page-break detection so we
// can inject the repeating header BEFORE the row that overflows
// (otherwise the header lands after the row, mid-page).
func (r *renderer) predictRowHeight(row docx.TableRow, widths []float64) float64 {
	cellHeights := make([]float64, len(row.Cells))
	col := 0
	for i, cell := range row.Cells {
		if col >= len(widths) {
			break
		}
		span := cell.GridSpan
		if span < 1 {
			span = 1
		}
		w := sumWidths(widths, col, span)
		if cell.VMerge == "continue" {
			cellHeights[i] = 0
		} else {
			cellHeights[i] = r.measureCell(cell, w)
		}
		col += span
	}
	rowH := 0.0
	for _, h := range cellHeights {
		if h > rowH {
			rowH = h
		}
	}
	if rowH < r.opts.DefaultFontSize*1.4 {
		rowH = r.opts.DefaultFontSize * 1.4
	}
	if row.HeightTwips > 0 {
		minH := float64(row.HeightTwips) / 20.0
		if row.HeightRuleExact || minH > rowH {
			rowH = minH
		}
	}
	return rowH
}

func (r *renderer) drawRow(row docx.TableRow, widths []float64) error {
	rowTop := r.cursorY
	cellHeights := make([]float64, len(row.Cells))
	col := 0
	for i, cell := range row.Cells {
		if col >= len(widths) {
			break
		}
		span := cell.GridSpan
		if span < 1 {
			span = 1
		}
		w := sumWidths(widths, col, span)
		if cell.VMerge == "continue" {
			cellHeights[i] = 0
		} else {
			cellHeights[i] = r.measureCell(cell, w)
		}
		col += span
	}
	rowH := 0.0
	for _, h := range cellHeights {
		if h > rowH {
			rowH = h
		}
	}
	if rowH < r.opts.DefaultFontSize*1.4 {
		rowH = r.opts.DefaultFontSize * 1.4
	}
	if row.HeightTwips > 0 {
		minH := float64(row.HeightTwips) / 20.0
		if row.HeightRuleExact || minH > rowH {
			rowH = minH
		}
	}

	// CantSplit: if the row won't fit on the current page, push it to the
	// next page intact rather than letting ensureRoom break it mid-row.
	// ensureRoom is already conservative when noPageBreak is set (header /
	// footer regions), so we only act when free flow is in effect.
	if row.CantSplit && !r.noPageBreak && r.cursorY+rowH > r.pageH-r.marB {
		r.drawFootnotesAtBottom()
		r.newPage()
		rowTop = r.cursorY
	} else {
		r.ensureRoom(rowH)
		if r.cursorY != rowTop {
			rowTop = r.cursorY
		}
	}

	r.pdf.SetLineWidth(0.5)
	r.pdf.SetStrokeColor(0, 0, 0)

	cellSpacingTwips := row.CellSpacingTwips
	if cellSpacingTwips == 0 {
		cellSpacingTwips = r.activeTableSpacing
	}
	cellSpacingPt := float64(cellSpacingTwips) / 20.0

	x := r.marL
	col = 0
	const defaultCellPad = 4.0
	for ci, cell := range row.Cells {
		if col >= len(widths) {
			break
		}
		span := cell.GridSpan
		if span < 1 {
			span = 1
		}
		w := sumWidths(widths, col, span)

		padL := cell.MarginLeftPt
		if padL == 0 {
			padL = defaultCellPad
		}
		padR := cell.MarginRightPt
		if padR == 0 {
			padR = defaultCellPad
		}
		padT := cell.MarginTopPt
		if padT == 0 {
			padT = defaultCellPad
		}
		padB := cell.MarginBottomPt
		if padB == 0 {
			padB = defaultCellPad
		}

		left, right := x, x+w
		top, bottom := rowTop, rowTop+rowH
		if cellSpacingPt > 0 {
			left += cellSpacingPt
			right -= cellSpacingPt
			top += cellSpacingPt
			bottom -= cellSpacingPt
		}

		if cell.Shading != "" {
			sr, sg, sb := parseHexColor(cell.Shading)
			r.pdf.SetFillColor(sr, sg, sb)
			r.pdf.Rectangle(left, top, right, bottom, "F", 0, 0)
		}

		// Continuation cells suppress the top edge so the vMerge region
		// looks like one connected box. resolveVMerge marks any cell
		// whose bottom edge is shared with a continue cell below; we
		// also suppress that edge so the merged region renders without
		// internal horizontal dividers.
		if cell.VMerge != "continue" {
			drawCellEdge(r, cell.Borders.Top, left, top, right, top)
		}
		if !cell.SuppressBottomBorder {
			drawCellEdge(r, cell.Borders.Bottom, left, bottom, right, bottom)
		}
		drawCellEdge(r, cell.Borders.Left, left, top, left, bottom)
		drawCellEdge(r, cell.Borders.Right, right, top, right, bottom)
		// Diagonal cell borders — common on "header corner" cells in row+
		// column-labeled matrices. They're drawn after the rectangular
		// edges so the diagonal sits on top of any shading.
		if cell.Borders.TL2BR.Has() {
			drawCellEdge(r, cell.Borders.TL2BR, left, top, right, bottom)
		}
		if cell.Borders.TR2BL.Has() {
			drawCellEdge(r, cell.Borders.TR2BL, right, top, left, bottom)
		}

		if cell.VMerge != "continue" {
			savedY := r.cursorY
			savedMarL := r.marL
			savedContentW := r.contentW
			r.marL = left + padL
			r.contentW = (right - left) - (padL + padR)
			if r.fields.tableCtx != nil {
				r.fields.tableCtx.col = col
			}
			// Content height is what measureCell returned MINUS the
			// pad-both-sides it added internally. Using the actual
			// measurement (not a one-line-per-paragraph stub) keeps
			// vAlign="center" from pushing wrapped content past the
			// row's bottom edge — multi-line cells in this row would
			// otherwise overflow when their column happens to be the
			// tallest.
			const cellPad = 4.0
			contentH := cellHeights[ci] - 2*cellPad
			if contentH < 0 {
				contentH = 0
			}
			// For a vMerge="restart" cell, the visual cell box spans
			// every row in the group, so vAlign centers/bottoms against
			// the merged-group height (set by resolveVMerge via the
			// pre-flight predictRowHeight pass). Fall back to the
			// current row's height when the cell isn't in a merge.
			boxH := rowH
			if cell.VMerge == "restart" && cell.MergedHeightPt > rowH {
				boxH = cell.MergedHeightPt
			}
			startY := rowTop + padT
			switch cell.VAlign {
			case "center":
				slack := boxH - contentH - (padT + padB)
				if slack > 0 {
					startY += slack / 2
				}
			case "bottom":
				slack := boxH - contentH - (padT + padB)
				if slack > 0 {
					startY += slack
				}
			}
			r.cursorY = startY
			clipped := row.HeightRuleExact && row.HeightTwips > 0
			if clipped {
				r.pdf.SaveGraphicsState()
				clipBox := []gopdf.Point{
					{X: left, Y: top},
					{X: right, Y: top},
					{X: right, Y: bottom},
					{X: left, Y: bottom},
				}
				r.pdf.ClipPolygon(clipBox)
			}
			for _, b := range cell.Blocks {
				switch v := b.(type) {
				case docx.Paragraph:
					if err := r.drawParagraph(v); err != nil {
						return err
					}
				case docx.Table:
					if err := r.drawTable(v); err != nil {
						return err
					}
				}
			}
			if clipped {
				r.pdf.RestoreGraphicsState()
			}
			r.marL = savedMarL
			r.contentW = savedContentW
			r.cursorY = savedY
		}

		x += w
		col += span
	}
	r.cursorY = rowTop + rowH
	return nil
}

// drawCellEdge draws one of a cell's four edges. An empty edge (zero
// BorderEdge) means "no border" — Word renders a table that lacks
// tblBorders/tcBorders without any lines, and we match that. Tables that
// want gridlines must declare tblBorders or tcBorders; the parser
// propagates tblBorders into each cell at parse time (see
// propagateTableBorders), so the renderer only needs to read CellBorders.
// Width is the Word-stored sz in points (1/8 pt units already converted
// upstream).
func drawCellEdge(r *renderer, e docx.BorderEdge, x1, y1, x2, y2 float64) {
	if !e.Has() || e.Style == "none" || e.Style == "nil" {
		return
	}
	width := e.Sz
	if width <= 0 {
		width = 0.5
	}
	if e.Color != "" && e.Color != "auto" {
		rr, gg, bb := parseHexColor(e.Color)
		r.pdf.SetStrokeColor(rr, gg, bb)
	} else {
		r.pdf.SetStrokeColor(0, 0, 0)
	}
	r.pdf.SetLineWidth(width)

	switch e.Style {
	case "double":
		offX, offY := 0.0, 0.0
		if y1 == y2 {
			offY = 1
		} else {
			offX = 1
		}
		r.pdf.Line(x1-offX, y1-offY, x2-offX, y2-offY)
		r.pdf.Line(x1+offX, y1+offY, x2+offX, y2+offY)
	case "dashed":
		drawDashedLine(r, x1, y1, x2, y2, 3, 2)
	case "dotted":
		drawDashedLine(r, x1, y1, x2, y2, 1, 2)
	default:
		r.pdf.Line(x1, y1, x2, y2)
	}
}

// drawDashedLine renders a dash pattern by stepping in fixed-length segments.
// gopdf has SetLineType but it's globally stateful and easy to leak — drawing
// the dashes ourselves keeps each call self-contained.
func drawDashedLine(r *renderer, x1, y1, x2, y2, dash, gap float64) {
	dx, dy := x2-x1, y2-y1
	length := dx*dx + dy*dy
	if length == 0 {
		return
	}
	if y1 == y2 {
		for x := x1; x < x2; x += dash + gap {
			end := x + dash
			if end > x2 {
				end = x2
			}
			r.pdf.Line(x, y1, end, y1)
		}
	} else if x1 == x2 {
		for y := y1; y < y2; y += dash + gap {
			end := y + dash
			if end > y2 {
				end = y2
			}
			r.pdf.Line(x1, y, x1, end)
		}
	}
}

// cellContentHeight estimates the rendered height of a cell's contents at
// the renderer's current contentW. Used for vAlign slack math.
func sumWidths(ws []float64, start, n int) float64 {
	sum := 0.0
	for i := start; i < start+n && i < len(ws); i++ {
		sum += ws[i]
	}
	return sum
}

// measureCell estimates rendered height for a cell at the given content
// width. Does a dry layout reusing the line-breaker math without drawing.
//
// runsToAtoms has the side effect of queuing footnote IDs onto
// pendingFootnotes, so we save and restore that slice — otherwise table
// cells with footnote refs would queue each note twice (once in measure,
// once in the real draw) and the page bottom would render duplicates.
func (r *renderer) measureCell(cell docx.TableCell, width float64) float64 {
	const cellPad = 4.0
	h := 2 * cellPad
	innerW := width - 2*cellPad
	savedLine := r.lineHeight
	savedFootnotes := r.pendingFootnotes
	defer func() {
		r.lineHeight = savedLine
		r.pendingFootnotes = savedFootnotes
	}()
	// Walk blocks in document order so nested tables and surrounding
	// paragraphs contribute their real height to the cell. The previous
	// implementation iterated only Paragraphs(), so a cell whose nested
	// table outgrew the implicit one-line height would overflow.
	for _, b := range cell.Blocks {
		switch v := b.(type) {
		case docx.Table:
			h += r.measureNestedTable(v, innerW)
			continue
		case docx.Paragraph:
			p := v
			r.lineHeight = p.LineHeight
			atoms := r.runsToAtoms(p.Runs)
			h += r.measureAtomsHeight(atoms, innerW)
			h += p.SpacingBefore + p.SpacingAfter
		}
	}
	return h
}

// measureNestedTable estimates the height of a table laid out inside a
// cell whose content area is `outerW` wide. We reuse the same column-
// width resolution and per-row measurement that drawTable would use,
// just without drawing. Borders aren't reasoned about explicitly — Word
// treats them as zero-width for layout.
func (r *renderer) measureNestedTable(t docx.Table, outerW float64) float64 {
	savedContentW := r.contentW
	r.contentW = outerW
	defer func() { r.contentW = savedContentW }()
	cols := 0
	for _, row := range t.Rows {
		if len(row.Cells) > cols {
			cols = len(row.Cells)
		}
	}
	if cols == 0 {
		return 0
	}
	r.resolveHMerge(&t)
	widths := r.resolveColumnWidths(t, cols)
	h := 0.0
	for _, row := range t.Rows {
		h += r.predictRowHeight(row, widths)
	}
	return h
}

// measureAtomsHeight is the dry-pass line breaker extracted from
// measureCell so both paragraph blocks and nested-table walks can reuse
// the same per-line accumulator. Returns the height contribution (does
// not include the cell's outer padding).
func (r *renderer) measureAtomsHeight(atoms []atom, innerW float64) float64 {
	h := 0.0
	lineW := 0.0
	lineH := r.applyLineHeight(r.opts.DefaultFontSize * 1.2)
	hadAny := false
	accumulate := func(a atom) {
		ah := atomHeight(a, r.opts.DefaultFontSize)
		if lineW+a.width > innerW && lineW > 0 {
			h += lineH
			lineW = 0
			lineH = r.applyLineHeight(r.opts.DefaultFontSize * 1.2)
			if a.kind == atomSpace {
				return
			}
		}
		lineW += a.width
		scaled := r.applyLineHeight(ah)
		if scaled > lineH {
			lineH = scaled
		}
		hadAny = true
	}
	for _, a := range atoms {
		if a.kind == atomBookmark {
			continue
		}
		if a.kind == atomBreak || a.kind == atomPageBreak {
			h += lineH
			lineW = 0
			lineH = r.applyLineHeight(r.opts.DefaultFontSize * 1.2)
			hadAny = false
			continue
		}
		if a.kind == atomWord && innerW > 0 && a.width > innerW && a.text != "" {
			if lineW > 0 {
				h += lineH
				lineW = 0
				lineH = r.applyLineHeight(r.opts.DefaultFontSize * 1.2)
			}
			if a.width > innerW {
				for _, sub := range r.splitWordAtomByRune(a) {
					accumulate(sub)
				}
				continue
			}
		}
		accumulate(a)
	}
	if hadAny || lineW > 0 || len(atoms) == 0 {
		h += lineH
	}
	return h
}

