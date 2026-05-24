package render

import (
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

func TestPickMargin(t *testing.T) {
	tests := []struct {
		override, base, want float64
	}{
		{0, 10, 10}, {5, 10, 5}, {0, 0, 0},
	}
	for _, tt := range tests {
		got := pickMargin(tt.override, tt.base)
		if got != tt.want {
			t.Errorf("pickMargin(%v,%v) = %v, want %v", tt.override, tt.base, got, tt.want)
		}
	}
}

func TestSumWidths(t *testing.T) {
	ws := []float64{10, 20, 30, 40}
	tests := []struct {
		start, n int
		want     float64
	}{
		{0, 2, 30}, {1, 2, 50}, {2, 5, 70}, {3, 0, 0},
	}
	for _, tt := range tests {
		got := sumWidths(ws, tt.start, tt.n)
		if got != tt.want {
			t.Errorf("sumWidths(%v,%d,%d) = %v, want %v", ws, tt.start, tt.n, got, tt.want)
		}
	}
}

func TestBorderColorWeight(t *testing.T) {
	tests := []struct {
		hex  string
		want int
	}{
		{"", 0},
		{"auto", 0},
		{"000000", 0},
		{"FFFFFF", 765},
		{"FF0000", 255},
		{"00FF00", 255},
		{"0000FF", 255},
		{"808080", 384},
		{"ABC", 0},
	}
	for _, tt := range tests {
		got := borderColorWeight(tt.hex)
		if got != tt.want {
			t.Errorf("borderColorWeight(%q) = %d, want %d", tt.hex, got, tt.want)
		}
	}
}

func TestPickBorder(t *testing.T) {
	thick := docx.BorderEdge{Style: "single", Sz: 3, Color: "000000"}
	thin := docx.BorderEdge{Style: "single", Sz: 1, Color: "000000"}
	none := docx.BorderEdge{Style: "none"}
	nilStyle := docx.BorderEdge{Style: "nil"}
	empty := docx.BorderEdge{}
	darker := docx.BorderEdge{Style: "single", Sz: 1, Color: "000000"}
	lighter := docx.BorderEdge{Style: "single", Sz: 1, Color: "CCCCCC"}
	tests := []struct {
		name string
		a, b docx.BorderEdge
		want docx.BorderEdge
	}{
		{"thick_wins", thick, thin, thick},
		{"thin_loses", thin, thick, thick},
		{"none_lets_other_win", none, thin, thin},
		{"nil_lets_other_win", nilStyle, thick, thick},
		{"empty_b_wins", empty, thin, thin},
		{"empty_a_wins", thin, empty, thin},
		{"same_size_darker_wins", lighter, darker, darker},
		{"same_size_darker_wins_b_first", darker, lighter, darker},
	}
	for _, tt := range tests {
		got := pickBorder(tt.a, tt.b)
		if got.Sz != tt.want.Sz || got.Color != tt.want.Color || got.Style != tt.want.Style {
			t.Errorf("%s: pickBorder = %+v, want %+v", tt.name, got, tt.want)
		}
	}
}

func TestHasCondition(t *testing.T) {
	ts := docx.TableStyle{
		Conditional: map[string]docx.TableCondPr{
			"firstRow": {},
		},
	}
	if !hasCondition(ts, "firstRow") {
		t.Error("hasCondition should find firstRow")
	}
	if hasCondition(ts, "lastRow") {
		t.Error("hasCondition should not find lastRow")
	}
	if hasCondition(docx.TableStyle{}, "firstRow") {
		t.Error("hasCondition on empty style should return false")
	}
}

func TestCellHasNonAtomContent(t *testing.T) {
	emptyCell := docx.TableCell{}
	if cellHasNonAtomContent(emptyCell) {
		t.Error("empty cell should not have non-atom content")
	}
	paraCell := docx.TableCell{
		Blocks: []docx.Block{
			docx.Paragraph{Runs: []docx.Run{{Text: "hello"}}},
		},
	}
	if cellHasNonAtomContent(paraCell) {
		t.Error("paragraph-only cell should not have non-atom content")
	}
	imageCell := docx.TableCell{
		Blocks: []docx.Block{
			docx.Paragraph{Runs: []docx.Run{{ImageID: "rId1"}}},
		},
	}
	if !cellHasNonAtomContent(imageCell) {
		t.Error("cell with image should have non-atom content")
	}
	tableCell := docx.TableCell{
		Blocks: []docx.Block{
			docx.Table{},
		},
	}
	if !cellHasNonAtomContent(tableCell) {
		t.Error("cell with nested table should have non-atom content")
	}
}
