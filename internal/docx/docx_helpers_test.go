package docx

import "testing"

func TestCnfStyleAny(t *testing.T) {
	tests := []struct {
		name string
		s    CnfStyle
		want bool
	}{
		{"empty", CnfStyle{}, false},
		{"first_row", CnfStyle{FirstRow: true}, true},
		{"last_col", CnfStyle{LastColumn: true}, true},
		{"all_false", CnfStyle{FirstRow: false, LastRow: false}, false},
	}
	for _, tt := range tests {
		got := tt.s.Any()
		if got != tt.want {
			t.Errorf("%s: CnfStyle.Any = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestBorderEdgeHas(t *testing.T) {
	tests := []struct {
		name string
		e    BorderEdge
		want bool
	}{
		{"empty", BorderEdge{}, false},
		{"style", BorderEdge{Style: "single"}, true},
		{"size", BorderEdge{Sz: 1}, true},
		{"color", BorderEdge{Color: "000000"}, true},
		{"art", BorderEdge{Art: "foo"}, true},
	}
	for _, tt := range tests {
		got := tt.e.Has()
		if got != tt.want {
			t.Errorf("%s: BorderEdge.Has = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestTableBordersHas(t *testing.T) {
	tests := []struct {
		name string
		b    TableBorders
		want bool
	}{
		{"empty", TableBorders{}, false},
		{"top", TableBorders{Top: BorderEdge{Style: "single"}}, true},
		{"bottom", TableBorders{Bottom: BorderEdge{Style: "single"}}, true},
		{"right", TableBorders{Right: BorderEdge{Style: "single"}}, true},
		{"inside_h", TableBorders{InsideH: BorderEdge{Sz: 1}}, true},
	}
	for _, tt := range tests {
		got := tt.b.Has()
		if got != tt.want {
			t.Errorf("%s: TableBorders.Has = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestTableCellParagraphs(t *testing.T) {
	tests := []struct {
		name string
		cell TableCell
		want int
	}{
		{"empty", TableCell{}, 0},
		{"one_para", TableCell{Blocks: []Block{Paragraph{Runs: []Run{{Text: "hi"}}}}}, 1},
		{"mixed", TableCell{Blocks: []Block{Paragraph{}, Table{}, Paragraph{}}}, 2},
	}
	for _, tt := range tests {
		got := tt.cell.Paragraphs()
		if len(got) != tt.want {
			t.Errorf("%s: Paragraphs returned %d, want %d", tt.name, len(got), tt.want)
		}
	}
}

func TestChartStyleSummaryHasAny(t *testing.T) {
	tests := []struct {
		name string
		s    ChartStyleSummary
		want bool
	}{
		{"empty", ChartStyleSummary{}, false},
		{"title_font", ChartStyleSummary{TitleFontSizePt: 12}, true},
		{"cat_axis", ChartStyleSummary{CatAxisFontSizePt: 10}, true},
		{"val_axis", ChartStyleSummary{ValAxisFontSizePt: 10}, true},
		{"axis_title", ChartStyleSummary{AxisTitleFontSizePt: 10}, true},
		{"data_label", ChartStyleSummary{DataLabelFontSizePt: 10}, true},
	}
	for _, tt := range tests {
		got := tt.s.HasAny()
		if got != tt.want {
			t.Errorf("%s: HasAny = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestRequiresSupported(t *testing.T) {
	tests := []struct {
		reqs string
		want bool
	}{
		{"", true},
		{"w14", true},
		{"w14 w15", true},
		{"w14 unknown", false},
		{"unknown", false},
	}
	for _, tt := range tests {
		got := requiresSupported(tt.reqs)
		if got != tt.want {
			t.Errorf("requiresSupported(%q) = %v, want %v", tt.reqs, got, tt.want)
		}
	}
}

func TestIsChartExRel(t *testing.T) {
	tests := []struct {
		rel  string
		want bool
	}{
		{"http://schemas.microsoft.com/office/2017/06/relationships/chartEx", true},
		{"http://schemas.microsoft.com/office/2017/06/relationships/chartex", true},
		{"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", false},
		{"", false},
	}
	for _, tt := range tests {
		got := isChartExRel(tt.rel)
		if got != tt.want {
			t.Errorf("isChartExRel(%q) = %v, want %v", tt.rel, got, tt.want)
		}
	}
}

func TestThemeFor(t *testing.T) {
	if th := themeFor(nil); len(th.Colors) != 0 {
		t.Error("themeFor(nil) should return empty Theme")
	}
	doc := &Document{Theme: Theme{Colors: map[string]string{"a": "FF0000"}}}
	th := themeFor(doc)
	if th.Colors["a"] != "FF0000" {
		t.Errorf("themeFor(doc) should return doc's Theme")
	}
}

func TestTrimHash(t *testing.T) {
	tests := []struct {
		in, want string
	}{
		{"#FF0000", "FF0000"},
		{"FF0000", "FF0000"},
		{"", ""},
		{"#", ""},
	}
	for _, tt := range tests {
		got := trimHash(tt.in)
		if got != tt.want {
			t.Errorf("trimHash(%q) = %q, want %q", tt.in, got, tt.want)
		}
	}
}

func TestFlattenEmbeddedDocxText(t *testing.T) {
	tests := []struct {
		name   string
		blocks []Block
		want   string
	}{
		{"empty", nil, ""},
		{"single_para", []Block{Paragraph{Runs: []Run{{Text: "hello"}}}}, "hello"},
		{"multi_para", []Block{
			Paragraph{Runs: []Run{{Text: "line1"}}},
			Paragraph{Runs: []Run{{Text: "line2"}}},
		}, "line1\nline2"},
		{"simple_table", []Block{
			Table{Rows: []TableRow{{Cells: []TableCell{
				{Blocks: []Block{Paragraph{Runs: []Run{{Text: "a"}}}}},
				{Blocks: []Block{Paragraph{Runs: []Run{{Text: "b"}}}}},
			}}}},
		}, "a\t\nb\n"},
	}
	for _, tt := range tests {
		got := flattenEmbeddedDocxText(tt.blocks)
		if got != tt.want {
			t.Errorf("%s: flattenEmbeddedDocxText = %q, want %q", tt.name, got, tt.want)
		}
	}
}
