package render

import (
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

func TestHasSeriesNames(t *testing.T) {
	tests := []struct {
		name   string
		series []docx.ChartSeries
		want   bool
	}{
		{"empty", nil, false},
		{"no_names", []docx.ChartSeries{{Values: []float64{1}}}, false},
		{"with_name", []docx.ChartSeries{{Name: "Sales"}}, true},
	}
	for _, tt := range tests {
		got := hasSeriesNames(tt.series)
		if got != tt.want {
			t.Errorf("%s: hasSeriesNames = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestSeriesValueRange(t *testing.T) {
	tests := []struct {
		name         string
		series       []docx.ChartSeries
		wantMin, wantMax float64
	}{
		{"empty", nil, 0, 1},
		{"single", []docx.ChartSeries{{Values: []float64{1, 2, 3}}}, 1, 3},
		{"mixed", []docx.ChartSeries{{Values: []float64{-5, 10, 3}}}, -5, 10},
		{"negative", []docx.ChartSeries{{Values: []float64{-10, -5}}}, -10, -5},
		{"multi_series", []docx.ChartSeries{
			{Values: []float64{1, 5}},
			{Values: []float64{3, 2}},
		}, 1, 5},
	}
	for _, tt := range tests {
		min, max := seriesValueRange(tt.series)
		if min != tt.wantMin || max != tt.wantMax {
			t.Errorf("%s: seriesValueRange = (%v,%v), want (%v,%v)",
				tt.name, min, max, tt.wantMin, tt.wantMax)
		}
	}
}

func TestStackedValueRange(t *testing.T) {
	tests := []struct {
		name    string
		series  []docx.ChartSeries
		percent bool
		wantMin, wantMax float64
	}{
		{"percent", nil, true, 0, 100},
		{"positive", []docx.ChartSeries{
			{Values: []float64{10, 20}},
			{Values: []float64{30, 5}},
		}, false, 0, 40},
		{"mixed_sign", []docx.ChartSeries{
			{Values: []float64{10, -5}},
		}, false, -5, 10},
	}
	for _, tt := range tests {
		min, max := stackedValueRange(tt.series, tt.percent)
		if min != tt.wantMin || max != tt.wantMax {
			t.Errorf("%s: stackedValueRange = (%v,%v), want (%v,%v)",
				tt.name, min, max, tt.wantMin, tt.wantMax)
		}
	}
}

func TestValueToY(t *testing.T) {
	got := valueToY(5, 0, 10, 0, 100)
	if got != 50 {
		t.Errorf("valueToY(5,0,10,0,100) = %v, want 50", got)
	}
	got = valueToY(10, 0, 10, 0, 100)
	if got != 0 {
		t.Errorf("valueToY(10,0,10,0,100) = %v, want 0", got)
	}
	got = valueToY(0, 0, 0, 0, 100)
	if got != 100 {
		t.Errorf("valueToY(0,0,0,0,100) = %v, want 100", got)
	}
}

func TestValueToX(t *testing.T) {
	got := valueToX(5, 0, 10, 0, 100)
	if got != 50 {
		t.Errorf("valueToX(5,0,10,0,100) = %v, want 50", got)
	}
	got = valueToX(0, 0, 10, 0, 100)
	if got != 0 {
		t.Errorf("valueToX(0,0,10,0,100) = %v, want 0", got)
	}
	got = valueToX(5, 5, 5, 0, 100)
	if got != 0 {
		t.Errorf("valueToX(5,5,5,0,100) = %v, want 0", got)
	}
}

func TestPaletteColor(t *testing.T) {
	tests := []struct {
		idx  int
		want string
	}{
		{0, "4472C4"}, {1, "ED7D31"}, {7, "9E480E"},
		{8, "4472C4"}, {-1, "4472C4"},
	}
	for _, tt := range tests {
		got := paletteColor(tt.idx)
		if got != tt.want {
			t.Errorf("paletteColor(%d) = %q, want %q", tt.idx, got, tt.want)
		}
	}
}

func TestFormatChartValue(t *testing.T) {
	tests := []struct {
		v    float64
		want string
	}{
		{0, "0"},
		{5, "5"},
		{5.5, "5.5"},
		{1234567, "1.23e+06"},
		{0.5, "0.5"},
	}
	for _, tt := range tests {
		got := formatChartValue(tt.v)
		if got != tt.want {
			t.Errorf("formatChartValue(%v) = %q, want %q", tt.v, got, tt.want)
		}
	}
}

func TestComposeDataLabel(t *testing.T) {
	opts := docx.DataLabelOptions{ShowVal: true, ShowCatName: true, ShowSerName: true}
	got := composeDataLabel(opts, "Ser1", "Cat1", 42.5, 0.3)
	want := "Ser1 Cat1 42.5"
	if got != want {
		t.Errorf("composeDataLabel = %q, want %q", got, want)
	}
	opts2 := docx.DataLabelOptions{ShowPercent: true}
	got2 := composeDataLabel(opts2, "", "", 0, 0.25)
	if got2 != "25%" {
		t.Errorf("composeDataLabel pct = %q, want 25%%", got2)
	}
}

func TestEffectiveDataLabels(t *testing.T) {
	def := docx.DataLabelOptions{ShowVal: true}
	ser := docx.DataLabelOptions{ShowCatName: true}
	c := &docx.ChartData{DataLabels: def}
	s := docx.ChartSeries{DataLabels: &ser}
	got := effectiveDataLabels(c, s)
	if !got.ShowCatName {
		t.Error("series-level labels should win")
	}
	sNoLabels := docx.ChartSeries{}
	got2 := effectiveDataLabels(c, sNoLabels)
	if !got2.ShowVal {
		t.Error("chart-level labels should be fallback")
	}
}

func TestMaxCategoryCount(t *testing.T) {
	c := &docx.ChartData{
		Categories: []string{"A", "B"},
		Series: []docx.ChartSeries{
			{Values: []float64{1, 2, 3}},
		},
	}
	if n := maxCategoryCount(c); n != 3 {
		t.Errorf("maxCategoryCount = %d, want 3", n)
	}
}

func TestCategoryAt(t *testing.T) {
	c := &docx.ChartData{Categories: []string{"A", "B", "C"}}
	tests := []struct {
		i    int
		want string
	}{
		{0, "A"}, {2, "C"}, {-1, ""}, {5, ""},
	}
	for _, tt := range tests {
		got := categoryAt(c, tt.i)
		if got != tt.want {
			t.Errorf("categoryAt(%d) = %q, want %q", tt.i, got, tt.want)
		}
	}
}

func TestSeriesColor(t *testing.T) {
	s1 := docx.ChartSeries{Color: "FF0000"}
	if c := seriesColor(s1, 0); c != "FF0000" {
		t.Errorf("seriesColor with explicit color = %q, want FF0000", c)
	}
	s2 := docx.ChartSeries{}
	if c := seriesColor(s2, 0); c != "4472C4" {
		t.Errorf("seriesColor fallback = %q, want 4472C4", c)
	}
}

func TestLightenHex(t *testing.T) {
	r, g, b := lightenHex("000000", 0.5)
	if r != 127 || g != 127 || b != 127 {
		t.Errorf("lightenHex(000000,0.5) = (%d,%d,%d), want (127,127,127)", r, g, b)
	}
	r, g, b = lightenHex("FFFFFF", 0)
	if r != 255 || g != 255 || b != 255 {
		t.Errorf("lightenHex(FFFFFF,0) = (%d,%d,%d), want (255,255,255)", r, g, b)
	}
}

func TestTruncateLabel(t *testing.T) {
	tests := []struct {
		s        string
		maxRunes int
		want     string
	}{
		{"hello", 5, "hello"},
		{"hello", 3, "he…"},
		{"hello", 0, ""},
		{"hello", 1, "h"},
		{"世界世界", 3, "世界…"},
	}
	for _, tt := range tests {
		got := truncateLabel(tt.s, tt.maxRunes)
		if got != tt.want {
			t.Errorf("truncateLabel(%q,%d) = %q, want %q", tt.s, tt.maxRunes, got, tt.want)
		}
	}
}

func TestUtf8len(t *testing.T) {
	tests := []struct {
		s    string
		want int
	}{
		{"hello", 5},
		{"世界", 2},
		{"", 0},
		{"héllo", 5},
	}
	for _, tt := range tests {
		got := utf8len(tt.s)
		if got != tt.want {
			t.Errorf("utf8len(%q) = %d, want %d", tt.s, got, tt.want)
		}
	}
}

func TestApplyColorMods(t *testing.T) {
	tests := []struct {
		name                 string
		hex                  string
		lumMod, lumOff, satMod, satOff float64
		want                 string
	}{
		{"no_mods", "FF0000", 0, 0, 0, 0, "FF0000"},
		{"lum_mod", "FF0000", 0.5, 0, 0, 0, "800000"},
		{"lum_off_pos", "FF0000", 0, 0.5, 0, 0, "FF8080"},
		{"sat_mod", "FF0000", 0, 0, 0.5, 0, "BF4040"},
	}
	for _, tt := range tests {
		got := applyColorMods(tt.hex, tt.lumMod, tt.lumOff, tt.satMod, tt.satOff)
		if got != tt.want {
			t.Errorf("%s: applyColorMods(%q,%v,%v,%v,%v) = %q, want %q",
				tt.name, tt.hex, tt.lumMod, tt.lumOff, tt.satMod, tt.satOff, got, tt.want)
		}
	}
}

func TestFormatChartValueEdge(t *testing.T) {
	if got := formatChartValue(1e6); got != "1.00e+06" {
		t.Errorf("formatChartValue(1e6) = %q, want 1.00e+06", got)
	}
}

func TestComposeDataLabelEmpty(t *testing.T) {
	opts := docx.DataLabelOptions{}
	got := composeDataLabel(opts, "Ser1", "Cat1", 42.5, 0.3)
	if got != "" {
		t.Errorf("composeDataLabel with empty opts = %q, want empty", got)
	}
}

func TestStackedValueRangeCats(t *testing.T) {
	series := []docx.ChartSeries{
		{Values: []float64{10, 20, 30}},
		{Values: []float64{5}},
	}
	min, max := stackedValueRange(series, false)
	if min != 0 || max != 30 {
		t.Errorf("stackedValueRange uneven cats = (%v,%v), want (0,30)", min, max)
	}
}

func TestMathModZero(t *testing.T) {
	if got := mathMod(0, 2.5); got != 0 {
		t.Errorf("mathMod(0,2.5) = %v, want 0", got)
	}
}

func TestClamp01Bounds(t *testing.T) {
	if got := clamp01(-0.001); got != 0 {
		t.Errorf("clamp01(-0.001) = %v, want 0", got)
	}
	if got := clamp01(255.001); got != 255 {
		t.Errorf("clamp01(255.001) = %v, want 255", got)
	}
}
