package render

import (
	"math"
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
	"github.com/signintech/gopdf"
)

func TestShapeBodyInsets(t *testing.T) {
	type result struct{ l, t, r, b float64 }
	tests := []struct {
		name string
		s    docx.VMLShape
		want result
	}{
		{"defaults", docx.VMLShape{}, result{7.2, 3.6, 7.2, 3.6}},
		{"custom_left", docx.VMLShape{TextLeftInsetPt: 10}, result{10, 3.6, 7.2, 3.6}},
		{"custom_all", docx.VMLShape{TextLeftInsetPt: 5, TextTopInsetPt: 5, TextRightInsetPt: 5, TextBottomInsetPt: 5},
			result{5, 5, 5, 5}},
	}
	for _, tc := range tests {
		l, tt, r, b := shapeBodyInsets(&tc.s)
		if l != tc.want.l || tt != tc.want.t || r != tc.want.r || b != tc.want.b {
			t.Errorf("%s: shapeBodyInsets = (%v,%v,%v,%v), want (%v,%v,%v,%v)",
				tc.name, l, tt, r, b, tc.want.l, tc.want.t, tc.want.r, tc.want.b)
		}
	}
}

func TestHexFromRGB(t *testing.T) {
	tests := []struct {
		r, g, b uint8
		want    string
	}{
		{0, 0, 0, "000000"},
		{255, 255, 255, "FFFFFF"},
		{255, 0, 0, "FF0000"},
		{0, 255, 0, "00FF00"},
		{0, 0, 255, "0000FF"},
		{128, 128, 128, "808080"},
	}
	for _, tt := range tests {
		got := hexFromRGB(tt.r, tt.g, tt.b)
		if got != tt.want {
			t.Errorf("hexFromRGB(%d,%d,%d) = %q, want %q", tt.r, tt.g, tt.b, got, tt.want)
		}
	}
}

func TestAbsF(t *testing.T) {
	tests := []struct {
		x    float64
		want float64
	}{
		{5, 5}, {-5, 5}, {0, 0}, {-3.14, 3.14},
	}
	for _, tt := range tests {
		got := absF(tt.x)
		if got != tt.want {
			t.Errorf("absF(%v) = %v, want %v", tt.x, got, tt.want)
		}
	}
}

func TestCosFloatSinFloat(t *testing.T) {
	if got := cosFloat(0); math.Abs(got-1) > 1e-10 {
		t.Errorf("cosFloat(0) = %v, want 1", got)
	}
	if got := sinFloat(0); math.Abs(got) > 1e-10 {
		t.Errorf("sinFloat(0) = %v, want 0", got)
	}
	if got := cosFloat(math.Pi); math.Abs(got+1) > 1e-10 {
		t.Errorf("cosFloat(pi) = %v, want -1", got)
	}
}

func TestInterpolateGradient(t *testing.T) {
	tests := []struct {
		name  string
		stops []docx.GradientStop
		t     float64
		want  string
	}{
		{"empty_stops", nil, 0.5, "000000"},
		{"before_first", []docx.GradientStop{{Pos: 0.2, Color: "FF0000"}, {Pos: 0.8, Color: "0000FF"}}, 0, "FF0000"},
		{"after_last", []docx.GradientStop{{Pos: 0.2, Color: "FF0000"}, {Pos: 0.8, Color: "0000FF"}}, 1, "0000FF"},
		{"exact_stop", []docx.GradientStop{{Pos: 0.5, Color: "00FF00"}}, 0.5, "00FF00"},
		{"midpoint", []docx.GradientStop{{Pos: 0, Color: "000000"}, {Pos: 1, Color: "FFFFFF"}}, 0.5, "7F7F7F"},
		{"three_stops", []docx.GradientStop{
			{Pos: 0, Color: "FF0000"}, {Pos: 0.5, Color: "00FF00"}, {Pos: 1, Color: "0000FF"},
		}, 0.25, "7F7F00"},
	}
	for _, tt := range tests {
		got := interpolateGradient(tt.stops, tt.t)
		if got != tt.want {
			t.Errorf("%s: interpolateGradient = %q, want %q", tt.name, got, tt.want)
		}
	}
}

func TestCrescentPoints(t *testing.T) {
	pts := crescentPoints(0, 0, 100, 100)
	// 19 outer + 19 inner = 38 points (segs=18, so 19 each)
	if len(pts) != 38 {
		t.Errorf("crescentPoints should return 38 points, got %d", len(pts))
	}
	if pts[0].X != 60 || pts[0].Y != 95 {
		t.Errorf("crescentPoints first point = (%v,%v), want (60,95)", pts[0].X, pts[0].Y)
	}
	// Verify points are of the right type
	if _, ok := interface{}(pts[0]).(gopdf.Point); !ok {
		t.Error("crescentPoints should return gopdf.Point slice")
	}
}
