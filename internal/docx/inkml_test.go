package docx

import "testing"

func TestSplitInkTrace(t *testing.T) {
	tests := []struct {
		raw  string
		want int
	}{
		{"100 200,300 400", 2},
		{"single", 1},
		{"", 1},
	}
	for _, tt := range tests {
		got := splitInkTrace(tt.raw)
		if len(got) != tt.want {
			t.Errorf("splitInkTrace(%q) = %d parts, want %d", tt.raw, len(got), tt.want)
		}
	}
}

func TestParseInkTrace(t *testing.T) {
	tests := []struct {
		raw     string
		wantLen int
		wantX   float64
		wantY   float64
	}{
		{"", 0, 0, 0},
		{"100 200,300 400", 2, 100, 200},
		{"!50 60,70 80", 2, 50, 60},
		{"invalid", 0, 0, 0},
	}
	for _, tt := range tests {
		s := parseInkTrace(tt.raw)
		if len(s.Points) != tt.wantLen {
			t.Errorf("parseInkTrace(%q): got %d points, want %d", tt.raw, len(s.Points), tt.wantLen)
			continue
		}
		if tt.wantLen > 0 {
			if s.Points[0].X != tt.wantX || s.Points[0].Y != tt.wantY {
				t.Errorf("parseInkTrace(%q): first point = (%v,%v), want (%v,%v)",
					tt.raw, s.Points[0].X, s.Points[0].Y, tt.wantX, tt.wantY)
			}
		}
	}
}

func TestInkStrokeBounds(t *testing.T) {
	tests := []struct {
		name               string
		strokes            []InkStroke
		wantOk             bool
		wantMinX, wantMaxX float64
	}{
		{"empty", nil, false, 0, 0},
		{"single_point", []InkStroke{{Points: []InkPoint{{X: 10, Y: 20}}}}, true, 10, 10},
		{"multiple_points", []InkStroke{{Points: []InkPoint{
			{X: 10, Y: 20}, {X: 100, Y: 200},
		}}}, true, 10, 100},
		{"multi_stroke", []InkStroke{
			{Points: []InkPoint{{X: 5, Y: 5}}},
			{Points: []InkPoint{{X: 50, Y: 50}}},
		}, true, 5, 50},
	}
	for _, tt := range tests {
		minX, minY, maxX, maxY, ok := InkStrokeBounds(tt.strokes)
		if ok != tt.wantOk {
			t.Errorf("%s: ok = %v, want %v", tt.name, ok, tt.wantOk)
		}
		if ok {
			if minX != tt.wantMinX || maxX != tt.wantMaxX {
				t.Errorf("%s: bounds x = (%v,%v), want (%v,%v)", tt.name, minX, maxX, tt.wantMinX, tt.wantMaxX)
			}
			if minY > maxY {
				t.Errorf("%s: minY (%v) > maxY (%v)", tt.name, minY, maxY)
			}
		}
	}
}

func TestInkStrokesToShape(t *testing.T) {
	// Empty strokes -> nil
	if s := inkStrokesToShape(nil); s != nil {
		t.Error("inkStrokesToShape(nil) should return nil")
	}
	// Single stroke
	strokes := []InkStroke{
		{Points: []InkPoint{{X: 0, Y: 0}, {X: 100, Y: 100}}},
	}
	s := inkStrokesToShape(strokes)
	if s == nil {
		t.Fatal("inkStrokesToShape should return non-nil")
	}
	if s.Kind != "ink" {
		t.Errorf("kind = %q, want ink", s.Kind)
	}
	if s.CustomPath == "" {
		t.Error("CustomPath should not be empty")
	}
	if s.StrokeColor != "000000" {
		t.Errorf("StrokeColor = %q, want 000000", s.StrokeColor)
	}
	if s.WidthPt <= 0 || s.HeightPt <= 0 {
		t.Errorf("bad dimensions: %v x %v", s.WidthPt, s.HeightPt)
	}
}
