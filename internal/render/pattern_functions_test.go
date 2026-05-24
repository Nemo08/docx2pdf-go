package render

import "testing"

func TestPatternPctValue(t *testing.T) {
	tests := []struct {
		preset string
		want   float64
	}{
		{"pct5", 5}, {"pct10", 10}, {"pct20", 20}, {"pct25", 25},
		{"pct30", 30}, {"pct40", 40}, {"pct50", 50}, {"pct60", 60},
		{"pct70", 70}, {"pct75", 75}, {"pct80", 80}, {"pct90", 90},
		{"pct0", 50}, {"pct100", 50}, {"unknown", 50}, {"", 50},
	}
	for _, tt := range tests {
		got := patternPctValue(tt.preset)
		if got != tt.want {
			t.Errorf("patternPctValue(%q) = %v, want %v", tt.preset, got, tt.want)
		}
	}
}
