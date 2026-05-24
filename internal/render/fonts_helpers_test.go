package render

import (
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

func TestIsRTL(t *testing.T) {
	tests := []struct {
		r    rune
		want bool
	}{
		{'A', false},
		{'1', false},
		{' ', false},
		{'\u0590', true},  // Hebrew
		{'\u05D0', true},  // Hebrew Aleph
		{'\u05FF', true},  // Hebrew end
		{'\u0600', true},  // Arabic
		{'\u0660', false}, // Arabic digit
		{'\u06F0', false}, // Arabic digit
		{'\u0700', true},  // Syriac
		{'\uFB1D', true},  // Hebrew/Arabic Forms
		{'\uFE70', true},  // Arabic Forms-B
		{'\uFEFF', true},  // Arabic Forms-B end
	}
	for _, tt := range tests {
		got := isRTL(tt.r)
		if got != tt.want {
			t.Errorf("isRTL(%U %q) = %v, want %v", tt.r, tt.r, got, tt.want)
		}
	}
}



func TestIsSymbolGlyph(t *testing.T) {
	tests := []struct {
		r    rune
		want bool
	}{
		{'A', false},
		{'\u2190', true}, // Arrows
		{'\u2200', true}, // Math Operators
		{'\u2300', true}, // Misc Technical
		{'\u2500', true}, // Box Drawing
		{'\u25A0', true}, // Geometric Shapes
		{'\u2600', true}, // Misc Symbols
		{'\u2700', true}, // Dingbats
		{'\u27BF', true}, // Dingbats end
	}
	for _, tt := range tests {
		got := isSymbolGlyph(tt.r)
		if got != tt.want {
			t.Errorf("isSymbolGlyph(%U %q) = %v, want %v", tt.r, tt.r, got, tt.want)
		}
	}
}

func TestIsPUA(t *testing.T) {
	tests := []struct {
		r    rune
		want bool
	}{
		{'A', false},
		{'\uE000', true},  // BMP PUA start
		{'\uF8FF', true},  // BMP PUA end
		{'\U000F0000', true}, // Suppl PUA-A start
		{'\U0010FFFD', true}, // Suppl PUA-B end
	}
	for _, tt := range tests {
		got := isPUA(tt.r)
		if got != tt.want {
			t.Errorf("isPUA(%U %q) = %v, want %v", tt.r, tt.r, got, tt.want)
		}
	}
}

func TestIsMajorThemeRole(t *testing.T) {
	tests := []struct {
		role string
		want bool
	}{
		{"majorAscii", true},
		{"majorEastAsia", true},
		{"minorAscii", false},
		{"", false},
	}
	for _, tt := range tests {
		got := isMajorThemeRole(tt.role)
		if got != tt.want {
			t.Errorf("isMajorThemeRole(%q) = %v, want %v", tt.role, got, tt.want)
		}
	}
}

func TestCharSpacingFor(t *testing.T) {
	tests := []struct {
		name     string
		props    docx.RunProps
		fontSize float64
		want     float64
	}{
		{"no_spacing", docx.RunProps{}, 12, 0},
		{"letter_spacing", docx.RunProps{LetterSpacingPt: 2}, 12, 2},
		{"char_scale", docx.RunProps{CharacterScale: 1.5}, 12, 3},
	}
	for _, tt := range tests {
		got := charSpacingFor(tt.props, tt.fontSize)
		if got != tt.want {
			t.Errorf("%s: charSpacingFor = %v, want %v", tt.name, got, tt.want)
		}
	}
}



func TestRGBToHSL(t *testing.T) {
	type hsv struct{ h, s, l float64 }
	tests := []struct {
		name   string
		r, g, b uint8
		want   hsv
	}{
		{"black", 0, 0, 0, hsv{0, 0, 0}},
		{"white", 255, 255, 255, hsv{0, 0, 1}},
		{"red", 255, 0, 0, hsv{0, 1, 0.5}},
		{"green", 0, 255, 0, hsv{1.0 / 3.0, 1, 0.5}},
		{"blue", 0, 0, 255, hsv{2.0 / 3.0, 1, 0.5}},
		{"gray", 128, 128, 128, hsv{0, 0, 128.0 / 255.0}},
	}
	closeEnough := func(a, b, eps float64) bool {
		diff := a - b
		if diff < 0 {
			diff = -diff
		}
		return diff <= eps
	}
	for _, tt := range tests {
		h, s, l := rgbToHSL(tt.r, tt.g, tt.b)
		if !closeEnough(h, tt.want.h, 0.01) || !closeEnough(s, tt.want.s, 0.01) || !closeEnough(l, tt.want.l, 0.01) {
			t.Errorf("%s: rgbToHSL(%d,%d,%d) = (%.4f,%.4f,%.4f), want (%.4f,%.4f,%.4f)",
				tt.name, tt.r, tt.g, tt.b, h, s, l, tt.want.h, tt.want.s, tt.want.l)
		}
	}
}

func TestHSLToRGB(t *testing.T) {
	type rgb struct{ r, g, b uint8 }
	tests := []struct {
		name     string
		h, s, l  float64
		want     rgb
	}{
		{"black", 0, 0, 0, rgb{0, 0, 0}},
		{"white", 0, 0, 1, rgb{255, 255, 255}},
		{"red", 0, 1, 0.5, rgb{255, 0, 0}},
		{"green", 1.0 / 3.0, 1, 0.5, rgb{0, 255, 0}},
		{"blue", 2.0 / 3.0, 1, 0.5, rgb{0, 0, 255}},
	}
	for _, tt := range tests {
		r, g, b := hslToRGB(tt.h, tt.s, tt.l)
		if r != tt.want.r || g != tt.want.g || b != tt.want.b {
			t.Errorf("%s: hslToRGB(%.4f,%.4f,%.4f) = (%d,%d,%d), want (%d,%d,%d)",
				tt.name, tt.h, tt.s, tt.l, r, g, b, tt.want.r, tt.want.g, tt.want.b)
		}
	}
}




