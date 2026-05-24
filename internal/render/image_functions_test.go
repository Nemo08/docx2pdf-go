package render

import (
	"image"
	"image/color"
	"testing"
)

func TestClamp01(t *testing.T) {
	tests := []struct {
		in   float64
		want float64
	}{
		{-1, 0}, {0, 0}, {128, 128}, {255, 255}, {300, 255},
	}
	for _, tt := range tests {
		got := clamp01(tt.in)
		if got != tt.want {
			t.Errorf("clamp01(%v) = %v, want %v", tt.in, got, tt.want)
		}
	}
}

func TestMathMod(t *testing.T) {
	tests := []struct {
		x, m, want float64
	}{
		{0, 1, 0}, {0.5, 1, 0.5}, {1.5, 1, 0.5},
		{-0.5, 1, 0.5}, {-1.5, 1, 0.5}, {5, 3, 2},
	}
	for _, tt := range tests {
		got := mathMod(tt.x, tt.m)
		if got != tt.want {
			t.Errorf("mathMod(%v, %v) = %v, want %v", tt.x, tt.m, got, tt.want)
		}
	}
}

func TestAdjustLum(t *testing.T) {
	type result struct{ r, g, b uint8 }
	tests := []struct {
		name           string
		r, g, b        uint8
		bright, contra float64
		want           result
	}{
		{"noop", 128, 128, 128, 0, 0, result{128, 128, 128}},
		{"bright_up", 100, 100, 100, 0.2, 0, result{151, 151, 151}},
		{"bright_down", 100, 100, 100, -0.2, 0, result{48, 48, 48}},
		{"contrast_up", 128, 128, 128, 0, 0.5, result{128, 128, 128}},
		{"contrast_bright", 64, 64, 64, 0.1, 0.3, result{70, 70, 70}},
	}
	for _, tt := range tests {
		r, g, b := adjustLum(tt.r, tt.g, tt.b, tt.bright, tt.contra)
		if r != tt.want.r || g != tt.want.g || b != tt.want.b {
			t.Errorf("%s: adjustLum(%d,%d,%d,%v,%v) = (%d,%d,%d), want (%d,%d,%d)",
				tt.name, tt.r, tt.g, tt.b, tt.bright, tt.contra,
				r, g, b, tt.want.r, tt.want.g, tt.want.b)
		}
	}
}

func TestMustHex(t *testing.T) {
	white := color.NRGBA{R: 255, G: 255, B: 255, A: 255}
	tests := []struct {
		hex      string
		fallback color.NRGBA
		want     color.NRGBA
	}{
		{"FF0000", white, color.NRGBA{R: 255, G: 0, B: 0, A: 255}},
		{"00FF00", white, color.NRGBA{R: 0, G: 255, B: 0, A: 255}},
		{"0000FF", white, color.NRGBA{R: 0, G: 0, B: 255, A: 255}},
		{"", white, white},
		{"ABC", white, white},
		{"G00000", white, white},
	}
	for _, tt := range tests {
		got := mustHex(tt.hex, tt.fallback)
		if got != tt.want {
			t.Errorf("mustHex(%q, %v) = %v, want %v", tt.hex, tt.fallback, got, tt.want)
		}
	}
}

func TestCropImage(t *testing.T) {
	img := image.NewNRGBA(image.Rect(0, 0, 100, 100))
	for x := 0; x < 100; x++ {
		for y := 0; y < 100; y++ {
			img.SetNRGBA(x, y, color.NRGBA{R: uint8(x), G: uint8(y), B: 128, A: 255})
		}
	}
	tests := []struct {
		name                 string
		top, bottom, l, r    float64
		wantW, wantH         int
	}{
		{"no_crop", 0, 0, 0, 0, 100, 100},
		{"crop_even", 10, 10, 10, 10, 80, 80},
		{"crop_100pct", 50, 50, 50, 50, 1, 1},
	}
	for _, tt := range tests {
		got := cropImage(img, tt.top, tt.bottom, tt.l, tt.r)
		b := got.Bounds()
		if b.Dx() != tt.wantW || b.Dy() != tt.wantH {
			t.Errorf("%s: cropImage bounds = %dx%d, want %dx%d",
				tt.name, b.Dx(), b.Dy(), tt.wantW, tt.wantH)
		}
	}
}
