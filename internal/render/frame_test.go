package render

import (
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

func TestResolveFrameX(t *testing.T) {
	r := &renderer{
		marL:     72,
		marR:     72,
		pageW:    612,
		contentW: 468,
	}
	tests := []struct {
		name   string
		fr     *docx.FrameInfo
		frameW float64
		want   float64
	}{
		{"default_margin_left", &docx.FrameInfo{}, 200, 72},
		{"page_h_anchor", &docx.FrameInfo{HAnchor: "page"}, 200, 0},
		{"text_anchor", &docx.FrameInfo{HAnchor: "text"}, 200, 72},
		{"x_align_center", &docx.FrameInfo{XAlign: "center"}, 200, 72 + (468-200)/2},
		{"x_align_right", &docx.FrameInfo{XAlign: "right"}, 200, 72 + 468 - 200},
		{"x_twips", &docx.FrameInfo{XTwips: 1440}, 200, 72 + 72},
	}
	for _, tt := range tests {
		got := resolveFrameX(r, tt.fr, tt.frameW)
		if got != tt.want {
			t.Errorf("%s: resolveFrameX = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestResolveFrameY(t *testing.T) {
	r := &renderer{
		marT:  72,
		marB:  72,
		pageH: 792,
	}
	tests := []struct {
		name string
		fr   *docx.FrameInfo
		want float64
	}{
		{"default_page_top", &docx.FrameInfo{}, 0},
		{"margin_anchor", &docx.FrameInfo{VAnchor: "margin"}, 72},
		{"text_anchor", &docx.FrameInfo{VAnchor: "text"}, 100},
		{"y_align_center", &docx.FrameInfo{YAlign: "center"}, (792 - 0) / 2},
		{"y_align_bottom", &docx.FrameInfo{YAlign: "bottom"}, 792},
		{"y_twips", &docx.FrameInfo{YTwips: 1440}, 72},
	}
	for _, tt := range tests {
		r.cursorY = 100
		r.cursorY = 100

		got := resolveFrameY(r, tt.fr)
		if got != tt.want {
			t.Errorf("%s: resolveFrameY = %v, want %v", tt.name, got, tt.want)
		}
	}
}
