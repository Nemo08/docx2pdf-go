package render

import (
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

func TestIsAllCapsWord(t *testing.T) {
	tests := []struct {
		s    string
		want bool
	}{
		{"HELLO", true},
		{"Hello", false},
		{"", false},
		{"ABC123", true},
		{"abc", false},
		{"HELLO_WORLD", true},
	}
	for _, tt := range tests {
		got := isAllCapsWord(tt.s)
		if got != tt.want {
			t.Errorf("isAllCapsWord(%q) = %v, want %v", tt.s, got, tt.want)
		}
	}
}

func TestTransformText(t *testing.T) {
	tests := []struct {
		name string
		s    string
		p    docx.RunProps
		want string
	}{
		{"no_caps", "hello", docx.RunProps{}, "hello"},
		{"caps", "hello", docx.RunProps{Caps: true}, "HELLO"},
		{"small_caps", "hello", docx.RunProps{SmallCaps: true}, "HELLO"},
		{"already_upper", "HELLO", docx.RunProps{Caps: true}, "HELLO"},
		{"mixed", "Hello World", docx.RunProps{Caps: true}, "HELLO WORLD"},
	}
	for _, tt := range tests {
		got := transformText(tt.s, tt.p)
		if got != tt.want {
			t.Errorf("%s: transformText(%q) = %q, want %q", tt.name, tt.s, got, tt.want)
		}
	}
}

func TestApplyDropCap(t *testing.T) {
	tests := []struct {
		name  string
		runs  []docx.Run
		lines int
		want  int
	}{
		{"empty_runs", nil, 3, 0},
		{"empty_text", []docx.Run{{Text: ""}}, 3, 1},
		{"single_char", []docx.Run{{Text: "A"}}, 3, 2},
		{"multi_char", []docx.Run{{Text: "Hello"}}, 2, 2},
		{"multi_run_skip_first", []docx.Run{{Text: ""}, {Text: "Hello"}}, 3, 3},
	}
	for _, tt := range tests {
		got := applyDropCap(tt.runs, tt.lines)
		if len(got) != tt.want {
			t.Errorf("%s: applyDropCap returned %d runs, want %d", tt.name, len(got), tt.want)
		}
	}
	// Verify drop cap properties
	runs := applyDropCap([]docx.Run{{Text: "Test", Props: docx.RunProps{FontSize: 12}}}, 3)
	if len(runs) != 2 {
		t.Fatal("drop cap should return 2 runs")
	}
	if runs[0].Text != "T" {
		t.Errorf("first run text = %q, want T", runs[0].Text)
	}
	if !runs[0].Props.Bold {
		t.Error("drop cap should be bold")
	}
	if runs[0].Props.FontSize <= 12 {
		t.Errorf("drop cap font size %v should be > 12", runs[0].Props.FontSize)
	}
}

func TestTransformTextEmpty(t *testing.T) {
	if got := transformText("", docx.RunProps{Caps: true}); got != "" {
		t.Errorf("transformText empty = %q, want empty", got)
	}
}

func TestIsAllCapsWordUnicode(t *testing.T) {
	if isAllCapsWord("ПРИВЕТ") {
		t.Error("isAllCapsWord should return false for non-Latin uppercase")
	}
	if !isAllCapsWord("ABC") {
		t.Error("isAllCapsWord should return true for ABC")
	}
}
