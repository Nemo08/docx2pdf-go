package docx

import (
	"archive/zip"
	"bytes"
	"testing"
)

func TestChartSibling_Found(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	must := func(name string) {
		_, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
	}
	must("word/charts/chart1.xml")
	must("word/charts/colors1.xml")
	must("word/charts/style1.xml")
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	files := map[string]*zip.File{}
	for _, zf := range zr.File {
		files[zf.Name] = zf
	}
	if got := chartSibling(files, "word/charts/chart1.xml", "colors"); got == nil || got.Name != "word/charts/colors1.xml" {
		t.Errorf("chartSibling(colors) = %v, want word/charts/colors1.xml", got)
	}
	if got := chartSibling(files, "word/charts/chart1.xml", "style"); got == nil || got.Name != "word/charts/style1.xml" {
		t.Errorf("chartSibling(style) = %v, want word/charts/style1.xml", got)
	}
}

func TestChartSibling_MissingWordPrefix(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	must := func(name string) {
		_, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
	}
	must("word/charts/chart1.xml")
	must("word/charts/colors1.xml")
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	files := map[string]*zip.File{}
	for _, zf := range zr.File {
		files[zf.Name] = zf
	}
	if got := chartSibling(files, "charts/chart1.xml", "colors"); got == nil || got.Name != "word/charts/colors1.xml" {
		t.Errorf("chartSibling(missing word/) = %v, want word/charts/colors1.xml", got)
	}
}

func TestChartSibling_NonChartName(t *testing.T) {
	files := map[string]*zip.File{}
	if got := chartSibling(files, "word/styles.xml", "colors"); got != nil {
		t.Errorf("expected nil for non-chart name, got %v", got)
	}
}

func TestChartSibling_NoMatchingSibling(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	_, _ = zw.Create("word/charts/chart1.xml")
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	zr, _ := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	files := map[string]*zip.File{}
	for _, zf := range zr.File {
		files[zf.Name] = zf
	}
	if got := chartSibling(files, "word/charts/chart1.xml", "colors"); got != nil {
		t.Errorf("expected nil for missing sibling, got %v", got)
	}
}
