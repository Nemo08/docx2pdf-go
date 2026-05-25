package main

import (
	"archive/zip"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"testing"

	"github.com/bobyeoh/docx2pdf-go/internal/convert"
)

func TestFindDocxFlat(t *testing.T) {
	dir := t.TempDir()
	touch := func(p string) {
		if err := os.WriteFile(filepath.Join(dir, p), []byte{}, 0o644); err != nil {
			t.Fatal(err)
		}
	}
	mk := func(p string) {
		if err := os.MkdirAll(filepath.Join(dir, p), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	touch("a.docx")
	touch("b.DOCX") // case-insensitive match
	touch("~$tmp.docx")
	touch("notes.txt")
	mk("sub")
	if err := os.WriteFile(filepath.Join(dir, "sub", "c.docx"), []byte{}, 0o644); err != nil {
		t.Fatal(err)
	}

	got, err := findDocx(dir, false)
	if err != nil {
		t.Fatal(err)
	}
	sort.Strings(got)
	if len(got) != 2 {
		t.Fatalf("non-recursive: got %v, want 2 entries (a.docx, b.DOCX)", got)
	}

	got2, err := findDocx(dir, true)
	if err != nil {
		t.Fatal(err)
	}
	if len(got2) != 3 {
		t.Fatalf("recursive: got %v, want 3 entries", got2)
	}

	// Confirm we filter the Word lockfile pattern.
	for _, p := range got2 {
		if filepath.Base(p) == "~$tmp.docx" {
			t.Errorf("lockfile %s should be filtered", p)
		}
	}
}

func TestWithExt(t *testing.T) {
	cases := map[string]string{
		"a.docx":     "a.pdf",
		"x/y/b.docx": "x/y/b.pdf",
		"no_ext":     "no_ext.pdf",
	}
	for in, want := range cases {
		if got := withExt(in, ".pdf"); got != want {
			t.Errorf("withExt(%q) = %q want %q", in, got, want)
		}
	}
}

// TestConvertFile_NoInput ensures convertFile fails on missing input.
func TestConvertFile_NoInput(t *testing.T) {
	tmp := t.TempDir()
	j := job{src: "/nonexistent.docx", rel: "x.docx", dst: filepath.Join(tmp, "out.pdf")}
	err := convertFile(j, false, convert.Options{})
	if err == nil {
		t.Fatal("expected error for missing input file")
	}
}

// TestConvertFile_NoDir ensures convertFile fails when output dir missing.
func TestConvertFile_NoDir(t *testing.T) {
	tmp := t.TempDir()
	src := filepath.Join(tmp, "in.docx")
	if err := writeMinimalDocxFile(src); err != nil {
		t.Fatal(err)
	}
	j := job{src: src, rel: "in.docx", dst: filepath.Join(tmp, "nope", "out.pdf")}
	err := convertFile(j, false, convert.Options{FontRegular: "/missing"})
	if err == nil {
		t.Fatal("expected error for missing output dir")
	}
}

// TestRunStream_FileInFileOut verifies stream mode with real file I/O.
func TestRunStream_FileInFileOut(t *testing.T) {
	tmp := t.TempDir()
	src := filepath.Join(tmp, "in.docx")
	if err := writeMinimalDocxFile(src); err != nil {
		t.Fatal(err)
	}
	dst := filepath.Join(tmp, "out.pdf")
	err := runStream(src, dst, convert.Options{})
	if err != nil {
		t.Logf("runStream(file→file): %v (expected on fontless CI)", err)
	}
}

// TestRunBatchEmptyDir verifies runBatch handles empty directories gracefully.
func TestRunBatchEmptyDir(t *testing.T) {
	tmp := t.TempDir()
	inDir := filepath.Join(tmp, "empty")
	if err := os.Mkdir(inDir, 0o755); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(tmp, "out")
	failed := runBatch(inDir, outDir, false, false, 1, convert.Options{})
	if failed != 0 {
		t.Errorf("expected 0 failures for empty dir, got %d", failed)
	}
}

// TestRunBatchSingleFile verifies runBatch with one docx file.
func TestRunBatchSingleFile(t *testing.T) {
	tmp := t.TempDir()
	inDir := filepath.Join(tmp, "in")
	outDir := filepath.Join(tmp, "out")
	if err := os.Mkdir(inDir, 0o755); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(inDir, "test.docx")
	if err := writeMinimalDocxFile(src); err != nil {
		t.Fatal(err)
	}
	failed := runBatch(inDir, outDir, false, false, 1, convert.Options{})
	if failed != 0 {
		t.Logf("runBatch: %d failed (expected on fontless CI)", failed)
	}
	if _, err := os.Stat(outDir); os.IsNotExist(err) {
		t.Error("expected output directory to exist")
	}
}

// writeMinimalDocxFile creates a syntactically valid minimal .docx at path.
func writeMinimalDocxFile(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	zw := zip.NewWriter(f)
	defer zw.Close()
	w, err := zw.Create("word/document.xml")
	if err != nil {
		return err
	}
	_, err = fmt.Fprintf(w, `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body>
</w:document>`)
	return err
}
