# Contributing to docx2pdf-go

Thanks for the interest! Issues and pull requests are welcome.

## Quick start

```bash
git clone https://github.com/bobyeoh/docx2pdf-go.git
cd docx2pdf-go
go test ./...
```

External tools the test suite needs (Linux: `apt-get install poppler-utils`;
macOS: `brew install poppler`):

- `pdftotext` — text-extraction assertions
- `pdfinfo`   — page-count / page-size assertions
- `pdftoppm`  — PNG snapshots for visual review

The suite is skip-friendly: tests that need a TTF at `testdata/font.ttf`
will `t.Skip()` if it's missing, so you can run `go test ./...` on a
fresh clone and most of the unit + integration tests still pass. To run
the full visual coverage, drop a sans-serif TTF at `testdata/font.ttf`.

## Project layout

```
docx2pdf.go              ← public API
cmd/docx2pdf/            ← CLI
internal/docx/           ← OOXML parser (zip → AST)
internal/render/         ← PDF renderer (AST → gopdf calls)
internal/convert/        ← thin orchestrator
internal/verify/         ← test harness; one file per "case batch"
testdata/                ← shared fixtures (font, sample docx)
```

The renderer is split by concern — `pdf.go` for entry points, `page.go`
for headers/footers, `text.go` for the atom-based line breaker, etc.
See the source-layout block in [README.md](README.md#architecture).

## Adding a feature or fixing a bug

The verify harness is the fastest feedback loop. Add a new case to
`internal/verify/verify_test.go` modeled after existing ones:

```go
func caseYourThing() verifyCase {
    return verifyCase{
        name:        "999_your_thing",
        description: "One-line description for the HTML report",
        build: func(t *testing.T, dir string) string {
            return newDocx().Body(`<w:p>…your XML…</w:p>`).Write(t, dir)
        },
        expectText: []string{"distinctive marker"},
        expectPages: 1,
    }
}
```

…then add it to the `cases()` slice and run:

```bash
go test ./internal/verify/ -run 'TestVerifyAll/999_your_thing' -v
```

The test produces a PDF and PNG snapshots under
`internal/verify/out/your_thing/` for visual review.

## Before submitting

1. `gofmt -w .` — required, CI rejects on diff
2. `go vet ./...` — required
3. `go test ./... -race` — must pass
4. `staticcheck ./...` — should pass (CI will warn but not fail)

Optional but appreciated:

- `GOLDEN=1 go test ./internal/verify/` — confirms PNG-level rendering
  matches the checked-in baselines
- Update [README.md](README.md) feature table if you added a new ✅/⚠️
- One regression test per feature/bug

## Commit conventions

Use clear, present-tense subject lines describing the change, not the
task ("Support legacy VML images" rather than "VML"). Body explains the
*why*, especially when fixing a non-obvious bug. Multi-paragraph bodies
are welcome.

## License

By contributing, you agree your contributions are released under the
[MIT License](LICENSE).
