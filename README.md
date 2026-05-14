# docx2pdf-go

[![test](https://github.com/bobyeoh/docx2pdf-go/actions/workflows/test.yml/badge.svg)](https://github.com/bobyeoh/docx2pdf-go/actions/workflows/test.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/bobyeoh/docx2pdf-go.svg)](https://pkg.go.dev/github.com/bobyeoh/docx2pdf-go)
[![Go Report Card](https://goreportcard.com/badge/github.com/bobyeoh/docx2pdf-go)](https://goreportcard.com/report/github.com/bobyeoh/docx2pdf-go)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

> Pure-Go library and CLI that converts Microsoft Word `.docx` files to PDF.
> **No JVM. No LibreOffice. No MS Word. No CGO.** Just a single static binary
> and a `go get`-able package.

```go
import docx2pdf "github.com/bobyeoh/docx2pdf-go"

// Simplest call — picks up a common system font automatically
// (Arial / Helvetica on macOS, DejaVu / Liberation / Noto on Linux).
err := docx2pdf.Convert("report.docx", "report.pdf", docx2pdf.Options{})

// Or be explicit (recommended for reproducible cross-machine output):
err = docx2pdf.Convert("report.docx", "report.pdf", docx2pdf.Options{
    FontRegular:  "/usr/share/fonts/noto/NotoSans-Regular.ttf",
    FontFallback: "/usr/share/fonts/noto/NotoSansCJK-Regular.ttc", // CJK
    PageNumbers:  true,
})
```

---

## Why docx2pdf-go

Pure-Go `.docx` → PDF is harder than it looks. Real-world options today:

| Approach | Trade-off |
|---|---|
| Shell out to **LibreOffice** | 500 MB+ container, slow fork-per-conversion, headless quirks |
| **unioffice** | Commercial license required for closed-source apps |
| **Pandoc + LaTeX** | Heavy toolchain, fragile layout for complex Office docs |
| Hand-roll it | Months of XML wrangling, font metrics, table layout… |

**docx2pdf-go** sits in the gap: a focused, MIT-licensed, pure-Go library that
covers the 90% of WordprocessingML real documents actually use, without
external runtimes.

### What you get

- **~3.5 MB static binary.** Runs in `scratch` / `distroless` / Alpine
  with no fonts installed — a 150 KB Latin face (Go fonts, MIT) is
  embedded as a last-resort fallback so the binary always has
  something to draw with.
- **~70 MB official Docker image** with Noto Sans + WenQuanYi Zen Hei
  baked in for Latin + CJK fallback, no Word installation needed.
- **Streaming `io.Reader` → `io.Writer` API** for HTTP handlers and
  in-memory pipelines.
- **CLI for batch processing** — point it at a directory tree, get a
  mirrored tree of PDFs.
- **CJK supported via fallback font** with per-character line-break
  opportunities, so Chinese/Japanese/Korean paragraphs wrap mid-text
  even without whitespace. Latin and CJK glyph shaping work; Arabic
  letter shaping (initial/medial/final) is not yet implemented.

---

## Quick start

### As a library

```bash
go get github.com/bobyeoh/docx2pdf-go
```

**1. File paths — the simplest case.** Empty Options auto-detects a
system font; pass `FontRegular` for reproducible output.

```go
import docx2pdf "github.com/bobyeoh/docx2pdf-go"

func writeReport() error {
    return docx2pdf.Convert("in.docx", "out.pdf", docx2pdf.Options{})
}
```

**2. Streaming — perfect for HTTP handlers.**

```go
import (
    "bytes"
    "io"
    "net/http"

    docx2pdf "github.com/bobyeoh/docx2pdf-go"
)

func handle(w http.ResponseWriter, r *http.Request) {
    body, _ := io.ReadAll(r.Body)
    w.Header().Set("Content-Type", "application/pdf")
    _ = docx2pdf.ConvertReader(
        bytes.NewReader(body), int64(len(body)), w,
        docx2pdf.Options{},
    )
}
```

**3. Parse → inspect / modify → render.**

```go
doc, err := docx2pdf.Open("in.docx")
if err != nil {
    return err
}
for _, b := range doc.Body {
    if p, ok := b.(docx2pdf.Paragraph); ok && len(p.Runs) > 0 {
        // walk the AST: redact, translate, reformat, ...
    }
}
return docx2pdf.Render(doc, "out.pdf", docx2pdf.Options{})
```

### As a CLI

```bash
# Simplest — system font auto-detected.
docx2pdf -in input.docx -out output.pdf

# Batch — walks a directory tree, mirrors structure to -out.
docx2pdf -in indir/ -out outdir/ -recursive -keep-going -page-numbers
```

Run `docx2pdf -help` for the full flag list (font overrides, parallel
workers, lenient mode, author override, etc.).

### Docker

```bash
# Simplest — fonts auto-detected from $DOCX2PDF_FONT and
# $DOCX2PDF_FONT_CJK baked into the image.
docker run --rm -v "$PWD":/work bobyeoh/docx2pdf-go \
    -in /work/in.docx -out /work/out.pdf -page-numbers
```

The official image (~70 MB) ships **Noto Sans** for Latin text and
**WenQuanYi Zen Hei** for CJK fallback. Noto Sans CJK is *not*
bundled because it uses CFF/PostScript outlines (`.ttc` with `OTTO`
faces) which gopdf's TrueType-only parser can't render; WQY Zen Hei
is a TrueType TTC that the runtime extracts face 0 from automatically.

The env vars `DOCX2PDF_FONT` and `DOCX2PDF_FONT_CJK` are honored by
the binary when no `-font` / `-font-fallback` flag is given, so the
container works out of the box.

#### Running in *any* container — including `scratch` / `distroless`

The binary embeds a small Latin font as a last-resort fallback, so
even minimal images with no fonts at all produce a valid PDF for
Latin-only documents:

```dockerfile
# Multi-stage build into distroless: total size ~5 MB.
FROM golang:alpine AS build
WORKDIR /src
COPY . .
RUN CGO_ENABLED=0 go build -trimpath -ldflags="-s -w" \
    -o /docx2pdf ./cmd/docx2pdf

FROM gcr.io/distroless/static-debian12
COPY --from=build /docx2pdf /docx2pdf
ENTRYPOINT ["/docx2pdf"]
```

```bash
docker run --rm -v "$PWD":/work my-image \
    -in /work/in.docx -out /work/out.pdf  # no -font needed
```

CJK content in a fontless image still needs a CJK TTF — mount it and
point `$DOCX2PDF_FONT_CJK` at it:

```bash
docker run --rm -v "$PWD":/work \
    -e DOCX2PDF_FONT_CJK=/work/SimSun.ttf \
    my-image -in /work/in.docx -out /work/out.pdf
```

---

## What it actually renders

> **TL;DR**: solid for reports / contracts / generated paperwork in
> Latin + CJK. Loses fidelity on floating-image wrap, math typesetting,
> and SmartArt. See the ⚠️ / ❌ rows for the precise list.

### Text & inline formatting

| Feature | Status |
|---|---|
| Bold / italic / underline / strikethrough / color / size / sub-super-script | ✅ |
| Caps / small-caps / vertical position (`w:position`) / character scale (`w:w`) | ✅ |
| Letter spacing (`w:spacing` in rPr) | ✅ |
| Highlight + shading background fills | ✅ |
| Hidden text (`w:vanish`) — suppressed from output | ✅ |
| Theme colors / theme fonts (Heading / Body roles) | ✅ |
| Hyperlinks (external URL + internal anchors) — real PDF annotations | ✅ |
| Soft-hyphen / non-breaking hyphen / `w:sym` symbol references | ✅ |

### Paragraph & page layout

| Feature | Status |
|---|---|
| Alignment: left / center / right / justify | ✅ |
| Indent: left / first-line / hanging | ✅ |
| Line spacing (single / 1.5x / double / exact / atLeast) | ✅ |
| Tab stops + tab leaders (dot / hyphen / underscore) | ✅ |
| Drop caps (`w:framePr w:dropCap`) | ✅ |
| Multi-column layout (`w:cols`) | ✅ |
| Multi-section docs (per-section page size / orientation / margins) | ✅ |
| Headers & footers (default / first / even) | ✅ |
| Mirror margins, gutter, page borders, line numbers | ✅ |
| Page background color (`w:background` gated by `displayBackgroundShape`) | ✅ |
| Explicit page breaks (`w:br type="page"` / `w:pageBreakBefore`) | ✅ |

### Lists & tables

| Feature | Status |
|---|---|
| Numbered lists: decimal / lower-upper letter / lower-upper roman | ✅ |
| Bullet lists with custom text or picture bullets | ✅ |
| Multi-level lists + legal numbering (`w:isLgl`) + custom `start` | ✅ |
| Tables: column widths, `gridSpan` merge, `vMerge` merge | ✅ |
| Multi-page tables with header-row repeat | ✅ |
| Nested tables (table inside a cell) | ✅ |
| Cell shading + per-edge borders (single / double / dashed / dotted) | ✅ |
| Row `cantSplit` (keep row intact across page break) | ✅ |
| Table styles (`w:tblStyle` + `w:tblLook` conditional emphasis) | ✅ |

### Images & graphics

| Feature | Status |
|---|---|
| Inline images: PNG / JPEG / GIF | ✅ |
| Image cropping (`a:srcRect`) and explicit extent | ✅ |
| Legacy VML images (`w:pict` / `v:imagedata`) — older Word, Excel/Outlook pastes | ✅ |
| Anchored images (`wp:anchor`) — rendered as inline best-effort | ⚠️ |
| Text wrap around floating images | ❌ |
| SmartArt diagrams | ❌ |
| Charts (`c:chart`) — title / labels extracted as `[Chart: …]` text | ⚠️ |

### Document structure & metadata

| Feature | Status |
|---|---|
| Paragraph styles with `basedOn` chains + `docDefaults` | ✅ |
| PDF outline / sidebar bookmarks from `Heading1..9` + `Title` styles | ✅ |
| Fields: `PAGE` / `NUMPAGES` / `DATE` / `AUTHOR` / `SEQ` / `REF` / `HYPERLINK` | ✅ |
| Both field encodings: `fldChar` complex + `fldSimple` compact | ✅ |
| Footnotes & endnotes (refs as `[N]`, bodies at page bottom / trailer) | ✅ |
| Comments (`comments.xml`) — surfaced as a trailing "Comments" section | ✅ |
| Tracked changes (`w:ins` / `w:del` / `w:moveFrom` / `w:moveTo`) — accept-all | ✅ |
| Content controls (`w:sdt`) — block + inline, transparent wrapper | ✅ |
| `mc:AlternateContent` — Choice over Fallback | ✅ |
| Doc properties (`docProps/core.xml` + `app.xml`) → PDF /Info | ✅ |
| Settings (`settings.xml`) — defaultTabStop / evenAndOddHeaders / displayBackgroundShape | ✅ |

### Advanced / partial

| Feature | Status |
|---|---|
| Math equations (`m:oMath` / `m:oMathPara`) — text extracted as italic, structure lost | ⚠️ |
| Text boxes (`wps:txbx`) — inline-extracted as italic; box geometry not preserved | ⚠️ |
| Floating frames (`w:framePr`) — anchored correctly; body text does **not** wrap around | ⚠️ |
| RTL (Hebrew / Arabic) — word order reversed, right-aligned; no UAX#9 mixed-direction | ⚠️ |
| OLE / embedded objects — emit `[Embedded object]` placeholder | ⚠️ |
| Arabic letter shaping (initial / medial / final) | ❌ |
| Form controls' interactive behavior | ❌ |
| Embedded fonts (`w:embedRegular`) loaded from package | ❌ |

If your document hinges on the "❌" rows and you need pixel-perfect
rendering, **fall back to a LibreOffice-backed service**. The "⚠️" rows
preserve content but lose some structural fidelity — check whether
that's acceptable for your use case.

---

## Architecture

```
.docx (zip)
  │
  ├─ word/styles.xml         ─▶ ParagraphStyle map + docDefaults
  ├─ word/numbering.xml      ─▶ list definitions
  ├─ word/header*.xml        ─▶ block-level header content
  ├─ word/footer*.xml        ─▶ block-level footer content
  ├─ word/_rels/...          ─▶ rId → media | hyperlink | part
  ├─ word/media/*            ─▶ image.Image objects
  └─ word/document.xml       ─▶ Section[] each carrying its own
                                Body, PageSize, Margins, H/F
                                          │
                                          ▼
                                  renderer (gopdf)
                                          │
                                          ▼
                                       .pdf
```

The pipeline mirrors **[docx4j](https://github.com/plutext/docx4j)**'s
load-then-visit architecture, but collapses the intermediate XSL-FO step:
we draw directly to PDF via [signintech/gopdf](https://github.com/signintech/gopdf).
That's why it stays small — no FOP, no JAXB, no schema binding.

### Source layout

```
docx2pdf.go                  ← public API (re-exports via type aliases)
example_test.go             ← external-package smoke tests
cmd/docx2pdf/main.go        ← CLI
internal/docx/              ← OOXML parser
internal/render/            ← PDF renderer (one file per concern):
                                pdf.go        — entry points + state
                                page.go       — H/F, page breaks, footnotes
                                paragraph.go  — paragraph + list markers
                                frame.go      — positioned (w:framePr) frames
                                text.go       — atom model + line layout
                                table.go      — drawTable, drawRow, borders
                                image.go      — fit / crop / draw
                                fonts.go      — font registration + CJK + RTL
                                fields.go     — w:fldChar / w:instrText
                                util.go       — twips / hex helpers
internal/convert/           ← thin orchestrator (parse → render)
internal/verify/            ← test harness (see below)
```

Everything under `internal/` stays private — the package boundary lets us
refactor freely. The root `docx2pdf` package re-exports types via aliases so
consumers can still `type-assert` against `docx2pdf.Paragraph`,
`docx2pdf.Table`, etc. without reaching into internals.

---

## Tested seriously

| Layer | Tests | Notes |
|---|---|---|
| Unit (parser + render) | 30+ | XML decoding, style resolution, list numbering, field codes, settings, VML, theme |
| Unit (CLI) | 2 | Directory walking, extension handling |
| Public API smoke | 3 | Library is importable from outside the module |
| End-to-end | 127 | `docx → PDF → pdftotext + pdfinfo + PNG` per case |
| Comprehensive integration | 1 | Single 30+ feature docx validated with pdftotext, bbox, PDF byte structure, and PNG pixel sampling |
| Real-world corpus | 6 | Real Word docs from the docx4j project |
| Crash resistance | 6 | Empty zip, malformed XML, circular `basedOn`, corrupt images, 500-deep nesting |
| Golden image diff | all cases | Opt-in (`GOLDEN=1`); detects visual regressions via mean L1 pixel distance |
| Fuzz | 2 | `FuzzDocxOpen` + `FuzzInMemoryDocx`, run via `-fuzz` flag |
| Benchmarks | 4 | Parser + full pipeline at small/large scale |

The end-to-end harness builds synthetic `.docx` files in memory, runs them
through `Convert`, extracts text with `pdftotext`, asserts both content
substrings and page geometry (via `pdfinfo`), and saves PNG snapshots for
visual review (rendered with `pdftoppm`).

```bash
# All deterministic tests (~17 s with the comprehensive case)
go test ./...

# All tests including golden image diff
GOLDEN=1 go test ./...

# Stress
go test ./... -race           # race-clean
go test ./internal/verify/... -bench=. -benchtime=2s

# Coverage (~78 % total, ~76 % in the verify package — exercises body
# render via the full pipeline)
go test -coverpkg=./... -coverprofile=cover.out ./...
go tool cover -html=cover.out
```

The verify suite has caught real bugs other tests missed — `w:br` page
breaks not advancing pages, JPEG images failing the PNG re-encode path,
footnotes being enqueued twice inside table cells, list markers
overlapping their text when numbering.xml omitted indent.

---

## Performance

Apple M-series (arm64), `go test -bench=. -benchtime=2s`:

| Benchmark | Time per op |
|---|---|
| Parse 2-paragraph docx | **22 µs** |
| Parse 500-paragraph docx | **534 µs** |
| Convert 2-paragraph → PDF | **13 ms** (font load dominates) |
| Convert 500-paragraph → PDF | **21 ms** |

The 500-paragraph render is ~8 ms slower than the small one — the line
breaker and table layout scale linearly with cheap constants, leaving
the per-doc font load as the bulk of small-doc cost.

---

## Status & non-goals

docx2pdf-go aims to be **good enough for content-driven documents**: reports,
contracts, generated paperwork, internal tooling. It is **not** trying to
become a pixel-perfect Word replacement.

If you need complex DTP, real shape layout, SmartArt diagrams, full bidi
shaping, or embedded-font support — use LibreOffice as a backend. This
library exists for everyone who would rather not ship a 500 MB office
suite next to their Go service.

---

## FAQ

**Does it support `.doc` (Word 97-2003 binary format)?**
No. `.docx` only. Convert legacy `.doc` to `.docx` first (LibreOffice's
`soffice --convert-to docx` does this reliably).

**Does it support `.pptx` / `.xlsx`?**
No. Word documents only. PowerPoint and Excel are separate OOXML
schemas with their own renderers (and very different layout models —
in particular slides are absolute-positioned, which is the opposite of
flow text).

**Encrypted / password-protected docx?**
No. The zip is opened with `archive/zip` which doesn't know about the
ECMA-376 Agile Encryption layer. Decrypt first with a dedicated tool.

**Will the output look identical to Word?**
Close, not pixel-identical. Fonts, hyphenation, and floating-image
wrap differ. For Latin / CJK content-driven docs the difference is
usually below "reader notices anything off". For complex layouts
(magazine-style multi-column with image wrap), it won't.

**Why doesn't Noto Sans CJK work as the fallback font?**
gopdf only renders TrueType outlines. Noto Sans CJK uses
CFF/PostScript outlines (the `.ttc` contains `OTTO`-tagged faces);
gopdf rejects them. Use a TrueType CJK font like WenQuanYi Zen Hei
(bundled in our Docker image) or Source Han Sans's TTF distribution.
A clear error mentions this if you try.

**Can the AST be modified before rendering?**
Yes — `Open` returns a `*Document` whose `Body` / `Sections` / `Styles`
fields are exported. The example in Quick start §3 walks the AST.
This is useful for redaction, translation, or template fill-in
pipelines.

**Is the public API stable?**
Pre-1.0, so the surface may shift between minor releases. The
function signatures of `Convert` / `ConvertReader` / `Open` / `Render`
are unlikely to change; the AST struct fields (Paragraph, Run, Table)
may gain fields as features land. Pin a tag in production.

---

## Contributing

Issues and PRs welcome. Highest-impact missing features (in roughly that
order):

1. Text wrap around floating images / frames — needs per-line shape
   exclusion in the layout pass (`wp:anchor` with wrap geometry, `w:framePr`
   currently positions but doesn't wrap)
2. Full UAX#9 bidi for mixed-direction lines (Latin embedded in Arabic)
3. Arabic letter shaping (initial / medial / final connected forms)
4. SmartArt rendering
5. Embedded fonts (`w:embedRegular`) loaded from the package

---

## Acknowledgements

Heavily indebted to **[plutext/docx4j](https://github.com/plutext/docx4j)**
for the OOXML knowledge it has codified over more than a decade — the
parser layout, style-resolution model, and many of the edge cases
docx2pdf-go handles were figured out by reading its source first.

PDF rendering is provided by **[signintech/gopdf](https://github.com/signintech/gopdf)**.

---

## License

MIT.
