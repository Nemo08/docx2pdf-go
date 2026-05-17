// Package render walks the docx AST and writes a PDF via github.com/signintech/gopdf.
//
// Design parallels docx4j's FOExporterVisitor: a visitor walks blocks and
// emits drawing operations. Unlike docx4j we do not go through an
// intermediate XSL-FO document — we draw to PDF directly.
//
// File map:
//
//	pdf.go        — entry points, Options, renderer struct, RenderWriter
//	page.go       — page decorations, headers/footers, page break, footnotes
//	paragraph.go  — drawParagraph + list marker resolution
//	text.go       — atom model, line layout, runs→atoms
//	table.go      — drawTable, drawRow, borders, cell measurement
//	image.go      — image fit/crop/draw
//	fonts.go      — font registration, CJK fallback, color resolution
//	fields.go     — w:fldChar / w:instrText flattening, field codes
//	util.go       — twips/hex/file helpers
package render

import (
	"bytes"
	"context"
	"fmt"
	"image"
	"io"
	"os"
	"path/filepath"
	"time"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
	"github.com/signintech/gopdf"
)

// Options controls font selection, page numbering, fields, and tracing.
type Options struct {
	// SourceFilename and Author are surfaced to FILENAME / AUTHOR fields.
	// Empty values cause those fields to fall through to their cached text.
	SourceFilename string
	Author         string
	// Logger receives one-line progress messages (instead of Verbose stdout
	// printf). When nil and Verbose is true, falls back to stdout.
	Logger func(string)
	// OnProgress is called with a fraction in [0,1] after each section and
	// at the start of each page-decoration pass.
	OnProgress func(fraction float64, stage string)
	// Lenient: keep going past per-paragraph errors and log them. Useful
	// for crawling corpora of files of unknown quality.
	Lenient bool
	// ctx is set internally by RenderWithContext / RenderWriterWithContext.
	// External callers should use those entry points instead of poking ctx
	// directly. Public users get cancellation via convert.ConvertContext.
	ctx context.Context
	// prepopulatedBookmarkPages, when non-nil, seeds the renderer's
	// bookmark→page map from a prior dry-run pass. Used for PAGEREF
	// forward-reference resolution. Internal-only.
	prepopulatedBookmarkPages map[string]int
	// skipWrite, when true, asks RenderWriter to do all layout work but
	// skip the final WriteTo. The dry pass uses this to populate
	// bookmarkPages without producing a usable PDF.
	skipWrite bool

	// FontRegular is the path to the TTF used for normal text. When
	// empty, resolution order is: $DOCX2PDF_FONT env var, then a list
	// of common system-font locations (Arial / Helvetica on macOS,
	// DejaVu / Liberation / Noto on Linux), then a small embedded Go
	// font that ships with the binary so scratch / distroless /
	// fontless containers still work. The embedded face is Latin only;
	// CJK documents still need an explicit FontFallback.
	FontRegular string
	FontBold    string // optional; falls back to FontRegular
	FontItalic  string // optional
	// FontHeading is an optional TTF used for runs that the theme tags with
	// a "major" font role (w:rFonts w:asciiTheme="majorHAnsi" etc.). When
	// empty, theme-major runs fall back to FontRegular — which means modern
	// Word templates render headings in the body face. Set this to e.g.
	// Cambria.ttf to get the visual distinction Office shows by default.
	FontHeading string
	// FontFallback is a TTF used for runes the regular font cannot render
	// (typically CJK). Recommended: Noto Sans CJK or similar. When empty,
	// $DOCX2PDF_FONT_CJK is consulted; missing it just means CJK glyphs
	// share the regular face (and likely render as boxes).
	FontFallback string
	// DefaultFontSize is the size in points used when the document does
	// not specify one. Word's default is 11pt.
	DefaultFontSize float64
	// PageNumbers, when true, draws "X / N" centered in the bottom margin
	// of every page after the body is rendered.
	PageNumbers bool
	Verbose     bool
	// ShowRevisions controls how tracked-changes runs are rendered.
	// When false (default), del/moveFrom runs are silently dropped from
	// the output (Word's "Accept All" semantics). When true, del/moveFrom
	// runs render with strikethrough in red and ins/moveTo runs render
	// with underline in blue — Word's "Show Markup" view.
	ShowRevisions bool
	// MergeData supplies values for MERGEFIELD fields. Keys are
	// case-insensitive field names (e.g. "FirstName"). When a MERGEFIELD
	// resolves through this map, the renderer substitutes the value
	// (after applying \b prefix, \f suffix, and \* format switches);
	// otherwise it falls back to Word's cached result region.
	MergeData map[string]string
}

// RenderWithContext is Render with cancellation.
func RenderWithContext(ctx context.Context, doc *docx.Document, outPath string, opts Options) error {
	f, err := os.Create(outPath)
	if err != nil {
		return fmt.Errorf("create pdf: %w", err)
	}
	if err := RenderWriterWithContext(ctx, doc, f, opts); err != nil {
		f.Close()
		_ = os.Remove(outPath)
		return err
	}
	return f.Close()
}

// RenderWriterWithContext is RenderWriter with cancellation.
func RenderWriterWithContext(ctx context.Context, doc *docx.Document, w io.Writer, opts Options) error {
	opts.ctx = ctx
	return RenderWriter(doc, w, opts)
}

// Render writes doc to outPath as a PDF.
func Render(doc *docx.Document, outPath string, opts Options) error {
	f, err := os.Create(outPath)
	if err != nil {
		return fmt.Errorf("create pdf: %w", err)
	}
	if err := RenderWriter(doc, f, opts); err != nil {
		f.Close()
		_ = os.Remove(outPath) // don't leave a half-written file behind
		return err
	}
	return f.Close()
}

// RenderWriter is the streaming variant — writes the produced PDF to w.
func RenderWriter(doc *docx.Document, w io.Writer, opts Options) error {
	// Two-pass layout for PAGEREF forward-references. When the doc has any
	// PAGEREF field and we're not already running the second pass, do a
	// throwaway render first to populate bookmark→page mapping, then run
	// the real render with the populated map seeded into Options.
	if opts.prepopulatedBookmarkPages == nil && needsForwardPageRefPass(doc) {
		seed := map[string]int{}
		seedOpts := opts
		seedOpts.prepopulatedBookmarkPages = seed
		seedOpts.skipWrite = true
		seedOpts.OnProgress = nil
		seedOpts.Logger = nil
		var discard bytes.Buffer
		if err := RenderWriter(doc, &discard, seedOpts); err != nil {
			return err
		}
		opts.prepopulatedBookmarkPages = seed
	}
	if opts.FontRegular == "" {
		// Resolution order when no explicit font was passed:
		//   1. DOCX2PDF_FONT env var (set by our Docker image, also a
		//      convenient knob for containerized deployments).
		//   2. findSystemFont(): a list of common /usr/share/fonts/
		//      and macOS / Windows paths.
		//   3. Embedded Go font (~150 KB Latin face bundled into the
		//      binary) — last resort so scratch / distroless / fontless
		//      containers still produce output.
		opts.FontRegular = resolveFontFromEnv(envFontRegular)
		if opts.FontRegular == "" {
			opts.FontRegular = findSystemFont() // never empty: falls back to embedded
		}
	}
	// Symmetric env-var fallback for the CJK / symbol fallback font.
	// Resolution: explicit Options.FontFallback → $DOCX2PDF_FONT_CJK →
	// system-CJK auto-detection (Hiragino on macOS, WQY on Linux).
	// No final embedded fallback because the Go font is Latin only —
	// it wouldn't actually cover the glyphs callers need a fallback
	// FOR (CJK + Dingbats + arrows + etc.).
	if opts.FontFallback == "" {
		opts.FontFallback = resolveFontFromEnv(envFontFallback)
	}
	if opts.FontFallback == "" {
		opts.FontFallback = findSystemCJKFont()
	}
	if opts.DefaultFontSize == 0 {
		opts.DefaultFontSize = 11
	}

	sections := doc.Sections
	if len(sections) == 0 {
		sections = []docx.Section{{
			Blocks:       doc.Body,
			PageSize:     doc.PageSize,
			Margins:      doc.Margins,
			HeaderBlocks: doc.HeaderBlocks,
			FooterBlocks: doc.FooterBlocks,
		}}
		if sections[0].PageSize.WidthTwips == 0 {
			sections[0].PageSize = docx.A4Twips
		}
		if sections[0].Margins == (docx.Margins{}) {
			sections[0].Margins = docx.DefaultMarginsTwips
		}
	}

	pdf := gopdf.GoPdf{}
	firstW := twipsToPt(sections[0].PageSize.WidthTwips)
	firstH := twipsToPt(sections[0].PageSize.HeightTwips)
	pdf.Start(gopdf.Config{PageSize: gopdf.Rect{W: firstW, H: firstH}})

	parseRFCDate := func(s string) time.Time {
		if s == "" {
			return time.Time{}
		}
		if t, err := time.Parse(time.RFC3339, s); err == nil {
			return t
		}
		if t, err := time.Parse("2006-01-02T15:04:05Z", s); err == nil {
			return t
		}
		if t, err := time.Parse("2006-01-02", s); err == nil {
			return t
		}
		return time.Time{}
	}
	// Prefer the doc's own creation date when present; the new-PDF time
	// is only a fallback. This preserves provenance information for
	// archive workflows that audit document age.
	creationDate := parseRFCDate(doc.Properties.CreateDate)
	if creationDate.IsZero() {
		creationDate = time.Now()
	}
	// Fold the doc's keywords + custom DOC-properties into the subject
	// when subject is empty — gopdf's PdfInfo doesn't expose a Keywords
	// slot, but the Subject field is widely surfaced by readers and
	// avoids dropping the metadata entirely. When the doc already has a
	// subject, we leave it alone and the keywords stay accessible via
	// DOCPROPERTY field expansion in the body.
	subject := doc.Properties.Subject
	if subject == "" && doc.Properties.Keywords != "" {
		subject = doc.Properties.Keywords
	}
	pdf.SetInfo(gopdf.PdfInfo{
		Title:        doc.Properties.Title,
		Subject:      subject,
		Author:       firstNonEmpty(opts.Author, doc.Properties.Author),
		Creator:      firstNonEmpty(doc.CustomProperties["Application"], "docx2pdf-go"),
		Producer:     "docx2pdf-go (gopdf)",
		CreationDate: creationDate,
	})
	r := &renderer{
		pdf:      &pdf,
		doc:      doc,
		opts:     opts,
		fonts:    map[string]bool{},
		counters: map[int]map[int]int{},
		fields: fieldVars{
			now:         time.Now(),
			filename:    filepath.Base(opts.SourceFilename),
			author:      firstNonEmpty(opts.Author, doc.Properties.Author),
			title:       doc.Properties.Title,
			subject:     doc.Properties.Subject,
			company:     doc.Properties.Company,
			keywords:    doc.Properties.Keywords,
			comments:    doc.Properties.Comments,
			numWords:    doc.Properties.Words,
			numChars:    doc.Properties.Characters,
			totalMinutes: doc.Properties.TotalTime,
			createDate:  parseRFCDate(doc.Properties.CreateDate),
			saveDate:    parseRFCDate(doc.Properties.ModifyDate),
			printDate:   parseRFCDate(doc.Properties.PrintDate),
			seqCounters: map[string]int{},
			bookmarks:   doc.Bookmarks,
			bookmarkPages: func() map[string]int {
				// Seed from a prior dry pass when supplied; otherwise
				// start empty and let atomBookmark populate it.
				if opts.prepopulatedBookmarkPages != nil {
					return opts.prepopulatedBookmarkPages
				}
				return map[string]int{}
			}(),
			docProperties: buildDocProperties(doc),
			docVars:         doc.DocVars,
			bibliography:    doc.Bibliography,
			headings:        collectHeadings(doc),
			styleParagraphs: collectStyleParagraphs(doc),
			mergeData:       opts.MergeData,
			glossary:        doc.Glossary,
			tcEntries:       collectTCEntries(doc),
			xeEntries:       collectXEEntries(doc),
		},
	}
	if err := r.registerFonts(); err != nil {
		return err
	}
	r.footnoteLabels = buildNoteLabels(doc, false)
	r.endnoteLabels = buildNoteLabels(doc, true)

	// Track which sections each PDF page belongs to so stampPageDecorations
	// can look up the right header/footer per page.
	sectionPageStart := make([]int, len(sections))

	logFn := opts.Logger
	if logFn == nil && opts.Verbose {
		logFn = func(s string) { fmt.Println(s) }
	}
	if logFn == nil {
		logFn = func(string) {}
	}
	progressFn := opts.OnProgress
	if progressFn == nil {
		progressFn = func(float64, string) {}
	}

	for i, sec := range sections {
		if opts.ctx != nil {
			if err := opts.ctx.Err(); err != nil {
				return err
			}
		}
		progressFn(float64(i)/float64(len(sections)), fmt.Sprintf("section %d/%d", i+1, len(sections)))
		r.pageW = twipsToPt(sec.PageSize.WidthTwips)
		r.pageH = twipsToPt(sec.PageSize.HeightTwips)
		marL := twipsToPt(sec.Margins.Left)
		marR := twipsToPt(sec.Margins.Right)
		marT := twipsToPt(sec.Margins.Top)
		marB := twipsToPt(sec.Margins.Bottom)
		marL += twipsToPt(sec.GutterTwips)
		r.marL, r.marR, r.marT, r.marB = marL, marR, marT, marB
		r.contentW = r.pageW - r.marL - r.marR
		applyPageBorderMargins(r, sec.Borders)
		r.lineNumCounter = sec.LineNumbering.Start
		if r.lineNumCounter < 1 {
			r.lineNumCounter = 1
		}

		// Section break TYPE is recorded on the section that's ENDING (it
		// describes how the NEXT section starts), so the decision for
		// whether section[i] starts on a new page comes from section[i-1].
		startsNewPage := true
		if i == 0 {
			startsNewPage = false
		} else if sections[i-1].Type == "continuous" {
			startsNewPage = false
		}
		switch {
		case i == 0:
			pdf.AddPage()
			r.cursorY = r.marT
			primeContentStream(&pdf)
		case !startsNewPage:
			// Continuous: stay on the same page, adopt new geometry.
		default:
			pdf.AddPageWithOption(gopdf.PageOption{
				PageSize: &gopdf.Rect{W: r.pageW, H: r.pageH},
			})
			r.cursorY = r.marT
			primeContentStream(&pdf)
		}
		sectionPageStart[i] = pdf.GetNumberOfPages()

		r.numColumns = float64(sec.Columns)
		if r.numColumns < 1 {
			r.numColumns = 1
		}
		r.colGap = twipsToPt(sec.ColumnSpaceTwips)
		r.colDrawSep = sec.ColumnSeparator
		r.colSpecs = nil
		if r.numColumns > 1 {
			full := r.pageW - r.marL - r.marR
			if !sec.ColumnEqualWidth && len(sec.ColumnSpecs) == int(r.numColumns) {
				// Unequal widths: derive each column's (x, w) from the spec.
				r.colSpecs = make([]columnRect, len(sec.ColumnSpecs))
				x := r.marL
				for i, c := range sec.ColumnSpecs {
					w := twipsToPt(c.WidthTwips)
					r.colSpecs[i] = columnRect{x: x, w: w}
					x += w
					gap := twipsToPt(c.SpaceTwips)
					if i == len(sec.ColumnSpecs)-1 {
						gap = 0
					}
					x += gap
				}
				r.colW = r.colSpecs[0].w
				r.contentW = r.colW
				r.colBaseX = r.colSpecs[0].x
			} else {
				r.colW = (full - r.colGap*(r.numColumns-1)) / r.numColumns
				r.contentW = r.colW
				r.colBaseX = r.marL
			}
			r.colIdx = 0
			if r.colDrawSep {
				drawColumnSeparators(r, sec)
			}
		} else {
			r.colW = 0
			r.colBaseX = r.marL
			r.colIdx = 0
		}

		// Section vAlign — when the section's content fits on a single
		// page, pre-measure it and shift the starting cursorY so the
		// content lands centered (or bottom-aligned). Cover pages
		// commonly use "center" to vertically center a title.
		if sec.VAlign == "center" || sec.VAlign == "bottom" {
			h := r.measureBlocks(sec.Blocks)
			avail := (r.pageH - r.marB) - r.cursorY
			if h > 0 && h <= avail {
				slack := avail - h
				if sec.VAlign == "center" {
					r.cursorY += slack / 2
				} else {
					r.cursorY += slack
				}
			}
		}
		for _, b := range sec.Blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				if err := r.drawParagraph(v); err != nil {
					if opts.Lenient {
						logFn(fmt.Sprintf("lenient: skip paragraph: %v", err))
						continue
					}
					return err
				}
			case docx.Table:
				if err := r.drawTable(v); err != nil {
					if opts.Lenient {
						logFn(fmt.Sprintf("lenient: skip table: %v", err))
						continue
					}
					return err
				}
			}
		}
	}
	r.drawFootnotesAtBottom()

	progressFn(1.0, "done")

	// Endnotes always go at document end as a trailer (Word puts them
	// there too). Footnotes were already rendered at each page's bottom.
	if err := r.appendNotesSection(doc.Endnotes, "Endnotes"); err != nil {
		return err
	}
	// Comments are reviewer markup; they're not part of the visible body
	// in Word's default print view, but dropping them silently loses
	// content. Surface them as a trailing section after endnotes so a
	// human can still see them in the produced PDF.
	if err := r.appendCommentsSection(doc); err != nil {
		return err
	}

	if err := r.stampPageDecorations(sections, sectionPageStart); err != nil {
		return err
	}
	if opts.PageNumbers {
		if err := r.stampPageNumbers(); err != nil {
			return err
		}
	}

	if opts.skipWrite {
		// Dry-run pass: layout is complete and bookmarkPages is now
		// populated; the caller seeded their copy of the map into
		// opts.prepopulatedBookmarkPages so it sees the updates. Skip
		// the final WriteTo — we never want the discard buffer's bytes.
		return nil
	}

	if _, err := pdf.WriteTo(w); err != nil {
		return fmt.Errorf("write pdf: %w", err)
	}
	return nil
}

// renderer carries the drawing state through one Render call. Methods on
// renderer live in the topic-specific files (page.go, paragraph.go, ...).
type renderer struct {
	pdf         *gopdf.GoPdf
	doc         *docx.Document
	opts        Options
	pageW       float64
	pageH       float64
	marL        float64
	marR        float64
	marT        float64
	marB        float64
	contentW    float64
	cursorY     float64
	fonts       map[string]bool     // registered font names
	counters    map[int]map[int]int // numId → level → next counter value
	noPageBreak bool                // when true, ensureRoom never adds pages
	// Multi-column layout (w:cols).
	numColumns float64
	colW       float64
	colGap     float64
	colBaseX   float64
	colIdx     int
	// colSpecs, when non-empty, carries the per-column (width, space) pair
	// (in points) for unequal-width sections. The renderer uses it to size
	// each column individually and to know whether to draw separators
	// between adjacent columns.
	colSpecs []columnRect
	// colDrawSep is set when w:cols w:sep="1" — the renderer paints thin
	// vertical rules between adjacent columns.
	colDrawSep bool

	// colSepPending records that we need to draw separators on every page
	// of the active multi-column section. Cleared at section end.
	colSepPending bool
	// Line numbering state: counter advances per visible body line; reset
	// to LineNumbering.Start at each section.
	lineNumCounter int
	// croppedCache stores cropped image instances keyed by "<origID>:crop".
	croppedCache map[string]image.Image
	// pendingFootnotes holds IDs queued for page-bottom render. ensureRoom
	// (and the end-of-body finalizer) drains this list before a page break.
	pendingFootnotes []pendingNote
	// drawingFootnotes prevents the page-bottom draw from re-triggering
	// itself when ensureRoom calls into the same code path.
	drawingFootnotes bool
	fields           fieldVars
	lineHeight       docx.LineHeight
	// prevStyleID is the StyleID of the paragraph just drawn — used by
	// contextualSpacing to detect "same style as previous sibling".
	prevStyleID string
	// pendingMarker, if non-nil, is drawn at the first line's baseline
	// during layoutLine.flush() — used for hanging list markers.
	pendingMarker *pendingMarker
	// firstLineHangPt, when > 0, outdents the first physical line of the
	// active paragraph by that many points (Word's w:ind w:hanging). Cleared
	// after the first flush so subsequent lines wrap at the normal margin.
	firstLineHangPt float64
	// paragraphRTL is set while drawing a right-to-left paragraph.
	// layoutLine consults it to reverse line-internal atom order before
	// drawing; runsToAtoms uses it to reverse the rune sequence inside
	// RTL word atoms. Cleared at paragraph end.
	paragraphRTL bool
	// paragraphKinsoku is true when the current paragraph honors East
	// Asian line-break rules (forbidden start/end punctuation). Mirrors
	// w:kinsoku — defaults on for paragraphs that did not opt out.
	paragraphKinsoku bool
	// paragraphOverflowPunct is true when trailing CJK punctuation may
	// overhang the right margin instead of wrapping (w:overflowPunct).
	// Defaults true; gates the "keep no-start atom on this line"
	// behavior inside layoutLine.
	paragraphOverflowPunct bool
	// paragraphWordWrap is true when long Latin words may be split at
	// arbitrary code points to fit the line (w:wordWrap). Defaults true.
	paragraphWordWrap bool
	// frameLastBottom captures the cursor Y at the end of a drawFrame
	// call so wrap-around frames can register a float band sized to
	// the rendered content.
	frameLastBottom float64
	// activeTabs is the active paragraph's tab stops, used by layoutLine
	// to snap atomTab atoms to the next stop.
	activeTabs []docx.TabStop
	// embeddedFamilies maps docx font name to its registered gopdf
	// family slots (populated from doc.EmbeddedFonts).
	embeddedFamilies map[string]embeddedFamily
	// activeTableSpacing is the table-level w:tblCellSpacing (twips) for
	// the table currently being drawn.
	activeTableSpacing int
	// footnoteLabels / endnoteLabels map a w:id to its formatted display
	// label (honoring section-level numFmt / numStart / numRestart). The
	// renderer rewrites the literal "[id]" text on reference runs and on
	// page-bottom marker paragraphs through these maps.
	footnoteLabels map[string]string
	endnoteLabels  map[string]string
	// floatBand, when non-nil, narrows the line/paragraph horizontal band
	// so subsequent text flows beside a floating image / shape instead of
	// stacking below it. Implements w:wrap="square" (and best-effort
	// "tight") for anchored drawings whose positionH aligns left or right.
	// The band auto-clears the first time the renderer notices cursorY
	// crossed bottomY, falling back to the natural full-width geometry.
	floatBand *floatWrapBand
}

// floatWrapBand is the active text-wrap exclusion zone for a single
// floating drawing. Coordinates are in renderer (page) space.
type floatWrapBand struct {
	leftX    float64 // image's left edge
	rightX   float64 // image's right edge
	bottomY  float64 // y where the band ends (image top + image height)
	side     string  // "left" or "right" — which side of the page the image hugs
	gapPt    float64 // horizontal padding between image and flowing text
}

// pendingNote is one queued note awaiting page-bottom render.
type pendingNote struct {
	id      string
	endnote bool
}

// pendingMarker carries the next list marker to be drawn at the start of
// the first physical line of a paragraph.
type pendingMarker struct {
	text  string      // text marker (decimal/bullet/letter/roman)
	image image.Image // picture-bullet marker (alternative to text)
	x     float64
	props docx.RunProps
}
