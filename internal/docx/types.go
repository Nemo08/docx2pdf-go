// Package docx defines the parsed AST of a .docx (OpenXML WordprocessingML) document.
//
// We deliberately model only the subset we can render: paragraphs, runs with
// basic character formatting, tables (single-level), inline images, and page
// breaks. This mirrors the pipeline docx4j uses (DOM-like tree → renderer),
// but stays narrow enough to keep in one head.
package docx

import "image"

// Document is the top-level parsed result of a .docx file.
type Document struct {
	Body         []Block
	Images       map[string]image.Image    // keyed by rId (relationship id)
	Hyperlink    map[string]string         // rId → external URL (from rels, TargetMode=External)
	Defaults     RunProps                  // default run properties from styles.xml docDefaults/rPrDefault
	ParaDefaults ParaDefaults              // default paragraph properties from styles.xml docDefaults/pPrDefault
	Styles       map[string]ParagraphStyle // styleId → resolved paragraph style (basedOn already flattened)
	CharStyles   map[string]RunProps       // styleId → resolved character style (basedOn flattened)
	PageSize     PageSize                  // from sectPr; falls back to A4
	Margins      Margins                   // from sectPr; falls back to defaults
	Numbering    Numbering                 // list definitions from numbering.xml
	HeaderBlocks []Block                   // default header content; rendered on every page
	FooterBlocks []Block                   // default footer content; rendered on every page
	Sections     []Section                 // body split at sectPr boundaries; len>=1
	// Footnotes / endnotes keyed by w:id, parsed from word/footnotes.xml and
	// word/endnotes.xml. The renderer optionally appends a "Footnotes" /
	// "Endnotes" trailing section at document end so the content survives.
	Footnotes map[string][]Block
	Endnotes  map[string][]Block
	// Comments are reviewer annotations from word/comments.xml, keyed by
	// w:id. The renderer appends a "Comments" trailing section so the
	// notes survive into the PDF rather than being silently dropped.
	Comments map[string][]Block
	// Charts maps relationship id → flattened text extracted from a
	// referenced chart part (word/charts/chartN.xml). We can't render
	// the data graphic but we surface titles, axis labels, and data
	// labels so the prose around the chart still makes sense.
	Charts map[string]string
	// Diagrams maps the dgm:relIds "r:dm" relationship id → flattened
	// text extracted from the SmartArt data part
	// (word/diagrams/dataN.xml). Each diagram surfaces as a list of
	// node texts joined with " → " so the conceptual structure
	// survives in the PDF text stream even though we don't render the
	// graphical shapes.
	Diagrams map[string]string
	// Theme is the parsed contents of word/theme/theme1.xml — color scheme +
	// font scheme — used to resolve w:themeColor / rFonts w:asciiTheme refs.
	Theme Theme
	// TableStyles maps a tblStyle ID to its flattened formatting block. The
	// renderer applies these when a w:tbl carries w:tblStyle.
	TableStyles map[string]TableStyle
	// Bookmarks captures the text body of each named bookmark so REF fields
	// can resolve to it. Populated as the parser walks paragraphs.
	Bookmarks map[string]string
	// Properties from core.xml / app.xml when present — used for AUTHOR /
	// TITLE fields and for /Info dictionary in the PDF.
	Properties Properties
	// Settings from word/settings.xml — doc-wide rendering knobs.
	Settings Settings
	// FootnoteSeparators captures any custom w:type="separator" /
	// "continuationSeparator" / "continuationNotice" footnote bodies.
	// Keys are the OOXML w:type values; absence means "use renderer
	// default" (a thin horizontal rule).
	FootnoteSeparators map[string][]Block
	// EndnoteSeparators is the endnote equivalent.
	EndnoteSeparators map[string][]Block
}

// Theme holds the bits of theme1.xml we read.
type Theme struct {
	Colors map[string]string // name (accent1/dk1/lt1/...) → 6-hex
	Fonts  map[string]string // role (majorAscii/minorAscii) → font name
}

// TableStyle is the flattened formatting for a named w:tblStyle.
type TableStyle struct {
	ID      string
	BasedOn string
	Run     RunProps
	// Table-level defaults applied to every cell unless the cell overrides.
	CellShading string
	Borders     CellBorders
	// TableBorders mirrors the style's <w:tblPr><w:tblBorders>. Used by
	// decodeTable to seed the table's effective tblBorders before the
	// table's own tblBorders override. Critical for the built-in
	// "TableGrid" style, which is how Word's default bordered tables
	// declare their grid lines — without applying this, those tables
	// render borderless.
	TableBorders TableBorders
	// Conditional formatting blocks keyed by Word's tblStylePr w:type.
	Conditional map[string]TableCondPr
}

// TableCondPr is one piece of conditional formatting (firstRow, lastRow,
// firstCol, lastCol, band1Horz, band2Horz, band1Vert, band2Vert, etc.).
type TableCondPr struct {
	Run         RunProps
	CellShading string
	Borders     CellBorders
}

// Settings captures the document-wide knobs from word/settings.xml that the
// renderer actually consumes. Everything in this struct is doc-level — per-
// section overrides live on Section. Unset (zero) values mean "fall back to
// the renderer's built-in default".
type Settings struct {
	// DefaultTabStopTwips is w:defaultTabStop — grid spacing for implicit
	// tabs when a paragraph defines no explicit w:tabs. Zero → renderer
	// uses 720 twips (half inch), the typical Word default.
	DefaultTabStopTwips int
	// EvenAndOddHeaders is w:evenAndOddHeaders — when true, sections may
	// supply distinct even-page header/footer references. When the setting
	// is absent, even-page references in sectPr should be ignored. We OR
	// this with the per-section flag so docs that only declare one of the
	// two still behave reasonably.
	EvenAndOddHeaders bool
	// DisplayBackgroundShape is w:displayBackgroundShape — the master
	// switch for w:background (page color). When absent, Word does NOT
	// paint the background even if a color is defined.
	DisplayBackgroundShape bool
}

// Properties mirrors a slice of word/docProps/core.xml + app.xml. Word/Office
// computes some counts (Pages/Words/Characters) and saves them into app.xml;
// we surface them for the PDF /Info dictionary.
type Properties struct {
	Title      string
	Author     string
	Subject    string
	Company    string
	Pages      int
	Words      int
	Characters int
	Lines      int
}

// ParaDefaults seeds every paragraph before its own pPr is applied. Modern
// Office writes spacing-after=8pt, line=1.08 here so unstyled paragraphs
// inherit the "Office 2010+ defaults" look.
type ParaDefaults struct {
	SpacingBefore float64
	SpacingAfter  float64
	LineHeight    LineHeight
}

// PageNumberType encodes w:pgNumType: starting value + numeric format.
type PageNumberType struct {
	Start int    // 0 = use natural (1, 2, 3 ...)
	Fmt   string // "decimal" (default), "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"
	// ChapStyle is the heading style level (1..9) prefixed before the
	// page number when w:chapStyle is set; ChapSep is the separator
	// character ("hyphen", "period", "colon", "emDash", "enDash").
	ChapStyle int
	ChapSep   string
}

// PageBorders encodes w:pgBorders — colored frame around the page.
type PageBorders struct {
	Top, Bottom, Left, Right BorderEdge
}

// FrameInfo encodes w:framePr's positioning attributes for a paragraph
// that should render as a floating frame rather than in the normal flow.
//
// Drop-cap framing is a special case handled separately on Paragraph
// (DropCap / DropCapLines); this struct is only populated when the frame
// has at least one positioning attribute (w:w, w:x, w:y, w:xAlign, ...).
type FrameInfo struct {
	WidthTwips  int    // w:w — frame width
	HeightTwips int    // w:h — frame height (0 = fit content)
	XTwips      int    // w:x — absolute horizontal offset from HAnchor
	YTwips      int    // w:y — absolute vertical offset from VAnchor
	HAnchor     string // "margin" (default), "page", "text"
	VAnchor     string // "margin", "page" (default), "text"
	XAlign      string // "", "left", "center", "right", "inside", "outside"
	YAlign      string // "", "top", "center", "bottom", "inside", "outside"
	Wrap        string // "auto" (default), "around", "tight", "through", "none", "notBeside"
	HRule       string // "auto", "exact", "atLeast" — applies to HeightTwips
}

// LineNumbering encodes w:lnNumType. The renderer paints at a fixed
// horizontal inset, so w:distance is intentionally not modeled — a per-doc
// value would drift from the actual draw position.
type LineNumbering struct {
	Start   int    // first line number (default 1)
	CountBy int    // every Nth line shown (default 1)
	Restart string // "newPage", "newSection", "continuous"
}

// Section represents one continuous range of body blocks that share the same
// page setup and header/footer references. A doc has at least one section; a
// new section starts wherever the body had an inline sectPr.
//
// Headers/footers come in three flavors per Word's titlePg / evenAndOddHeaders
// settings: "default" applies to every page where a more specific one isn't
// set; "first" overrides on page 1 (when TitlePg is true); "even" overrides
// on even pages (when EvenAndOddHeaders is true).
type Section struct {
	// Type is one of "nextPage" (default), "continuous", "evenPage",
	// "oddPage", "nextColumn". Only nextPage and continuous are honored;
	// the others fall back to nextPage.
	Type              string
	Blocks            []Block
	PageSize          PageSize
	Margins           Margins
	HeaderBlocks      []Block // default header
	FooterBlocks      []Block // default footer
	HeaderFirstBlocks []Block // first-page header (when TitlePg=true)
	FooterFirstBlocks []Block
	HeaderEvenBlocks  []Block // even-page header (when EvenAndOddHeaders=true)
	FooterEvenBlocks  []Block
	TitlePg           bool // honor HeaderFirstBlocks / FooterFirstBlocks on page 1
	EvenAndOddHeaders bool // honor *Even* blocks on even pages
	PageNumber        PageNumberType
	Borders           PageBorders   // page-perimeter frame
	BackgroundColor   string        // w:background w:color (hex)
	MirrorMargins     bool          // mirror left/right on facing pages
	GutterTwips       int           // additional inside margin (binding)
	LineNumbering     LineNumbering // w:lnNumType
	Columns           int           // w:cols w:num (1 = no columns)
	ColumnSpaceTwips  int           // w:cols w:space between columns
	// VAlign is w:vAlign — vertical page alignment for the section:
	// "" (default top), "center", "both" (justify), "bottom". Cover pages
	// often set this to "center" for the title.
	VAlign string
	// DocGrid is w:docGrid — the CJK line/character grid. When type is
	// "lines" or "linesAndChars" the renderer enforces an exact line height
	// derived from linePitch (1/20 pt), giving the per-page line-count look
	// that East-Asian docs expect.
	DocGrid DocGrid
	// FormProt is w:formProt — section is form-protected.
	FormProt bool
	// RtlGutter mirrors the gutter onto the right side for RTL languages.
	RtlGutter bool
	// FootnotePr / EndnotePr override doc-level note configuration for
	// this section: position, numbering format, restart policy, start at N.
	FootnotePr *NoteConfig
	EndnotePr  *NoteConfig
}

// DocGrid captures w:docGrid's three knobs.
type DocGrid struct {
	Type      string // "", "lines", "linesAndChars", "snapToChars", "default"
	LinePitch int    // 1/20 pt per line
	CharSpace int    // 1/100 pt added per char (only for linesAndChars)
}

// NoteConfig captures w:footnotePr / w:endnotePr settings.
type NoteConfig struct {
	Pos      string // "pageBottom", "beneathText", "sectEnd", "docEnd"
	NumFmt   string // "decimal", "upperRoman", ...
	NumStart int    // start at N
	Restart  string // "continuous", "eachSect", "eachPage"
}

// ParagraphStyle is a flattened paragraph-style definition from styles.xml.
// `basedOn` chains are resolved at load time so consumers never have to walk
// the inheritance graph themselves.
type ParagraphStyle struct {
	ID            string
	BasedOn       string
	Run           RunProps
	Alignment     Alignment
	HasAlignment  bool // discriminates AlignLeft default from explicit-left
	SpacingBefore float64
	SpacingAfter  float64
	LineHeight    LineHeight
}

// Block is either a Paragraph or a Table.
type Block interface{ isBlock() }

// Paragraph is a w:p element.
type Paragraph struct {
	Runs      []Run
	Alignment Alignment
	// SpacingBefore / SpacingAfter are extra vertical space in points.
	SpacingBefore float64
	SpacingAfter  float64
	PageBreak     bool      // a leading page break (w:br w:type="page" in first run)
	List          *ListInfo // non-nil if paragraph is a list item
	// IndentLeftPt is the body-text left indent in points (w:ind w:left).
	IndentLeftPt float64
	// IndentFirstLinePt is the first-line offset relative to IndentLeftPt
	// (w:ind w:firstLine for positive, w:ind w:hanging for negative).
	IndentFirstLinePt float64
	// LineHeight encodes w:spacing w:line + w:lineRule. Zero value means
	// "fall back to the renderer's default" (single-spacing with the natural
	// font line height).
	LineHeight LineHeight
	// KeepNext: prefer keeping this paragraph on the same page as the next.
	KeepNext bool
	// KeepLines: prefer not breaking lines of this paragraph across pages.
	KeepLines bool
	// ContextualSpacing: suppress SpacingBefore/After if the adjacent
	// paragraph uses the same style (typical for list items).
	ContextualSpacing bool
	// Bidi: paragraph reads right-to-left. We currently mark it but don't
	// perform RTL line layout — text is preserved in source order.
	Bidi bool
	// StyleID is the w:pStyle reference (resolved at decode time but kept
	// here so the renderer can apply contextualSpacing per-style sibling.)
	StyleID string
	// Tabs is the parsed w:tabs list — sorted by Pos.
	Tabs []TabStop
	// DropCap is "drop" or "margin" when w:framePr declares drop-cap on the
	// paragraph; "" otherwise. We render the first character at ~3× size as
	// an approximation (real wrap-around layout is out of scope).
	DropCap string
	// DropCapLines is the number of body lines the drop-cap visually spans.
	// Pulled from w:framePr w:lines; defaults to 3 when DropCap is set.
	DropCapLines int
	// Frame, when non-nil, declares this paragraph is a positioned frame
	// (w:framePr with placement attributes — distinct from the drop-cap
	// variant). The renderer draws at the anchored position without
	// advancing the document cursor; surrounding body text is NOT
	// reflowed around the frame, so wrapping with `wrap="around"` may
	// visually overlap.
	Frame *FrameInfo
	// Borders holds the four edges of <w:pBdr>. Markdown-style "---"
	// thematic breaks are commonly encoded as an empty paragraph with
	// only the bottom edge set. We reuse CellBorders because the shape
	// (Top / Bottom / Left / Right BorderEdge) is identical.
	Borders CellBorders
	// WidowControl — preserve at least 2 lines of the paragraph on a page
	// (default true in Word). Parsed but not yet honored at layout time.
	WidowControl *bool
	// MirrorIndents flips IndentLeftPt/Right on mirrored (verso) pages.
	MirrorIndents bool
	// AdjustRightInd — auto-adjust the right indent for East-Asian wrap.
	AdjustRightInd bool
	// SnapToGrid — when false, the paragraph opts out of the section's
	// docGrid line snapping.
	SnapToGrid *bool
	// OutlineLvl is w:outlineLvl (0-9). When >= 0 this paragraph contributes
	// to the PDF outline even if its style is not Heading*.
	OutlineLvl int
	// TextDirection is one of "lrTb" (default), "tbRl", "btLr", "lrTbV",
	// "tbRlV", "tbLrV". When non-default the renderer rotates the
	// paragraph's drawing region (vertical CJK text).
	TextDirection string
	// TextAlignment is the vertical baseline anchor inside a line:
	// "top", "center", "baseline", "bottom", "auto" (default "auto").
	TextAlignment string
	// CJK line-break / spacing controls.
	Kinsoku       *bool // w:kinsoku — honor line-break rules for CJK
	WordWrap      *bool // w:wordWrap — allow word break for Latin
	OverflowPunct *bool // w:overflowPunct — punctuation may hang outside text area
	TopLinePunct  *bool // w:topLinePunct — leading punctuation compression
	AutoSpaceDE   *bool // w:autoSpaceDE — auto-space CJK/Latin
	AutoSpaceDN   *bool // w:autoSpaceDN — auto-space CJK/numeric

	// endsSection is set when this paragraph's pPr contained an inline sectPr.
	// Internal-only: the parser uses it to know when to close out a section.
	endsSection bool
}

// TabStop is one entry of a paragraph's w:tabs definition.
//
//	Pos is the absolute position in points from the paragraph's left edge.
//	Val is one of "left" (default), "center", "right", "decimal", "clear".
//	Leader is "", "dot", "hyphen", "underscore", or "middleDot".
type TabStop struct {
	Pos    float64
	Val    string
	Leader string
}

// LineHeight encodes the w:spacing/@w:line attribute together with its
// @w:lineRule discriminator.
//
//   - Rule = "auto" (the default in Word): Mul is the multiplier; 1.0 = single,
//     1.5 = one-and-a-half, 2.0 = double. Pt is unused.
//   - Rule = "exact": Pt is the line height in points. Mul is unused.
//   - Rule = "atLeast": Pt is a minimum line height in points; the renderer
//     uses max(Pt, natural).
//   - Empty Rule = unset; the renderer keeps its default.
type LineHeight struct {
	Rule string
	Pt   float64
	Mul  float64
}

// ListInfo is a reference into Document.Numbering for a paragraph.
type ListInfo struct {
	NumID int
	Level int
}

func (Paragraph) isBlock() {}

// Run is one inline atom inside a paragraph. The decoder emits Runs in
// document order; a Run carries exactly one piece of content (text, image,
// break) OR exactly one field-structure marker. The renderer collapses field
// markers into resolved text via a small state machine — keeping field state
// out of the AST lets the parser stay stateless.
type Run struct {
	Text       string
	IsBreak    bool   // soft line break (w:br without page type)
	ImageID    string // rId if this run is a w:drawing/pic image
	LinkURL    string // hyperlink rId → external URL (resolved by renderer)
	LinkAnchor string // hyperlink w:anchor → internal bookmark name target
	Bookmark   string // when set, this is a marker placing a named anchor here
	// Explicit image size in points (from wp:extent in EMU). Zero means
	// "use the image's native dimensions scaled to content width if too big."
	ImageWidthPt, ImageHeightPt float64
	// Image source-rect crop in PERCENT (a:srcRect attrs are 1/1000 of percent
	// from each edge). E.g. CropTop=10000 = 10%. Zero = no crop on that side.
	CropTopPct, CropBottomPct, CropLeftPct, CropRightPct float64
	// ImageAnchored is true if the run comes from a wp:anchor (floating
	// image) rather than wp:inline. Renderer still draws inline as a
	// best-effort fallback; AnchorAlignH/V capture the requested anchor
	// alignment ("left", "center", "right", "inside", "outside") so the
	// inline placement can at least approximate the source location.
	ImageAnchored                    bool
	AnchorAlignH, AnchorAlignV       string
	AnchorOffsetXPt, AnchorOffsetYPt float64
	AnchorWrap                       string // "", "none", "square", "tight", "through", "topAndBottom"
	// FootnoteID, when non-empty, tags this run as a footnote / endnote
	// reference site. The visible Text is still drawn (typically as a
	// superscript marker); the renderer also queues the corresponding note
	// body for the current page's bottom area.
	FootnoteID string
	// HorizontalRule marks a run that should render as a horizontal
	// separator line at the paragraph's vertical position. Word emits
	// these as <w:pict><v:rect o:hr="t"/></w:pict> — Office's HTML-
	// compatibility way of representing markdown's "---" thematic break.
	HorizontalRule bool
	// IsEndnote distinguishes endnote refs from footnote refs (different
	// lookup map on the renderer side).
	IsEndnote bool
	Props     RunProps

	// --- Field structure (w:fldChar / w:instrText) ---
	// Exactly one of FieldBegin/FieldSep/FieldEnd is set on a marker run, OR
	// InstrText is non-empty for an instruction-text run. Marker runs carry
	// no visible content.
	FieldBegin bool
	FieldSep   bool // w:fldChar w:fldCharType="separate" — code → result boundary
	FieldEnd   bool
	InstrText  string
}

// RunProps captures character-level formatting we honor.
type RunProps struct {
	Bold       bool
	Italic     bool
	Underline  bool
	Strike     bool    // w:strike — single-line strikethrough
	Caps       bool    // w:caps — render as uppercase
	SmallCaps  bool    // w:smallCaps — lowercase rendered as small upper-case
	FontSize   float64 // half-points in docx; we store points
	FontFamily string
	Color      string // hex without leading '#', e.g. "FF0000"
	// Highlight is one of Word's predefined names (yellow, green, cyan, ...).
	// When non-empty the renderer paints a colored rect under the run.
	Highlight string
	// Shading is a 6-hex background color (run-level w:shd w:fill).
	Shading string
	// VertAlign is "superscript", "subscript", or "" (normal). Renderer
	// shifts the y baseline and reduces the font size to ~60%.
	VertAlign string
	// StyleID points at a named character style (w:rStyle val). Resolved at
	// run-construction time by merging the style's props underneath this run.
	StyleID string
	// Vanish suppresses the run from rendering (w:vanish). The text is still
	// kept in the AST so callers can introspect it.
	Vanish bool
	// PositionPt is w:position in points — raises/lowers baseline. Positive
	// = up. Distinct from VertAlign which also changes size.
	PositionPt float64
	// CharacterScale is w:w as a fraction (1.0 = 100% width). Applied at
	// draw time as text-matrix horizontal scale.
	CharacterScale float64
	// ThemeColor names the theme color slot (e.g. "accent1", "text1"). When
	// non-empty the renderer resolves it through Document.Theme.Colors.
	ThemeColor string
	// LumMod / LumOff are Word's HSL luminance adjustments derived from
	// w:themeShade and w:themeTint. LumMod < 1 darkens; LumOff > 0 lightens
	// toward white. Both in [0, 1].
	LumMod, LumOff float64
	// ThemeFontRole is "majorAscii", "minorAscii", "majorEastAsia", ... .
	// Resolved at draw time via Document.Theme.Fonts.
	ThemeFontRole string
	// LetterSpacingPt widens every glyph's advance by this many points
	// (w:spacing in rPr — Word stores 1/20 pt; we convert).
	LetterSpacingPt float64
	// TextEffect is "emboss", "imprint", "outline", or "". Renderer draws
	// emboss/imprint with a faint highlight stroke, outline with text fill
	// none + a stroke. These are approximations.
	TextEffect string
	// Em is the CJK emphasis mark (w:em) — "dot" / "circle" / "underDot" /
	// "comma" etc. The renderer draws a small mark above each glyph of the
	// run. Empty = no emphasis mark.
	Em string
	// Lang carries explicit language hints (w:lang) for the latin / CJK /
	// complex-script halves of the run. The renderer uses Lang.EastAsia
	// to bias the CJK fallback font selection even when the character itself
	// is ambiguous (a full-width digit is "0030 ZERO" — its language tells
	// us whether to draw it with the latin or the CJK face).
	Lang RunLang
	// RTL = w:rtl: the run reads right-to-left (Hebrew/Arabic). When
	// combined with Bidi on the paragraph the renderer reverses glyph
	// order for the run.
	RTL bool
	// CS = w:cs: this run is a complex-script run (Arabic/Hebrew/Thai). The
	// complex-script bold/italic/size attributes (BCs/ICs/SzCs) override the
	// regular B/I/Sz when CS is set.
	CS       bool
	BCs, ICs bool
	SzCs     float64
	// NoProof suppresses spellcheck flags; render-noop, but parsed for
	// completeness so consumers can introspect it.
	NoProof bool
	// WebHidden mirrors Word's "hide in web view" toggle. We treat it like
	// w:vanish for print output (web-hidden text shouldn't appear in PDF).
	WebHidden bool
	// KernThresholdPt is w:kern in half-points → points (font kerning
	// activates above this size threshold). Stored for completeness; the
	// underlying gopdf doesn't expose a kerning toggle so render is a no-op.
	KernThresholdPt float64
	// FitTextID + FitTextWidthPt are w:fitText — squeeze N runs into a
	// fixed width. We don't currently implement the squeeze; stored so the
	// caller can detect it.
	FitTextID      int
	FitTextWidthPt float64
	// TextBorder mirrors w:bdr (a border around the run's text — distinct
	// from paragraph or table borders). Empty Style means "no border".
	TextBorder BorderEdge
}

// RunLang carries the language hints from w:lang.
type RunLang struct {
	Latin    string // w:val — e.g. "en-US"
	EastAsia string // w:eastAsia — e.g. "zh-CN", "ja-JP"
	Bidi     string // w:bidi — e.g. "ar-SA", "he-IL"
}

// Alignment maps w:jc values.
type Alignment int

const (
	AlignLeft Alignment = iota
	AlignCenter
	AlignRight
	AlignJustify
)

// Table is a w:tbl element.
type Table struct {
	Rows []TableRow
	// ColumnWidthsTwips: column widths in twentieths of a point (twips).
	// 1440 twips = 1 inch.
	ColumnWidthsTwips []int
	// StyleID points at a named table style (w:tblPr/w:tblStyle val). The
	// parser flattens that style's tblPr / tcPr defaults into the table's
	// own properties at decode time.
	StyleID string
	// Look encodes w:tblLook flags from tblPr — which conditional formatting
	// blocks should apply (firstRow / lastRow / firstColumn / lastColumn,
	// banding etc.).
	Look TableLook
	// Borders carries <w:tblBorders> straight from tblPr. After parsing
	// is complete these are propagated into each cell's CellBorders (with
	// outer rows/columns taking the outer edge and interior cells taking
	// insideH/insideV) so the renderer only has to read CellBorders.
	Borders TableBorders
	// Layout is "auto" (default — column widths adjust to content) or
	// "fixed" (column widths are strictly the values in ColumnWidthsTwips
	// / tcW). The renderer is currently fixed-only; we record this so it
	// can be honored if/when an auto-fit pass lands.
	Layout string
	// TableWidthTwips / TableWidthType captures w:tblW (table total width).
	// Type is one of "", "auto", "dxa", "pct", "nil". Twips holds the dxa
	// value when Type=="dxa", or the pct value (0..5000 = 0..100%) when
	// Type=="pct".
	TableWidthTwips int
	TableWidthType  string
	// FloatPos carries w:tblPr/w:tblpPr — the floating-table anchor. When
	// non-nil the table is positioned absolutely; the renderer currently
	// still draws it in-flow (no text wrap), but the anchor is preserved
	// for callers that introspect the AST.
	FloatPos *TableFloatPos
	// Overlap is w:tblOverlap: "" or "never".
	Overlap string
	// Caption and Description mirror w:tblCaption / w:tblDescription —
	// accessibility metadata that surfaces in tagged PDFs.
	Caption     string
	Description string
	// DefaultCellMargins captures w:tblCellMar — default margins applied
	// to every cell unless that cell sets w:tcMar.
	DefaultCellMargins CellMargins
}

// CellMargins in points.
type CellMargins struct {
	Top, Bottom, Left, Right float64
}

// TableFloatPos captures w:tblpPr — anchor coordinates and text-wrap
// behavior for a floating table. Mirrors the FrameInfo shape used for
// paragraph frames.
type TableFloatPos struct {
	HAnchor                                                                      string // "margin", "page", "text"
	VAnchor                                                                      string // "margin", "page", "text"
	XAlign                                                                       string // "left", "center", "right", "inside", "outside"
	YAlign                                                                       string // "top", "center", "bottom", "inside", "outside"
	XTwips                                                                       int
	YTwips                                                                       int
	LeftFromTextTwips, RightFromTextTwips, TopFromTextTwips, BottomFromTextTwips int
}

// TableLook is the parsed w:tblLook bitfield.
type TableLook struct {
	FirstRow    bool
	LastRow     bool
	FirstColumn bool
	LastColumn  bool
	NoHBand     bool // suppress horizontal banding
	NoVBand     bool // suppress vertical banding
}

// Block interface for Table.
// (already satisfied below; declared here for clarity at the type definition.)

func (Table) isBlock() {}

type TableRow struct {
	Cells []TableCell
	// IsHeader marks a row with w:trPr/w:tblHeader. Headers repeat on every
	// page the table spans across.
	IsHeader bool
	// HeightTwips is the minimum row height from w:trHeight (rule="atLeast").
	// Zero means "natural" — use content height.
	HeightTwips int
	// HeightRuleExact means w:trHeight rule="exact" — render at exactly the
	// given height, clipping if content overflows.
	HeightRuleExact bool
	// CantSplit means the row must be drawn intact — if it won't fit on the
	// current page, push it to the next page first.
	CantSplit bool
	// WBeforeTwips / WAfterTwips are w:wBefore / w:wAfter — extra leading
	// or trailing column space (a fake first/last column whose width is
	// these twips). Parsed but not currently rendered.
	WBeforeTwips, WAfterTwips int
	// CellSpacingTwips is w:tblCellSpacing — space between cells. Parsed
	// but not rendered (we render with the standard zero-gap layout).
	CellSpacingTwips int
}

type TableCell struct {
	// A cell may contain paragraphs OR nested tables, in document order.
	// We use the same Block interface as the body so nesting Just Works.
	Blocks []Block
	// GridSpan is the number of columns this cell spans (default 1).
	GridSpan int
	// VMerge is "restart", "continue", or "" (no vertical merge).
	VMerge string
	// HMerge is "restart" or "continue" — Word's deprecated horizontal
	// merge predates GridSpan. Parsed for completeness; the renderer
	// resolves it the same way as GridSpan continuation cells (consumed
	// by the preceding cell's span).
	HMerge string
	// Shading is the 6-hex background fill color (w:shd w:fill).
	Shading string
	// VAlign is "top", "center", "bottom", or "" (default top).
	VAlign string
	// TextDirection rotates the cell's text. One of "" (default lrTb),
	// "tbRl" (90° clockwise, top-to-bottom right-to-left — Chinese/Japanese
	// vertical table headers), "btLr" (270° / 90° counter-clockwise — the
	// classic English rotated header), "lrTbV", "tbRlV", "tbLrV".
	TextDirection string
	// NoWrap suppresses line breaks: cell content stays on a single line
	// even if it overflows the column. Word uses this for narrow numeric
	// columns.
	NoWrap bool
	// HideMark: hide the paragraph-end mark inside this cell so it doesn't
	// contribute to row height. Parsed but unused by the renderer.
	HideMark bool
	// FitText: scale text horizontally to fit the column width.
	FitText bool
	// CellWidthType is "", "auto", "dxa", "pct", "nil"; CellWidthTwips
	// carries the dxa or pct value (pct stored as twips-equivalent).
	CellWidthType  string
	CellWidthTwips int
	// Borders, when set, override the default thin black per-edge borders.
	Borders CellBorders
	// Margins (w:tcMar) in points, defaulting to {Top: 0, Bottom: 0, Left: 4, Right: 4}
	// when zero. We only honor symmetric defaults; per-cell overrides take
	// precedence.
	MarginTopPt, MarginBottomPt, MarginLeftPt, MarginRightPt float64
}

// Paragraphs returns the paragraph-typed blocks in document order — kept as
// a convenience method so existing callers that iterated cell.Paragraphs
// can still do so (they now do `for _, p := range cell.Paragraphs() { ... }`).
func (c TableCell) Paragraphs() []Paragraph {
	out := make([]Paragraph, 0, len(c.Blocks))
	for _, b := range c.Blocks {
		if p, ok := b.(Paragraph); ok {
			out = append(out, p)
		}
	}
	return out
}

// CellBorders carries the four per-edge border specs. A zero Edge means
// "no border on that edge" — table-level borders are propagated into
// cells at parse time (see propagateTableBorders), so by the time the
// renderer sees a cell every meaningful edge is filled in.
type CellBorders struct {
	Top, Bottom, Left, Right BorderEdge
}

// TableBorders mirrors <w:tblBorders>. It carries the four outer edges
// plus the two "inside" edges that apply between cells (insideH between
// rows, insideV between columns).
type TableBorders struct {
	Top, Bottom, Left, Right BorderEdge
	InsideH, InsideV         BorderEdge
}

// Has reports whether any of the six edges is non-zero.
func (b TableBorders) Has() bool {
	return b.Top.Has() || b.Bottom.Has() || b.Left.Has() || b.Right.Has() ||
		b.InsideH.Has() || b.InsideV.Has()
}

// BorderEdge holds the style, width (points), and color (hex) for one edge.
//
//	Style examples Word writes: "single" (default), "double", "dashed",
//	"dotted", "thick", "none".
type BorderEdge struct {
	Style string
	Sz    float64 // line thickness in points (Word stores 1/8 pt; we convert)
	Color string  // 6-hex; empty = auto/black
}

// Has reports whether the edge carries any styling info.
func (e BorderEdge) Has() bool { return e.Style != "" || e.Sz != 0 || e.Color != "" }

// PageSize in twips. 1 pt = 20 twips.
type PageSize struct {
	WidthTwips  int
	HeightTwips int
}

// Margins in twips.
type Margins struct {
	Top    int
	Bottom int
	Left   int
	Right  int
}

// A4Twips is the standard A4 page size in twips. Used as a fallback when the
// document does not declare a w:sectPr/w:pgSz.
var A4Twips = PageSize{WidthTwips: 11906, HeightTwips: 16838}

// DefaultMarginsTwips is the typical 1-inch margin (1440 twips).
var DefaultMarginsTwips = Margins{Top: 1440, Bottom: 1440, Left: 1440, Right: 1440}

// Numbering is the parsed contents of word/numbering.xml.
//
// Word stores list definitions in two layers:
//   - abstractNum: a reusable template with one Level per indent depth.
//   - num: a concrete list instance pointing at an abstractNumId. Multiple
//     w:num entries can share an abstractNum (e.g. when a doc has many
//     separate bullet lists that all look the same).
//
// A paragraph references a list via w:numId + w:ilvl.
type Numbering struct {
	Abstract map[int]AbstractNum // abstractNumId → definition
	NumToAbs map[int]int         // numId → abstractNumId
	// PicBullets maps w:numPicBulletId → image rId. A level whose
	// w:lvlPicBulletId names one of these renders the image as its marker.
	PicBullets map[int]string
	// Overrides keyed by numId → ilvl → NumOverride. Captures
	// w:lvlOverride inside a w:num: per-numId level swaps and start
	// overrides. Empty map is fine.
	Overrides map[int]map[int]NumOverride
}

type AbstractNum struct {
	Levels map[int]NumLevel // ilvl → level definition
}

// NumLevel describes how one indent level is rendered.
type NumLevel struct {
	Format       string
	Text         string
	Start        int
	LeftTwips    int
	HangingTwips int
	// IsLgl is w:isLgl — Word forces all lvlText substitutions for THIS
	// level (and below) to render in decimal, regardless of their original
	// numFmt. Used for "1.2.3.4" legal-style numbering.
	IsLgl bool
	// PicBulletID, when > 0, names a w:numPicBullet whose image should be
	// used as this level's bullet marker. Resolved via Numbering.PicBullets.
	PicBulletID int
	// Suff is w:suff — what comes between the marker and the body:
	// "tab" (default), "space", or "nothing".
	Suff string
	// LvlRestart is w:lvlRestart — when set, this level restarts whenever
	// the level value is reached. 0 means "never restart"; negative means
	// "use default". Zero default = renderer uses Word's rule.
	LvlRestart int
	// PStyleLink is w:pStyle — a paragraph style that's linked to this
	// level. Paragraphs carrying that style number from this level
	// automatically.
	PStyleLink string
}

// NumOverride captures w:lvlOverride for a concrete num. We attach this to
// the Numbering map by storing overrides keyed by (numId, ilvl).
type NumOverride struct {
	// StartOverride > 0 forces a different starting value for this level
	// on this specific numId.
	StartOverride int
	// LvlReplace optionally swaps in a brand-new level definition.
	LvlReplace *NumLevel
}
