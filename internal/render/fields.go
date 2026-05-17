package render

import (
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/bobyeoh/docx2pdf-go/internal/docx"
)

// firstNonEmpty returns the first non-empty string among its args.
func firstNonEmpty(ss ...string) string {
	for _, s := range ss {
		if s != "" {
			return s
		}
	}
	return ""
}

// formatPageNumber returns the page-number string in the requested format
// after applying the section's "start at N" offset.
func formatPageNumber(page int, fmt string) string {
	if page < 1 {
		page = 1
	}
	switch fmt {
	case "upperRoman":
		return roman(page, true)
	case "lowerRoman":
		return roman(page, false)
	case "upperLetter":
		return alphaLabel(page, true)
	case "lowerLetter":
		return alphaLabel(page, false)
	}
	return strconv.Itoa(page)
}

// fieldVars supplies values for w:instrText codes. The zero value means
// "use the docx-cached field result as-is" — the body's default behavior
// since Word keeps a snapshot of the rendered text so even unsupported
// fields look right enough. Header/footer rendering overrides these per page.
type fieldVars struct {
	page     int
	numPages int
	pageFmt  string

	now      time.Time
	filename string
	author   string
	title    string
	subject  string
	keywords string
	comments string
	company  string
	username string

	// Document-level metadata used by NUMWORDS / NUMCHARS / EDITTIME.
	numWords     int
	numChars     int
	totalMinutes int
	createDate   time.Time
	saveDate     time.Time
	printDate    time.Time

	seqCounters map[string]int
	bookmarks   map[string]string
	// bookmarkPages maps bookmark name → 1-based PDF page number where it
	// landed. Populated as the renderer walks bookmark markers; used by
	// PAGEREF for cross-references that fall after the body has been
	// laid out (i.e., during page-decoration stamping).
	bookmarkPages map[string]int
	// docProperties indexes custom + standard doc properties so the
	// DOCPROPERTY field can resolve `{ DOCPROPERTY "AppVersion" }`.
	docProperties map[string]string
	// docVars indexes settings.xml/w:docVars entries.
	docVars map[string]string
	// bibliography exposes parsed b:Source entries.
	bibliography map[string]docx.BibSource
	// headings carries every Heading 1-9 / Title paragraph for TOC.
	headings []tocEntry
	// setVars carries values that SET fields have assigned.
	setVars map[string]string
	// listNumCounters tracks per-LISTNUM-list counters.
	listNumCounters map[string]int
	// tableCtx is non-nil while drawing inside a table cell. FORMULA uses
	// it to resolve =SUM(ABOVE) / explicit A1 refs.
	tableCtx *tableContext
	// styleParagraphs indexes body paragraphs by their StyleID — used by
	// the STYLEREF field to surface "the current chapter" text.
	styleParagraphs map[string][]string
	// footnoteRefs maps bookmark name → footnote ID. Used by NOTEREF.
	footnoteRefs map[string]string
	// mergeData supplies MERGEFIELD values.
	mergeData map[string]string
	// glossary maps the docPart names from word/glossary/document.xml to
	// their plain-text payload. AUTOTEXT / GLOSSARY fields resolve their
	// first argument against this table.
	glossary map[string]string
	// tcEntries collects TC field markers — explicit TOC entries that the
	// document author placed outside heading styles.
	tcEntries []tocEntry
	// xeEntries collects XE field markers — explicit Index entries.
	xeEntries []string
}

// parseTCInstr parses a TC field instruction like
//
//	TC "My Custom Entry" \l 2 \f t
//
// returning the entry text and outline level (default 1). Returns ok=false
// when the instruction has no title.
func parseTCInstr(instrFull string) (tocEntry, bool) {
	s := strings.TrimSpace(instrFull)
	if !strings.HasPrefix(strings.ToUpper(s), "TC") {
		return tocEntry{}, false
	}
	s = strings.TrimSpace(s[2:])
	// First quoted token is the title; if unquoted, take up to the first \ switch.
	var title string
	if strings.HasPrefix(s, `"`) {
		if end := strings.Index(s[1:], `"`); end >= 0 {
			title = s[1 : 1+end]
			s = s[1+end+1:]
		}
	} else {
		end := strings.Index(s, `\`)
		if end < 0 {
			title = strings.TrimSpace(s)
			s = ""
		} else {
			title = strings.TrimSpace(s[:end])
			s = s[end:]
		}
	}
	if title == "" {
		return tocEntry{}, false
	}
	level := 1
	parts := strings.Fields(s)
	for i := 0; i < len(parts)-1; i++ {
		if parts[i] == `\l` {
			if n, err := strconv.Atoi(strings.Trim(parts[i+1], `"`)); err == nil && n >= 1 && n <= 9 {
				level = n
			}
		}
	}
	return tocEntry{Level: level, Text: title}, true
}

// parseXEInstr extracts the visible title from an XE field instruction.
// Subentries separated by ':' are flattened to a single "Major:Minor" string.
func parseXEInstr(instrFull string) string {
	s := strings.TrimSpace(instrFull)
	if !strings.HasPrefix(strings.ToUpper(s), "XE") {
		return ""
	}
	s = strings.TrimSpace(s[2:])
	if strings.HasPrefix(s, `"`) {
		if end := strings.Index(s[1:], `"`); end >= 0 {
			return s[1 : 1+end]
		}
	}
	end := strings.Index(s, `\`)
	if end < 0 {
		return strings.TrimSpace(s)
	}
	return strings.TrimSpace(s[:end])
}

// tableContext locates the cell currently being drawn so FORMULA fields
// can reach into sibling cells. Row/Col are post-gridSpan logical coords.
type tableContext struct {
	table *docx.Table
	row   int
	col   int
}

// tocEntry is one heading + outline level for TOC synthesis.
type tocEntry struct {
	Level int
	Text  string
}

// needsForwardPageRefPass reports whether the document contains any PAGEREF
// (or TOC) field whose body emits page numbers that depend on layout. When
// true, RenderWriter does an initial dry pass to populate the
// bookmark→page map so the real pass can substitute resolved values.
func needsForwardPageRefPass(doc *docx.Document) bool {
	if doc == nil {
		return false
	}
	var scan func(blocks []docx.Block) bool
	scan = func(blocks []docx.Block) bool {
		for _, b := range blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				for _, r := range v.Runs {
					if r.InstrText == "" {
						continue
					}
					instr := strings.ToUpper(r.InstrText)
					// PAGEREF: explicit forward ref. TOC: synthesizes a
					// PAGEREF list internally; same need.
					if strings.Contains(instr, "PAGEREF") || strings.Contains(instr, " TOC ") || strings.HasPrefix(strings.TrimSpace(instr), "TOC") {
						return true
					}
				}
			case docx.Table:
				for _, row := range v.Rows {
					for _, cell := range row.Cells {
						if scan(cell.Blocks) {
							return true
						}
					}
				}
			}
		}
		return false
	}
	for _, sec := range doc.Sections {
		if scan(sec.Blocks) {
			return true
		}
		if scan(sec.HeaderBlocks) || scan(sec.FooterBlocks) ||
			scan(sec.HeaderFirstBlocks) || scan(sec.FooterFirstBlocks) ||
			scan(sec.HeaderEvenBlocks) || scan(sec.FooterEvenBlocks) {
			return true
		}
	}
	if scan(doc.Body) {
		return true
	}
	if scan(doc.HeaderBlocks) || scan(doc.FooterBlocks) {
		return true
	}
	return false
}

// fieldCodeAndArgs splits an instrText like ` SEQ Figure \* ARABIC ` into the
// code ("SEQ") and the first non-switch argument.
func fieldCodeAndArgs(s string) (code, primary string) {
	s = strings.TrimSpace(s)
	parts := strings.Fields(s)
	if len(parts) == 0 {
		return "", ""
	}
	code = strings.ToUpper(parts[0])
	for _, p := range parts[1:] {
		if strings.HasPrefix(p, "\\") {
			continue
		}
		p = strings.Trim(p, `"`)
		if p != "" {
			primary = p
			break
		}
	}
	return code, primary
}

// hyperlinkFieldInstr decodes a HYPERLINK instrText into (target, isAnchor).
// `\l` means the primary arg is an internal bookmark name, not a URL.
func hyperlinkFieldInstr(s string) (target string, isAnchor bool) {
	s = strings.TrimSpace(s)
	parts := strings.Fields(s)
	skipNext := false
	for i := 1; i < len(parts); i++ {
		if skipNext {
			skipNext = false
			continue
		}
		p := parts[i]
		switch p {
		case "\\l":
			isAnchor = true
			continue
		case "\\o", "\\t", "\\m", "\\n":
			skipNext = true
			continue
		}
		if strings.HasPrefix(p, "\\") {
			continue
		}
		target = strings.Trim(p, `"`)
		return target, isAnchor
	}
	return "", isAnchor
}

// flattenFields walks a paragraph's raw Run stream and resolves
// w:fldChar / w:instrText structure into plain text runs.
func flattenFields(runs []docx.Run, vars fieldVars) []docx.Run {
	type frame struct {
		instr       strings.Builder
		inResult    bool
		code        string
		arg         string
		instrFull   string
		substituted bool
		linkURL     string
		linkAnchor  string
		suppress    bool
		formField   *docx.FormFieldInfo
	}
	var stack []*frame
	top := func() *frame {
		if len(stack) == 0 {
			return nil
		}
		return stack[len(stack)-1]
	}

	out := make([]docx.Run, 0, len(runs))
	for _, r := range runs {
		switch {
		case r.FieldBegin:
			stack = append(stack, &frame{formField: r.FormField})
		case r.FieldSep:
			if f := top(); f != nil {
				f.instrFull = f.instr.String()
				f.code, f.arg = fieldCodeAndArgs(f.instrFull)
				f.inResult = true
				switch f.code {
				case "HYPERLINK":
					target, isAnchor := hyperlinkFieldInstr(f.instrFull)
					if isAnchor {
						f.linkAnchor = target
					} else {
						f.linkURL = target
					}
				case "REF", "PAGEREF", "NOTEREF":
					// \h turns the cross-reference into an internal link
					// that jumps to the named bookmark. Word's other ref
					// switches (\n paragraph number, \w full paragraph
					// number, \r relative number) reach into the numbering
					// engine and are out of scope; we surface the cached
					// result for those.
					if f.arg != "" && hasFlagSwitch(f.instrFull, "h") {
						f.linkAnchor = f.arg
					}
				case "SET":
					name, value := setFieldInstr(f.instrFull)
					if name != "" {
						if vars.setVars == nil {
							vars.setVars = map[string]string{}
						}
						vars.setVars[name] = value
					}
					f.suppress = true
				case "ADVANCE":
					f.suppress = true
				case "TC", "XE", "RD", "PRIVATE":
					// TC: TOC entry marker. XE: Index entry marker.
					// RD: Reference document. PRIVATE: app-specific data.
					// All have no visible result; recorded so the caller can
					// later mine them. We harvest below in vars.
					if vars.tcEntries == nil {
						vars.tcEntries = []tocEntry{}
					}
					if f.code == "TC" {
						if entry, ok := parseTCInstr(f.instrFull); ok {
							vars.tcEntries = append(vars.tcEntries, entry)
						}
					} else if f.code == "XE" {
						if title := parseXEInstr(f.instrFull); title != "" {
							vars.xeEntries = append(vars.xeEntries, title)
						}
					}
					f.suppress = true
				case "FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN":
					// Form fields: synthesize visible output from the
					// parsed ffData, replacing whatever Word cached.
					if v, ok := formFieldOutput(f.formField, f.code); ok {
						sample := docx.Run{Text: v}
						// Look ahead for the next non-marker run's props
						// to inherit font/size; otherwise fall back to
						// default props. We use a placeholder Run carrying
						// just text — applied by the substitute branch.
						_ = sample
						f.substituted = false
						// Treat as "value supplied at result time": we'll
						// catch the visible run below and override its text.
						// Mark suppress=true so cached glyphs are dropped;
						// we then emit a single synthetic run after.
						f.suppress = true
						// Emit the synthesized run immediately so it
						// renders even if the cached result region was
						// empty (common for FORMCHECKBOX).
						out = append(out, docx.Run{Text: v, Props: r.Props})
					}
				}
			}
		case r.FieldEnd:
			if n := len(stack); n > 0 {
				f := stack[n-1]
				// Form fields with no SEPARATE phase: synthesize here.
				if f.formField != nil && f.code == "" {
					_, fallbackCode := formFieldKindCode(f.formField)
					if v, ok := formFieldOutput(f.formField, fallbackCode); ok {
						out = append(out, docx.Run{Text: v, Props: r.Props})
					}
				}
				stack = stack[:n-1]
			}
		case r.InstrText != "":
			if f := top(); f != nil && !f.inResult {
				f.instr.WriteString(r.InstrText)
			}
		default:
			f := top()
			if f == nil {
				out = append(out, r)
				continue
			}
			if !f.inResult {
				continue
			}
			if f.suppress {
				continue
			}
			if value, ok := lookupFieldValueFull(f.code, f.arg, f.instrFull, vars); ok {
				if !f.substituted {
					// Apply \* general-format switches (Upper/Lower/
					// roman/Hex/Ordinal/...) and SYMBOL \f font.
					value = applyGeneralFormatSwitch(value, f.instrFull)
					props := r.Props
					if f.code == "SYMBOL" {
						if fontName := symbolFontSwitch(f.instrFull); fontName != "" {
							props.FontFamily = fontName
						}
						if sz := symbolFontSizeSwitch(f.instrFull); sz > 0 {
							props.FontSize = sz
						}
					}
					if strings.Contains(value, "\n") {
						lines := strings.Split(value, "\n")
						for i, line := range lines {
							if i > 0 {
								out = append(out, docx.Run{IsBreak: true, Props: props})
							}
							rr := r
							rr.Text = line
							rr.Props = props
							out = append(out, rr)
						}
					} else {
						rr := r
						rr.Text = value
						rr.Props = props
						out = append(out, rr)
					}
					f.substituted = true
				}
				continue
			}
			if f.linkURL != "" || f.linkAnchor != "" {
				rr := r
				if f.linkURL != "" {
					rr.LinkURL = f.linkURL
				}
				if f.linkAnchor != "" {
					rr.LinkAnchor = f.linkAnchor
				}
				out = append(out, rr)
				continue
			}
			out = append(out, r)
		}
	}
	return out
}

// lookupFieldValueWith is the legacy entry point that drops the full
// instrText. Kept for tests that don't need switches.
func lookupFieldValueWith(code, arg string, vars fieldVars) (string, bool) {
	return lookupFieldValueFull(code, arg, "", vars)
}

// lookupFieldValueFull resolves one field reference to its rendered value.
// instrFull is the entire instrText (e.g. "SYMBOL 61472 \\f Wingdings"); a
// few field codes (SYMBOL, FORMULA, REF) need switches beyond the primary
// arg. Returning (_, false) lets the caller fall back to the cached Word
// result.
func lookupFieldValueFull(code, arg, instrFull string, vars fieldVars) (string, bool) {
	switch code {
	case "PAGE":
		if vars.page > 0 {
			return formatPageNumber(vars.page, vars.pageFmt), true
		}
	case "NUMPAGES":
		if vars.numPages > 0 {
			return formatPageNumber(vars.numPages, vars.pageFmt), true
		}
	case "DATE":
		if !vars.now.IsZero() {
			return formatFieldDateTime(vars.now, instrFull, "2006-01-02"), true
		}
	case "TIME":
		if !vars.now.IsZero() {
			return formatFieldDateTime(vars.now, instrFull, "15:04"), true
		}
	case "CREATEDATE":
		when := vars.createDate
		if when.IsZero() {
			when = vars.now
		}
		if !when.IsZero() {
			return formatFieldDateTime(when, instrFull, "2006-01-02"), true
		}
	case "SAVEDATE":
		when := vars.saveDate
		if when.IsZero() {
			when = vars.now
		}
		if !when.IsZero() {
			return formatFieldDateTime(when, instrFull, "2006-01-02"), true
		}
	case "PRINTDATE":
		when := vars.printDate
		if when.IsZero() {
			when = vars.now
		}
		if !when.IsZero() {
			return formatFieldDateTime(when, instrFull, "2006-01-02"), true
		}
	case "EDITTIME":
		// w:TotalTime in minutes — surfaced from app.xml. Honor a
		// \# format switch when present (e.g. "h:mm").
		if vars.totalMinutes > 0 {
			h := vars.totalMinutes / 60
			m := vars.totalMinutes % 60
			if strings.Contains(instrFull, "\\#") {
				if v := formatNumericSwitch(float64(vars.totalMinutes), instrFull); v != "" {
					return v, true
				}
			}
			if h > 0 {
				return strconv.Itoa(h) + "h " + strconv.Itoa(m) + "m", true
			}
			return strconv.Itoa(vars.totalMinutes) + "m", true
		}
		return "", false
	case "NUMWORDS":
		if vars.numWords > 0 {
			return formatNumericValue(float64(vars.numWords), instrFull), true
		}
		return "", false
	case "NUMCHARS":
		if vars.numChars > 0 {
			return formatNumericValue(float64(vars.numChars), instrFull), true
		}
		return "", false
	case "FILENAME":
		if vars.filename != "" {
			return vars.filename, true
		}
	case "USERNAME":
		if vars.username != "" {
			return vars.username, true
		}
		// Fall through to the author when USERNAME is unset — close enough
		// for most templates.
		if vars.author != "" {
			return vars.author, true
		}
	case "USERINITIALS":
		// Approximate initials from the username/author.
		if name := firstNonEmpty(vars.username, vars.author); name != "" {
			return initialsOf(name), true
		}
	case "AUTHOR":
		if vars.author != "" {
			return vars.author, true
		}
	case "LASTSAVEDBY":
		if vars.author != "" {
			return vars.author, true
		}
	case "SEQ":
		if arg != "" && vars.seqCounters != nil {
			// Switches:
			//   \r N   — reset counter to N (then return N)
			//   \c     — repeat last value, do not increment
			//   \h     — increment but emit no visible text
			//   \n     — explicit "next" (default)
			//   \* fmt — formatted via applyGeneralFormatSwitch later
			if n, ok := seqResetSwitch(instrFull); ok {
				vars.seqCounters[arg] = n
				return strconv.Itoa(n), true
			}
			if seqHasFlag(instrFull, "c") {
				v := vars.seqCounters[arg]
				if v == 0 {
					v = 1
				}
				return strconv.Itoa(v), true
			}
			vars.seqCounters[arg]++
			if seqHasFlag(instrFull, "h") {
				return "", true
			}
			return strconv.Itoa(vars.seqCounters[arg]), true
		}
	case "REF":
		// REF consults SET-assigned variables first, then bookmarks.
		if arg != "" {
			if vars.setVars != nil {
				if v, ok := vars.setVars[arg]; ok && v != "" {
					return v, true
				}
			}
			if vars.bookmarks != nil {
				if text, ok := vars.bookmarks[arg]; ok && text != "" {
					return text, true
				}
			}
		}
	case "PAGEREF":
		// PAGEREF resolves to the page number of a bookmark. Prefer the
		// bookmarkPages index (populated as the body is laid out); the
		// `\h` switch makes it a hyperlink — the linking is handled by
		// the surrounding HYPERLINK or by a separate annotation, so we
		// just emit the number here.
		if arg != "" {
			if vars.bookmarkPages != nil {
				if pg, ok := vars.bookmarkPages[arg]; ok && pg > 0 {
					return strconv.Itoa(pg), true
				}
			}
			if vars.bookmarks != nil {
				if text, ok := vars.bookmarks[arg]; ok && text != "" {
					return text, true
				}
			}
		}
		return "", false
	case "NOTEREF":
		// NOTEREF resolves to a footnote/endnote reference number. We
		// surface the bookmark text when possible (the bookmark's
		// content typically IS the note ID), then fall back to the
		// cached result.
		if arg != "" {
			if vars.footnoteRefs != nil {
				if id, ok := vars.footnoteRefs[arg]; ok && id != "" {
					return id, true
				}
			}
			if vars.bookmarks != nil {
				if text, ok := vars.bookmarks[arg]; ok && text != "" {
					return text, true
				}
			}
		}
		return "", false
	case "STYLEREF":
		// STYLEREF prints the most-recent text styled with the named
		// style. The ideal implementation needs per-page state we don't
		// track; instead we return the FIRST paragraph that uses the
		// named style, which is the typical "current chapter" answer for
		// headers on every page of a single-chapter section.
		if arg != "" && vars.styleParagraphs != nil {
			if texts, ok := vars.styleParagraphs[arg]; ok && len(texts) > 0 {
				return texts[0], true
			}
		}
		return "", false
	case "TITLE":
		if vars.title != "" {
			return vars.title, true
		}
	case "SUBJECT":
		if vars.subject != "" {
			return vars.subject, true
		}
	case "KEYWORDS":
		if vars.keywords != "" {
			return vars.keywords, true
		}
	case "COMMENTS":
		if vars.comments != "" {
			return vars.comments, true
		}
	case "COMPANY":
		if vars.company != "" {
			return vars.company, true
		}
	case "DOCPROPERTY":
		if arg != "" && vars.docProperties != nil {
			if v, ok := vars.docProperties[arg]; ok && v != "" {
				return v, true
			}
		}
		return "", false
	case "DOCVARIABLE":
		if arg != "" && vars.docVars != nil {
			if v, ok := vars.docVars[arg]; ok && v != "" {
				return v, true
			}
		}
		return "", false
	case "CITATION":
		if arg != "" && vars.bibliography != nil {
			if src, ok := vars.bibliography[arg]; ok {
				return formatCitation(src), true
			}
		}
		return "", false
	case "BIBLIOGRAPHY":
		if vars.bibliography != nil && len(vars.bibliography) > 0 {
			return formatBibliography(vars.bibliography), true
		}
		return "", false
	case "MERGEFIELD":
		// MERGEFIELD names a mail-merge column. When the caller supplied
		// Options.MergeData, look up the value (case-insensitive). With
		// no MergeData, fall through to the cached result so templates
		// already pre-merged by Word render correctly.
		if arg != "" && vars.mergeData != nil {
			v, ok := mergeDataLookup(vars.mergeData, arg)
			if !ok {
				return "", false
			}
			pre, post := mergeFieldAffixes(instrFull)
			return pre + v + post, true
		}
		return "", false
	case "FORMTEXT":
		// FORMTEXT shows the result region's content as-is — return ""
		// + false so the result region's text streams through normally.
		return "", false
	case "FORMCHECKBOX":
		// Checkbox: cached result is empty (Word draws the box from a
		// separate FFData blob). Surface ☐ as a visible placeholder.
		return "☐", true
	case "FORMDROPDOWN":
		// Dropdown: same situation as FORMCHECKBOX — surface ▾ as the
		// "selected value" placeholder when no result was cached.
		return "▾", true
	case "QUOTE":
		// QUOTE simply emits its argument as text.
		if arg != "" {
			return strings.Trim(arg, `"`), true
		}
	case "IF":
		// IF is a conditional expression of the shape
		//   IF <expr1> <op> <expr2> "trueText" "falseText"
		// where op is = / <> / < / > / <= / >=. Word also allows the wildcard
		// pattern "* / ?". We evaluate the comparison and return the chosen
		// branch text; if the instruction can't be parsed we fall back to the
		// cached result.
		if v, ok := evaluateIfField(instrFull); ok {
			return v, true
		}
		return "", false
	case "INCLUDETEXT":
		// INCLUDETEXT references an external file or rel target. We can't
		// safely read arbitrary host-filesystem paths from PDF rendering,
		// but we DO honor the "bookmark" reference form
		//   INCLUDETEXT <path> <bookmark>
		// when the bookmark resolves locally — useful for self-referential
		// templates. Falls back to the cached result otherwise.
		toks := tokenizeFieldArgs(arg)
		if len(toks) >= 2 {
			if v, ok := vars.bookmarks[toks[1]]; ok {
				return v, true
			}
		}
		return "", false
	case "INCLUDEPICTURE":
		// External picture: we can't open arbitrary paths. The result
		// region (a w:drawing) already carries the image — let it through.
		return "", false
	case "TOC":
		entries := vars.headings
		// Merge TC field entries (explicit author-added TOC marks) so they
		// appear alongside the heading-derived ones in source order.
		// We don't try to splice by position; TC entries just get appended.
		entries = append(entries, vars.tcEntries...)
		if len(entries) > 0 {
			return formatTOC(entries), true
		}
		return "", false
	case "INDEX":
		// INDEX synthesizes an index from XE entries. We emit a simple
		// alphabetical list when XE markers were found.
		if len(vars.xeEntries) > 0 {
			return formatIndex(vars.xeEntries), true
		}
		return "", false
	case "TOA":
		return "", false
	case "AUTOTEXT", "GLOSSARY":
		// Both fields look up a docPart by name. AUTOTEXT and GLOSSARY
		// take the docPart name as the first arg; we hand back the
		// parsed plain-text body. Fall through to the cached result
		// when the name isn't in the glossary or the package shipped
		// without one.
		name := strings.TrimSpace(arg)
		name = strings.Trim(name, "\"")
		if name == "" || vars.glossary == nil {
			return "", false
		}
		if v, ok := vars.glossary[name]; ok {
			return v, true
		}
		return "", false
	case "ADDRESSBLOCK", "GREETINGLINE", "MACROBUTTON", "AUTOTEXTLIST":
		// Mail-merge / interactive elements with cached display text.
		return "", false
	case "EQ":
		// Legacy equation field — Word stores the typeset glyphs in the
		// result region. Let those through.
		return "", false
	case "HYPERLINK":
		return "", false
	case "SYMBOL":
		// SYMBOL embeds a single glyph by code point + font.
		if cp, ok := parseSymbolCodePointWithSwitches(arg, instrFull); ok {
			return string(cp), true
		}
		return "", false
	case "LISTNUM":
		listName := arg
		if listName == "" {
			listName = "__default__"
		}
		start, hasStart := listNumStart(instrFull)
		if vars.listNumCounters == nil {
			vars.listNumCounters = map[string]int{}
		}
		if hasStart {
			vars.listNumCounters[listName] = start
		} else {
			vars.listNumCounters[listName]++
		}
		return strconv.Itoa(vars.listNumCounters[listName]) + ")", true
	}
	if isFormulaCode(code) {
		if vars.tableCtx == nil {
			// Pure arithmetic still works without a table context.
			expr := formulaExpression(code, arg, instrFull)
			if expr == "" {
				return "", false
			}
			v, ok := evalTableFormula(expr, nil)
			if !ok {
				return "", false
			}
			return formatFormulaNumber(v), true
		}
		expr := formulaExpression(code, arg, instrFull)
		if expr == "" {
			return "", false
		}
		v, ok := evalTableFormula(expr, vars.tableCtx)
		if !ok {
			return "", false
		}
		return formatFormulaNumber(v), true
	}
	return "", false
}

// setFieldInstr parses ` SET name "value" ` into its name/value pair.
func setFieldInstr(s string) (name, value string) {
	s = strings.TrimSpace(s)
	parts := strings.Fields(s)
	if len(parts) == 0 || strings.ToUpper(parts[0]) != "SET" {
		return "", ""
	}
	if len(parts) >= 2 {
		name = parts[1]
	}
	if i := strings.Index(s, name); i >= 0 {
		rest := strings.TrimSpace(s[i+len(name):])
		if strings.HasPrefix(rest, `"`) {
			if j := strings.Index(rest[1:], `"`); j >= 0 {
				value = rest[1 : 1+j]
				return name, value
			}
		}
		if rest != "" {
			value = strings.Fields(rest)[0]
		}
	}
	return name, value
}

// parseSymbolCodePoint decodes a SYMBOL field's primary arg into a rune.
func parseSymbolCodePoint(arg string) (rune, bool) {
	arg = strings.TrimSpace(arg)
	if arg == "" {
		return 0, false
	}
	base := 10
	if strings.HasPrefix(arg, "0x") || strings.HasPrefix(arg, "0X") {
		arg = arg[2:]
		base = 16
	}
	n, err := strconv.ParseInt(arg, base, 32)
	if err != nil || n <= 0 {
		return 0, false
	}
	r := rune(n)
	if !utf8.ValidRune(r) {
		return 0, false
	}
	return r, true
}

// parseSymbolCodePointWithSwitches is parseSymbolCodePoint but also consults
// instrFull for `\h` (force hex parse) and `\u` (force unicode interpretation
// — same as no switch since we already store runes as code points).
func parseSymbolCodePointWithSwitches(arg, instrFull string) (rune, bool) {
	arg = strings.TrimSpace(arg)
	if arg == "" {
		return 0, false
	}
	// `\h` forces hex parse on a bare digit string (no 0x prefix).
	if hexSymbolSwitch(instrFull) && !(strings.HasPrefix(arg, "0x") || strings.HasPrefix(arg, "0X")) {
		n, err := strconv.ParseInt(arg, 16, 32)
		if err != nil || n <= 0 {
			return 0, false
		}
		r := rune(n)
		if !utf8.ValidRune(r) {
			return 0, false
		}
		return r, true
	}
	return parseSymbolCodePoint(arg)
}

// mergeDataLookup does a case-insensitive lookup on a string map.
func mergeDataLookup(m map[string]string, key string) (string, bool) {
	if v, ok := m[key]; ok {
		return v, true
	}
	keyLow := strings.ToLower(key)
	for k, v := range m {
		if strings.ToLower(k) == keyLow {
			return v, true
		}
	}
	return "", false
}

// mergeFieldAffixes parses \b "prefix" and \f "suffix" from a MERGEFIELD
// instrText. These are added around the value ONLY when the value is
// non-empty (Word's "If field is not empty" rule).
func mergeFieldAffixes(instrFull string) (prefix, suffix string) {
	prefix = readQuotedSwitch(instrFull, `\b`)
	suffix = readQuotedSwitch(instrFull, `\f`)
	return
}

func readQuotedSwitch(instrFull, tag string) string {
	i := strings.Index(instrFull, tag)
	if i < 0 {
		return ""
	}
	rest := strings.TrimLeft(instrFull[i+len(tag):], " \t")
	if strings.HasPrefix(rest, `"`) {
		if end := strings.Index(rest[1:], `"`); end >= 0 {
			return rest[1 : 1+end]
		}
		return rest[1:]
	}
	// Unquoted: take to next whitespace.
	for j, c := range rest {
		if c == ' ' || c == '\t' || c == '\\' {
			return rest[:j]
		}
	}
	return rest
}

// seqResetSwitch parses ` SEQ Figure \r 4 ` and returns the reset value.
func seqResetSwitch(instrFull string) (int, bool) {
	parts := strings.Fields(instrFull)
	for i := 0; i < len(parts)-1; i++ {
		if parts[i] == `\r` {
			if n, err := strconv.Atoi(parts[i+1]); err == nil {
				return n, true
			}
		}
	}
	return 0, false
}

// seqHasFlag reports whether a SEQ field carries a no-argument switch
// like \h or \c. The check is case-sensitive (Word writes lowercase).
func seqHasFlag(instrFull, flag string) bool {
	return hasFlagSwitch(instrFull, flag)
}

// hasFlagSwitch is the shared no-arg switch detector used by SEQ / REF /
// PAGEREF / NOTEREF. Word writes switches as lowercase tokens (\h \p \n)
// preceded by whitespace; we compare token-by-token so substrings inside
// quoted picture switches (e.g. `\@ "h:mm"`) don't false-match.
func hasFlagSwitch(instrFull, flag string) bool {
	target := `\` + flag
	for _, p := range strings.Fields(instrFull) {
		if p == target {
			return true
		}
	}
	return false
}

// collectStyleParagraphs indexes the document's body paragraphs by their
// w:pStyle ID so STYLEREF can surface "the first paragraph with style X".
func collectStyleParagraphs(doc *docx.Document) map[string][]string {
	out := map[string][]string{}
	var walk func(blocks []docx.Block)
	walk = func(blocks []docx.Block) {
		for _, b := range blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				if v.StyleID == "" {
					continue
				}
				if txt := paragraphPlainText(v); txt != "" {
					out[v.StyleID] = append(out[v.StyleID], txt)
				}
			case docx.Table:
				for _, row := range v.Rows {
					for _, cell := range row.Cells {
						walk(cell.Blocks)
					}
				}
			}
		}
	}
	if len(doc.Sections) > 0 {
		for _, sec := range doc.Sections {
			walk(sec.Blocks)
		}
	} else {
		walk(doc.Body)
	}
	return out
}

// listNumStart returns the explicit start value from a LISTNUM field's
// \s switch when present.
func listNumStart(instrFull string) (int, bool) {
	parts := strings.Fields(instrFull)
	for i := 0; i < len(parts)-1; i++ {
		if parts[i] == "\\s" {
			if n, err := strconv.Atoi(parts[i+1]); err == nil {
				return n, true
			}
		}
	}
	return 0, false
}

// buildDocProperties merges the standard core/app props with any custom
// properties from docProps/custom.xml.
func buildDocProperties(doc *docx.Document) map[string]string {
	out := map[string]string{
		"Title":   doc.Properties.Title,
		"Author":  doc.Properties.Author,
		"Subject": doc.Properties.Subject,
		"Company": doc.Properties.Company,
		"Pages":   strconv.Itoa(doc.Properties.Pages),
		"Words":   strconv.Itoa(doc.Properties.Words),
		"Lines":   strconv.Itoa(doc.Properties.Lines),
	}
	for k, v := range doc.CustomProperties {
		if v != "" {
			out[k] = v
		}
	}
	return out
}

// collectHeadings flattens the document body into a list of {level, text}
// entries for TOC synthesis.
func collectHeadings(doc *docx.Document) []tocEntry {
	var out []tocEntry
	var walk func(blocks []docx.Block)
	walk = func(blocks []docx.Block) {
		for _, b := range blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				lvl := headingLevel(v, doc)
				if lvl > 0 {
					txt := paragraphPlainText(v)
					if txt != "" {
						out = append(out, tocEntry{Level: lvl, Text: txt})
					}
				}
			case docx.Table:
				for _, row := range v.Rows {
					for _, cell := range row.Cells {
						walk(cell.Blocks)
					}
				}
			}
		}
	}
	if len(doc.Sections) > 0 {
		for _, sec := range doc.Sections {
			walk(sec.Blocks)
		}
	} else {
		walk(doc.Body)
	}
	return out
}

// collectTCEntries walks the body looking for TC field instruction runs
// and parses each into a tocEntry. Done as a pre-pass so the TOC field —
// which usually appears near the start of the doc — can include marks
// defined later.
func collectTCEntries(doc *docx.Document) []tocEntry {
	var out []tocEntry
	var walk func(blocks []docx.Block)
	walk = func(blocks []docx.Block) {
		for _, b := range blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				for _, r := range v.Runs {
					if r.InstrText == "" {
						continue
					}
					if entry, ok := parseTCInstr(r.InstrText); ok {
						out = append(out, entry)
					}
				}
			case docx.Table:
				for _, row := range v.Rows {
					for _, cell := range row.Cells {
						walk(cell.Blocks)
					}
				}
			}
		}
	}
	if len(doc.Sections) > 0 {
		for _, sec := range doc.Sections {
			walk(sec.Blocks)
		}
	} else {
		walk(doc.Body)
	}
	return out
}

// collectXEEntries gathers XE field titles for INDEX synthesis.
func collectXEEntries(doc *docx.Document) []string {
	var out []string
	var walk func(blocks []docx.Block)
	walk = func(blocks []docx.Block) {
		for _, b := range blocks {
			switch v := b.(type) {
			case docx.Paragraph:
				for _, r := range v.Runs {
					if r.InstrText == "" {
						continue
					}
					if title := parseXEInstr(r.InstrText); title != "" {
						out = append(out, title)
					}
				}
			case docx.Table:
				for _, row := range v.Rows {
					for _, cell := range row.Cells {
						walk(cell.Blocks)
					}
				}
			}
		}
	}
	if len(doc.Sections) > 0 {
		for _, sec := range doc.Sections {
			walk(sec.Blocks)
		}
	} else {
		walk(doc.Body)
	}
	return out
}

// headingLevel returns 1..9 if p is a heading paragraph.
func headingLevel(p docx.Paragraph, doc *docx.Document) int {
	_ = doc
	if p.OutlineLvl >= 1 && p.OutlineLvl <= 9 {
		return p.OutlineLvl
	}
	if p.StyleID != "" {
		id := strings.ToLower(p.StyleID)
		if id == "title" {
			return 1
		}
		if strings.HasPrefix(id, "heading") {
			tail := strings.TrimPrefix(id, "heading")
			if n, err := strconv.Atoi(tail); err == nil && n >= 1 && n <= 9 {
				return n
			}
		}
	}
	return 0
}

// paragraphPlainText collapses runs into a single string for TOC entries.
func paragraphPlainText(p docx.Paragraph) string {
	var b strings.Builder
	for _, r := range p.Runs {
		if r.FieldBegin || r.FieldSep || r.FieldEnd || r.InstrText != "" {
			continue
		}
		if r.IsBreak {
			b.WriteByte(' ')
			continue
		}
		b.WriteString(r.Text)
	}
	return strings.TrimSpace(b.String())
}

// formatTOC renders a multi-line TOC from the heading list. Each line is
// indented by the heading level and padded with a dot-leader filler so
// the output reads like a real Word TOC even before Word re-cooks it.
// When vars.bookmarkPages knows where each heading landed we suffix the
// page number; otherwise the trailing column is omitted.
func formatTOC(entries []tocEntry) string {
	const lineWidth = 60
	var b strings.Builder
	for i, e := range entries {
		if i > 0 {
			b.WriteByte('\n')
		}
		depth := e.Level - 1
		if depth < 0 {
			depth = 0
		}
		indent := ""
		for j := 0; j < depth; j++ {
			indent += "  "
		}
		title := strings.TrimSpace(e.Text)
		body := indent + title
		// Pad with dot leaders to lineWidth columns. Each character is one
		// "column" — purely visual; the actual layout is up to the renderer
		// but this gives a recognizable TOC look in plain text dumps and
		// reflows tolerably when re-wrapped at narrower widths.
		if r := lineWidth - len(body); r > 4 {
			b.WriteString(body)
			b.WriteByte(' ')
			for k := 0; k < r-2; k++ {
				b.WriteByte('.')
			}
			b.WriteByte(' ')
		} else {
			b.WriteString(body)
		}
	}
	return b.String()
}

// formatIndex synthesizes a simple alphabetical index from XE entries.
// Duplicates collapse; "Major:Minor" entries indent the minor part under
// the major heading.
func formatIndex(entries []string) string {
	if len(entries) == 0 {
		return ""
	}
	type indexLine struct {
		major string
		minor []string
	}
	// Stable de-dup then alphabetise.
	seen := map[string]map[string]bool{}
	majorOrder := []string{}
	for _, raw := range entries {
		major, minor, _ := strings.Cut(raw, ":")
		major = strings.TrimSpace(major)
		minor = strings.TrimSpace(minor)
		if major == "" {
			continue
		}
		if _, ok := seen[major]; !ok {
			seen[major] = map[string]bool{}
			majorOrder = append(majorOrder, major)
		}
		if minor != "" && !seen[major][minor] {
			seen[major][minor] = true
		}
	}
	// Sort majors alphabetically (Go-style ascending bytes — good enough).
	for i := 1; i < len(majorOrder); i++ {
		for j := i; j > 0 && majorOrder[j] < majorOrder[j-1]; j-- {
			majorOrder[j], majorOrder[j-1] = majorOrder[j-1], majorOrder[j]
		}
	}
	var b strings.Builder
	for i, m := range majorOrder {
		if i > 0 {
			b.WriteByte('\n')
		}
		b.WriteString(m)
		// Sort minors for stability.
		mins := make([]string, 0, len(seen[m]))
		for k := range seen[m] {
			mins = append(mins, k)
		}
		for i := 1; i < len(mins); i++ {
			for j := i; j > 0 && mins[j] < mins[j-1]; j-- {
				mins[j], mins[j-1] = mins[j-1], mins[j]
			}
		}
		for _, mn := range mins {
			b.WriteString("\n  ")
			b.WriteString(mn)
		}
	}
	return b.String()
}

// formatCitation produces an APA-style "(Author, Year)" string.
func formatCitation(s docx.BibSource) string {
	author := ""
	if len(s.Authors) > 0 {
		author = s.Authors[0]
	}
	switch {
	case author != "" && s.Year != "":
		return "(" + author + ", " + s.Year + ")"
	case author != "":
		return "(" + author + ")"
	case s.Year != "":
		return "(" + s.Year + ")"
	case s.Title != "":
		return "(" + s.Title + ")"
	}
	return "(" + s.Tag + ")"
}

// formatBibliography emits a newline-joined list of full entries.
func formatBibliography(sources map[string]docx.BibSource) string {
	tags := make([]string, 0, len(sources))
	for t := range sources {
		tags = append(tags, t)
	}
	for i := 1; i < len(tags); i++ {
		for j := i; j > 0 && tags[j] < tags[j-1]; j-- {
			tags[j], tags[j-1] = tags[j-1], tags[j]
		}
	}
	var b strings.Builder
	for i, tag := range tags {
		if i > 0 {
			b.WriteByte('\n')
		}
		b.WriteString(formatBibEntry(sources[tag]))
	}
	return b.String()
}

func formatBibEntry(s docx.BibSource) string {
	var b strings.Builder
	if len(s.Authors) > 0 {
		b.WriteString(strings.Join(s.Authors, ", "))
	}
	if s.Year != "" {
		if b.Len() > 0 {
			b.WriteByte(' ')
		}
		b.WriteByte('(')
		b.WriteString(s.Year)
		b.WriteByte(')')
	}
	if s.Title != "" {
		if b.Len() > 0 {
			b.WriteString(". ")
		}
		b.WriteString(s.Title)
	}
	if s.JournalName != "" {
		b.WriteString(". ")
		b.WriteString(s.JournalName)
	}
	if s.Publisher != "" {
		b.WriteString(". ")
		b.WriteString(s.Publisher)
	}
	if s.City != "" {
		b.WriteString(", ")
		b.WriteString(s.City)
	}
	if s.Pages != "" {
		b.WriteString(", ")
		b.WriteString(s.Pages)
	}
	if s.URL != "" {
		b.WriteString(". ")
		b.WriteString(s.URL)
	}
	b.WriteByte('.')
	return b.String()
}

// formFieldOutput returns the synthetic glyph for a legacy form field.
// FORMCHECKBOX → ☒/☐; FORMDROPDOWN → currently-selected choice;
// FORMTEXT → ffData.Default (or "" if nothing to show).
func formFieldOutput(ff *docx.FormFieldInfo, code string) (string, bool) {
	if ff == nil {
		return "", false
	}
	kind := ff.Kind
	if kind == "" {
		// Infer from the field code when ffData didn't say.
		switch code {
		case "FORMCHECKBOX":
			kind = "checkbox"
		case "FORMDROPDOWN":
			kind = "dropdown"
		case "FORMTEXT":
			kind = "text"
		}
	}
	switch kind {
	case "checkbox":
		if ff.Checked {
			return "☒", true
		}
		return "☐", true
	case "dropdown":
		if ff.Selected >= 0 && ff.Selected < len(ff.Choices) {
			return ff.Choices[ff.Selected], true
		}
		if len(ff.Choices) > 0 {
			return ff.Choices[0], true
		}
		return "▾", true
	case "text":
		if ff.Default != "" {
			return ff.Default, true
		}
	}
	return "", false
}

// formFieldKindCode derives the field code from a FormFieldInfo when
// the instrText didn't supply one (some FORMFIELDs ship with empty
// instrText and just the ffData blob).
func formFieldKindCode(ff *docx.FormFieldInfo) (string, string) {
	if ff == nil {
		return "", ""
	}
	switch ff.Kind {
	case "checkbox":
		return ff.Kind, "FORMCHECKBOX"
	case "dropdown":
		return ff.Kind, "FORMDROPDOWN"
	case "text":
		return ff.Kind, "FORMTEXT"
	}
	return "", ""
}

// formatFieldDateTime applies a `\@ "format"` switch to t. When no switch
// is present, fallback is used as a sensible default. Supported tokens:
// yyyy, yy, MMMM, MMM, MM, M, dddd, ddd, dd, d, HH, H, hh, h, mm, m, ss,
// s, AM/PM, am/pm.
func formatFieldDateTime(t time.Time, instrFull, fallback string) string {
	layout := fieldDateLayoutSwitch(instrFull)
	if layout == "" {
		return t.Format(fallback)
	}
	return applyWordDateLayout(t, layout)
}

// fieldDateLayoutSwitch extracts the quoted body of a `\@ "format"`
// switch. Returns "" when no such switch is present.
func fieldDateLayoutSwitch(instrFull string) string {
	i := strings.Index(instrFull, `\@`)
	if i < 0 {
		return ""
	}
	rest := instrFull[i+2:]
	rest = strings.TrimLeft(rest, " \t")
	if !strings.HasPrefix(rest, `"`) {
		// Unquoted form: `\@ yyyy/MM/dd` until end-of-string or next `\`.
		if j := strings.Index(rest, " \\"); j >= 0 {
			return strings.TrimSpace(rest[:j])
		}
		return strings.TrimSpace(rest)
	}
	end := strings.Index(rest[1:], `"`)
	if end < 0 {
		return rest[1:]
	}
	return rest[1 : 1+end]
}

// applyWordDateLayout converts a Word format string ("yyyy/MM/dd h:mm")
// into the corresponding rendered time. We process longer tokens first
// so "MMMM" doesn't get matched as four "M"s. Literal tokens (slashes,
// colons, the words "AM"/"PM") pass through.
func applyWordDateLayout(t time.Time, layout string) string {
	type repl struct {
		tok string
		val string
	}
	year, month, day := t.Date()
	hour, minute, second := t.Clock()
	weekday := t.Weekday()
	twoDigit := func(n int) string {
		if n < 10 {
			return "0" + strconv.Itoa(n)
		}
		return strconv.Itoa(n)
	}
	hour12 := hour % 12
	if hour12 == 0 {
		hour12 = 12
	}
	monthLong := []string{"", "January", "February", "March", "April", "May", "June",
		"July", "August", "September", "October", "November", "December"}
	monthShort := []string{"", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
		"Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	dayLong := []string{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
	dayShort := []string{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"}
	tokens := []repl{
		{"yyyy", strconv.Itoa(year)},
		{"yy", twoDigit(year % 100)},
		{"MMMM", monthLong[int(month)]},
		{"MMM", monthShort[int(month)]},
		{"MM", twoDigit(int(month))},
		{"M", strconv.Itoa(int(month))},
		{"dddd", dayLong[weekday]},
		{"ddd", dayShort[weekday]},
		{"dd", twoDigit(day)},
		{"d", strconv.Itoa(day)},
		{"HH", twoDigit(hour)},
		{"H", strconv.Itoa(hour)},
		{"hh", twoDigit(hour12)},
		{"h", strconv.Itoa(hour12)},
		{"mm", twoDigit(minute)},
		{"ss", twoDigit(second)},
		{"s", strconv.Itoa(second)},
		{"AM/PM", func() string {
			if hour < 12 {
				return "AM"
			}
			return "PM"
		}()},
		{"am/pm", func() string {
			if hour < 12 {
				return "am"
			}
			return "pm"
		}()},
		{"tt", func() string {
			if hour < 12 {
				return "am"
			}
			return "pm"
		}()},
	}
	// We need to consume tokens left-to-right with longest-first matching,
	// so a single sweep with prioritized comparison.
	var b strings.Builder
	for i := 0; i < len(layout); {
		matched := false
		for _, tk := range tokens {
			if strings.HasPrefix(layout[i:], tk.tok) {
				b.WriteString(tk.val)
				i += len(tk.tok)
				matched = true
				break
			}
		}
		if !matched {
			b.WriteByte(layout[i])
			i++
		}
	}
	return b.String()
}

// formatNumericValue applies a `\# "format"` switch to v. Returns the
// decimal string when no switch is present.
func formatNumericValue(v float64, instrFull string) string {
	if s := formatNumericSwitch(v, instrFull); s != "" {
		return s
	}
	if v == float64(int64(v)) {
		return strconv.FormatInt(int64(v), 10)
	}
	return strconv.FormatFloat(v, 'f', -1, 64)
}

// formatNumericSwitch implements Word's `\#` numeric picture-format
// switch. Recognized format chars: '0' = digit required, '#' = digit
// optional, '.' = decimal separator, ',' = thousands separator,
// '%' = percent (value gets multiplied by 100 before formatting).
// Any other characters before / after the numeric block (or between
// thousands and decimal) are kept as literal prefix / suffix so
// currency symbols like '$' and unit suffixes pass through.
//
// A semicolon splits the picture into positive ; negative sub-formats
// (e.g. `0.00;(0.00)` shows negative values in parens).
//
// Returns "" when no `\#` switch is present.
func formatNumericSwitch(v float64, instrFull string) string {
	i := strings.Index(instrFull, `\#`)
	if i < 0 {
		return ""
	}
	rest := instrFull[i+2:]
	rest = strings.TrimLeft(rest, " \t")
	picture := ""
	if strings.HasPrefix(rest, `"`) {
		end := strings.Index(rest[1:], `"`)
		if end < 0 {
			picture = rest[1:]
		} else {
			picture = rest[1 : 1+end]
		}
	} else {
		if j := strings.Index(rest, " \\"); j >= 0 {
			picture = strings.TrimSpace(rest[:j])
		} else {
			picture = strings.TrimSpace(rest)
		}
	}
	if picture == "" {
		return ""
	}
	// Word allows a single-quote escape around literal chars
	// ("\# '$'#,##0"). Strip the quotes — the contents stay literal
	// in the picture.
	picture = strings.ReplaceAll(picture, "'", "")
	posPic, negPic, hasNeg := strings.Cut(picture, ";")
	chosen := posPic
	negative := v < 0
	abs := v
	if negative {
		abs = -v
		if hasNeg && negPic != "" {
			chosen = negPic
			// The negative format already encodes the sign, so suppress
			// the implicit leading minus when applying it.
			return applyNumericPicture(abs, chosen, false)
		}
	}
	out := applyNumericPicture(abs, chosen, negative)
	return out
}

// applyNumericPicture renders v into picture, treating non-format runes
// as literal text. addMinus prepends a '-' to the numeric block when
// the caller hasn't already supplied a negative sub-format.
func applyNumericPicture(v float64, picture string, addMinus bool) string {
	if picture == "" {
		return ""
	}
	if strings.Contains(picture, "%") {
		v *= 100
	}
	// Find the numeric block (first run of [0#.,]).
	start := -1
	end := -1
	for i := 0; i < len(picture); i++ {
		c := picture[i]
		if c == '0' || c == '#' || c == '.' || c == ',' {
			if start < 0 {
				start = i
			}
			end = i + 1
		} else if start >= 0 {
			// Allow commas/decimal already covered; any other char
			// after the block closes it.
			break
		}
	}
	if start < 0 {
		// No format chars at all — return the picture unchanged.
		return picture
	}
	prefix := picture[:start]
	suffix := picture[end:]
	numPic := picture[start:end]

	intPart, fracPart, hasFrac := strings.Cut(numPic, ".")
	intDigitsNeeded := strings.Count(intPart, "0")
	fracDigits := strings.Count(fracPart, "0") + strings.Count(fracPart, "#")
	if fracDigits > 0 {
		mul := 1.0
		for i := 0; i < fracDigits; i++ {
			mul *= 10
		}
		v = float64(int64(v*mul+0.5)) / mul
	} else {
		v = float64(int64(v + 0.5))
	}
	intVal := int64(v)
	intStr := strconv.FormatInt(intVal, 10)
	for len(intStr) < intDigitsNeeded {
		intStr = "0" + intStr
	}
	if strings.Contains(intPart, ",") {
		var b strings.Builder
		n := len(intStr)
		for i, c := range intStr {
			if i > 0 && (n-i)%3 == 0 {
				b.WriteByte(',')
			}
			b.WriteRune(c)
		}
		intStr = b.String()
	}
	numStr := intStr
	if hasFrac && fracDigits > 0 {
		fracStr := ""
		fracVal := v - float64(intVal)
		for i := 0; i < fracDigits; i++ {
			fracVal *= 10
			d := int(fracVal)
			if d > 9 {
				d = 9
			}
			fracStr += strconv.Itoa(d)
			fracVal -= float64(d)
		}
		for i := len(fracStr) - 1; i >= 0; i-- {
			if i >= len(fracPart) {
				break
			}
			if fracPart[i] == '#' && fracStr[i] == '0' {
				fracStr = fracStr[:i]
				continue
			}
			break
		}
		if fracStr != "" {
			numStr += "." + fracStr
		}
	}
	if addMinus {
		numStr = "-" + numStr
	}
	return prefix + numStr + suffix
}

// initialsOf extracts a 2-3 letter initials string from a full name.
// "Alice Wonder Land" → "AWL". Falls back to the whole name if it has
// no spaces.
func initialsOf(name string) string {
	parts := strings.Fields(name)
	if len(parts) == 0 {
		return ""
	}
	var b strings.Builder
	for _, p := range parts {
		if p == "" {
			continue
		}
		r := []rune(p)
		b.WriteRune(r[0])
	}
	return strings.ToUpper(b.String())
}

// tokenizeFieldArgs splits a field's argument list honoring double-quoted
// strings. "a b" c → ["a b", "c"]. Switches (\…) and their operands stay
// separate; the caller can filter them. Whitespace inside quotes is
// preserved verbatim.
func tokenizeFieldArgs(s string) []string {
	var out []string
	var cur strings.Builder
	inQ := false
	flush := func() {
		if cur.Len() == 0 {
			return
		}
		out = append(out, cur.String())
		cur.Reset()
	}
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c == '"' {
			inQ = !inQ
			continue
		}
		if !inQ && (c == ' ' || c == '\t') {
			flush()
			continue
		}
		cur.WriteByte(c)
	}
	flush()
	return out
}

// evaluateIfField parses and evaluates a Word IF field instruction:
//
//	IF <e1> <op> <e2> "true" "false"
//
// op ∈ {=, <>, !=, <, >, <=, >=}. The operands are quoted strings, numbers,
// or unquoted identifiers (treated as case-insensitive strings). Returns
// the chosen branch text + ok=true on a successful evaluation; ok=false
// when the instruction can't be parsed (caller falls back to cached
// result).
func evaluateIfField(instrFull string) (string, bool) {
	s := strings.TrimSpace(instrFull)
	upper := strings.ToUpper(s)
	if !strings.HasPrefix(upper, "IF") {
		return "", false
	}
	s = strings.TrimSpace(s[2:])
	toks := tokenizeFieldArgs(s)
	if len(toks) < 5 {
		return "", false
	}
	left, op, right := toks[0], toks[1], toks[2]
	truePart := strings.Trim(toks[3], `"`)
	falsePart := strings.Trim(toks[4], `"`)
	pass := ifCompare(left, op, right)
	if pass {
		return truePart, true
	}
	return falsePart, true
}

func ifCompare(left, op, right string) bool {
	// Try numeric comparison when both sides parse as numbers.
	lf, lok := strconv.ParseFloat(strings.Trim(left, `"`), 64)
	rf, rok := strconv.ParseFloat(strings.Trim(right, `"`), 64)
	if lok == nil && rok == nil {
		switch op {
		case "=":
			return lf == rf
		case "<>", "!=":
			return lf != rf
		case "<":
			return lf < rf
		case ">":
			return lf > rf
		case "<=":
			return lf <= rf
		case ">=":
			return lf >= rf
		}
	}
	// Fall back to string compare (case-insensitive — matches Word).
	l := strings.ToLower(strings.Trim(left, `"`))
	r := strings.ToLower(strings.Trim(right, `"`))
	switch op {
	case "=":
		return l == r
	case "<>", "!=":
		return l != r
	case "<":
		return l < r
	case ">":
		return l > r
	case "<=":
		return l <= r
	case ">=":
		return l >= r
	}
	return false
}
