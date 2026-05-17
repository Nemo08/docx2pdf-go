package docx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"path"
	"strconv"
	"strings"
)

const pi180 = math.Pi / 180.0

func cosF(rad float64) float64 { return math.Cos(rad) }
func sinF(rad float64) float64 { return math.Sin(rad) }

// resolveRelTarget resolves a rel target against the source part's
// directory. Targets are expressed relative to the rels file's
// containing part, so word/_rels/document.xml.rels uses "word/" as the
// base. "../customXml/item1.xml" → "customXml/item1.xml".
func resolveRelTarget(base, target string) string {
	target = strings.TrimPrefix(target, "/")
	if strings.HasPrefix(target, "../") || strings.Contains(target, "/../") {
		return path.Clean(path.Join(base, target))
	}
	return base + target
}

// parseDocVars reads w:docVars/w:docVar entries from word/settings.xml into
// the destination map. Called from inside parseSettings's element loop — it
// receives the StartElement and consumes up to the matching EndElement.
func decodeDocVars(dec *xml.Decoder, start xml.StartElement, out map[string]string) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "docVar" {
				name := attr(t, "name")
				val := attr(t, "val")
				if name != "" {
					out[name] = val
				}
				_ = dec.Skip()
			} else {
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
}

// parseCustomProperties reads docProps/custom.xml. Each <property name="X">
// element holds a typed child (vt:lpwstr / vt:i4 / vt:bool / vt:filetime ...)
// whose text content is the value.
func parseCustomProperties(f *zip.File, out map[string]string) error {
	rc, err := openZipFile(f)
	if err != nil {
		return err
	}
	defer rc.Close()
	dec := xml.NewDecoder(rc)
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if se.Name.Local != "property" {
			continue
		}
		name := attr(se, "name")
		// The value lives one element deeper, as text inside a vt:* node.
		var val string
		depth := 1
		for depth > 0 {
			tok, err := dec.Token()
			if err != nil {
				return err
			}
			switch t := tok.(type) {
			case xml.StartElement:
				depth++
				// Accept the first vt:* descendant as the value source.
				txt, e := readElementText(dec, t)
				if e != nil {
					return e
				}
				if val == "" {
					val = strings.TrimSpace(txt)
				}
				depth-- // readElementText consumed the EndElement
			case xml.EndElement:
				depth--
			}
		}
		if name != "" {
			out[name] = val
		}
	}
}

// extractAltChunkText pulls plain text out of an AlternativeFormatInputPart
// (HTML / XHTML / plain text / RTF). We don't try to lay out the foreign
// markup — we strip tags and return paragraph-broken text.
func extractAltChunkText(f *zip.File) (string, error) {
	rc, err := openZipFile(f)
	if err != nil {
		return "", err
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return "", err
	}
	s := string(data)
	low := strings.ToLower(strings.TrimSpace(s))
	switch {
	case strings.HasPrefix(low, "{\\rtf"):
		return stripRTF(s), nil
	case strings.HasPrefix(low, "<!doctype html") ||
		strings.HasPrefix(low, "<html") ||
		strings.HasPrefix(low, "<body") ||
		strings.Contains(low, "<p>") ||
		strings.Contains(low, "<div"):
		return stripHTML(s), nil
	}
	return s, nil
}

// stripHTML produces a paragraph-broken plain-text version of an HTML
// fragment. Block-level closes become \n; whitespace collapses; entities
// resolve to the common five.
func stripHTML(s string) string {
	var b strings.Builder
	b.Grow(len(s))
	inTag := false
	inComment := false
	lastWasSpace := true // suppress leading whitespace
	emit := func(r rune) {
		if r == ' ' || r == '\t' || r == '\r' || r == '\n' {
			if !lastWasSpace {
				b.WriteRune(' ')
				lastWasSpace = true
			}
			return
		}
		b.WriteRune(r)
		lastWasSpace = false
	}
	emitBreak := func() {
		// Collapse a run of breaks to one newline boundary.
		if b.Len() == 0 {
			return
		}
		out := b.String()
		if strings.HasSuffix(out, "\n") {
			return
		}
		b.WriteByte('\n')
		lastWasSpace = true
	}
	for i := 0; i < len(s); i++ {
		c := s[i]
		if inComment {
			if c == '-' && i+2 < len(s) && s[i+1] == '-' && s[i+2] == '>' {
				inComment = false
				i += 2
			}
			continue
		}
		if !inTag && c == '<' && i+3 < len(s) && s[i+1] == '!' && s[i+2] == '-' && s[i+3] == '-' {
			inComment = true
			i += 3
			continue
		}
		if c == '<' {
			// Look ahead to spot block-closing tags so we can insert breaks.
			end := strings.IndexByte(s[i:], '>')
			if end < 0 {
				break
			}
			tag := strings.ToLower(strings.TrimSpace(s[i+1 : i+end]))
			tag = strings.TrimPrefix(tag, "/")
			cut := strings.IndexAny(tag, " \t\r\n")
			if cut > 0 {
				tag = tag[:cut]
			}
			switch tag {
			case "p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
				"blockquote", "section", "article", "header", "footer", "pre",
				"table", "thead", "tbody", "tfoot", "ul", "ol":
				emitBreak()
			}
			i += end
			inTag = false
			continue
		}
		if c == '&' {
			semi := strings.IndexByte(s[i:], ';')
			if semi > 0 && semi < 12 {
				name := s[i+1 : i+semi]
				if r, ok := htmlEntity(name); ok {
					emit(r)
					i += semi
					continue
				}
			}
			emit(rune(c))
			continue
		}
		emit(rune(c))
	}
	return strings.TrimSpace(b.String())
}

func htmlEntity(name string) (rune, bool) {
	switch strings.ToLower(name) {
	case "amp":
		return '&', true
	case "lt":
		return '<', true
	case "gt":
		return '>', true
	case "quot":
		return '"', true
	case "apos":
		return '\'', true
	case "nbsp":
		return ' ', true
	}
	if strings.HasPrefix(name, "#") {
		// Numeric character reference. We don't decode every form; bail.
		return 0, false
	}
	return 0, false
}

// stripRTF returns the text content of an RTF document. Best-effort:
// it removes control words ({\xxx ...}), unescapes \\ \{ \} and \'hh hex
// escapes, and joins everything that's left. Lossy for tables, fonts,
// colors, and most formatting — but the prose survives.
func stripRTF(s string) string {
	var b strings.Builder
	b.Grow(len(s) / 2)
	depth := 0
	skipGroup := -1 // ignore nested groups deeper than this when in \fonttbl etc.
	i := 0
	for i < len(s) {
		c := s[i]
		switch c {
		case '{':
			depth++
			i++
		case '}':
			if depth == skipGroup {
				skipGroup = -1
			}
			if depth > 0 {
				depth--
			}
			i++
		case '\\':
			// Control word or escape.
			if i+1 < len(s) {
				n := s[i+1]
				if n == '\\' || n == '{' || n == '}' {
					b.WriteByte(n)
					i += 2
					continue
				}
				if n == '\'' && i+3 < len(s) {
					// Hex escape \'hh
					var hexv byte
					ok := true
					for k := 0; k < 2; k++ {
						hc := s[i+2+k]
						var v byte
						switch {
						case hc >= '0' && hc <= '9':
							v = hc - '0'
						case hc >= 'a' && hc <= 'f':
							v = hc - 'a' + 10
						case hc >= 'A' && hc <= 'F':
							v = hc - 'A' + 10
						default:
							ok = false
						}
						hexv = hexv*16 + v
					}
					if ok {
						b.WriteByte(hexv)
					}
					i += 4
					continue
				}
			}
			// Control word: consume word chars + optional numeric arg + one delimiter.
			j := i + 1
			for j < len(s) && ((s[j] >= 'a' && s[j] <= 'z') || (s[j] >= 'A' && s[j] <= 'Z')) {
				j++
			}
			word := s[i+1 : j]
			// Optional numeric argument.
			for j < len(s) && (s[j] == '-' || (s[j] >= '0' && s[j] <= '9')) {
				j++
			}
			// One optional space delimiter is consumed as part of the word.
			if j < len(s) && s[j] == ' ' {
				j++
			}
			// Word-emitting destinations: \par / \line / \tab translate to whitespace.
			switch word {
			case "par", "line", "pard", "sect":
				if b.Len() > 0 && !strings.HasSuffix(b.String(), "\n") {
					b.WriteByte('\n')
				}
			case "tab":
				b.WriteByte('\t')
			case "fonttbl", "colortbl", "stylesheet", "info", "pict", "header", "footer":
				// Suppress contents of this group.
				skipGroup = depth
			}
			i = j
		default:
			if skipGroup > 0 && depth >= skipGroup {
				i++
				continue
			}
			if c == '\r' || c == '\n' {
				i++
				continue
			}
			b.WriteByte(c)
			i++
		}
	}
	return strings.TrimSpace(b.String())
}

// loadCustomXMLParts walks the rels map for entries pointing at
// customXml/itemN.xml and reads their bytes into doc.CustomXMLRoots. Also
// detects the bibliography namespace and populates doc.Bibliography.
func loadCustomXMLParts(rels map[string]relEntry, files map[string]*zip.File, doc *Document) {
	for _, e := range rels {
		if !isCustomXMLRel(e.Type) {
			continue
		}
		full := resolveRelTarget("word/", e.Target)
		zf, ok := files[full]
		if !ok {
			zf, ok = files[strings.TrimPrefix(e.Target, "/")]
		}
		if !ok {
			continue
		}
		rc, err := openZipFile(zf)
		if err != nil {
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			continue
		}
		// Companion itemPropsN.xml carries the {GUID} that SDTs use to
		// pick a specific data store via w:storeItemID. Path layout is
		// the standard pair: word/customXml/item1.xml ↔ word/customXml/itemProps1.xml.
		storeGuid := readStoreItemGUID(files, full)
		doc.CustomXMLRoots = append(doc.CustomXMLRoots, CustomXMLPart{
			PartName:    full,
			Data:        data,
			StoreItemID: storeGuid,
		})
		// Inspect for bibliography namespace.
		if strings.Contains(string(data), "/officeDocument/2006/bibliography") {
			parseBibliography(data, doc)
		}
		// Inspect for OpenDoPE xpaths namespace.
		if strings.Contains(string(data), "opendope/xpaths") {
			if table := parseOpenDoPEXPaths(data); len(table) > 0 {
				if doc.OpenDoPEXPaths == nil {
					doc.OpenDoPEXPaths = map[string]string{}
				}
				for k, v := range table {
					doc.OpenDoPEXPaths[k] = v
				}
			}
		}
	}
}

func isCustomXMLRel(t string) bool {
	return strings.HasSuffix(t, "/customXml") ||
		strings.HasSuffix(t, "/customXmlProps") ||
		strings.Contains(t, "customXml")
}

// parseBibliography walks a customXml store for <b:Source> entries and
// stores them by tag.
func parseBibliography(data []byte, doc *Document) {
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if se.Name.Local != "Source" {
			continue
		}
		src := BibSource{}
		// Walk the source's children.
		depth := 1
		var curField string
		var inAuthorName bool
		for depth > 0 {
			tok, err := dec.Token()
			if err != nil {
				return
			}
			switch t := tok.(type) {
			case xml.StartElement:
				depth++
				curField = t.Name.Local
				if curField == "Author" || curField == "Editor" {
					inAuthorName = false
				}
				if curField == "NameList" || curField == "Person" {
					inAuthorName = true
				}
			case xml.CharData:
				v := strings.TrimSpace(string(t))
				if v == "" {
					continue
				}
				switch curField {
				case "Tag":
					src.Tag = v
				case "SourceType":
					src.SourceType = v
				case "Title":
					src.Title = v
				case "Year":
					src.Year = v
				case "Publisher":
					src.Publisher = v
				case "City":
					src.City = v
				case "JournalName":
					src.JournalName = v
				case "Pages":
					src.Pages = v
				case "URL":
					src.URL = v
				case "Last", "First", "Middle":
					if inAuthorName {
						if len(src.Authors) == 0 || curField == "Last" {
							src.Authors = append(src.Authors, v)
						} else {
							// Append given name to most recent surname.
							src.Authors[len(src.Authors)-1] = v + " " + src.Authors[len(src.Authors)-1]
						}
					}
				}
			case xml.EndElement:
				depth--
				if t.Name.Local == "Person" || t.Name.Local == "NameList" {
					inAuthorName = false
				}
			}
		}
		if src.Tag != "" {
			if doc.Bibliography == nil {
				doc.Bibliography = map[string]BibSource{}
			}
			doc.Bibliography[src.Tag] = src
		}
	}
}

// isAltChunkRel matches the AlternativeFormatInputPart relationship type.
func isAltChunkRel(t string) bool {
	return strings.HasSuffix(t, "/aFChunk") ||
		strings.HasSuffix(t, "/altChunk") ||
		strings.HasSuffix(t, "/afChunk")
}

// loadAltChunks reads every AlternativeFormatInputPart referenced in rels
// and parses its content into a Block tree stored under doc.AltChunks keyed
// by rId. HTML content is parsed by parseHTMLAltChunk so bold/italic/
// headings/lists/links survive into the body; RTF and plain text fall back
// to flat paragraphs.
func loadAltChunks(rels map[string]relEntry, files map[string]*zip.File, doc *Document) {
	for rid, e := range rels {
		if !isAltChunkRel(e.Type) {
			continue
		}
		full := resolveRelTarget("word/", e.Target)
		zf, ok := files[full]
		if !ok {
			zf, ok = files[strings.TrimPrefix(e.Target, "/")]
		}
		if !ok {
			continue
		}
		rc, err := openZipFile(zf)
		if err != nil {
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			continue
		}
		s := string(data)
		low := strings.ToLower(strings.TrimSpace(s))
		var blocks []Block
		switch {
		case strings.HasPrefix(low, "{\\rtf"):
			// RTF: strip to plain text — we don't model RTF semantics here.
			txt := stripRTF(s)
			blocks = flatTextToParagraphs(txt, doc.Defaults)
		case strings.HasPrefix(low, "<!doctype html") ||
			strings.HasPrefix(low, "<html") ||
			strings.HasPrefix(low, "<body") ||
			strings.Contains(low, "<p>") ||
			strings.Contains(low, "<p ") ||
			strings.Contains(low, "<div") ||
			strings.Contains(low, "<h1") ||
			strings.Contains(low, "<h2") ||
			strings.Contains(low, "<ul") ||
			strings.Contains(low, "<ol"):
			blocks = parseHTMLAltChunk(s, doc.Defaults)
		default:
			blocks = flatTextToParagraphs(s, doc.Defaults)
		}
		if len(blocks) == 0 {
			continue
		}
		if doc.AltChunks == nil {
			doc.AltChunks = map[string][]Block{}
		}
		doc.AltChunks[rid] = blocks
	}
}

// flatTextToParagraphs turns a newline-delimited string into Paragraph
// blocks, one per non-empty line.
func flatTextToParagraphs(s string, defaults RunProps) []Block {
	var blocks []Block
	for _, line := range strings.Split(s, "\n") {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		blocks = append(blocks, Paragraph{
			Runs: []Run{{Text: line, Props: defaults}},
		})
	}
	return blocks
}

// resolveXPath does a very small subset of XPath against a custom-xml store.
// Supports:
//   - "/ns:Root/ns:Foo/ns:Bar"             element-suffix match → first text
//   - "/ns:Root/ns:Foo/@attr"              attribute selector → that attr's value
//   - "/ns:Root/ns:Foo[2]"                 positional predicate (1-based)
//   - "/ns:Root/ns:Foo[@k='v']/ns:Bar"     attribute-equality predicate
//
// Namespace prefixes are stripped throughout — the resolver is name-only
// because Word's stored XPaths often reference custom prefixes that aren't
// declared in the same scope as the data. Returns the first matching text.
func resolveXPath(parts []CustomXMLPart, xpath string) (string, bool) {
	xpath = strings.TrimSpace(xpath)
	if xpath == "" {
		return "", false
	}
	// Attribute selector at the tail of the path: "/foo/bar/@attr".
	attrSel := ""
	if i := strings.LastIndex(xpath, "/@"); i >= 0 {
		attrSel = xpath[i+2:]
		if j := strings.IndexAny(attrSel, "[/"); j >= 0 {
			attrSel = attrSel[:j]
		}
		xpath = xpath[:i]
	}
	if attrSel != "" {
		if j := strings.IndexByte(attrSel, ':'); j >= 0 {
			attrSel = attrSel[j+1:]
		}
	}
	rawSteps := strings.Split(strings.TrimPrefix(xpath, "/"), "/")
	steps := make([]xpathStep, 0, len(rawSteps))
	for _, s := range rawSteps {
		st := parseXPathStep(s)
		if st.name == "" {
			continue
		}
		steps = append(steps, st)
	}
	if len(steps) == 0 && attrSel == "" {
		return "", false
	}
	for _, part := range parts {
		if v, ok := walkXMLForPath(part.Data, steps, attrSel); ok {
			return v, true
		}
	}
	return "", false
}

// xpathStep is one parsed segment of an XPath: tag name plus an optional
// predicate. We currently honor either a positional "[N]" predicate or a
// single attribute-equality "[@a='v']" predicate; other forms are ignored.
type xpathStep struct {
	name string
	// position > 0 selects the Nth match (1-based). 0 means "any".
	position int
	// attrName/attrVal are non-empty for attribute-equality predicates.
	attrName, attrVal string
}

func parseXPathStep(s string) xpathStep {
	var step xpathStep
	if i := strings.IndexByte(s, ':'); i >= 0 {
		s = s[i+1:]
	}
	if i := strings.IndexByte(s, '['); i >= 0 {
		pred := s[i+1:]
		s = s[:i]
		pred = strings.TrimSuffix(pred, "]")
		if n, err := strconv.Atoi(pred); err == nil && n > 0 {
			step.position = n
		} else if strings.HasPrefix(pred, "@") {
			pred = strings.TrimPrefix(pred, "@")
			if eq := strings.IndexByte(pred, '='); eq >= 0 {
				step.attrName = strings.TrimSpace(pred[:eq])
				v := strings.TrimSpace(pred[eq+1:])
				v = strings.Trim(v, `'"`)
				step.attrVal = v
			}
		}
	}
	step.name = s
	return step
}

// xpathFrame is one element on the walker's stack.
type xpathFrame struct {
	name      string
	attrs     map[string]string
	childPos  map[string]int // how many times each child tag has been seen
}

// resolveXPathInStore is the store-scoped variant of resolveXPath: when
// storeItemID is non-empty, only the matching part is searched; when
// empty, every part is searched (legacy behavior). Returns the value
// and a found flag.
func resolveXPathInStore(parts []CustomXMLPart, storeItemID, xpath string) (string, bool) {
	if storeItemID == "" {
		return resolveXPath(parts, xpath)
	}
	guid := strings.Trim(storeItemID, "{}")
	for _, p := range parts {
		pg := strings.Trim(p.StoreItemID, "{}")
		if !strings.EqualFold(pg, guid) {
			continue
		}
		if v, ok := resolveXPath([]CustomXMLPart{p}, xpath); ok {
			return v, true
		}
		return "", false
	}
	// Fallback: GUID not found among loaded stores (older docs sometimes
	// drop the itemProps file). Search every store as a recovery.
	return resolveXPath(parts, xpath)
}

// readStoreItemGUID locates customXml/itemPropsN.xml that pairs with
// itemDataPath and pulls the ds:itemID GUID from its
// <ds:datastoreItem> root. Returns "" when the props file is absent.
func readStoreItemGUID(files map[string]*zip.File, itemDataPath string) string {
	propsPath := storePropsCompanion(itemDataPath)
	zf, ok := files[propsPath]
	if !ok {
		return ""
	}
	rc, err := openZipFile(zf)
	if err != nil {
		return ""
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return ""
	}
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	for {
		tok, err := dec.Token()
		if err != nil {
			return ""
		}
		if se, ok := tok.(xml.StartElement); ok {
			if se.Name.Local == "datastoreItem" {
				for _, a := range se.Attr {
					if a.Name.Local == "itemID" {
						return strings.TrimSpace(a.Value)
					}
				}
				return ""
			}
		}
	}
}

// storePropsCompanion turns "word/customXml/item3.xml" into
// "word/customXml/itemProps3.xml". For non-conforming names it returns
// the original path unchanged (so the lookup harmlessly misses).
func storePropsCompanion(itemPath string) string {
	dir := itemPath
	base := ""
	if i := strings.LastIndex(itemPath, "/"); i >= 0 {
		dir = itemPath[:i+1]
		base = itemPath[i+1:]
	}
	if !strings.HasPrefix(base, "item") {
		return itemPath
	}
	return dir + "itemProps" + strings.TrimPrefix(base, "item")
}

// countXPathMatches returns how many element-matches an XPath has across
// all custom XML stores. Used by the OpenDoPE repeat resolver so it
// doesn't have to probe positional predicates one-at-a-time. Matches
// are counted by *element start* — a path that ends with /@attr is
// counted by attribute presence on each match.
func countXPathMatches(parts []CustomXMLPart, xpath string) int {
	xpath = strings.TrimSpace(xpath)
	if xpath == "" {
		return 0
	}
	attrSel := ""
	if i := strings.LastIndex(xpath, "/@"); i >= 0 {
		attrSel = xpath[i+2:]
		if j := strings.IndexAny(attrSel, "[/"); j >= 0 {
			attrSel = attrSel[:j]
		}
		xpath = xpath[:i]
	}
	if attrSel != "" {
		if j := strings.IndexByte(attrSel, ':'); j >= 0 {
			attrSel = attrSel[j+1:]
		}
	}
	rawSteps := strings.Split(strings.TrimPrefix(xpath, "/"), "/")
	steps := make([]xpathStep, 0, len(rawSteps))
	for _, s := range rawSteps {
		st := parseXPathStep(s)
		if st.name == "" {
			continue
		}
		steps = append(steps, st)
	}
	if len(steps) == 0 {
		return 0
	}
	total := 0
	for _, part := range parts {
		total += countMatchesInPart(part.Data, steps, attrSel)
	}
	return total
}

func countMatchesInPart(data []byte, steps []xpathStep, attrSel string) int {
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	var stack []xpathFrame
	count := 0
	for {
		tok, err := dec.Token()
		if err != nil {
			return count
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if len(stack) > 0 {
				if stack[len(stack)-1].childPos == nil {
					stack[len(stack)-1].childPos = map[string]int{}
				}
				stack[len(stack)-1].childPos[t.Name.Local]++
			}
			attrs := map[string]string{}
			for _, a := range t.Attr {
				attrs[a.Name.Local] = a.Value
			}
			stack = append(stack, xpathFrame{name: t.Name.Local, attrs: attrs})
			if matchSuffixWithPredicates(stack, steps) {
				if attrSel != "" {
					if _, ok := attrs[attrSel]; ok {
						count++
					}
				} else {
					count++
				}
			}
		case xml.EndElement:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
}

// walkXMLForPath streams through data and returns the first text node that
// satisfies a suffix match of steps. When attrSel is non-empty the walker
// instead returns the named attribute's value on the deepest matching
// element.
func walkXMLForPath(data []byte, steps []xpathStep, attrSel string) (string, bool) {
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	var stack []xpathFrame
	for {
		tok, err := dec.Token()
		if err != nil {
			return "", false
		}
		switch t := tok.(type) {
		case xml.StartElement:
			// Increment the parent's child-count for this tag (used by
			// positional predicates on the upcoming frame).
			if len(stack) > 0 {
				if stack[len(stack)-1].childPos == nil {
					stack[len(stack)-1].childPos = map[string]int{}
				}
				stack[len(stack)-1].childPos[t.Name.Local]++
			}
			attrs := map[string]string{}
			for _, a := range t.Attr {
				attrs[a.Name.Local] = a.Value
			}
			stack = append(stack, xpathFrame{name: t.Name.Local, attrs: attrs})
			if attrSel != "" && matchSuffixWithPredicates(stack, steps) {
				if v, ok := attrs[attrSel]; ok {
					return strings.TrimSpace(v), true
				}
			}
		case xml.EndElement:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if attrSel != "" {
				continue
			}
			if matchSuffixWithPredicates(stack, steps) {
				v := strings.TrimSpace(string(t))
				if v != "" {
					return v, true
				}
			}
		}
	}
}

// matchSuffixWithPredicates returns true when the last len(steps) stack
// frames satisfy their corresponding steps' tag-name and optional predicate.
func matchSuffixWithPredicates(stack []xpathFrame, steps []xpathStep) bool {
	if len(stack) < len(steps) {
		return false
	}
	off := len(stack) - len(steps)
	for i, st := range steps {
		f := stack[off+i]
		if st.name != "" && st.name != f.name {
			return false
		}
		if st.position > 0 {
			// Position checks against the parent frame's child-count
			// for this tag at the moment this frame was created. We
			// approximate by reading the live child-count which equals
			// the cumulative number of f.name children opened so far —
			// matching st.position when this is the Nth.
			var parentCount int
			if off+i-1 >= 0 {
				parentCount = stack[off+i-1].childPos[f.name]
			}
			if parentCount != st.position {
				return false
			}
		}
		if st.attrName != "" && f.attrs[st.attrName] != st.attrVal {
			return false
		}
	}
	return true
}


// parseGradFill parses a DrawingML <a:gradFill> element into a list of
// color stops, the gradient angle (in degrees), and a kind ("linear" or
// "radial"). The XML schema is:
//
//	<a:gradFill ...>
//	  <a:gsLst>
//	    <a:gs pos="0">    <a:srgbClr val="…"/> </a:gs>
//	    <a:gs pos="100000"><a:srgbClr val="…"/> </a:gs>
//	  </a:gsLst>
//	  <a:lin ang="5400000" />   <!-- linear angle in 60000ths of a degree -->
//	  OR
//	  <a:path path="circle"/>    <!-- radial -->
//	</a:gradFill>
func parseGradFill(dec *xml.Decoder, start xml.StartElement) (stops []GradientStop, angleDeg float64, kind string, err error) {
	kind = "linear"
	for {
		tok, e := dec.Token()
		if e != nil {
			return stops, angleDeg, kind, e
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "gsLst":
				if err := parseGsLst(dec, t, &stops); err != nil {
					return stops, angleDeg, kind, err
				}
			case "lin":
				kind = "linear"
				if v := attr(t, "ang"); v != "" {
					if a, e := strconv.ParseFloat(v, 64); e == nil {
						angleDeg = a / 60000.0
					}
				}
				_ = dec.Skip()
			case "path":
				kind = "radial"
				_ = dec.Skip()
			default:
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return stops, angleDeg, kind, nil
			}
		}
	}
}

func parseGsLst(dec *xml.Decoder, start xml.StartElement, stops *[]GradientStop) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "gs" {
				stop := GradientStop{Alpha: 1}
				if v := attr(t, "pos"); v != "" {
					if x, err := strconv.ParseFloat(v, 64); err == nil {
						stop.Pos = x / 100000.0
					}
				}
				if c := scanSolidFillColor(dec, t); c != "" {
					stop.Color = c
				}
				*stops = append(*stops, stop)
			} else {
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
}

// extractChartStruct parses a chart part (word/charts/chartN.xml) into a
// structured ChartData ready for the renderer. The parser only recognizes
// the chart families we can paint (bar / column / pie / doughnut / line /
// scatter); other chart types return an empty Kind so the caller can fall
// back to the existing text-extraction path.
func extractChartStruct(f *zip.File) (ChartData, error) {
	var out ChartData
	rc, err := openZipFile(f)
	if err != nil {
		return out, err
	}
	defer rc.Close()
	dec := xml.NewDecoder(rc)
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return out, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "title":
			out.Title = extractChartTitle(dec, se)
		case "barChart":
			out.Kind = "column"
			for _, a := range se.Attr {
				_ = a
			}
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
			// barChart sub-discriminator: <c:barDir val="bar"/> ↔ horizontal.
			// Our column kind is the vertical default; horizontal is recorded
			// at parseChartTypeBody time when it sees barDir.
		case "lineChart":
			out.Kind = "line"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "pieChart", "pie3DChart":
			out.Kind = "pie"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "doughnutChart":
			out.Kind = "doughnut"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "scatterChart":
			out.Kind = "scatter"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "areaChart", "area3DChart":
			out.Kind = "area"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "bubbleChart":
			out.Kind = "bubble"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "radarChart":
			out.Kind = "radar"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "stockChart":
			// Stock chart is rendered as a line chart for the closing series.
			out.Kind = "line"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		case "surfaceChart", "surface3DChart":
			// Surface charts have no good 2D analogue; fall through to line.
			out.Kind = "line"
			if err := parseChartTypeBody(dec, se, &out); err != nil {
				return out, err
			}
		}
	}
}

// extractChartTitle pulls the visible text out of a c:title element.
func extractChartTitle(dec *xml.Decoder, start xml.StartElement) string {
	var sb strings.Builder
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return strings.TrimSpace(sb.String())
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		case xml.CharData:
			s := strings.TrimSpace(string(t))
			if s != "" {
				if sb.Len() > 0 {
					sb.WriteByte(' ')
				}
				sb.WriteString(s)
			}
		}
	}
	return strings.TrimSpace(sb.String())
}

// parseChartTypeBody walks the children of a chart type element, picking
// up each c:ser sub-tree. Also detects barDir="bar" so the renderer can
// distinguish horizontal bars from vertical columns.
func parseChartTypeBody(dec *xml.Decoder, start xml.StartElement, out *ChartData) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "barDir":
				if attr(t, "val") == "bar" && out.Kind == "column" {
					out.Kind = "bar"
				}
				_ = dec.Skip()
			case "grouping":
				// c:grouping discriminates clustered (default) / stacked /
				// percentStacked for bar+column, and standard / stacked /
				// percentStacked for area+line. The renderer reads this to
				// pick between side-by-side and stacked bar layouts.
				if v := attr(t, "val"); v != "" && out.Grouping == "" {
					out.Grouping = v
				}
				_ = dec.Skip()
			case "ser":
				ser, cats, err := parseChartSeries(dec, t)
				if err != nil {
					return err
				}
				if ser.Name != "" || len(ser.Values) > 0 {
					out.Series = append(out.Series, ser)
				}
				if len(cats) > len(out.Categories) {
					out.Categories = cats
				}
			default:
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
}

// parseChartSeries reads one c:ser element. Returns the series plus any
// category labels it carried (categories live on the series in OOXML —
// they typically repeat identically across series, so the caller picks
// the longest list).
func parseChartSeries(dec *xml.Decoder, start xml.StartElement) (ChartSeries, []string, error) {
	var ser ChartSeries
	var cats []string
	for {
		tok, err := dec.Token()
		if err != nil {
			return ser, cats, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tx":
				ser.Name = extractFirstText(dec, t)
			case "cat", "xVal":
				cats = parseChartRefStrings(dec, t)
			case "val", "yVal":
				ser.Values = parseChartRefNumbers(dec, t)
			case "spPr":
				ser.Color = parseFirstSolidFill(dec, t)
			default:
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return ser, cats, nil
			}
		}
	}
}

// extractFirstText returns the first non-empty CharData inside the
// subtree. Used for series name + axis label captures.
func extractFirstText(dec *xml.Decoder, start xml.StartElement) string {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return ""
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		case xml.CharData:
			s := strings.TrimSpace(string(t))
			if s != "" {
				// Drain remaining tokens so the caller's element loop
				// sees the matching EndElement at the correct depth.
				for depth > 0 {
					tk, e := dec.Token()
					if e != nil {
						return s
					}
					switch tk.(type) {
					case xml.StartElement:
						depth++
					case xml.EndElement:
						depth--
					}
				}
				return s
			}
		}
	}
	return ""
}

// parseChartRefStrings collects the visible text of every <c:pt><c:v>
// child anywhere under start. Categories arrive in document order; idx
// attributes are not honored beyond what the source already provides.
func parseChartRefStrings(dec *xml.Decoder, start xml.StartElement) []string {
	var out []string
	depth := 1
	inV := false
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return out
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "v" {
				inV = true
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "v" {
				inV = false
			}
		case xml.CharData:
			if inV {
				s := strings.TrimSpace(string(t))
				if s != "" {
					out = append(out, s)
				}
			}
		}
	}
	return out
}

// parseChartRefNumbers is the numeric twin of parseChartRefStrings. Non-
// numeric entries are silently dropped (Word writes "#N/A" for missing
// data points; we treat those as zero so the chart still renders).
func parseChartRefNumbers(dec *xml.Decoder, start xml.StartElement) []float64 {
	var out []float64
	depth := 1
	inV := false
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return out
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "v" {
				inV = true
			}
		case xml.EndElement:
			depth--
			if t.Name.Local == "v" {
				inV = false
			}
		case xml.CharData:
			if inV {
				s := strings.TrimSpace(string(t))
				if s == "" {
					continue
				}
				if v, err := strconv.ParseFloat(s, 64); err == nil {
					out = append(out, v)
				} else {
					out = append(out, 0)
				}
			}
		}
	}
	return out
}

// parseFirstSolidFill returns the first <a:srgbClr val="…"/> color found
// anywhere in the subtree. Used for series color discovery.
func parseFirstSolidFill(dec *xml.Decoder, start xml.StartElement) string {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return ""
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "srgbClr" {
				v := attr(t, "val")
				if v != "" {
					// Drain remaining tokens at correct depth before return.
					for depth > 0 {
						tk, e := dec.Token()
						if e != nil {
							return v
						}
						switch tk.(type) {
						case xml.StartElement:
							depth++
						case xml.EndElement:
							depth--
						}
					}
					return strings.ToUpper(v)
				}
			}
		case xml.EndElement:
			depth--
		}
	}
	return ""
}

// parsePattFill parses a DrawingML <a:pattFill> element and returns an
// approximated solid color: the per-channel average of <a:fgClr> and
// <a:bgClr>. We don't render the actual pattern tile; the average gives
// a sensible mid-tone the eye reads as the pattern's overall shade.
//
// XML schema:
//
//	<a:pattFill prst="…">
//	  <a:fgClr><a:srgbClr val="…"/></a:fgClr>
//	  <a:bgClr><a:srgbClr val="…"/></a:bgClr>
//	</a:pattFill>
func parsePattFill(dec *xml.Decoder, start xml.StartElement) (string, error) {
	var fg, bg string
	for {
		tok, err := dec.Token()
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "fgClr":
				if c := scanSolidFillColor(dec, t); c != "" {
					fg = c
				}
			case "bgClr":
				if c := scanSolidFillColor(dec, t); c != "" {
					bg = c
				}
			default:
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				switch {
				case fg != "" && bg != "":
					return averageHexColor(fg, bg), nil
				case fg != "":
					return fg, nil
				case bg != "":
					return bg, nil
				}
				return "", nil
			}
		}
	}
}

// averageHexColor returns the per-channel arithmetic mean of two 6-hex
// colors. Inputs are tolerated with or without a leading '#'. Returns
// "" if either side fails to parse.
func averageHexColor(a, b string) string {
	a = strings.TrimPrefix(a, "#")
	b = strings.TrimPrefix(b, "#")
	if len(a) != 6 || len(b) != 6 {
		return ""
	}
	parse := func(s string) (int, int, int, bool) {
		x, err := strconv.ParseUint(s, 16, 32)
		if err != nil {
			return 0, 0, 0, false
		}
		return int(x>>16) & 0xff, int(x>>8) & 0xff, int(x) & 0xff, true
	}
	ar, ag, ab, ok1 := parse(a)
	br, bg, bb, ok2 := parse(b)
	if !ok1 || !ok2 {
		return ""
	}
	return fmt.Sprintf("%02X%02X%02X", (ar+br)/2, (ag+bg)/2, (ab+bb)/2)
}

// parseEffectList scans <a:effectLst> for the first outer-shadow effect
// and returns its parameters. Inner-shadow, glow, reflection, and other
// effects are ignored.
func parseEffectList(dec *xml.Decoder, start xml.StartElement) (*ShadowEffect, error) {
	var out *ShadowEffect
	for {
		tok, err := dec.Token()
		if err != nil {
			return out, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "outerShdw" && out == nil {
				eff := &ShadowEffect{Alpha: 1, Color: "000000"}
				// blurRad in EMU
				if v := attr(t, "blurRad"); v != "" {
					if x, err := strconv.ParseFloat(v, 64); err == nil {
						eff.BlurPt = x / emuPerPt
					}
				}
				// dist + dir: distance in EMU, direction in 60000ths-of-a-degree.
				dist := 0.0
				dirDeg := 0.0
				if v := attr(t, "dist"); v != "" {
					if x, err := strconv.ParseFloat(v, 64); err == nil {
						dist = x / emuPerPt
					}
				}
				if v := attr(t, "dir"); v != "" {
					if x, err := strconv.ParseFloat(v, 64); err == nil {
						dirDeg = x / 60000.0
					}
				}
				rad := dirDeg * pi180
				eff.OffsetXPt = dist * cosF(rad)
				eff.OffsetYPt = dist * sinF(rad)
				if c := scanSolidFillColor(dec, t); c != "" {
					eff.Color = c
				}
				out = eff
			} else {
				_ = dec.Skip()
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return out, nil
			}
		}
	}
}
