package docx

import (
	"archive/zip"
	"encoding/hex"
	"encoding/xml"
	"errors"
	"io"
	"strings"
)

// parseFontTable reads word/fontTable.xml + the embedded font parts it
// references, deobfuscates each per ECMA-376 §17.8.1, and populates
// doc.EmbeddedFonts.
func parseFontTable(f *zip.File, fontRels map[string]relEntry, files map[string]*zip.File, doc *Document) error {
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
		if se.Name.Local != "font" {
			continue
		}
		name := attr(se, "name")
		if name == "" {
			_ = dec.Skip()
			continue
		}
		ef := EmbeddedFont{Name: name}
		if err := decodeFontEntry(dec, se, &ef, fontRels, files); err != nil {
			return err
		}
		if ef.Regular == nil && ef.Bold == nil && ef.Italic == nil && ef.BoldItalic == nil {
			continue
		}
		if doc.EmbeddedFonts == nil {
			doc.EmbeddedFonts = map[string]EmbeddedFont{}
		}
		doc.EmbeddedFonts[name] = ef
		if ef.AltName != "" {
			doc.EmbeddedFonts[ef.AltName] = ef
		}
	}
}

func decodeFontEntry(dec *xml.Decoder, start xml.StartElement, ef *EmbeddedFont, fontRels map[string]relEntry, files map[string]*zip.File) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "altName":
				ef.AltName = attr(t, "val")
				_ = dec.Skip()
			case "embedRegular":
				ef.Regular = loadEmbeddedFontPart(t, fontRels, files)
				_ = dec.Skip()
			case "embedBold":
				ef.Bold = loadEmbeddedFontPart(t, fontRels, files)
				_ = dec.Skip()
			case "embedItalic":
				ef.Italic = loadEmbeddedFontPart(t, fontRels, files)
				_ = dec.Skip()
			case "embedBoldItalic":
				ef.BoldItalic = loadEmbeddedFontPart(t, fontRels, files)
				_ = dec.Skip()
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

func loadEmbeddedFontPart(se xml.StartElement, fontRels map[string]relEntry, files map[string]*zip.File) []byte {
	rid := attr(se, "id")
	key := attr(se, "fontKey")
	if rid == "" || key == "" {
		return nil
	}
	rel, ok := fontRels[rid]
	if !ok {
		return nil
	}
	full := resolveRelTarget("word/", rel.Target)
	zf, ok := files[full]
	if !ok {
		return nil
	}
	rc, err := zf.Open()
	if err != nil {
		return nil
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return nil
	}
	out, err := deobfuscateFont(data, key)
	if err != nil {
		return nil
	}
	return out
}

// deobfuscateFont reverses ECMA-376 §17.8.1 obfuscation.
func deobfuscateFont(data []byte, fontKey string) ([]byte, error) {
	key, err := guidToBytes(fontKey)
	if err != nil {
		return nil, err
	}
	out := make([]byte, len(data))
	copy(out, data)
	n := 32
	if len(out) < n {
		n = len(out)
	}
	for i := 0; i < n; i++ {
		out[i] ^= key[15-(i%16)]
	}
	return out, nil
}

func guidToBytes(s string) ([]byte, error) {
	s = strings.TrimSpace(s)
	s = strings.TrimPrefix(s, "{")
	s = strings.TrimSuffix(s, "}")
	parts := strings.Split(s, "-")
	if len(parts) != 5 {
		return nil, errors.New("docx: malformed font GUID")
	}
	expected := []int{8, 4, 4, 4, 12}
	for i, p := range parts {
		if len(p) != expected[i] {
			return nil, errors.New("docx: malformed font GUID component")
		}
	}
	a, err := hex.DecodeString(parts[0])
	if err != nil {
		return nil, err
	}
	b, err := hex.DecodeString(parts[1])
	if err != nil {
		return nil, err
	}
	c, err := hex.DecodeString(parts[2])
	if err != nil {
		return nil, err
	}
	d, err := hex.DecodeString(parts[3])
	if err != nil {
		return nil, err
	}
	e, err := hex.DecodeString(parts[4])
	if err != nil {
		return nil, err
	}
	out := make([]byte, 16)
	for i := 0; i < 4; i++ {
		out[i] = a[3-i]
	}
	for i := 0; i < 2; i++ {
		out[4+i] = b[1-i]
	}
	for i := 0; i < 2; i++ {
		out[6+i] = c[1-i]
	}
	copy(out[8:10], d)
	copy(out[10:16], e)
	return out, nil
}
