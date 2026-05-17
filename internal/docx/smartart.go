package docx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

// diagramSiblingDrawing returns the drawing*.xml zip entry that pairs
// with a given data*.xml target path, or nil when no match exists.
// Word writes them with matching numeric suffixes (data1.xml ↔
// drawing1.xml) in the same directory.
func diagramSiblingDrawing(files map[string]*zip.File, dataTarget string) *zip.File {
	// dataTarget is relative to word/ (e.g. "diagrams/data1.xml").
	// Replace "data" with "drawing" in the file segment.
	const dataPrefix = "data"
	const drawingPrefix = "drawing"
	full := "word/" + dataTarget
	slash := strings.LastIndex(full, "/")
	if slash < 0 {
		return nil
	}
	dir, fname := full[:slash+1], full[slash+1:]
	if !strings.HasPrefix(fname, dataPrefix) {
		return nil
	}
	guess := dir + drawingPrefix + fname[len(dataPrefix):]
	if zf, ok := files[guess]; ok {
		return zf
	}
	return nil
}

// extractDiagramDrawing parses word/diagrams/drawingN.xml — the
// pre-rendered DrawingML shape tree Word writes alongside the SmartArt
// data part. We translate each dsp:sp into a VMLShape positioned in the
// group's coordinate space; the renderer projects those into the
// outer bounding rect at paint time.
//
// Returns nil when the part contains no useful shapes — callers should
// fall back to the text-only diagram surface in Document.Diagrams.
func extractDiagramDrawing(f *zip.File) (*VMLShape, error) {
	rc, err := openZipFile(f)
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	dec := xml.NewDecoder(rc)
	var shapes []VMLShape
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if se.Name.Local == "sp" {
			sh, ok := decodeDiagramSp(dec, se)
			if ok {
				shapes = append(shapes, sh)
			}
		}
	}
	if len(shapes) == 0 {
		return nil, nil
	}
	// Compute bounding box so the group's coord size matches the
	// child layout. Without this the renderer would have nothing to
	// project against.
	maxX, maxY := 0.0, 0.0
	for _, sh := range shapes {
		if sh.OffsetXPt+sh.WidthPt > maxX {
			maxX = sh.OffsetXPt + sh.WidthPt
		}
		if sh.OffsetYPt+sh.HeightPt > maxY {
			maxY = sh.OffsetYPt + sh.HeightPt
		}
	}
	if maxX <= 0 || maxY <= 0 {
		return nil, nil
	}
	return &VMLShape{
		Kind:       "group",
		Children:   shapes,
		CoordSizeW: maxX,
		CoordSizeH: maxY,
	}, nil
}

// decodeDiagramSp parses one dsp:sp element. Returns (shape, true) when
// the element produced something we can draw; (_, false) when the entry
// only carried setup data (typically a missing prstGeom).
func decodeDiagramSp(dec *xml.Decoder, start xml.StartElement) (VMLShape, bool) {
	var sh VMLShape
	var hasGeom bool
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return sh, false
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "xfrm":
				ox, oy, cx, cy := decodeDiagramXfrm(dec, t)
				sh.OffsetXPt = ox
				sh.OffsetYPt = oy
				sh.WidthPt = cx
				sh.HeightPt = cy
				depth--
			case "prstGeom":
				prst := attr(t, "prst")
				sh.Kind = shapeKindForPrst(prst)
				hasGeom = true
				_ = dec.Skip()
				depth--
			case "solidFill":
				c := scanSolidFillColor(dec, t)
				if c != "" && sh.FillColor == "" {
					sh.FillColor = c
				}
				depth--
			case "gradFill":
				stops, angle, kind, err := parseGradFill(dec, t)
				if err == nil && len(stops) > 0 {
					sh.GradientKind = kind
					sh.GradientAngle = angle
					sh.GradientStops = stops
				}
				depth--
			case "ln":
				// Stroke width + color. <a:ln w="N"> N is in EMU.
				if v := attr(t, "w"); v != "" {
					if w, err := strconv.ParseInt(v, 10, 64); err == nil {
						sh.StrokeWeightPt = float64(w) / emuPerPt
					}
				}
				if c := scanSolidFillColor(dec, t); c != "" {
					sh.StrokeColor = c
				}
				depth--
			case "txBody":
				txt := extractTxBodyText(dec, t)
				if txt != "" {
					sh.TextBox = txt
				}
				depth--
			default:
				// nothing — fall through to depth handling
			}
		case xml.EndElement:
			depth--
		}
	}
	if !hasGeom && sh.WidthPt <= 0 && sh.HeightPt <= 0 {
		return sh, false
	}
	if !hasGeom {
		sh.Kind = "rect"
	}
	return sh, true
}

// decodeDiagramXfrm reads <a:xfrm> → <a:off x= y=>/<a:ext cx= cy=> and
// returns offset+extent in points. Returns zeros when the element didn't
// contain explicit values.
func decodeDiagramXfrm(dec *xml.Decoder, start xml.StartElement) (offX, offY, cx, cy float64) {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "off":
				if v := attr(t, "x"); v != "" {
					if x, err := strconv.ParseInt(v, 10, 64); err == nil {
						offX = float64(x) / emuPerPt
					}
				}
				if v := attr(t, "y"); v != "" {
					if y, err := strconv.ParseInt(v, 10, 64); err == nil {
						offY = float64(y) / emuPerPt
					}
				}
			case "ext":
				if v := attr(t, "cx"); v != "" {
					if x, err := strconv.ParseInt(v, 10, 64); err == nil {
						cx = float64(x) / emuPerPt
					}
				}
				if v := attr(t, "cy"); v != "" {
					if y, err := strconv.ParseInt(v, 10, 64); err == nil {
						cy = float64(y) / emuPerPt
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}
	return
}

// synthesizeSmartArtLayout builds a horizontal "process" diagram out of a
// flat node-text string when Word didn't pre-render the SmartArt visuals.
// The text comes in the form "Node1 → Node2 → Node3" from
// extractDiagramText; we split it on the arrow, draw one rounded rect per
// node, and connect them with thin arrow lines. Returns nil when the
// text doesn't split into >=2 nodes — single-node diagrams aren't worth
// the synthetic frame and fall back to the placeholder rect.
//
// We deliberately ignore the underlying SmartArt layout type (cycle vs.
// hierarchy vs. matrix) — every preset reduces to a recognizable boxes-
// with-arrows render at this density, and inferring the right layout
// algorithm from data.xml alone is unreliable without colors/style/quick-
// style parts.
func synthesizeSmartArtLayout(text string, widthPt, heightPt float64) *VMLShape {
	if text == "" {
		return nil
	}
	nodes := splitDiagramNodes(text)
	if len(nodes) < 2 {
		return nil
	}
	if widthPt <= 0 {
		widthPt = 480
	}
	if heightPt <= 0 {
		heightPt = 96
	}
	const (
		boxPad     = 8.0  // horizontal gap between boxes (also arrow span)
		minBoxW    = 64.0 // smallest acceptable box width
		strokeColr = "808080"
		fillColr   = "EEEEEE"
		arrowColr  = "606060"
	)
	n := float64(len(nodes))
	// Geometry: boxes share the row, with boxPad between them. Reserve
	// half a box-pad on each side as the outer margin so the leftmost
	// arrow has room to land on the box edge.
	boxW := (widthPt - (n-1)*boxPad) / n
	if boxW < minBoxW {
		// Too many nodes to fit horizontally — clip to fit. Caller can
		// still consume the synthetic group; nodes that overflow will
		// just be drawn at the actual computed width.
		boxW = minBoxW
	}
	boxH := heightPt * 0.7
	if boxH < 24 {
		boxH = 24
	}
	boxY := (heightPt - boxH) / 2
	var children []VMLShape
	for i, name := range nodes {
		x := float64(i) * (boxW + boxPad)
		children = append(children, VMLShape{
			Kind:           "roundrect",
			WidthPt:        boxW,
			HeightPt:       boxH,
			OffsetXPt:      x,
			OffsetYPt:      boxY,
			FillColor:      fillColr,
			StrokeColor:    strokeColr,
			StrokeWeightPt: 0.75,
			CornerArc:      6,
			TextBox:        name,
		})
		if i < len(nodes)-1 {
			// Connector arrow: a 1pt line from this box's right edge to
			// the next box's left edge. We render this as a v:polyline
			// with two points so the existing VML painter draws it.
			arrowFromX := x + boxW
			arrowToX := arrowFromX + boxPad
			arrowY := heightPt / 2
			children = append(children, VMLShape{
				Kind:           "polyline",
				StrokeColor:    arrowColr,
				StrokeWeightPt: 0.75,
				Points:         formatPolyPoints(arrowFromX, arrowY, arrowToX, arrowY),
			})
			// Small arrow-head: short two-segment polyline forming a
			// chevron at the destination end.
			children = append(children, VMLShape{
				Kind:           "polyline",
				StrokeColor:    arrowColr,
				StrokeWeightPt: 0.75,
				Points: formatPolyPoints(
					arrowToX-3, arrowY-3,
					arrowToX, arrowY,
					arrowToX-3, arrowY+3,
				),
			})
		}
	}
	return &VMLShape{
		Kind:       "group",
		WidthPt:    widthPt,
		HeightPt:   heightPt,
		Children:   children,
		CoordSizeW: widthPt,
		CoordSizeH: heightPt,
	}
}

// splitDiagramNodes splits an extractDiagramText output on the " → "
// separator. We also tolerate a comma fallback for diagrams whose flat
// text lacks the arrow (shouldn't happen with our extractor, but be
// defensive against future writers).
func splitDiagramNodes(text string) []string {
	sep := " → "
	if !strings.Contains(text, sep) {
		if strings.Contains(text, ", ") {
			sep = ", "
		} else {
			return []string{text}
		}
	}
	parts := strings.Split(text, sep)
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			out = append(out, p)
		}
	}
	return out
}

// formatPolyPoints formats a flat float list as the "x,y x,y …" string
// shape the VML painter expects for polyline Points.
func formatPolyPoints(coords ...float64) string {
	var b strings.Builder
	for i := 0; i+1 < len(coords); i += 2 {
		if i > 0 {
			b.WriteByte(' ')
		}
		b.WriteString(strconv.FormatFloat(coords[i], 'f', -1, 64))
		b.WriteByte(',')
		b.WriteString(strconv.FormatFloat(coords[i+1], 'f', -1, 64))
	}
	return b.String()
}

// extractTxBodyText concatenates the text inside an <a:txBody> element —
// the same shape Word uses for chart titles and dsp:sp captions. We
// preserve paragraph breaks as single spaces.
func extractTxBodyText(dec *xml.Decoder, start xml.StartElement) string {
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
			if t.Name.Local == "p" && sb.Len() > 0 {
				sb.WriteByte(' ')
			}
		case xml.EndElement:
			depth--
		case xml.CharData:
			s := string(t)
			if s != "" {
				sb.WriteString(s)
			}
		}
	}
	return strings.TrimSpace(sb.String())
}
