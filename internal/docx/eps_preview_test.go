package docx

import "testing"

func TestExtractEPSPreviewBytes_TooShort(t *testing.T) {
	img, ok := extractEPSPreviewBytes([]byte{0xC5, 0xD0, 0xD3})
	if ok {
		t.Error("expected false for short data")
	}
	if img != nil {
		t.Error("expected nil for short data")
	}
}

func TestExtractEPSPreviewBytes_WrongMagic(t *testing.T) {
	data := make([]byte, 30)
	data[0] = 0x00
	data[1] = 0x00
	data[2] = 0x00
	data[3] = 0x00
	img, ok := extractEPSPreviewBytes(data)
	if ok {
		t.Error("expected false for wrong magic")
	}
	if img != nil {
		t.Error("expected nil for wrong magic")
	}
}

func TestExtractEPSPreviewBytes_TIFFOffsetOutOfBounds(t *testing.T) {
	data := make([]byte, 30)
	data[0] = 0xC5
	data[1] = 0xD0
	data[2] = 0xD3
	data[3] = 0xC6
	data[20] = 0xFF
	data[21] = 0xFF
	data[22] = 0xFF
	data[23] = 0xFF
	data[24] = 0x01
	img, ok := extractEPSPreviewBytes(data)
	if ok {
		t.Error("expected false for out-of-bounds TIFF")
	}
	if img != nil {
		t.Error("expected nil for out-of-bounds TIFF")
	}
}
