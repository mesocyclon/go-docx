package image

import (
	"bytes"
	"encoding/binary"
	"errors"
	"testing"
)

func TestStreamReader_ReadByte(t *testing.T) {
	data := []byte{0x00, 0x42, 0xFF}
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	b, err := sr.ReadByteAt(0, 1)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if b != 0x42 {
		t.Errorf("ReadByte(0,1) = %#x, want 0x42", b)
	}
}

func TestStreamReader_ReadShort_BigEndian(t *testing.T) {
	data := make([]byte, 4)
	binary.BigEndian.PutUint16(data[2:], 0x1234)
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	v, err := sr.ReadShort(2, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != 0x1234 {
		t.Errorf("ReadShort big-endian = %#x, want 0x1234", v)
	}
}

func TestStreamReader_ReadShort_LittleEndian(t *testing.T) {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint16(data[0:], 0xABCD)
	sr := NewStreamReader(bytes.NewReader(data), false, 0)

	v, err := sr.ReadShort(0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != 0xABCD {
		t.Errorf("ReadShort little-endian = %#x, want 0xABCD", v)
	}
}

func TestStreamReader_ReadLong_BigEndian(t *testing.T) {
	data := make([]byte, 8)
	binary.BigEndian.PutUint32(data[4:], 0xDEADBEEF)
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	v, err := sr.ReadLong(4, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != 0xDEADBEEF {
		t.Errorf("ReadLong big-endian = %#x, want 0xDEADBEEF", v)
	}
}

func TestStreamReader_ReadLong_LittleEndian(t *testing.T) {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint32(data[0:], 0x12345678)
	sr := NewStreamReader(bytes.NewReader(data), false, 0)

	v, err := sr.ReadLong(0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != 0x12345678 {
		t.Errorf("ReadLong little-endian = %#x, want 0x12345678", v)
	}
}

func TestStreamReader_ReadStr(t *testing.T) {
	data := []byte("Hello, World!")
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	s, err := sr.ReadStr(5, 0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if s != "Hello" {
		t.Errorf("ReadStr = %q, want %q", s, "Hello")
	}
}

func TestStreamReader_ReadStr_UTF8(t *testing.T) {
	data := []byte("Héllo")
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	s, err := sr.ReadStr(len(data), 0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if s != "Héllo" {
		t.Errorf("ReadStr UTF-8 = %q, want %q", s, "Héllo")
	}
}

func TestStreamReader_ReadPastEOF(t *testing.T) {
	data := []byte{0x01, 0x02}
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	_, err := sr.ReadLong(0, 0) // needs 4 bytes, only 2 available
	if err == nil {
		t.Fatal("expected error for read past EOF")
	}
	if !errors.Is(err, ErrUnexpectedEOF) {
		t.Errorf("expected ErrUnexpectedEOF, got %v", err)
	}
}

func TestStreamReader_BaseOffset(t *testing.T) {
	data := []byte{0x00, 0x00, 0x00, 0x42, 0xFF}
	sr := NewStreamReader(bytes.NewReader(data), true, 2) // base offset 2

	b, err := sr.ReadByteAt(1, 0) // reads at position 2+1+0 = 3
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if b != 0x42 {
		t.Errorf("ReadByte with base offset = %#x, want 0x42", b)
	}
}

func TestStreamReader_SetByteOrder(t *testing.T) {
	data := make([]byte, 4)
	binary.BigEndian.PutUint16(data[0:], 0x0102)
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	v, _ := sr.ReadShort(0, 0)
	if v != 0x0102 {
		t.Errorf("big-endian ReadShort = %#x, want 0x0102", v)
	}

	sr.SetByteOrder(false) // switch to little-endian
	v, _ = sr.ReadShort(0, 0)
	if v != 0x0201 {
		t.Errorf("little-endian ReadShort = %#x, want 0x0201", v)
	}
}

func TestStreamReader_Tell(t *testing.T) {
	data := []byte{0x01, 0x02, 0x03}
	sr := NewStreamReader(bytes.NewReader(data), true, 0)

	sr.ReadByteAt(2, 0)
	pos, err := sr.Tell()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if pos != 3 { // after reading byte at position 2, cursor is at 3
		t.Errorf("Tell = %d, want 3", pos)
	}
}
