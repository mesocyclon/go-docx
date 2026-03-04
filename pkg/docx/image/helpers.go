package image

import (
	"encoding/binary"
	"fmt"
	"io"
)

// StreamReader wraps an io.ReadSeeker to provide access to structured data
// from a binary file. Byte-order is configurable. baseOffset is added to
// any base value provided to calculate actual location for reads.
//
// Mirrors Python helpers.StreamReader exactly.
type StreamReader struct {
	r          io.ReadSeeker
	byteOrder  binary.ByteOrder
	baseOffset int64
}

// NewStreamReader returns a new StreamReader wrapping r. If bigEndian is true,
// multi-byte integers are read as big-endian; otherwise little-endian.
func NewStreamReader(r io.ReadSeeker, bigEndian bool, baseOffset int64) *StreamReader {
	order := binary.ByteOrder(binary.LittleEndian)
	if bigEndian {
		order = binary.BigEndian
	}
	return &StreamReader{r: r, byteOrder: order, baseOffset: baseOffset}
}

// ReadByteAt returns the uint8 value at baseOffset + base + offset.
func (sr *StreamReader) ReadByteAt(base, offset int64) (byte, error) {
	buf, err := sr.readBytes(1, base, offset)
	if err != nil {
		return 0, err
	}
	return buf[0], nil
}

// ReadShort returns the uint16 value at baseOffset + base + offset,
// interpreted using the configured byte order.
func (sr *StreamReader) ReadShort(base, offset int64) (uint16, error) {
	buf, err := sr.readBytes(2, base, offset)
	if err != nil {
		return 0, err
	}
	return sr.byteOrder.Uint16(buf), nil
}

// ReadLong returns the uint32 value at baseOffset + base + offset,
// interpreted using the configured byte order.
func (sr *StreamReader) ReadLong(base, offset int64) (uint32, error) {
	buf, err := sr.readBytes(4, base, offset)
	if err != nil {
		return 0, err
	}
	return sr.byteOrder.Uint32(buf), nil
}

// ReadStr returns a string of charCount bytes at baseOffset + base + offset.
func (sr *StreamReader) ReadStr(charCount int, base, offset int64) (string, error) {
	buf, err := sr.readBytes(charCount, base, offset)
	if err != nil {
		return "", err
	}
	return string(buf), nil
}

// SeekTo positions the underlying stream to baseOffset + base + offset.
func (sr *StreamReader) SeekTo(base, offset int64) error {
	location := sr.baseOffset + base + offset
	_, err := sr.r.Seek(location, io.SeekStart)
	if err != nil {
		return fmt.Errorf("image: seek to %d: %w", location, err)
	}
	return nil
}

// Read is a pass-through to the underlying stream's Read method.
func (sr *StreamReader) Read(buf []byte) (int, error) {
	return sr.r.Read(buf)
}

// Tell returns the current position in the underlying stream.
func (sr *StreamReader) Tell() (int64, error) {
	return sr.r.Seek(0, io.SeekCurrent)
}

// SetByteOrder changes the byte order used for multi-byte reads.
func (sr *StreamReader) SetByteOrder(bigEndian bool) {
	if bigEndian {
		sr.byteOrder = binary.BigEndian
	} else {
		sr.byteOrder = binary.LittleEndian
	}
}

// readBytes reads count bytes from baseOffset + base + offset.
// Returns ErrUnexpectedEOF if fewer than count bytes are available.
func (sr *StreamReader) readBytes(count int, base, offset int64) ([]byte, error) {
	if err := sr.SeekTo(base, offset); err != nil {
		return nil, err
	}
	buf := make([]byte, count)
	_, err := io.ReadFull(sr.r, buf)
	if err != nil {
		return nil, fmt.Errorf("%w: wanted %d bytes at offset %d",
			ErrUnexpectedEOF, count, sr.baseOffset+base+offset)
	}
	return buf, nil
}
