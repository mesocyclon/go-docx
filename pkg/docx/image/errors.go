// Package image provides objects that can characterize image streams.
//
// That characterization is as to content type and size, as a required step
// in including them in a document.
package image

import "errors"

var (
	// ErrUnrecognizedImage is returned when the image stream format cannot
	// be identified from its magic bytes.
	ErrUnrecognizedImage = errors.New("image: unrecognized image format")

	// ErrInvalidImageStream is returned when a recognized image stream
	// appears to be corrupted.
	ErrInvalidImageStream = errors.New("image: invalid or corrupted image stream")

	// ErrUnexpectedEOF is returned when EOF is unexpectedly encountered
	// while reading an image stream.
	ErrUnexpectedEOF = errors.New("image: unexpected end of file")
)
