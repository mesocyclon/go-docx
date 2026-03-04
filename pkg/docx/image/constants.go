package image

// MIME content types for images.
const (
	MimeBMP  = "image/bmp"
	MimeGIF  = "image/gif"
	MimeJPEG = "image/jpeg"
	MimePNG  = "image/png"
	MimeTIFF = "image/tiff"
)

// JPEG marker codes.
const (
	markerTEM  = 0x01
	markerSOF0 = 0xC0
	markerSOF1 = 0xC1
	markerSOF2 = 0xC2
	markerSOF3 = 0xC3
	markerDHT  = 0xC4
	markerSOF5 = 0xC5
	markerSOF6 = 0xC6
	markerSOF7 = 0xC7
	markerJPG  = 0xC8
	markerSOF9 = 0xC9
	markerSOFA = 0xCA
	markerSOFB = 0xCB
	markerDAC  = 0xCC
	markerSOFD = 0xCD
	markerSOFE = 0xCE
	markerSOFF = 0xCF

	markerRST0 = 0xD0
	markerRST1 = 0xD1
	markerRST2 = 0xD2
	markerRST3 = 0xD3
	markerRST4 = 0xD4
	markerRST5 = 0xD5
	markerRST6 = 0xD6
	markerRST7 = 0xD7

	markerSOI = 0xD8
	markerEOI = 0xD9
	markerSOS = 0xDA
	markerDQT = 0xDB
	markerDNL = 0xDC
	markerDRI = 0xDD
	markerDHP = 0xDE
	markerEXP = 0xDF

	markerAPP0 = 0xE0
	markerAPP1 = 0xE1
)

// sofMarkerCodes is the set of JPEG Start-Of-Frame marker codes.
var sofMarkerCodes = map[byte]bool{
	markerSOF0: true,
	markerSOF1: true,
	markerSOF2: true,
	markerSOF3: true,
	markerSOF5: true,
	markerSOF6: true,
	markerSOF7: true,
	markerSOF9: true,
	markerSOFA: true,
	markerSOFB: true,
	markerSOFD: true,
	markerSOFE: true,
	markerSOFF: true,
}

// standaloneMarkers is the set of markers that have no following segment.
var standaloneMarkers = map[byte]bool{
	markerTEM:  true,
	markerSOI:  true,
	markerEOI:  true,
	markerRST0: true,
	markerRST1: true,
	markerRST2: true,
	markerRST3: true,
	markerRST4: true,
	markerRST5: true,
	markerRST6: true,
	markerRST7: true,
}

// PNG chunk type names.
const (
	pngChunkIHDR = "IHDR"
	pngChunkPHYs = "pHYs"
	pngChunkIEND = "IEND"
)

// TIFF IFD field types.
const (
	tiffFieldBYTE     = 1
	tiffFieldASCII    = 2
	tiffFieldSHORT    = 3
	tiffFieldLONG     = 4
	tiffFieldRATIONAL = 5
)

// TIFF tag codes.
const (
	tiffTagImageWidth    = 0x0100
	tiffTagImageLength   = 0x0101
	tiffTagXResolution   = 0x011A
	tiffTagYResolution   = 0x011B
	tiffTagResolutionUnit = 0x0128
)
