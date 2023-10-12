package excel

import (
	"archive/zip"
	"encoding/xml"
	"io"

	convert "github.com/szyhf/go-convert"
)

type SharedStringsReader struct {
	readCloser     io.ReadCloser
	decoder        *xml.Decoder
	index          int
	sharedStrings  []string
	readBytesLimit int64
}

func NewSharedStringsReader(sharedStringsFile *zip.File, readBytesLimit int64) (*SharedStringsReader, error) {
	if sharedStringsFile == nil {
		// No shared strings file, so no shared strings.
		return &SharedStringsReader{
			readCloser:     nil,
			decoder:        nil,
			index:          0,
			sharedStrings:  []string{},
			readBytesLimit: readBytesLimit,
		}, nil
	}

	reader, err := sharedStringsFile.Open()
	if err != nil {
		return nil, err
	}

	decoder := xml.NewDecoder(reader)

	return &SharedStringsReader{
		readCloser:     reader,
		decoder:        decoder,
		index:          0,
		sharedStrings:  []string{},
		readBytesLimit: readBytesLimit,
	}, nil
}

func (ssr *SharedStringsReader) Close() error {
	return ssr.readCloser.Close()
}

func (ssr *SharedStringsReader) GetString(index int) (string, error) {
	if ssr.readCloser == nil || ssr.decoder == nil {
		return "", ErrNoSharedStringsFile
	}

	if index >= ssr.index {
		ssr.readSharedStringsUntil(index)
	}

	if index >= len(ssr.sharedStrings) {
		return "", nil
	}

	return ssr.sharedStrings[index], nil
}

func (ssr *SharedStringsReader) PassedMemLimit() bool {
	return ssr.readBytesLimit != -1 && ssr.decoder.InputOffset() >= ssr.readBytesLimit
}

func (ssr *SharedStringsReader) readSharedStringsUntil(index int) {
	tStart, rStart := false, false

	for t, err := ssr.decoder.Token(); err == nil && !ssr.PassedMemLimit(); t, err = ssr.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case _SI:
				// don't enter default ...
			case _T:
				tStart = true
			case _R:
				rStart = true
			case _SST:
				count := 0
				unqCount := 0
				for _, attr := range token.Attr {
					switch attr.Name.Local {
					case _Count:
						count = convert.MustInt(attr.Value)
					case _UniqueCount:
						unqCount = convert.MustInt(attr.Value)
					}
				}
				if unqCount != 0 {
					ssr.sharedStrings = make([]string, unqCount)
				} else {
					ssr.sharedStrings = make([]string, count)
				}
			default:
				_ = ssr.decoder.Skip()
			}
		case xml.EndElement:
			switch token.Name.Local {
			case _SI:
				ssr.index++
			case _T:
				tStart = false
			case _R:
				rStart = false
			}
		case xml.CharData:
			if tStart {
				if rStart {
					str := ssr.sharedStrings[ssr.index]
					ssr.sharedStrings[ssr.index] = str + string(token)
				} else {
					ssr.sharedStrings[ssr.index] = string(token)
				}
			}
		}
	}
}
