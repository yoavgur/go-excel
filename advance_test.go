package excel_test

import (
	"archive/zip"
	"os"
	"reflect"
	"testing"

	"github.com/google/go-cmp/cmp"
	convert "github.com/szyhf/go-convert"
	excel "github.com/szyhf/go-excel"
)

var expectAdvanceList = []Advance{
	{
		ID:      1,
		Name:    "Andy",
		NamePtr: strPtr("Andy"),
		Age:     1,
		Slice:   []int{1, 2},
		Temp: &Temp{
			Foo: "Andy",
		},
	},
	{
		ID:      2,
		Name:    "Leo",
		NamePtr: strPtr("Leo"),
		Age:     2,
		Slice:   []int{2, 3, 4},
		Temp: &Temp{
			Foo: "Leo",
		},
	},
	{
		ID:      3,
		Name:    "",
		NamePtr: strPtr("Ben"),
		Age:     180, //  using default
		Slice:   []int{3, 4, 5, 6},
		Temp: &Temp{
			Foo: "Ben",
		},
	},
	{
		ID:      4,
		Name:    "Ming",
		NamePtr: strPtr("Ming"),
		Age:     4,
		Slice:   []int{1},
		Temp: &Temp{
			Foo: "Default",
		},
	},
}

type Advance struct {
	// use field name as default column name
	ID int
	// column means to map the column name, and skip cell that value equal to "Ben"
	Name string `xlsx:"column(NameOf);nil(Ben);req();"`
	// you can map a column into more than one field
	NamePtr *string `xlsx:"column(NameOf);req();"`
	// omit `column` if only want to map to column name, it's equal to `column(AgeOf)`
	// use 180 as default if cell is empty.
	Age int `xlsx:"column(AgeOf);default(180);req();"`
	// split means to split the string into slice by the `|`
	Slice []int `xlsx:"split(|);req();"`
	// use default also can marshal to struct
	Temp *Temp `xlsx:"column(UnmarshalString);default({\"Foo\":\"Default\"});req();"`
	// use '-' to ignore.
	WantIgnored string `xlsx:"-"`
	// By default, required tag req is not set
	NotRequired string
}

func TestRead(t *testing.T) {
	// file
	conn := excel.NewConnecter()
	err := conn.Open(filePath)
	if err != nil {
		t.Error(err)
	}
	defer conn.Close()

	rd, err := conn.NewReaderByConfig(&excel.Config{
		// Sheet name as string or sheet model as object or as slice of objecg.
		Sheet: advSheetName,
		// Use the index row as title, every row before title-row will be ignore, default is 0.
		TitleRowIndex: 1,
		// Skip n row after title, default is 0 (not skip), empty row is not counted.
		Skip: 1,
		// Auto prefix to sheet name.
		Prefix: "",
		// Auto suffix to sheet name.
		Suffix: advSheetSuffix,
	})
	if err != nil {
		t.Error(err)
	}
	defer rd.Close()

	expectedSheetSize := uint64(2036)
	if rd.GetSheetSize() != expectedSheetSize {
		t.Errorf("unexpect sheet size: %d", rd.GetSheetSize())
	}

	idx := 0
	for rd.Next() {
		var a Advance
		err := rd.Read(&a)
		if err != nil {
			t.Error(err)
			return
		}
		expect := expectAdvanceList[idx]
		if !reflect.DeepEqual(expect, a) {
			t.Errorf("unexpect advance at %d = \n%s", idx, convert.MustJsonPrettyString(a))
		}

		idx++
	}

	if rd.InputOffset() != int64(expectedSheetSize) {
		t.Errorf("unexpect input offset: %d", rd.InputOffset())
	}
}

func testReadAllWithOpenedConnector(t *testing.T, conn excel.Connecter) {
	rd, err := conn.NewReaderByConfig(&excel.Config{
		// Sheet name as string or sheet model as object or a slice of object.
		Sheet: advSheetName,
		// Use the index row as title, every row before title-row will be ignore, default is 0.
		TitleRowIndex: 1,
		// Skip n row after title, default is 0 (not skip), empty row is not counted.
		Skip: 1,
		// Auto prefix to sheet name.
		Prefix: "",
		// Auto suffix to sheet name.
		Suffix: advSheetSuffix,
	})
	if err != nil {
		t.Error(err)
	}
	defer conn.Close()

	var slc []Advance
	err = rd.ReadAll(&slc)
	if err != nil {
		t.Error(err)
		return
	}
	if !reflect.DeepEqual(slc, expectAdvanceList) {
		t.Errorf("unexpect advance list: \n%s", convert.MustJsonPrettyString(slc))
	}
}

func TestEmptyWorkbook(t *testing.T) {
	conn := excel.NewConnecter()

	// see the Advancd.suffix sheet in simple.xlsx
	err := conn.Open(emptyFilePath)
	if err != nil {
		t.Error(err)
	}

	rd, err := conn.NewReader(conn.GetSheetNames()[0])
	if err != nil {
		t.Error(err)
	}

	if len(rd.GetTitles()) != 0 {
		t.Error("empty workbook should have no titles")
	}

	if rd.Next() {
		t.Error("empty workbook should have no rows")
	}
}

func TestReadAll(t *testing.T) {
	conn := excel.NewConnecter()

	t.Run("ReadAll", func(t *testing.T) {
		// see the Advancd.suffix sheet in simple.xlsx
		err := conn.Open(filePath)
		if err != nil {
			t.Error(err)
		}

		testReadAllWithOpenedConnector(t, conn)
	})

	t.Run("ReadAllByReader", func(t *testing.T) {
		zipReader, err := zip.OpenReader(filePath)
		if err != nil {
			t.Error(err)
		}

		err = conn.OpenReader(&zipReader.Reader)
		if err != nil {
			t.Error(err)
		}

		testReadAllWithOpenedConnector(t, conn)
	})

	t.Run("ReadAllByBinary", func(t *testing.T) {
		// Read data from file
		binaryData, err := os.ReadFile(filePath)
		if err != nil {
			t.Error(err)
		}

		err = conn.OpenBinary(binaryData)
		if err != nil {
			t.Error(err)
		}

		testReadAllWithOpenedConnector(t, conn)
	})
}

func TestReadLimitSharedStrings(t *testing.T) {
	tests := []struct {
		name           string
		limit          int64
		expectedTitles []string
		expectedOutput [][][]byte
		passedMemLimit bool
	}{
		{
			name:           "no limit",
			limit:          -1,
			expectedTitles: []string{"ID", "NameOf", "AgeOf", "Slice", "UnmarshalString"},
			expectedOutput: [][][]byte{
				{[]byte("1"), []byte("Andy"), []byte("1"), []byte("1|2"), []byte("{\"Foo\":\"Andy\"}")},
				{[]byte("2"), []byte("Leo"), []byte("2"), []byte("2|3|4"), []byte("{\"Foo\":\"Leo\"}")},
				{[]byte("3"), []byte("Ben"), nil, []byte("3|4|5|6"), []byte("{\"Foo\":\"Ben\"}")},
				{[]byte("4"), []byte("Ming"), []byte("4"), []byte("1"), nil},
			},
			passedMemLimit: false,
		},
		{
			name:           "zero bytes",
			limit:          0,
			expectedTitles: []string{"", "", "", "", ""},
			expectedOutput: [][][]byte{
				{[]byte("1"), []byte{}, []byte("1"), []byte{}, []byte{}},
				{[]byte("2"), []byte{}, []byte("2"), []byte{}, []byte{}},
				{[]byte("3"), []byte{}, nil, []byte{}, []byte{}},
				{[]byte("4"), []byte{}, []byte("4"), []byte("1"), nil},
			},
			passedMemLimit: true,
		},
		{
			name:           "partial headers",
			limit:          200,
			expectedTitles: []string{"ID", "NameOf", "", "", ""},
			expectedOutput: [][][]byte{
				{[]byte("1"), []byte{}, []byte("1"), []byte{}, []byte{}},
				{[]byte("2"), []byte{}, []byte("2"), []byte{}, []byte{}},
				{[]byte("3"), []byte{}, nil, []byte{}, []byte{}},
				{[]byte("4"), []byte{}, []byte("4"), []byte("1"), nil},
			},
			passedMemLimit: true,
		},
		{
			name:           "partial data",
			limit:          630,
			expectedTitles: []string{"ID", "NameOf", "AgeOf", "Slice", "UnmarshalString"},
			expectedOutput: [][][]byte{
				{[]byte("1"), []byte("Andy"), []byte("1"), []byte("1|2"), []byte("{\"Foo\":\"Andy\"}")},
				{[]byte("2"), []byte("Leo"), []byte("2"), []byte("2|3|4"), []byte("{\"Foo\":\"Leo\"}")},
				{[]byte("3"), []byte("Ben"), nil, []byte{}, []byte{}},
				{[]byte("4"), []byte{}, []byte("4"), []byte("1"), nil},
			},
			passedMemLimit: true,
		},
		{
			name:           "big limit",
			limit:          1500,
			expectedTitles: []string{"ID", "NameOf", "AgeOf", "Slice", "UnmarshalString"},
			expectedOutput: [][][]byte{
				{[]byte("1"), []byte("Andy"), []byte("1"), []byte("1|2"), []byte("{\"Foo\":\"Andy\"}")},
				{[]byte("2"), []byte("Leo"), []byte("2"), []byte("2|3|4"), []byte("{\"Foo\":\"Leo\"}")},
				{[]byte("3"), []byte("Ben"), nil, []byte("3|4|5|6"), []byte("{\"Foo\":\"Ben\"}")},
				{[]byte("4"), []byte("Ming"), []byte("4"), []byte("1"), nil},
			},
			passedMemLimit: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			conn := excel.NewConnecter()
			err := conn.OpenByConfig(filePath, excel.ConnecterConfig{MaxSharedStringsBytesToRead: test.limit})
			if err != nil {
				t.Error(err)
			}

			rd, err := conn.NewReaderByConfig(&excel.Config{
				Sheet:         advSheetName,
				TitleRowIndex: 1,
				Skip:          1,
				Prefix:        "",
				Suffix:        advSheetSuffix,
			})
			if err != nil {
				t.Error(err)
			}

			if !reflect.DeepEqual(rd.GetTitles(), test.expectedTitles) {
				t.Errorf("unexpect titles: \n%s", convert.MustJsonPrettyString(rd.GetTitles()))
			}

			var output [][][]byte
			err = rd.ReadAll(&output)
			if err != nil {
				t.Error(err)
			}

			if !cmp.Equal(test.expectedOutput, output) {
				t.Error(cmp.Diff(test.expectedOutput, output))
			}

			if conn.PassedSharedStringsLimit() != test.passedMemLimit {
				t.Errorf("unexpect passedMemLimit: %t", conn.PassedSharedStringsLimit())
			}

			rd.Close()
			conn.Close()
		})
	}

}
