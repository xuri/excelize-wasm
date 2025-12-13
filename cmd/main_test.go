package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"math/rand"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"syscall/js"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

var MacintoshCyrillicCharset = []byte{0x8F, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2, 0x20, 0xEC, 0xE8, 0xF0}

func TestRegInteropFunc(t *testing.T) {
	js.Global().Set("excelize", map[string]interface{}{})
	regFuncs()
}

func TestInTypeSlice(t *testing.T) {
	assert.Equal(t, -1, inTypeSlice(nil, js.TypeBoolean))
	assert.Equal(t, 0, inTypeSlice([]js.Type{js.TypeBoolean}, js.TypeBoolean))
}

func TestPrepareArgs(t *testing.T) {
	assert.ErrorIs(t, prepareArgs(nil, []argsRule{{}}), errArgNum)
	assert.ErrorIs(t, prepareArgs([]js.Value{js.ValueOf("true")},
		[]argsRule{{types: []js.Type{js.TypeBoolean}}}), errArgType)
	assert.NoError(t, prepareArgs([]js.Value{js.ValueOf(true)},
		[]argsRule{
			{types: []js.Type{js.TypeBoolean}},
			{types: []js.Type{js.TypeBoolean}, opts: true},
		}), errArgType)
	assert.NoError(t, prepareArgs([]js.Value{js.ValueOf(true), js.ValueOf(true)},
		[]argsRule{
			{types: []js.Type{js.TypeBoolean}},
			{types: []js.Type{js.TypeBoolean}, opts: true},
		}), errArgType)
}

func TestCellNameToCoordinates(t *testing.T) {
	ret := CellNameToCoordinates(js.Value{}, []js.Value{js.ValueOf("A1")})
	assert.Equal(t, 1, ret.(js.Value).Get("col").Int())
	assert.Equal(t, 1, ret.(js.Value).Get("row").Int())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = CellNameToCoordinates(js.Value{}, []js.Value{})
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = CellNameToCoordinates(js.Value{}, []js.Value{js.ValueOf("A")})
	assert.Equal(t, "cannot convert cell \"A\" to coordinates: invalid cell name \"A\"", ret.(js.Value).Get("error").String())
}

func TestColumnNameToNumber(t *testing.T) {
	ret := ColumnNameToNumber(js.Value{}, []js.Value{js.ValueOf("A")})
	assert.Equal(t, 1, ret.(js.Value).Get("col").Int())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = ColumnNameToNumber(js.Value{}, []js.Value{})
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = ColumnNameToNumber(js.Value{}, []js.Value{js.ValueOf("-")})
	assert.Equal(t, "invalid column name \"-\"", ret.(js.Value).Get("error").String())
}

func TestColumnNumberToName(t *testing.T) {
	ret := ColumnNumberToName(js.Value{}, []js.Value{js.ValueOf(1)})
	assert.Equal(t, "A", ret.(js.Value).Get("col").String())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = ColumnNumberToName(js.Value{}, []js.Value{})
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = ColumnNumberToName(js.Value{}, []js.Value{js.ValueOf(0)})
	assert.EqualError(t, excelize.ErrColumnNumber, ret.(js.Value).Get("error").String())
}

func TestCoordinatesToCellName(t *testing.T) {
	ret := CoordinatesToCellName(js.Value{}, []js.Value{js.ValueOf(1), js.ValueOf(1)})
	assert.Equal(t, "A1", ret.(js.Value).Get("cell").String())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = CoordinatesToCellName(js.Value{}, []js.Value{js.ValueOf(1), js.ValueOf(1), js.ValueOf(true)})
	assert.Equal(t, "$A$1", ret.(js.Value).Get("cell").String())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = CoordinatesToCellName(js.Value{}, []js.Value{})
	assert.Equal(t, "", ret.(js.Value).Get("cell").String())
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = CoordinatesToCellName(js.Value{}, []js.Value{js.ValueOf(0), js.ValueOf(1)})
	assert.Equal(t, "invalid cell reference [0, 1]", ret.(js.Value).Get("error").String())
}

func TestHSLToRGB(t *testing.T) {
	ret := HSLToRGB(js.Value{}, []js.Value{js.ValueOf(0), js.ValueOf(1), js.ValueOf(.4)})
	assert.Equal(t, 204, ret.(js.Value).Get("r").Int())
	assert.Equal(t, 0, ret.(js.Value).Get("g").Int())
	assert.Equal(t, 0, ret.(js.Value).Get("b").Int())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = HSLToRGB(js.Value{}, []js.Value{})
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())
}

func TestJoinCellName(t *testing.T) {
	ret := JoinCellName(js.Value{}, []js.Value{js.ValueOf("A"), js.ValueOf(1)})
	assert.Equal(t, "A1", ret.(js.Value).Get("cell").String())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = JoinCellName(js.Value{}, []js.Value{})
	assert.Equal(t, "", ret.(js.Value).Get("cell").String())
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = JoinCellName(js.Value{}, []js.Value{js.ValueOf("-"), js.ValueOf(1)})
	assert.Equal(t, "invalid column name \"-\"", ret.(js.Value).Get("error").String())
}

func TestRGBToHSL(t *testing.T) {
	ret := RGBToHSL(js.Value{}, []js.Value{js.ValueOf(0), js.ValueOf(255), js.ValueOf(255)})
	assert.Equal(t, 0.5, ret.(js.Value).Get("h").Float())
	assert.Equal(t, 1.0, ret.(js.Value).Get("s").Float())
	assert.Equal(t, 0.5, ret.(js.Value).Get("l").Float())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = RGBToHSL(js.Value{}, []js.Value{})
	assert.Equal(t, 0.0, ret.(js.Value).Get("h").Float())
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())
}

func TestSplitCellName(t *testing.T) {
	ret := SplitCellName(js.Value{}, []js.Value{js.ValueOf("A1")})
	assert.Equal(t, "A", ret.(js.Value).Get("col").String())
	assert.Equal(t, 1, ret.(js.Value).Get("row").Int())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = SplitCellName(js.Value{}, []js.Value{})
	assert.Equal(t, "", ret.(js.Value).Get("col").String())
	assert.Equal(t, 0, ret.(js.Value).Get("row").Int())
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())

	ret = SplitCellName(js.Value{}, []js.Value{js.ValueOf("A")})
	assert.Equal(t, "invalid cell name \"A\"", ret.(js.Value).Get("error").String())
}

func TestThemeColor(t *testing.T) {
	ret := ThemeColor(js.Value{}, []js.Value{js.ValueOf("000000"), js.ValueOf(-0.1)})
	assert.Equal(t, "FF000000", ret.(js.Value).Get("color").String())
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = ThemeColor(js.Value{}, []js.Value{})
	assert.Equal(t, "", ret.(js.Value).Get("color").String())
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())
}

func TestNewFile(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	f = NewFile(js.Value{}, []js.Value{js.ValueOf(map[string]interface{}{
		"ShortDatePattern": "yyyy/m/d",
	})})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	f = NewFile(js.Value{}, []js.Value{js.ValueOf(map[string]interface{}{
		"ShortDatePattern": "yyyy/m/d",
	}), js.ValueOf(true)})
	assert.EqualError(t, errArgNum, f.(js.Value).Get("error").String())

	f = NewFile(js.Value{}, []js.Value{js.ValueOf(map[string]interface{}{
		"ShortDatePattern": true,
	})})
	assert.EqualError(t, errArgType, f.(js.Value).Get("error").String())
}

func TestOpenReader(t *testing.T) {
	buf, err := excelize.NewFile().WriteToBuffer()
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array")
	ret := OpenReader(js.Value{}, []js.Value{uint8Array})
	assert.EqualError(t, errArgType, ret.(js.Value).Get("error").String())

	uint8Array = js.Global().Get("Uint8Array").New(js.ValueOf(buf.Len()))
	for k, v := range buf.Bytes() {
		uint8Array.SetIndex(k, v)
	}
	ret = OpenReader(js.Value{}, []js.Value{uint8Array})
	assert.True(t, ret.(js.Value).Get("error").IsNull())

	ret = OpenReader(js.Value{}, []js.Value{uint8Array, js.ValueOf(map[string]interface{}{
		"Password": false,
	})})
	assert.EqualError(t, errArgType, ret.(js.Value).Get("error").String())

	buf = new(bytes.Buffer)
	_, err = excelize.NewFile().WriteTo(buf, excelize.Options{Password: "passwd"})
	assert.NoError(t, err)
	uint8Array = js.Global().Get("Uint8Array").New(js.ValueOf(buf.Len()))
	for k, v := range buf.Bytes() {
		uint8Array.SetIndex(k, v)
	}
	ret = OpenReader(js.Value{}, []js.Value{uint8Array, js.ValueOf(map[string]interface{}{
		"Password": "invalid",
	})})
	assert.EqualError(t, excelize.ErrWorkbookPassword, ret.(js.Value).Get("error").String())

	ret = OpenReader(js.Value{}, []js.Value{uint8Array})
	assert.EqualError(t, zip.ErrFormat, ret.(js.Value).Get("error").String())

	ret = OpenReader(js.Value{}, []js.Value{js.ValueOf(map[string]interface{}{})})
	assert.EqualError(t, excelize.ErrParameterInvalid, ret.(js.Value).Get("error").String())

	ret = OpenReader(js.Value{}, []js.Value{js.ValueOf(map[string]interface{}{})})
	assert.EqualError(t, excelize.ErrParameterInvalid, ret.(js.Value).Get("error").String())

	ret = OpenReader(js.Value{}, []js.Value{})
	assert.EqualError(t, errArgNum, ret.(js.Value).Get("error").String())
}

func TestAddChart(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	lineChart := js.ValueOf(map[string]interface{}{
		"Type": int(excelize.Line),
		"Series": []interface{}{
			map[string]interface{}{
				"Name":       "Sheet1!$A$2",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$2:$D$2",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$3",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$3:$D$3",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$4",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$4:$D$4",
			},
		},
		"Title": []interface{}{
			js.ValueOf(map[string]interface{}{"Text": "Fruit 3D Clustered Column Chart"}),
		},
	})
	colChart := js.ValueOf(map[string]interface{}{
		"Type": int(excelize.Col3DClustered),
		"Series": []interface{}{
			map[string]interface{}{
				"Name":       "Sheet1!$A$2",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$2:$D$2",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$3",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$3:$D$3",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$4",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$4:$D$4",
			},
		},
		"Title": []interface{}{
			js.ValueOf(map[string]interface{}{"Text": "Fruit 3D Clustered Column Chart"}),
		},
	})
	ret := f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), lineChart)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), lineChart, colChart)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Type": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Type": int(excelize.Col)}), js.ValueOf(map[string]interface{}{"Type": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Type": int(excelize.Col)}), js.ValueOf(map[string]interface{}{}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Type": 65}))
	assert.Equal(t, "unsupported chart type 65", ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Type": int(excelize.Col)}), js.ValueOf(map[string]interface{}{"Type": 65}))
	assert.Equal(t, "unsupported chart type 65", ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("SheetN"), js.ValueOf("A1"),
		js.ValueOf(map[string]interface{}{"Type": int(excelize.Col3DClustered)}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAddChartSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	lineChart := js.ValueOf(map[string]interface{}{
		"Type": int(excelize.Line),
		"Series": []interface{}{
			map[string]interface{}{
				"Name":       "Sheet1!$A$2",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$2:$D$2",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$3",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$3:$D$3",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$4",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$4:$D$4",
			},
		},
		"Title": []interface{}{
			js.ValueOf(map[string]interface{}{"Text": "Fruit 3D Clustered Column Chart"}),
		},
	})
	colChart := js.ValueOf(map[string]interface{}{
		"Type": int(excelize.Col3DClustered),
		"Series": []interface{}{
			map[string]interface{}{
				"Name":       "Sheet1!$A$2",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$2:$D$2",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$3",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$3:$D$3",
			},
			map[string]interface{}{
				"Name":       "Sheet1!$A$4",
				"Categories": "Sheet1!$B$1:$D$1",
				"Values":     "Sheet1!$B$4:$D$4",
			},
		},
		"Title": []interface{}{
			js.ValueOf(map[string]interface{}{"Text": "Fruit 3D Clustered Column Chart"}),
		},
	})
	ret := f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet2"), lineChart)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet3"), lineChart, colChart)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChartSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet4"), js.ValueOf(map[string]interface{}{"Type": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet2"), js.ValueOf(map[string]interface{}{"Type": int(excelize.Col)}), js.ValueOf(map[string]interface{}{"Type": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet4"), js.ValueOf(map[string]interface{}{"Type": 65}))
	assert.Equal(t, "unsupported chart type 65", ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet5"), js.ValueOf(map[string]interface{}{"Type": int(excelize.Col)}), js.ValueOf(map[string]interface{}{"Type": 65}))
	assert.Equal(t, "unsupported chart type 65", ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{}))
	assert.EqualError(t, excelize.ErrExistsSheet, ret.Get("error").String())
}

func TestComments(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	comment := js.ValueOf(map[string]interface{}{
		"Cell":   "A12",
		"Author": "Excelize",
		"Paragraph": []interface{}{
			map[string]interface{}{
				"Text": "Excelize: ",
				"Font": map[string]interface{}{"Bold": true},
			},
			map[string]interface{}{"Text": "This is a comment."},
		},
	})

	ret := f.(js.Value).Call("AddComment", js.ValueOf("Sheet1"), comment)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddComment")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddComment", js.ValueOf("SheetN"), js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddComment", js.ValueOf("Sheet1"), map[string]interface{}{"Cell": true})
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddComment", js.ValueOf("SheetN"), comment)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetComments")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetComments", js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetComments", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, comment.Get("Author").String(), ret.Get("comments").Index(0).Get("Author").String())

	ret = f.(js.Value).Call("GetComments", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDataValidation(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	dv := js.ValueOf(map[string]interface{}{})

	ret := f.(js.Value).Call("AddDataValidation", js.ValueOf("Sheet1"), dv)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddDataValidation")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddDataValidation", js.ValueOf("Sheet1"), js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddDataValidation", js.ValueOf("Sheet1"), map[string]interface{}{"Type": true})
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddDataValidation", js.ValueOf("SheetN"), dv)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetDataValidations", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, ret.Get("dataValidations").Length(), 1)

	ret = f.(js.Value).Call("GetDataValidations")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetDataValidations", js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetDataValidations", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestFormControl(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddFormControl", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{
		"Cell": "A1", "Type": int(excelize.FormControlButton),
	}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetFormControls", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("formControls").Length())

	ret = f.(js.Value).Call("DeleteFormControl", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddFormControl")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddFormControl", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{
		"Cell": "A1", "Type": true,
	}))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddFormControl", js.ValueOf("SheetN"), js.ValueOf(map[string]interface{}{
		"Cell": "A1", "Type": int(excelize.FormControlButton),
	}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetFormControls")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetFormControls", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetFormControls", js.ValueOf(true))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteFormControl")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteFormControl", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteFormControl", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAddHeaderFooterImage(t *testing.T) {
	buf, err := os.ReadFile(filepath.Join("..", "chart.png"))
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(len(buf)))
	for k, v := range buf {
		uint8Array.SetIndex(k, v)
	}

	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	opts := js.ValueOf(map[string]interface{}{
		"File":      js.ValueOf(uint8Array),
		"IsFooter":  true,
		"FirstPage": true,
		"Extension": ".png",
		"Width":     "50pt",
		"Height":    "32pt",
	})
	ret := f.(js.Value).Call("AddHeaderFooterImage", js.ValueOf("Sheet1"), opts)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddHeaderFooterImage", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"File": js.ValueOf(uint8Array), "Extension": "png"}),
	)
	assert.Equal(t, "unsupported image extension", ret.Get("error").String())

	ret = f.(js.Value).Call("AddHeaderFooterImage")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddHeaderFooterImage", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddHeaderFooterImage", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Extension": true}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddHeaderFooterImage", js.ValueOf("SheetN"), opts)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAddIgnoredErrors(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddIgnoredErrors", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(int(excelize.IgnoredErrorsEvalError)))
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddIgnoredErrors", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf(int(excelize.IgnoredErrorsEvalError)))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("AddIgnoredErrors", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddIgnoredErrors", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestAddPictureFromBytes(t *testing.T) {
	buf, err := os.ReadFile(filepath.Join("..", "chart.png"))
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(len(buf)))
	for k, v := range buf {
		uint8Array.SetIndex(k, v)
	}

	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	pic := js.ValueOf(map[string]interface{}{
		"Extension": ".png",
		"File":      js.ValueOf(uint8Array),
		"Format": map[string]interface{}{
			"AltText": "Picture 1",
		},
	})
	ret := f.(js.Value).Call("AddPictureFromBytes", js.ValueOf("Sheet1"), js.ValueOf("A1"), pic)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictures", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	assert.Equal(t, 1, ret.Get("pictures").Length())
	assert.Equal(t, uint8Array.Length(), ret.Get("pictures").Index(0).Get("File").Length())

	ret = f.(js.Value).Call("GetPictureCells", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	assert.Equal(t, 1, ret.Get("cells").Length())
	assert.Equal(t, "A1", ret.Get("cells").Index(0).String())

	ret = f.(js.Value).Call("AddPictureFromBytes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddPictureFromBytes", js.ValueOf("Sheet1"), js.ValueOf("A1"),
		js.ValueOf(map[string]interface{}{
			"Extension": ".png",
			"File":      uint8Array,
			"Format": map[string]interface{}{
				"Locked": 1,
			},
		}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictureCells", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddPictureFromBytes", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"Extension": "png", "File": uint8Array, "Format": map[string]interface{}{}}))
	assert.EqualError(t, excelize.ErrImgExt, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictures", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictureCells")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictures", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetPictureCells", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestPivotTable(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	year := []int{2017, 2018, 2019}
	types := []string{"Meat", "Dairy", "Beverages", "Produce"}
	region := []string{"East", "West", "North", "South"}
	ret := f.(js.Value).Call("SetSheetRow",
		js.ValueOf("Sheet1"), js.ValueOf("A1"),
		js.ValueOf([]interface{}{"Month", "Year", "Type", "Sales", "Region"}),
	)
	assert.True(t, ret.Get("error").IsNull())
	for row := 2; row < 32; row++ {
		ret = f.(js.Value).Call("SetCellValue",
			js.ValueOf("Sheet1"), js.ValueOf(fmt.Sprintf("A%d", row)), js.ValueOf(month[rand.Intn(12)]),
		)
		assert.True(t, ret.Get("error").IsNull())
		ret = f.(js.Value).Call("SetCellValue",
			js.ValueOf("Sheet1"), js.ValueOf(fmt.Sprintf("B%d", row)), js.ValueOf(year[rand.Intn(3)]),
		)
		assert.True(t, ret.Get("error").IsNull())
		ret = f.(js.Value).Call("SetCellValue",
			js.ValueOf("Sheet1"), js.ValueOf(fmt.Sprintf("C%d", row)), js.ValueOf(types[rand.Intn(4)]),
		)
		assert.True(t, ret.Get("error").IsNull())
		ret = f.(js.Value).Call("SetCellValue",
			js.ValueOf("Sheet1"), js.ValueOf(fmt.Sprintf("D%d", row)), js.ValueOf(rand.Intn(5000)),
		)
		assert.True(t, ret.Get("error").IsNull())
		ret = f.(js.Value).Call("SetCellValue",
			js.ValueOf("Sheet1"), js.ValueOf(fmt.Sprintf("E%d", row)), js.ValueOf(region[rand.Intn(4)]),
		)
		assert.True(t, ret.Get("error").IsNull())
	}
	opts := js.ValueOf(map[string]interface{}{
		"DataRange":       "Sheet1!$A$1:$E$31",
		"PivotTableRange": "Sheet1!$G$2:$M$34",
		"Rows":            []interface{}{map[string]interface{}{"Data": "Month", "DefaultSubtotal": true}, map[string]interface{}{"Data": "Year"}},
		"Filter":          []interface{}{map[string]interface{}{"Data": "Region"}},
		"Columns":         []interface{}{map[string]interface{}{"Data": "Type", "DefaultSubtotal": true}},
		"Data":            []interface{}{map[string]interface{}{"Data": "Sales", "Subtotal": "Sum", "Name": "Summarize by Sum"}},
		"RowGrandTotals":  true,
		"ColGrandTotals":  true,
		"ShowDrill":       true,
		"ShowRowHeaders":  true,
		"ShowColHeaders":  true,
		"ShowLastColumn":  true,
		"ShowError":       true,
	})
	ret = f.(js.Value).Call("AddPivotTable", opts)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetPivotTables", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddPivotTable")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddPivotTable", js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddPivotTable", map[string]interface{}{"ShowError": 1})
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddPivotTable", js.ValueOf(map[string]interface{}{}))
	assert.Equal(t, "parameter 'PivotTableRange' parsing error: parameter is required", ret.Get("error").String())

	ret = f.(js.Value).Call("GetPivotTables")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPivotTables", js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetPivotTables", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAddShape(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddShape", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Cell": "C30", "Type": "rect", "Paragraph": map[string]interface{}{}}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddShape")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddShape", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Cell": "C30", "Type": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddShape", js.ValueOf("SheetN"), js.ValueOf(map[string]interface{}{"Cell": "C30", "Type": "rect"}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("AddShape", js.ValueOf("Sheet1"), js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())
}

func TestSlicer(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddTable", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Name": "Table1", "Range": "A1:D5"}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddSlicer", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Name":       "Column1",
			"Cell":       "E1",
			"TableSheet": "Sheet1",
			"TableName":  "Table1",
			"Caption":    "Column1",
		}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetSlicers", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("slicers").Length())
	assert.Equal(t, "Column1", ret.Get("slicers").Index(0).Get("Name").String())

	ret = f.(js.Value).Call("DeleteSlicer", js.ValueOf("Column1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddSlicer")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddSlicer", js.ValueOf("Sheet1"), js.ValueOf(nil))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddSlicer", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Name": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddSlicer", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{
			"Name":       "Column1",
			"Cell":       "E1",
			"TableSheet": "SheetN",
			"TableName":  "Table1",
			"Caption":    "Column1",
		}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer", js.ValueOf("X"))
	assert.Equal(t, "slicer X does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer", js.ValueOf(nil))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSlicers")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSlicers", js.ValueOf(nil))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSlicers", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer", js.ValueOf("X"))
	assert.Equal(t, "slicer X does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSlicer", js.ValueOf(nil))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestAddSparkline(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddSparkline", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Location": []interface{}{"A2"},
			"Range":    []interface{}{"Sheet1!B1:J1"},
		}),
	)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddSparkline", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddSparkline", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Location": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddSparkline", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Location": []interface{}{"A2"},
			"Range":    []interface{}{"Sheet1!B1:J1"},
			"Style":    -1,
		}),
	)
	assert.Equal(t, "parameter 'Style' value must be an integer from 0 to 35", ret.Get("error").String(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddSparkline", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{
			"Location": []interface{}{"A2"},
			"Range":    []interface{}{"Sheet1!B1:J1"},
		}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestTable(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddTable", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"Range": "B26:A21"}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetTables", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("tables").Length())
	assert.Equal(t, "A21:B26", ret.Get("tables").Index(0).Get("Range").String())

	ret = f.(js.Value).Call("DeleteTable", js.ValueOf("Table1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddTable")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddTable", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"Name": true, "Range": "B26:A21"}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AddTable", js.ValueOf("SheetN"), js.ValueOf(map[string]interface{}{"Range": "B26:A21"}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetTables")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetTables", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetTables", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteTable")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteTable", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteTable", js.ValueOf("X"))
	assert.Equal(t, "table X does not exist", ret.Get("error").String())
}

func TestAddVBAProject(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	oleIdentifier := []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1}
	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(len(oleIdentifier)))
	for k, v := range oleIdentifier {
		uint8Array.SetIndex(k, v)
	}
	ret := f.(js.Value).Call("AddVBAProject", js.ValueOf(uint8Array))
	assert.True(t, ret.Get("error").IsNull())

	uint8Array = js.Global().Get("Uint8Array").New(js.ValueOf(1))
	ret = f.(js.Value).Call("AddVBAProject", js.ValueOf(uint8Array))
	assert.Equal(t, excelize.ErrAddVBAProject.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("AddVBAProject")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddVBAProject", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestAutoFilter(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AutoFilter", js.ValueOf("Sheet1"), js.ValueOf("D4:B1"), js.ValueOf([]interface{}{map[string]interface{}{}}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AutoFilter")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AutoFilter", js.ValueOf("Sheet1"), js.ValueOf("D4:B1"), js.ValueOf([]interface{}{map[string]interface{}{"Column": 1}}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("AutoFilter", js.ValueOf("SheetN"), js.ValueOf("D4:B1"), js.ValueOf([]interface{}{map[string]interface{}{}}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestCalcCellValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("CalcCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("CalcCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"RawCellValue": true}))
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("CalcCellValue")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("CalcCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"RawCellValue": 1}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("CalcCellValue", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestCalcProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCalcProps",
		js.ValueOf(map[string]interface{}{
			"FullCalcOnLoad":        true,
			"CalcID":                122211,
			"ConcurrentManualCount": 5,
			"IterateCount":          10,
			"ConcurrentCalc":        true,
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetCalcProps")
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	assert.True(t, ret.Get("props").Get("FullCalcOnLoad").Bool())
	assert.Equal(t, 122211, ret.Get("props").Get("CalcID").Int())
	assert.Equal(t, 5, ret.Get("props").Get("ConcurrentManualCount").Int())
	assert.Equal(t, 10, ret.Get("props").Get("IterateCount").Int())
	assert.True(t, ret.Get("props").Get("ConcurrentCalc").Bool())
	assert.True(t, ret.Get("props").Get("ForceFullCalc").IsUndefined())

	ret = f.(js.Value).Call("GetCalcProps", js.ValueOf(map[string]interface{}{"CalcMode": true}))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCalcProps", js.ValueOf(map[string]interface{}{"RefMode": "a1"}))
	assert.Equal(t, "invalid RefMode value \"a1\", acceptable value should be one of A1, R1C1", ret.Get("error").String())

	ret = f.(js.Value).Call("SetCalcProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCalcProps",
		js.ValueOf(map[string]interface{}{"CalcMode": true}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestCopySheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("CopySheet", js.ValueOf(0), ret.Get("index").Int())
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("CopySheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("CopySheet", js.ValueOf(-1), js.ValueOf(-1))
	assert.EqualError(t, excelize.ErrSheetIdx, ret.Get("error").String())
}

func TestDeleteChart(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DeleteChart", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteChart")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteChart", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDeleteComment(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DeleteComment", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteComment")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteComment", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDeleteDataValidation(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DeleteDataValidation", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteDataValidation", js.ValueOf("Sheet1"), js.ValueOf("A1:A2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteDataValidation")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteDataValidation", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDeleteDefinedName(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetDefinedName", js.ValueOf(map[string]interface{}{
		"Name":     "Amount",
		"RefersTo": "Sheet1!$A$2:$D$5",
		"Comment":  "defined name comment",
	}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteDefinedName", js.ValueOf(map[string]interface{}{
		"Name": "Amount",
	}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteDefinedName", js.ValueOf(map[string]interface{}{
		"Name": "No Exist Defined Name",
	}))
	assert.EqualError(t, excelize.ErrDefinedNameScope, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteDefinedName")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteDefinedName", js.ValueOf(map[string]interface{}{
		"Name": true,
	}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestDeletePicture(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DeletePicture", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeletePicture")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DeletePicture", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDeleteSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DeleteSheet", js.ValueOf("Sheet:1"))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("DeleteSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestDuplicateRow(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DuplicateRow", js.ValueOf("Sheet1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DuplicateRow")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DuplicateRow", js.ValueOf("SheetN"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestDuplicateRowTo(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("DuplicateRowTo", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("DuplicateRowTo")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("DuplicateRowTo", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestGetActiveSheetIndex(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetActiveSheetIndex")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("index").Int())

	ret = f.(js.Value).Call("GetActiveSheetIndex", js.ValueOf(1))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("index").Int())
}

func TestGetBaseColor(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetBaseColor", js.ValueOf("FFFFFF"), js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "FFFFFF", ret.Get("color").String())

	ret = f.(js.Value).Call("GetBaseColor", js.ValueOf("FFFFFF"), js.ValueOf(0), js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "FFFFFF", ret.Get("color").String())

	ret = f.(js.Value).Call("GetBaseColor", js.ValueOf("FFFFFF"), js.ValueOf(0), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "000000", ret.Get("color").String())

	ret = f.(js.Value).Call("GetBaseColor")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetBaseColor", js.ValueOf("FFFFFF"), js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())
}

func TestGetAppProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetAppProps")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Go Excelize", ret.Get("props").Get("Application").String())

	ret = f.(js.Value).Call("GetAppProps", js.ValueOf(1))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.True(t, ret.Get("props").IsUndefined())
}

func TestGetCellFormula(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellFormula", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("formula").String())

	ret = f.(js.Value).Call("GetCellFormula")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellFormula", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, "", ret.Get("formula").String())
}

func TestGetCellHyperLink(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellHyperLink", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.False(t, ret.Get("ok").Bool())

	ret = f.(js.Value).Call("GetCellHyperLink")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellHyperLink", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.False(t, ret.Get("ok").Bool())
}

func TestGetCellStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellStyle", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("style").Int())

	ret = f.(js.Value).Call("GetCellStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellStyle", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("style").Int())
}

func TestGetCellType(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellType", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("cellType").Int())

	ret = f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetCellType", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("cellType").Int())

	ret = f.(js.Value).Call("GetCellType")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellType", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("cellType").Int())
}

func TestGetCellValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("GetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"RawCellValue": true}))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("GetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(map[string]interface{}{"RawCellValue": "true"}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("GetCellValue")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellValue", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, "", ret.Get("value").String())
}

func TestGetColOutlineLevel(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetColOutlineLevel", js.ValueOf("Sheet1"), js.ValueOf("A"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("level").Int())

	ret = f.(js.Value).Call("GetColOutlineLevel")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetColOutlineLevel", js.ValueOf("SheetN"), js.ValueOf("A"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("level").Int())
}

func TestGetColStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetColStyle", js.ValueOf("Sheet1"), js.ValueOf("A"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("style").Int())

	ret = f.(js.Value).Call("GetColStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetColStyle", js.ValueOf("SheetN"), js.ValueOf("A"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("style").Int())
}

func TestGetColVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetColVisible", js.ValueOf("Sheet1"), js.ValueOf("A"))
	assert.True(t, ret.Get("error").IsNull())
	assert.True(t, ret.Get("visible").Bool())

	ret = f.(js.Value).Call("GetColVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetColVisible", js.ValueOf("SheetN"), js.ValueOf("A"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.False(t, ret.Get("visible").Bool())
}

func TestGetColWidth(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetColWidth", js.ValueOf("Sheet1"), js.ValueOf("A"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 9.140625, ret.Get("width").Float())

	ret = f.(js.Value).Call("GetColWidth")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetColWidth", js.ValueOf("SheetN"), js.ValueOf("A"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 9.140625, ret.Get("width").Float())
}

func TestGetCols(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"RawCellValue": true}))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"RawCellValue": "true"}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetCols")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())
}

func TestGetDefaultFont(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetDefaultFont")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Calibri", ret.Get("fontName").String())

	ret = f.(js.Value).Call("GetDefaultFont", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestGetMergeCells(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("value"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("MergeCell", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("C3"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetMergeCells", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	mergeCell := ret.Get("mergeCells").Index(0)
	assert.Equal(t, "value", mergeCell.Call("GetCellValue").String())
	assert.Equal(t, "A1", mergeCell.Call("GetStartAxis").String())
	assert.Equal(t, "C3", mergeCell.Call("GetEndAxis").String())

	// Test get merged cells without cell values
	ret = f.(js.Value).Call("GetMergeCells", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	mergeCell = ret.Get("mergeCells").Index(0)
	assert.Empty(t, mergeCell.Call("GetCellValue").String())
	assert.Equal(t, "A1", mergeCell.Call("GetStartAxis").String())
	assert.Equal(t, "C3", mergeCell.Call("GetEndAxis").String())

	ret = f.(js.Value).Call("GetMergeCells", js.ValueOf(1))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetMergeCells")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetMergeCells", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("mergeCells").Length())
}

func TestGetRowHeight(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetRowHeight", js.ValueOf("Sheet1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 15.0, ret.Get("height").Float())

	ret = f.(js.Value).Call("GetRowHeight")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetRowHeight", js.ValueOf("SheetN"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 15.0, ret.Get("height").Float())
}

func TestGetRowOutlineLevel(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetRowOutlineLevel", js.ValueOf("Sheet1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("level").Int())

	ret = f.(js.Value).Call("GetRowOutlineLevel")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetRowOutlineLevel", js.ValueOf("SheetN"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("level").Int())
}

func TestGetRowVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetRowVisible", js.ValueOf("Sheet1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())
	assert.False(t, ret.Get("visible").Bool())

	ret = f.(js.Value).Call("GetRowVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetRowVisible", js.ValueOf("SheetN"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.False(t, ret.Get("visible").Bool())
}

func TestSheetDimension(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetDimension", js.ValueOf("Sheet1"), js.ValueOf("A1:D5"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetSheetDimension", js.ValueOf("Sheet1"))
	assert.Equal(t, "A1:D5", ret.Get("dimension").String())
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetDimension", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetDimension")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetDimension", js.ValueOf("SheetN"), js.ValueOf("A1:D5"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetDimension", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetDimension")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetDimension", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestGetRows(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"RawCellValue": true}))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"RawCellValue": "true"}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetRows")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())
}

func TestGetSheetIndex(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetSheetIndex", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("index").Int())

	ret = f.(js.Value).Call("GetSheetIndex")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetIndex", js.ValueOf("Sheet:1"))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetIndex", js.ValueOf("SheetN"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, -1, ret.Get("index").Int())
}

func TestGetSheetList(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetSheetList")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("list").Length())

	ret = f.(js.Value).Call("GetSheetList", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("list").Length())
}

func TestGetSheetMap(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetSheetMap")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Sheet1", ret.Get("sheets").Get("1").String())

	ret = f.(js.Value).Call("GetSheetMap", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("sheets").Length())
}

func TestGetSheetName(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetSheetName", js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Sheet1", ret.Get("name").String())

	ret = f.(js.Value).Call("GetSheetName")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, "", ret.Get("name").String())

	ret = f.(js.Value).Call("GetSheetName", js.ValueOf(-1))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("name").String())
}

func TestGetSheetVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetSheetVisible", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.True(t, ret.Get("visible").Bool())

	ret = f.(js.Value).Call("GetSheetVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetVisible", js.ValueOf("Sheet:1"))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetVisible", js.ValueOf("SheetN"))
	assert.True(t, ret.Get("error").IsNull())
	assert.False(t, ret.Get("visible").Bool())
}

func TestGroupSheets(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("NewSheet", js.ValueOf("Sheet3"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GroupSheets", js.ValueOf([]interface{}{"Sheet2", "Sheet3"}))
	assert.EqualError(t, excelize.ErrGroupSheets, ret.Get("error").String())

	ret = f.(js.Value).Call("GroupSheets", js.ValueOf([]interface{}{"Sheet1", "Sheet2"}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GroupSheets")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestInsertCols(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("InsertCols", js.ValueOf("Sheet1"), js.ValueOf("A"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("InsertCols")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("InsertCols", js.ValueOf("SheetN"), js.ValueOf("A"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestInsertPageBreak(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("InsertPageBreak", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("InsertPageBreak")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("InsertPageBreak", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestInsertRows(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("InsertRows", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("InsertRows")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("InsertRows", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestMergeCell(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("MergeCell", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("B2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("MergeCell")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("MergeCell", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("B2"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestMoveSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("MoveSheet", js.ValueOf("Sheet2"), js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("MoveSheet", js.ValueOf("Sheet1"), js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("MoveSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("MoveSheet", js.ValueOf("Sheet1"), js.ValueOf(nil))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestNewConditionalStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewConditionalStyle",
		js.ValueOf(map[string]interface{}{
			"Fill": map[string]interface{}{
				"Type":    "pattern",
				"Color":   []interface{}{"FEEAA0"},
				"Pattern": 1,
			},
		}),
	)
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	styleID := ret.Get("style")
	assert.Equal(t, 0, styleID.Int())

	ret = f.(js.Value).Call("GetConditionalStyle", styleID)
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "pattern", ret.Get("style").Get("Fill").Get("Type").String())

	ret = f.(js.Value).Call("GetConditionalStyle", js.ValueOf(2))
	assert.Equal(t, "invalid style ID 2", ret.Get("error").String())

	ret = f.(js.Value).Call("NewConditionalStyle",
		js.ValueOf(map[string]interface{}{
			"Font": map[string]interface{}{
				"Size": excelize.MaxFontSize + 1,
			},
		}),
	)
	assert.EqualError(t, excelize.ErrFontSize, ret.Get("error").String())

	ret = f.(js.Value).Call("NewConditionalStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetConditionalStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("NewConditionalStyle", js.ValueOf(map[string]interface{}{"Fill": 1}))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("style").Int())

	ret = f.(js.Value).Call("GetConditionalStyle", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestNewSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("index").Int())

	ret = f.(js.Value).Call("NewSheet", js.ValueOf("Sheet:1"))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("NewSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("index").Int())
}

func TestStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewStyle", js.ValueOf(map[string]interface{}{
		"NumFmt":        1,
		"DecimalPlaces": 2,
		"CustomNumFmt":  "0.00",
		"NegRed":        true,
		"Border": []interface{}{
			map[string]interface{}{"Type": "left", "Color": "000000", "Style": 1},
		},
		"Fill": map[string]interface{}{
			"Type":    "gradient",
			"Color":   []interface{}{"FFFFFF", "E0EBF5"},
			"Shading": 1,
		},
		"Alignment": map[string]interface{}{
			"Horizontal":      "left",
			"Indent":          1,
			"JustifyLastLine": true,
			"ReadingOrder":    1,
			"RelativeIndent":  1,
			"ShrinkToFit":     true,
			"TextRotation":    90,
			"Vertical":        "center",
			"WrapText":        true,
		},
		"Font": map[string]interface{}{
			"Bold":      true,
			"Italic":    true,
			"Underline": "single",
			"Family":    "Calibri",
			"Size":      12,
			"Strike":    true,
			"Color":     "000000",
			"VertAlign": "superscript",
		},
		"Protection": map[string]interface{}{
			"Hidden": true,
			"Locked": true,
		},
	}))
	assert.True(t, ret.Get("error").IsNull(), ret.Get("error").String())
	assert.Equal(t, 1, ret.Get("style").Int())

	ret = f.(js.Value).Call("GetStyle", ret.Get("style"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "0.00", ret.Get("style").Get("CustomNumFmt").String())
	assert.True(t, ret.Get("style").Get("Font").Get("Bold").Bool())
	assert.Equal(t, "single", ret.Get("style").Get("Font").Get("Underline").String())

	for _, arg := range []map[string]interface{}{
		{"NumFmt": "1"},
		{"DecimalPlaces": "2"},
		{"CustomNumFmt": true},
		{"NegRed": "true"},
		{"Border": true},
		{"Border": []interface{}{map[string]interface{}{"Type": true}}},
		{"Border": []interface{}{map[string]interface{}{"Color": true}}},
		{"Border": []interface{}{map[string]interface{}{"Style": "1"}}},
		{"Fill": true},
		{"Fill": map[string]interface{}{"Type": true}},
		{"Fill": map[string]interface{}{"Color": true}},
		{"Fill": map[string]interface{}{"Color": []interface{}{true}}},
		{"Fill": map[string]interface{}{"Shading": "1"}},
		{"Alignment": true},
		{"Alignment": map[string]interface{}{"Horizontal": true}},
		{"Alignment": map[string]interface{}{"Indent": "1"}},
		{"Alignment": map[string]interface{}{"JustifyLastLine": "true"}},
		{"Alignment": map[string]interface{}{"ReadingOrder": "1"}},
		{"Alignment": map[string]interface{}{"RelativeIndent": "1"}},
		{"Alignment": map[string]interface{}{"ShrinkToFit": "true"}},
		{"Alignment": map[string]interface{}{"TextRotation": "90"}},
		{"Alignment": map[string]interface{}{"Vertical": true}},
		{"Alignment": map[string]interface{}{"WrapText": "true"}},
		{"Font": true},
		{"Font": map[string]interface{}{"Bold": "true"}},
		{"Font": map[string]interface{}{"Italic": "true"}},
		{"Font": map[string]interface{}{"Underline": true}},
		{"Font": map[string]interface{}{"Family": true}},
		{"Font": map[string]interface{}{"Size": "12"}},
		{"Font": map[string]interface{}{"Strike": "true"}},
		{"Font": map[string]interface{}{"Color": true}},
		{"Font": map[string]interface{}{"VertAlign": true}},
		{"Protection": true},
		{"Protection": map[string]interface{}{"Hidden": "true"}},
		{"Protection": map[string]interface{}{"Locked": "true"}},
	} {
		ret = f.(js.Value).Call("NewStyle", js.ValueOf(arg))
		assert.EqualError(t, errArgType, ret.Get("error").String())
	}

	ret = f.(js.Value).Call("NewStyle",
		js.ValueOf(map[string]interface{}{
			"Font": map[string]interface{}{
				"Size": excelize.MaxFontSize + 1,
			},
		}),
	)
	assert.EqualError(t, excelize.ErrFontSize, ret.Get("error").String())

	ret = f.(js.Value).Call("NewStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetStyle", js.ValueOf(-1))
	assert.Equal(t, "invalid style ID -1", ret.Get("error").String())

	ret = f.(js.Value).Call("GetStyle", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestProtectSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("ProtectSheet", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Password": "password",
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("ProtectSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("ProtectSheet", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{
			"Password": "password",
		}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("ProtectSheet", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Password": true,
		}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestProtectWorkbook(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("ProtectWorkbook",
		js.ValueOf(map[string]interface{}{
			"Password": "password",
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("ProtectWorkbook")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("ProtectWorkbook",
		js.ValueOf(map[string]interface{}{
			"Password": strings.Repeat("s", excelize.MaxFieldLength+1),
		}),
	)
	assert.EqualError(t, excelize.ErrPasswordLengthInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("ProtectWorkbook",
		js.ValueOf(map[string]interface{}{
			"Password": true,
		}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestRemoveCol(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("RemoveCol", js.ValueOf("Sheet1"), js.ValueOf("A"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("RemoveCol")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("RemoveCol", js.ValueOf("SheetN"), js.ValueOf("A"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestRemovePageBreak(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("RemovePageBreak", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("RemovePageBreak")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("RemovePageBreak", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestRemoveRow(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("RemoveRow", js.ValueOf("Sheet1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("RemoveRow")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("RemoveRow", js.ValueOf("SheetN"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSearchSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SearchSheet", js.ValueOf("Sheet1"), js.ValueOf("foo"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("SearchSheet", js.ValueOf("Sheet1"), js.ValueOf("foo"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("SearchSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SearchSheet", js.ValueOf("SheetN"), js.ValueOf("foo"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())
}

func TestSetActiveSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetActiveSheet", js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetActiveSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestSetAppProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetAppProps",
		js.ValueOf(map[string]interface{}{
			"Company": "Company Name",
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetAppProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetAppProps",
		js.ValueOf(map[string]interface{}{
			"Application": true,
		}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestSetCellBool(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellBool", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellBool")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellBool", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf(true))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellDefault(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellDefault", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellDefault")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellDefault", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellFloat(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellFloat", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(123.42), js.ValueOf(1), js.ValueOf(64))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellFloat")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellFloat", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf(123.42), js.ValueOf(1), js.ValueOf(64))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellFormula(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellFormula", js.ValueOf("Sheet1"), js.ValueOf("A3"), js.ValueOf("=SUM(A1,B1)"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellFormula", js.ValueOf("Sheet1"), js.ValueOf("A3"), js.ValueOf("=A1+B1"), js.ValueOf(map[string]interface{}{"Type": "shared", "Ref": "C1:C5"}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellFormula")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellFormula", js.ValueOf("SheetN"), js.ValueOf("A3"), js.ValueOf("=SUM(A1,B1)"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellFormula", js.ValueOf("Sheet1"), js.ValueOf("A3"), js.ValueOf("=A1+B1"), js.ValueOf(map[string]interface{}{"Type": true, "Ref": "C1:C5"}))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())
}

func TestSetCellHyperLink(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	display, tooltip := "https://github.com/xuri/excelize-wasm", "excelize-wasm on GitHub"

	ret := f.(js.Value).Call("SetCellHyperLink", js.ValueOf("Sheet1"), js.ValueOf("A3"), js.ValueOf(display), js.ValueOf("External"), js.ValueOf(map[string]interface{}{"Display": display, "Tooltip": tooltip}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellHyperLink")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellHyperLink", js.ValueOf("SheetN"), js.ValueOf("A3"), js.ValueOf(display), js.ValueOf("External"), js.ValueOf(map[string]interface{}{"Display": display, "Tooltip": tooltip}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellHyperLink", js.ValueOf("Sheet1"), js.ValueOf("A3"), js.ValueOf(display), js.ValueOf("External"), js.ValueOf(map[string]interface{}{"Display": true, "Tooltip": tooltip}))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())
}

func TestSetCellInt(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellInt", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellInt")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellInt", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestCellRichText(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	runs := js.ValueOf([]interface{}{
		map[string]interface{}{
			"Text": "bold",
			"Font": map[string]interface{}{
				"Bold":   true,
				"Color":  "2354e8",
				"Family": "Times New Roman",
			},
		},
	})
	ret := f.(js.Value).Call("SetCellRichText", js.ValueOf("Sheet1"), js.ValueOf("A1"), runs)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetCellRichText", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, runs.Length(), ret.Get("runs").Length())
	assert.True(t, ret.Get("runs").Index(0).Get("Font").Get("Bold").Bool())

	ret = f.(js.Value).Call("SetCellRichText")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellRichText", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf([]interface{}{map[string]interface{}{"Text": true}}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellRichText", js.ValueOf("SheetN"), js.ValueOf("A1"), runs)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellRichText")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellRichText", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCellRichText", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellStr(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellStr", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellStr")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellStr", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellStyle", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("B2"), js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellStyle", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("B2"), js.ValueOf(0))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetCellValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A2"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetCellValue")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCellValue", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("foo"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetColOutlineLevel(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetColOutlineLevel", js.ValueOf("Sheet1"), js.ValueOf("A"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetColOutlineLevel")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetColOutlineLevel", js.ValueOf("SheetN"), js.ValueOf("A"), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetColStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetColStyle", js.ValueOf("Sheet1"), js.ValueOf("A"), js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetColStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetColStyle", js.ValueOf("SheetN"), js.ValueOf("A"), js.ValueOf(0))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetColVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetColVisible", js.ValueOf("Sheet1"), js.ValueOf("A"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetColVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetColVisible", js.ValueOf("SheetN"), js.ValueOf("A"), js.ValueOf(true))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetColWidth(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetColWidth", js.ValueOf("Sheet1"), js.ValueOf("A"), js.ValueOf("B"), js.ValueOf(10))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetColWidth")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetColWidth", js.ValueOf("SheetN"), js.ValueOf("A"), js.ValueOf("B"), js.ValueOf(10))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetConditionalFormat(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	condFmt := js.ValueOf([]interface{}{
		map[string]interface{}{
			"Type":     "top",
			"Criteria": "=",
			"Format":   0,
		},
	})
	ret := f.(js.Value).Call("SetConditionalFormat", js.ValueOf("Sheet1"), js.ValueOf("A1:B2"), condFmt)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetConditionalFormat")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetConditionalFormat", js.ValueOf("Sheet1"), js.ValueOf("A1:B2"), js.ValueOf([]interface{}{map[string]interface{}{"Type": true}}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetConditionalFormat", js.ValueOf("SheetN"), js.ValueOf("A1:B2"), condFmt)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestCustomProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	for _, prop := range []interface{}{
		map[string]interface{}{"Name": "Text Prop", "Value": "text"},
		map[string]interface{}{"Name": "Boolean Prop 1", "Value": true},
		map[string]interface{}{"Name": "Boolean Prop 2", "Value": false},
		map[string]interface{}{"Name": "Number Prop 1", "Value": -123.456},
		map[string]interface{}{"Name": "Number Prop 2", "Value": 1},
		map[string]interface{}{"Name": "Number Prop 2", "Value": nil},
	} {
		ret := f.(js.Value).Call("SetCustomProps", js.ValueOf(prop))
		assert.True(t, ret.Get("error").IsNull())
	}

	ret := f.(js.Value).Call("GetCustomProps")
	assert.Equal(t, ret.Get("props").Length(), 4)
	assert.Equal(t, ret.Get("props").Index(0).Get("Value").String(), "text")
	assert.True(t, ret.Get("props").Index(1).Get("Value").Bool())
	assert.False(t, ret.Get("props").Index(2).Get("Value").Bool())
	assert.Equal(t, ret.Get("props").Index(3).Get("Value").Float(), -123.456)

	ret = f.(js.Value).Call("SetCustomProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCustomProps", js.ValueOf(1))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetCustomProps", js.ValueOf(map[string]interface{}{"Name": 1}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCustomProps", js.ValueOf(1))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	// Test get custom property with unsupported charset
	wb := excelize.NewFile()
	wb.Sheet.Delete("docProps/custom.xml")
	wb.Pkg.Store("docProps/custom.xml", MacintoshCyrillicCharset)
	buf, err := wb.WriteToBuffer()
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(buf.Len()))
	for k, v := range buf.Bytes() {
		uint8Array.SetIndex(k, v)
	}
	f = OpenReader(js.Value{}, []js.Value{uint8Array})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	ret = f.(js.Value).Call("GetCustomProps")
	assert.Equal(t, ret.Get("error").String(), "XML syntax error on line 1: invalid UTF-8")

	// Test set custom property with unsupported charset
	ret = f.(js.Value).Call("SetCustomProps", js.ValueOf(map[string]interface{}{"Name": "Text Prop", "Value": "text"}))
	assert.Equal(t, ret.Get("error").String(), "XML syntax error on line 1: invalid UTF-8")
}

func TestSetDefaultFont(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetDefaultFont", js.ValueOf("Arial"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetDefaultFont")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestDefinedName(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetDefinedName", js.ValueOf(map[string]interface{}{
		"Name":     "Amount",
		"RefersTo": "Sheet1!$A$2:$D$5",
		"Comment":  "defined name comment",
	}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetDefinedName")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Amount", ret.Get("definedNames").Index(0).Get("Name").String())

	ret = f.(js.Value).Call("SetDefinedName")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetDefinedName", js.ValueOf(map[string]interface{}{
		"Name":     true,
		"RefersTo": "Sheet1!$A$2:$D$5",
	}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	// Test set defined name without name
	ret = f.(js.Value).Call("SetDefinedName", js.ValueOf(map[string]interface{}{
		"RefersTo": "Sheet1!$A$2:$D$5",
	}))
	assert.EqualError(t, excelize.ErrParameterInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("GetDefinedName", js.ValueOf(true))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestDocProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetDocProps", js.ValueOf(map[string]interface{}{
		"Category": "category",
	}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetDocProps")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "category", ret.Get("props").Get("Category").String())

	ret = f.(js.Value).Call("SetDocProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetDocProps", js.ValueOf(map[string]interface{}{
		"Category": true,
	}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetDocProps", js.ValueOf(nil))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestHeaderFooter(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetHeaderFooter", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"OddHeader": "header"}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetHeaderFooter", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "header", ret.Get("opts").Get("OddHeader").String())

	ret = f.(js.Value).Call("SetHeaderFooter")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetHeaderFooter", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"OddHeader": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	// Test set header and footer with illegal setting
	ret = f.(js.Value).Call("SetHeaderFooter", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"OddHeader": strings.Repeat("c", excelize.MaxFieldLength+1),
		}),
	)
	assert.Equal(t, "field OddHeader must be less than or equal to 255 characters", ret.Get("error").String())

	ret = f.(js.Value).Call("GetHeaderFooter")
	assert.Equal(t, errArgNum.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetHeaderFooter", js.ValueOf(nil))
	assert.Equal(t, errArgType.Error(), ret.Get("error").String())

	ret = f.(js.Value).Call("GetHeaderFooter", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestPageLayout(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetPageLayout", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Size":            1,
			"Orientation":     "landscape",
			"FirstPageNumber": 1,
			"AdjustTo":        120,
			"FitToHeight":     2,
			"FitToWidth":      2,
			"BlackAndWhite":   true,
			"PageOrder":       "overThenDown",
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetPageLayout")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPageLayout", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Size": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPageLayout", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{"Size": 1}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetPageLayout", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	assert.Equal(t, "landscape", ret.Get("opts").Get("Orientation").String())
	assert.Equal(t, 120, ret.Get("opts").Get("AdjustTo").Int())
	assert.True(t, ret.Get("opts").Get("BlackAndWhite").Bool())

	ret = f.(js.Value).Call("GetPageLayout")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPageLayout", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestPageMargins(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetPageMargins", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Bottom":       1.0,
			"Footer":       1.0,
			"Header":       1.0,
			"Left":         1.0,
			"Right":        1.0,
			"Top":          1.0,
			"Horizontally": true,
			"Vertically":   true,
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetPageMargins")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPageMargins", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"Bottom": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPageMargins", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{"Bottom": 1}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetPageMargins", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	assert.Equal(t, 1, ret.Get("opts").Get("Top").Int())
	assert.True(t, ret.Get("opts").Get("Vertically").Bool())

	ret = f.(js.Value).Call("GetPageMargins")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPageMargins", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestPanes(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetPanes", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Freeze": false,
			"Split":  false,
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetPanes", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.False(t, ret.Get("panes").Get("Freeze").Bool())
	assert.False(t, ret.Get("panes").Get("Split").Bool())

	ret = f.(js.Value).Call("SetPanes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPanes", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"Freeze": 0,
		}),
	)
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("GetPanes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPanes", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{
			"Freeze": false,
			"Split":  false,
		}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetPanes", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetRowHeight(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetRowHeight", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(10))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetRowHeight")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetRowHeight", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(10))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetRowOutlineLevel(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetRowOutlineLevel", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetRowOutlineLevel")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetRowOutlineLevel", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetRowStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetRowStyle", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(1), js.ValueOf(0))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetRowStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetRowStyle", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(1), js.ValueOf(0))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetRowVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetRowVisible", js.ValueOf("Sheet1"), js.ValueOf(1), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetRowVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetRowVisible", js.ValueOf("SheetN"), js.ValueOf(1), js.ValueOf(true))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetSheetBackgroundFromBytes(t *testing.T) {
	buf, err := os.ReadFile(filepath.Join("..", "chart.png"))
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(len(buf)))
	for k, v := range buf {
		uint8Array.SetIndex(k, v)
	}

	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetBackgroundFromBytes", js.ValueOf("Sheet1"), js.ValueOf(".png"), js.ValueOf(uint8Array))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetBackgroundFromBytes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetBackgroundFromBytes", js.ValueOf("Sheet1"), js.ValueOf(".images"), js.ValueOf(uint8Array))
	assert.EqualError(t, excelize.ErrImgExt, ret.Get("error").String())
}

func TestSetSheetCol(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetCol", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf([]interface{}{"foo", 1, true, nil}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetCol")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetCol", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf([]interface{}{"foo", 1, true, nil}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetSheetName(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetName", js.ValueOf("Sheet1"), js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetName")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetName", js.ValueOf("Sheet:1"), js.ValueOf("Sheet2"))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetName", js.ValueOf("SheetN"), js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())
}

func TestSheetProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetProps", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{
			"CodeName":                          "code",
			"EnableFormatConditionsCalculation": true,
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetSheetProps", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.True(t, ret.Get("props").Get("EnableFormatConditionsCalculation").Bool())
	assert.Equal(t, "code", ret.Get("props").Get("CodeName").String())

	ret = f.(js.Value).Call("SetSheetProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetProps", js.ValueOf("Sheet1"),
		js.ValueOf(map[string]interface{}{"CodeName": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetProps", js.ValueOf("SheetN"),
		js.ValueOf(map[string]interface{}{"CodeName": "code"}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetProps", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetSheetRow(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetRow", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf([]interface{}{"foo", 1, true, nil}))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetRow")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetRow", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf([]interface{}{"foo", 1, true, nil}))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSheetView(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetView", js.ValueOf("Sheet1"), js.ValueOf(-1),
		js.ValueOf(map[string]interface{}{
			"DefaultGridColor":  false,
			"RightToLeft":       false,
			"ShowFormulas":      false,
			"ShowGridLines":     false,
			"ShowRowColHeaders": false,
			"ShowRuler":         false,
			"ShowZeros":         false,
			"TopLeftCell":       "A1",
			"View":              "normal",
			"ZoomScale":         120,
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetSheetView", js.ValueOf("Sheet1"), js.ValueOf(-1))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 120, ret.Get("opts").Get("ZoomScale").Int())

	ret = f.(js.Value).Call("SetSheetView")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetView", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetView", js.ValueOf("Sheet1"), js.ValueOf(-1),
		js.ValueOf(map[string]interface{}{"View": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetView", js.ValueOf("SheetN"), js.ValueOf(-1),
		js.ValueOf(map[string]interface{}{"View": "normal"}),
	)
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())

	ret = f.(js.Value).Call("GetSheetView", js.ValueOf("SheetN"), js.ValueOf(-1))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetSheetVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetVisible", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetVisible", js.ValueOf("Sheet:1"), js.ValueOf(true))
	assert.EqualError(t, excelize.ErrSheetNameInvalid, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetVisible", js.ValueOf("SheetN"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())
}

func TestWorkbookProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetWorkbookProps",
		js.ValueOf(map[string]interface{}{
			"Date1904":      true,
			"FilterPrivacy": true,
			"CodeName":      "code",
		}),
	)
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetWorkbookProps")
	assert.True(t, ret.Get("error").IsNull())
	assert.True(t, ret.Get("props").Get("Date1904").Bool())
	assert.Equal(t, "code", ret.Get("props").Get("CodeName").String())

	ret = f.(js.Value).Call("SetWorkbookProps")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetWorkbookProps", js.ValueOf(true))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetWorkbookProps",
		js.ValueOf(map[string]interface{}{"CodeName": true}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
}

func TestUngroupSheets(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UngroupSheets")
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UngroupSheets", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestUnmergeCell(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UnmergeCell", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf("B2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UnmergeCell")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("UnmergeCell", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("B2"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestUnprotectSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UnprotectSheet", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UnprotectSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("UnprotectSheet", js.ValueOf("Sheet1"), js.ValueOf("password"))
	assert.Equal(t, "worksheet has set no protect", ret.Get("error").String())
}

func TestUnprotectWorkbook(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UnprotectWorkbook")
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UnprotectWorkbook", js.ValueOf(true))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("UnprotectWorkbook", js.ValueOf("password"))
	assert.Equal(t, "workbook has set no protect", ret.Get("error").String())
}

func TestUnsetConditionalFormat(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UnsetConditionalFormat", js.ValueOf("Sheet1"), js.ValueOf("A1:B2"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UnsetConditionalFormat")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("UnsetConditionalFormat", js.ValueOf("SheetN"), js.ValueOf("A1:B2"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestUpdateLinkedValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("UpdateLinkedValue")
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("UpdateLinkedValue", js.ValueOf("Sheet1"))
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	// Test unsupported charset
	wb := excelize.NewFile()
	wb.Sheet.Delete("xl/worksheets/sheet1.xml")
	wb.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	buf, err := wb.WriteToBuffer()
	assert.NoError(t, err)

	uint8Array := js.Global().Get("Uint8Array").New(js.ValueOf(buf.Len()))
	for k, v := range buf.Bytes() {
		uint8Array.SetIndex(k, v)
	}
	f = OpenReader(js.Value{}, []js.Value{uint8Array})
	assert.True(t, f.(js.Value).Get("error").IsNull())
	ret = f.(js.Value).Call("UpdateLinkedValue")
	assert.Equal(t, ret.Get("error").String(), "XML syntax error on line 1: invalid UTF-8")
}

func TestWriteToBuffer(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("WriteToBuffer")
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(map[string]interface{}{
		"Password": "password",
	}))
	assert.Equal(t, js.TypeObject, ret.Type())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(map[string]interface{}{
		"Password": true,
	}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(true), js.ValueOf(true))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestJsValueToGo(t *testing.T) {
	type T1 struct{}
	type T2 struct {
		F1 []*T1
		F2 []*string
	}
	type T3 struct {
		F1 []*uint32
	}
	type T4 struct {
		F1 []*T3
	}

	_, err := jsValueToGo(js.ValueOf(map[string]interface{}{
		"F1": []interface{}{map[string]interface{}{}},
		"F2": []interface{}{"f2"},
	}), reflect.TypeOf(T2{}))
	assert.NoError(t, err)
	_, err = jsValueToGo(js.ValueOf(map[string]interface{}{
		"F1": true,
		"F2": []interface{}{"f2"},
	}), reflect.TypeOf(T2{}))
	assert.EqualError(t, err, errArgType.Error())
	_, err = jsValueToGo(js.ValueOf(map[string]interface{}{
		"F1": []interface{}{map[string]interface{}{}},
		"F2": true,
	}), reflect.TypeOf(T2{}))
	assert.EqualError(t, err, errArgType.Error())
	_, err = jsValueToGo(js.ValueOf(map[string]interface{}{
		"F1": []interface{}{0},
	}), reflect.TypeOf(T3{}))
	assert.EqualError(t, err, errArgType.Error())
	_, err = jsValueToGo(js.ValueOf(map[string]interface{}{
		"F1": []interface{}{map[string]interface{}{
			"F1": []interface{}{0},
		}},
	}), reflect.TypeOf(T4{}))
	assert.EqualError(t, err, errArgType.Error())
}

func TestJsToGoBaseType(t *testing.T) {
	_, err := jsToGoBaseType(js.ValueOf(0), reflect.Uint)
	assert.NoError(t, err)
	_, err = jsToGoBaseType(js.ValueOf(0), reflect.Int64)
	assert.NoError(t, err)
	_, err = jsToGoBaseType(js.ValueOf(true), reflect.Uint)
	assert.EqualError(t, err, errArgType.Error())
	_, err = jsToGoBaseType(js.ValueOf(true), reflect.Int64)
	assert.EqualError(t, err, errArgType.Error())
}

func TestGoValueToJS(t *testing.T) {
	enable, exp := true, "exp"
	result, err := goValueToJS(reflect.ValueOf(excelize.Chart{
		Format: excelize.GraphicOptions{PrintObject: &enable},
	}), reflect.TypeOf(excelize.Chart{}))
	assert.NoError(t, err)
	assert.True(t, js.ValueOf(result).Get("Format").Get("PrintObject").Bool())

	type T1 struct {
		F1 []*excelize.DataValidation
		F2 []*int64
		F3 []uint
	}
	var num int64 = 1
	result, err = goValueToJS(reflect.ValueOf(T1{
		F1: []*excelize.DataValidation{{AllowBlank: true}, {}},
		F2: []*int64{&num},
		F3: []uint{1},
	}), reflect.TypeOf(T1{}))
	assert.NoError(t, err)
	assert.True(t, js.ValueOf(result).Get("F1").Index(0).Get("AllowBlank").Bool())
	assert.Equal(t, 1, js.ValueOf(result).Get("F2").Index(0).Int())
	assert.Equal(t, 1, js.ValueOf(result).Get("F3").Index(0).Int())

	result, err = goValueToJS(reflect.ValueOf(excelize.Style{
		NumFmt:       1,
		CustomNumFmt: &exp,
		Alignment:    &excelize.Alignment{Indent: 1},
		Border:       []excelize.Border{{Type: "left"}, {Type: "top"}},
	}), reflect.TypeOf(excelize.Style{}))
	assert.NoError(t, err)
	assert.Equal(t, 1, js.ValueOf(result).Get("NumFmt").Int())
	assert.Equal(t, exp, js.ValueOf(result).Get("CustomNumFmt").String())
	assert.Equal(t, 1, js.ValueOf(result).Get("Alignment").Get("Indent").Int())
	assert.Equal(t, "left", js.ValueOf(result).Get("Border").Index(0).Get("Type").String())
	assert.Equal(t, "top", js.ValueOf(result).Get("Border").Index(1).Get("Type").String())

	type T2 struct{ F1 string }
	type T3 struct{ F1 bool }
	_, err = goValueToJS(reflect.ValueOf(T2{
		F1: "foo",
	}), reflect.TypeOf(T3{}))
	assert.EqualError(t, err, errArgType.Error())

	type T4 struct{ F1 *T2 }
	type T5 struct{ F1 *T3 }
	_, err = goValueToJS(reflect.ValueOf(T4{
		F1: &T2{F1: "foo"},
	}), reflect.TypeOf(T5{}))
	assert.EqualError(t, err, errArgType.Error())

	type T6 struct{ F1 *bool }
	type T7 struct{ F1 *string }
	_, err = goValueToJS(reflect.ValueOf(T6{
		F1: &enable,
	}), reflect.TypeOf(T7{}))
	assert.EqualError(t, err, errArgType.Error())

	type T8 struct{ F1 T6 }
	type T9 struct{ F1 T7 }
	_, err = goValueToJS(reflect.ValueOf(T8{
		F1: T6{F1: &enable},
	}), reflect.TypeOf(T9{}))
	assert.EqualError(t, err, errArgType.Error())

	type T10 struct{ F1 []*T2 }
	type T11 struct{ F1 []*T3 }
	_, err = goValueToJS(reflect.ValueOf(T10{
		F1: []*T2{{F1: "foo"}},
	}), reflect.TypeOf(T11{}))
	assert.EqualError(t, err, errArgType.Error())

	type T12 struct{ F1 []*string }
	type T13 struct{ F1 []*bool }
	_, err = goValueToJS(reflect.ValueOf(T12{
		F1: []*string{&exp},
	}), reflect.TypeOf(T13{}))
	assert.EqualError(t, err, errArgType.Error())

	type T14 struct{ F1 []T2 }
	type T15 struct{ F1 []T3 }
	_, err = goValueToJS(reflect.ValueOf(T14{
		F1: []T2{{F1: "foo"}},
	}), reflect.TypeOf(T15{}))
	assert.EqualError(t, err, errArgType.Error())

	type T16 struct{ F1 []string }
	type T17 struct{ F1 []bool }
	_, err = goValueToJS(reflect.ValueOf(T16{
		F1: []string{exp},
	}), reflect.TypeOf(T17{}))
	assert.EqualError(t, err, errArgType.Error())

	type T18 struct{ F1 uint8 }
	_, err = goValueToJS(reflect.ValueOf(T16{
		F1: []string{exp},
	}), reflect.TypeOf(T18{}))
	assert.EqualError(t, err, errArgType.Error())
}

func TestGoBaseTypeToJS(t *testing.T) {
	for _, typ := range []reflect.Kind{reflect.Bool, reflect.Bool, reflect.Int64} {
		_, err := goBaseTypeToJS(reflect.ValueOf(0), typ)
		assert.EqualError(t, err, errArgType.Error())
	}
	for _, typ := range []reflect.Kind{
		reflect.Uint, reflect.Uint64, reflect.Int, reflect.Int64,
		reflect.Float64, reflect.String, reflect.Complex128,
	} {
		_, err := goBaseTypeToJS(reflect.ValueOf(true), typ)
		assert.EqualError(t, err, errArgType.Error())
	}
}
