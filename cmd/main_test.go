package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"math/rand"
	"os"
	"path/filepath"
	"syscall/js"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize"
)

func TestRegInteropFunc(t *testing.T) {
	regFuncs()
}

func TestInTypeSlice(t *testing.T) {
	assert.Equal(t, -1, inTypeSlice(nil, js.TypeBoolean))
	assert.Equal(t, 0, inTypeSlice([]js.Type{js.TypeBoolean}, js.TypeBoolean))
}

func TestPrepareOptions(t *testing.T) {
	opts, err := prepareOptions(js.ValueOf(map[string]interface{}{
		"password":       "passwd",
		"raw_cell_value": true}))
	assert.NoError(t, err)
	assert.Equal(t, excelize.Options{Password: "passwd", RawCellValue: true}, opts)

	_, err = prepareOptions(js.ValueOf(map[string]interface{}{"password": false}))
	assert.ErrorIs(t, err, errArgType)

	_, err = prepareOptions(js.ValueOf(map[string]interface{}{"raw_cell_value": "true"}))
	assert.ErrorIs(t, err, errArgType)
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

	f = NewFile(js.Value{}, []js.Value{js.ValueOf(true)})
	assert.EqualError(t, errArgNum, f.(js.Value).Get("error").String())
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
		"password": false,
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
		"password": "invalid",
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

	ret := f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(`{"type":"col3DClustered","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(`{"type":"line","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`), js.ValueOf(`{"type":"col","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChart")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChart", js.ValueOf("SheetN"), js.ValueOf("A1"), js.ValueOf("{}"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAddChartSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet2"), js.ValueOf(`{"type":"col3DClustered","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet3"), js.ValueOf(`{"type":"line","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`), js.ValueOf(`{"type":"col","series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddChartSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddChartSheet", js.ValueOf("Sheet1"), js.ValueOf("{}"))
	assert.Equal(t, "the same name worksheet already exists", ret.Get("error").String())
}

func TestAddComment(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddComment", js.ValueOf("Sheet1"), js.ValueOf("A30"), js.ValueOf(`{"author":"Excelize: ","text":"This is a comment."}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddComment")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddComment", js.ValueOf("SheetN"), js.ValueOf("A30"), js.ValueOf(`{"author":"Excelize: ","text":"This is a comment."}`))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
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

	ret := f.(js.Value).Call("AddPictureFromBytes", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(""), js.ValueOf("Picture 1"), js.ValueOf(".png"), js.ValueOf(uint8Array))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddPictureFromBytes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddPictureFromBytes", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(""), js.ValueOf("Picture 1"), js.ValueOf("png"), js.ValueOf(uint8Array))
	assert.EqualError(t, excelize.ErrImgExt, ret.Get("error").String())
}

func TestAddPivotTable(t *testing.T) {
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
	opts, err := json.Marshal(excelize.PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$2:$M$34",
		Rows:            []excelize.PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Filter:          []excelize.PivotTableField{{Data: "Region"}},
		Columns:         []excelize.PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []excelize.PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Summarize by Sum"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
		ShowError:       true,
	})
	assert.NoError(t, err)
	ret = f.(js.Value).Call("AddPivotTable", js.ValueOf(string(opts)))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddPivotTable")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddPivotTable", js.ValueOf(""))
	assert.Equal(t, "unexpected end of JSON input", ret.Get("error").String())

	ret = f.(js.Value).Call("AddPivotTable", js.ValueOf("{}"))
	assert.Equal(t, "parameter 'PivotTableRange' parsing error: parameter is required", ret.Get("error").String())
}

func TestAddShape(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddShape", js.ValueOf("Sheet1"), js.ValueOf("C30"), js.ValueOf(`{"type":"rect","paragraph":[]}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddShape")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddShape", js.ValueOf("Sheet1"), js.ValueOf("C30"), js.ValueOf(""))
	assert.Equal(t, "unexpected end of JSON input", ret.Get("error").String())
}

func TestAddTable(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AddTable", js.ValueOf("Sheet1"), js.ValueOf("B26"), js.ValueOf("A21"), js.ValueOf("{}"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AddTable")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AddTable", js.ValueOf("SheetN"), js.ValueOf("B26"), js.ValueOf("A21"), js.ValueOf("{}"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestAutoFilter(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("AutoFilter", js.ValueOf("Sheet1"), js.ValueOf("D4"), js.ValueOf("B1"), js.ValueOf(""))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("AutoFilter")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("AutoFilter", js.ValueOf("SheetN"), js.ValueOf("D4"), js.ValueOf("B1"), js.ValueOf(""))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestCalcCellValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("CalcCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "", ret.Get("value").String())

	ret = f.(js.Value).Call("CalcCellValue")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("CalcCellValue", js.ValueOf("SheetN"), js.ValueOf("A1"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
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

func TestGetAppProps(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetAppProps")
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, "Go Excelize", ret.Get("application").String())

	ret = f.(js.Value).Call("GetAppProps", js.ValueOf(1))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, "", ret.Get("application").String())
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

func TestGetCellValue(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("GetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"))
	assert.True(t, ret.Get("error").IsNull())
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

	ret = f.(js.Value).Call("GetCols", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"raw_cell_value": true}))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"raw_cell_value": "true"}))
	assert.EqualError(t, errArgType, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetCols")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("GetCols", js.ValueOf("SheetN"))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("result").Length())
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

func TestGetRows(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetCellValue", js.ValueOf("Sheet1"), js.ValueOf("A1"), js.ValueOf(1))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"raw_cell_value": true}))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("result").Length())

	ret = f.(js.Value).Call("GetRows", js.ValueOf("Sheet1"), js.ValueOf(map[string]interface{}{"raw_cell_value": "true"}))
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

func TestNewConditionalStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewConditionalStyle", js.ValueOf(`{"fill":{"type":"pattern","color":["#FEEAA0"],"pattern":1}}`))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 0, ret.Get("style").Int())

	ret = f.(js.Value).Call("NewConditionalStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("NewConditionalStyle", js.ValueOf(""))
	assert.Equal(t, "unexpected end of JSON input", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("style").Int())
}

func TestNewSheet(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewSheet", js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("index").Int())

	ret = f.(js.Value).Call("NewSheet")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("index").Int())
}

func TestNewStyle(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("NewStyle", js.ValueOf(`{"number_format":1}`))
	assert.True(t, ret.Get("error").IsNull())
	assert.Equal(t, 1, ret.Get("style").Int())

	ret = f.(js.Value).Call("NewStyle")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("NewStyle", js.ValueOf(""))
	assert.Equal(t, "unexpected end of JSON input", ret.Get("error").String())
	assert.Equal(t, 0, ret.Get("style").Int())
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

	ret := f.(js.Value).Call("SetConditionalFormat", js.ValueOf("Sheet1"), js.ValueOf("A1:B2"), js.ValueOf(`[{"type":"top","criteria":"=","format":0}]`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetConditionalFormat")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetConditionalFormat", js.ValueOf("SheetN"), js.ValueOf("A1:B2"), js.ValueOf(`[{"type":"top","criteria":"=","format":0}]`))
	assert.Equal(t, "sheet SheetN does not exist", ret.Get("error").String())
}

func TestSetDefaultFont(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetDefaultFont", js.ValueOf("Arial"))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetDefaultFont")
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}

func TestSetPanes(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetPanes", js.ValueOf("Sheet1"), js.ValueOf(`{"freeze":false,"split":false}`))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetPanes")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetPanes", js.ValueOf("SheetN"), js.ValueOf(`{"freeze":false,"split":false}`))
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

	ret = f.(js.Value).Call("SetSheetName", js.ValueOf("SheetN"), js.ValueOf("Sheet2"))
	assert.True(t, ret.Get("error").IsNull())
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

func TestSetSheetVisible(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("SetSheetVisible", js.ValueOf("Sheet1"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())

	ret = f.(js.Value).Call("SetSheetVisible")
	assert.EqualError(t, errArgNum, ret.Get("error").String())

	ret = f.(js.Value).Call("SetSheetVisible", js.ValueOf("SheetN"), js.ValueOf(true))
	assert.True(t, ret.Get("error").IsNull())
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
}

func TestWriteToBuffer(t *testing.T) {
	f := NewFile(js.Value{}, []js.Value{})
	assert.True(t, f.(js.Value).Get("error").IsNull())

	ret := f.(js.Value).Call("WriteToBuffer")
	assert.Equal(t, js.TypeObject, ret.Type())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(map[string]interface{}{
		"password": "password",
	}))
	assert.Equal(t, js.TypeObject, ret.Type())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(map[string]interface{}{
		"password": true,
	}))
	assert.EqualError(t, errArgType, ret.Get("error").String())

	ret = f.(js.Value).Call("WriteToBuffer", js.ValueOf(true), js.ValueOf(true))
	assert.EqualError(t, errArgNum, ret.Get("error").String())
}
