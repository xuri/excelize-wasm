// Type definitions for excelize-wasm

declare module 'excelize-wasm' {

  /**
   * AppProperties directly maps the document application properties.
   */
  export type AppProperties = {
    application?:        string;
    scale_crop?:         boolean;
    doc_security?:       number;
    company?:            string;
    links_up_to_date?:   boolean;
    hyperlinks_changed?: boolean;
    app_version?:        string;
    error?:              string | null;
  };

  /**
   * Options define the options for open and reading spreadsheet.
   */
  export type Options = {
    max_calc_iterations?:  number;
    password?:             string;
    raw_cell_value?:       boolean;
    unzip_size_limit?:     number;
    unzip_xml_size_limit?: number;
  };

  /**
   * Border directly maps the border settings of the cells.
   */
  export type Border = {
    type?:  string;
    color?: string;
    style?: number;
  };

  /**
   * Fill directly maps the fill settings of the cells.
   */
  export type Fill = {
    type?:    string;
    pattern?: number;
    color?:   string[];
    shading?: number;
  };

  /**
   * Font directly maps the font settings of the fonts.
   */
  export type Font = {
    bold?:       boolean;
    italic?:     boolean;
    underline?:  string;
    family?:     string;
    size?:       number;
    strike?:     boolean;
    color?:      string;
    vert_align?: string;
  };

  /**
   * Alignment directly maps the alignment settings of the cells.
   */
  export type Alignment = {
    horizontal?:        string;
    indent?:            number;
    justify_last_line?: boolean;
    reading_order?:     number;
    relative_indent?:   number;
    shrink_to_fit?:     boolean;
    text_rotation?:     number;
    vertical?:          string;
    wrap_text?:         boolean;
  };

  /**
   * Protection directly maps the protection settings of the cells.
   */
  export type Protection = {
    hidden?: boolean;
    locked?: boolean;
  };

  /**
   * Style directly maps the style settings of the cells.
   */
  export type Style = {
    border?:               Border[];
    fill?:                 Fill;
    font?:                 Font;
    alignment?:            Alignment;
    protection?:           Protection;
    number_format?:        number;
    decimal_places?:       number;
    custom_number_format?: string;
    lang?:                 string;
    negred?:               boolean;
  }

  /**
   * CellNameToCoordinates converts alphanumeric cell name to [X, Y]
   * coordinates or returns an error.
   * @param cell The cell reference
   */
  export function CellNameToCoordinates(cell: string): { col: number, row: number, error: string | null }

  /**
   * ColumnNameToNumber provides a function to convert Excel sheet column name
   * (case-insensitive) to int. The function returns an error if column name
   * incorrect.
   * @param name The column name
   */
  export function ColumnNameToNumber(name: string): { col: number, error: string | null }

  /**
   * ColumnNumberToName provides a function to convert the integer to Excel
   * sheet column title.
   * @param num The column name
   */
  export function ColumnNumberToName(num: number): { col: string, error: string | null }

  /**
   * CoordinatesToCellName converts [X, Y] coordinates to alpha-numeric cell
   * name or returns an error.
   * @param col The column number
   * @param row The row number
   * @param abs Specifies the absolute cell references
   */
  export function CoordinatesToCellName(col: number, row: number, abs?: boolean): { cell: string, error: string | null }

  /**
   * HSLToRGB converts an HSL triple to a RGB triple.
   * @param h Hue
   * @param s Saturation
   * @param l Lightness
   */
  export function HSLToRGB(h: number, s: number, l: number): { r: number, g: number, b: number, error: string | null }

  /**
   * JoinCellName joins cell name from column name and row number.
   * @param col The column name
   * @param row The row number
   */
  export function JoinCellName(col: string, row: number): { cell: string, error: string | null }

  /**
   * RGBToHSL converts an RGB triple to a HSL triple.
   * @param r Red
   * @param g Green
   * @param b Blue
   */
  export function RGBToHSL(r: number, g: number, b: number): { h: number, s: number, l: number, error: string | null }

  /**
   * SplitCellName splits cell name to column name and row number.
   * @param cell The cell reference
   */
  export function SplitCellName(cell: string): { col: string, row: number, error: string | null }

  /**
   * ThemeColor applied the color with tint value.
   * @param baseColor Base color in hex format
   * @param tint A mixture of a color with white
   */
  export function ThemeColor(baseColor: string, tint: number): { color: string, error: string | null }

  /**
   * NewFile provides a function to create new file by default template.
   */
  export function NewFile(): NewFile;

  /**
   * OpenReader read data stream from buffer and return a populated spreadsheet
   * file.
   * @param r The contents buffer of the file
   * @param opts The options for open and reading spreadsheet
   */
  export function OpenReader(r: Uint8Array[], opts?: Options): NewFile;

  /**
   * @constructor
   */
  export class NewFile {
    /**
     * AddChart provides the method to add chart in a sheet by given chart
     * format set (such as offset, scale, aspect ratio setting and print
     * settings) and properties set.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param opts The chart options
     * @param combo Specifies the create a chart that combines two or more
     *  chart types in a single chart
     */
    AddChart(sheet: string, cell: string, opts: string, combo?: string): { error: string | null }

    /**
     * AddChartSheet provides the method to create a chartsheet by given chart
     * format set (such as offset, scale, aspect ratio setting and print
     * settings) and properties set. In Excel a chartsheet is a worksheet that
     * only contains a chart.
     * @param sheet The worksheet name
     * @param opts The chart options
     * @param combo Specifies the create a chart that combines two or more
     *  chart types in a single chart
     */
    AddChartSheet(sheet: string, opts: string, combo?: string): { error: string | null }

    /**
     * AddComment provides the method to add comment in a sheet by given
     * worksheet index, cell and format set (such as author and text). Note
     * that the max author length is 255 and the max text length is 32512.
     * @param sheet The worksheet name
     * @param opts The comment options
     */
    AddComment(sheet: string, opts: string): { error: string | null }

    /**
     * AddPictureFromBytes provides the method to add picture in a sheet by
     * given picture format set (such as offset, scale, aspect ratio setting
     * and print settings), file base name, extension name and file bytes.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param opts The picture options
     * @param name The picture name
     * @param extension The extension name
     * @param file The contents buffer of the file
     */
    AddPictureFromBytes(sheet: string, cell: string, opts: string, name: string, extension: string, file: Uint8Array[]): { error: string | null }

    /**
     * AddPivotTable provides the method to add pivot table by given pivot
     * table options. Note that the same fields can not in Columns, Rows and
     * Filter fields at the same time.
     * @param opt The pivot table option
     */
    AddPivotTable(opt: string): { error: string | null }

    /**
     * AddShape provides the method to add shape in a sheet by given worksheet
     * index, shape format set (such as offset, scale, aspect ratio setting
     * and print settings) and properties set.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param opts The shape options
     */
    AddShape(sheet: string, cell: string, opts: string): { error: string | null }

    /**
     * AddTable provides the method to add table in a worksheet by given
     * worksheet name, range reference and format set.
     * @param sheet The worksheet name
     * @param hCell The top-left cell reference
     * @param vCell The right-bottom cell reference
     * @param opts The table options
     */
    AddTable(sheet: string, hCell: string, vCell: string, opts: string): { error: string | null }

    /**
     * AutoFilter provides the method to add auto filter in a worksheet by
     * given worksheet name, range reference and settings. An auto filter in
     * Excel is a way of filtering a 2D range of data based on some simple
     * criteria.
     * @param sheet The worksheet name
     * @param hCell The top-left cell reference
     * @param vCell The right-bottom cell reference
     * @param opts The auto filter options
     */
    AutoFilter(sheet: string, hCell: string, vCell: string, opts: string): { error: string | null }

    /**
     * CalcCellValue provides a function to get calculated cell value. This
     * feature is currently in working processing. Iterative calculation,
     * implicit intersection, explicit intersection, array formula, table
     * formula and some other formulas are not supported currently.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    CalcCellValue(sheet: string, cell: string): { value: string, error: string | null }

    /**
     * CopySheet provides a function to duplicate a worksheet by gave source
     * and target worksheet index. Note that currently doesn't support
     * duplicate workbooks that contain tables, charts or pictures.
     * @param from Source sheet index
     * @param to Target sheet index
     */
    CopySheet(from: number, to: number): { error: string | null }

    /**
     * DeleteChart provides a function to delete chart in spreadsheet by given
     * worksheet name and cell reference.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    DeleteChart(sheet: string, cell: string): { error: string | null }

    /**
     * DeleteComment provides the method to delete comment in a sheet by given
     * worksheet name.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    DeleteComment(sheet: string, cell: string): { error: string | null }

    /**
     * DeleteDataValidation delete data validation by given worksheet name and
     * reference sequence. All data validations in the worksheet will be
     * deleted if not specify reference sequence parameter.
     * @param sheet The worksheet name
     * @param sqref The cell reference sequence
     */
    DeleteDataValidation(sheet: string, sqref?: string): { error: string | null }

    /**
     * DeletePicture provides a function to delete charts in spreadsheet by
     * given worksheet name and cell reference. Note that the image file won't
     * be deleted from the document currently.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    DeletePicture(sheet: string, cell: string): { error: string | null }

    /**
     * DeleteSheet provides a function to delete worksheet in a workbook by
     * given worksheet name. Use this method with caution, which will affect
     * changes in references such as formulas, charts, and so on. If there is
     * any referenced value of the deleted worksheet, it will cause a file
     * error when you open it. This function will be invalid when only one
     * worksheet is left.
     * @param sheet The worksheet name
     */
    DeleteSheet(sheet: string): { error: string | null }

    /**
     * DuplicateRow inserts a copy of specified row (by its Excel row number)
     * below. Use this method with caution, which will affect changes in
     * references such as formulas, charts, and so on. If there is any
     * referenced value of the worksheet, it will cause a file error when you
     * open it. The excelize only partially updates these references
     * currently.
     * @param sheet The worksheet name
     * @param row The row number
     */
    DuplicateRow(sheet: string, row: number): { error: string | null }

    /**
     * DuplicateRowTo inserts a copy of specified row by it Excel number to
     * specified row position moving down exists rows after target position.
     * Use this method with caution, which will affect changes in references
     * such as formulas, charts, and so on. If there is any referenced value
     * of the worksheet, it will cause a file error when you open it. The
     * excelize only partially updates these references currently.
     * @param sheet The worksheet name
     * @param row The source row number
     * @param row2 The target row number
     */
    DuplicateRowTo(sheet: string, row: number, row2: number): { error: string | null }

    /**
     * GetActiveSheetIndex provides a function to get active sheet index of the
     * spreadsheet. If not found the active sheet will be return integer 0.
     */
    GetActiveSheetIndex(): { index: number, error: string | null }

    /**
     * GetAppProps provides a function to get document application properties.
     * @return This is the document application properties.
     */
    GetAppProps(): AppProperties;

    /**
     * GetCellFormula provides a function to get formula from cell by given
     * worksheet name and cell reference in spreadsheet.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    GetCellFormula(sheet: string, cell: string): { formula: string, error: string | null }

    /**
     * GetCellHyperLink gets a cell hyperlink based on the given worksheet name
     * and cell reference. If the cell has a hyperlink, it will return 'true'
     * and the link address, otherwise it will return 'false' and an empty
     * link address.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    GetCellHyperLink(sheet: string, cell: string): { ok: boolean, location: string, error: string | null }

    /**
     * GetCellStyle provides a function to get cell style index by given
     * worksheet name and cell reference.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    GetCellStyle(sheet: string, cell: string): { style: number, error: string | null }

    /**
     * GetCellValue provides a function to get formatted value from cell by
     * given worksheet name and cell reference in spreadsheet. The return value
     * is converted to the 'string' data type. If the cell format can be
     * applied to the value of a cell, the applied value will be returned,
     * otherwise the original value will be returned. All cells' values will be
     * the same in a merged range.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    GetCellValue(sheet: string, cell: string): { value: string, error: string | null }

    /**
     * GetColOutlineLevel provides a function to get outline level of a single
     * column by given worksheet name and column name.
     * @param sheet The worksheet name
     * @param col The column name
     */
    GetColOutlineLevel(sheet: string, col: string): { level: number, error: string | null }

    /**
     * GetColStyle provides a function to get column style ID by given
     * worksheet name and column name.
     * @param sheet The worksheet name
     * @param col The column name
     */
    GetColStyle(sheet: string, col: string): { style: number, error: string | null }

    /**
     * GetColVisible provides a function to get visible of a single column by
     * given worksheet name and column name.
     * @param sheet The worksheet name
     * @param col The column name
     */
    GetColVisible(sheet: string, col: string): { visible: boolean, error: string | null }

    /**
     * GetColWidth provides a function to get column width by given worksheet
     * name and column name.
     * @param sheet The worksheet name
     * @param col The column name
     */
    GetColWidth(sheet: string, col: string): { width: number, error: string | null }

    /**
     * GetCols gets the value of all cells by columns on the worksheet based on
     * the given worksheet name, returned as a two-dimensional array, where
     * the value of the cell is converted to the `string` type. If the cell
     * format can be applied to the value of the cell, the applied value will
     * be used, otherwise the original value will be used.
     * @param sheet The worksheet name
     * @param opts
     */
    GetCols(sheet: string, opts?: Options): { result: string[][], error: string | null }

    /**
     * GetRowHeight provides a function to get row height by given worksheet
     * name and row number.
     * @param sheet The worksheet name
     * @param row The row number
     */
    GetRowHeight(sheet: string, row: number): { height: number, error: string | null }

    /**
     * GetRowOutlineLevel provides a function to get outline level number of a
     * single row by given worksheet name and Excel row number.
     * @param sheet The worksheet name
     * @param row The row number
     */
    GetRowOutlineLevel(sheet: string, row: number): { level: number, error: string | null }

    /**
     * GetRowVisible provides a function to get visible of a single row by
     * given worksheet name and Excel row number.
     * @param sheet The worksheet name
     * @param row The row number
     */
    GetRowVisible(sheet: string, row: number): { visible: boolean, error: string | null }

    /**
     * GetRows return all the rows in a sheet by given worksheet name, returned
     * as a two-dimensional array, where the value of the cell is converted to
     * the string type. If the cell format can be applied to the value of the
     * cell, the applied value will be used, otherwise the original value will
     * be used. GetRows fetched the rows with value or formula cells, the
     * continually blank cells in the tail of each row will be skipped, so the
     * length of each row may be inconsistent.
     * @param sheet The worksheet name
     * @param opts The options for get rows
     */
    GetRows(sheet: string, opts?: Options): { result: string[][], error: string | null }

    /**
     * GetSheetIndex provides a function to get a sheet index of the workbook
     * by the given sheet name. If the given sheet name is invalid or sheet
     * doesn't exist, it will return an integer type value -1.
     * @param sheet The worksheet name
     */
    GetSheetIndex(sheet: string): { index: number, error: string | null }

    /**
     * GetSheetList provides a function to get worksheets, chart sheets, and
     * dialog sheets name list of the workbook.
     */
    GetSheetList(): { list: string[] }

    /**
     * GetSheetMap provides a function to get worksheets, chart sheets, dialog
     * sheets ID and name map of the workbook.
     */
    GetSheetMap(): { sheets: Map<string,string>, error: string | null }

    /**
     * GetSheetName provides a function to get the sheet name of the workbook
     * by the given sheet index. If the given sheet index is invalid, it will
     * return an empty string.
     * @param index The sheet index
     */
    GetSheetName(index: number): { name: string, error: string | null }

    /**
     * GetSheetVisible provides a function to get worksheet visible by given
     * worksheet name.
     * @param sheet The worksheet name
     */
    GetSheetVisible(sheet: string): { visible: boolean, error: string | null }

    /**
     * GroupSheets provides a function to group worksheets by given worksheets
     * name. Group worksheets must contain an active worksheet.
     * @param sheets The worksheet names
     */
    GroupSheets(sheets: string[]): { error: string | null }

    /**
     * InsertCols provides a function to insert new columns before the given
     * column name and number of columns.
     *
     * Use this method with caution, which will affect changes in references
     * such as formulas, charts, and so on. If there is any referenced value
     * of the worksheet, it will cause a file error when you open it. The
     * excelize only partially updates these references currently.
     * @param sheet The worksheet name
     * @param col The base column name
     * @param n The insert columns count
     */
    InsertCols(sheet: string, col: string, n: number): { error: string | null }

    /**
     * InsertPageBreak create a page break to determine where the printed page
     * ends and where begins the next one by given worksheet name and cell, so
     * the content before the page break will be printed on one page and after
     * the page break on another.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    InsertPageBreak(sheet: string, cell: string): { error: string | null }

    /**
     * InsertRows provides a function to insert new rows after the given Excel
     * row number starting from 1 and number of rows.
     *
     * Use this method with caution, which will affect changes in references
     * such as formulas, charts, and so on. If there is any referenced value
     * of the worksheet, it will cause a file error when you open it. The
     * excelize only partially updates these references currently.
     * @param sheet The worksheet name
     * @param row The base row number
     * @param n Insert rows count
     */
    InsertRows(sheet: string, row: number, n: number): { error: string | null }

    /**
     * MergeCell provides a function to merge cells by given range reference
     * and sheet name. Merging cells only keeps the upper-left cell value, and
     * discards the other values.
     *
     * If you create a merged cell that overlaps with another existing merged
     * cell, those merged cells that already exist will be removed. The cell
     * references tuple after merging in the following range will be: A1
     * (x3,y1) D1(x2,y1) A8(x3,y4) D8(x2,y4)
     *
     *                  B1(x1,y1)      D1(x2,y1)
     *                +------------------------+
     *                |                        |
     *      A4(x3,y3) |    C4(x4,y3)           |
     *     +------------------------+          |
     *     |          |             |          |
     *     |          |B5(x1,y2)    | D5(x2,y2)|
     *     |          +------------------------+
     *     |                        |
     *     |A8(x3,y4)      C8(x4,y4)|
     *     +------------------------+
     * @param sheet The worksheet name
     * @param hCell The top-left cell reference
     * @param vCell The right-bottom cell reference
     */
    MergeCell(sheet: string, hCell: string, vCell: string): { error: string | null }

    /**
     * NewConditionalStyle provides a function to create style for conditional
     * format by given style format. The parameters are the same with the
     * NewStyle function. Note that the color field uses RGB color code and
     * only support to set font, fills, alignment and borders currently.
     * @param style
     */
    NewConditionalStyle(style: string): { style: number, error: string | null }

    /**
     * NewSheet provides the function to create a new sheet by given a
     * worksheet name and returns the index of the sheets in the workbook
     * after it appended. Note that when creating a new workbook, the default
     * worksheet named `Sheet1` will be created.
     * @param sheet The worksheet name
     */
    NewSheet(sheet: string): { index: number, error: string | null }

    /**
     * NewStyle provides a function to create the style for cells by given
     * options. Note that the color field uses RGB color code.
     * @param style The style options
     */
    NewStyle(style: Style): { style: number, error: string | null }

    /**
     * RemoveCol provides a function to remove single column by given worksheet
     * name and column index.
     *
     * Use this method with caution, which will affect changes in references
     * such as formulas, charts, and so on. If there is any referenced value
     * of the worksheet, it will cause a file error when you open it. The
     * excelize only partially updates these references currently.
     * @param sheet The worksheet name
     * @param col The column name
     */
    RemoveCol(sheet: string, col: string): { error: string | null }

    /**
     * RemovePageBreak remove a page break by given worksheet name and cell
     * reference.
     * @param sheet The worksheet name
     * @param cell The cell reference
     */
    RemovePageBreak(sheet: string, cell: string): { error: string | null }

    /**
     * RemoveRow provides a function to remove single row by given worksheet
     * name and Excel row number.
     *
     * Use this method with caution, which will affect changes in references
     * such as formulas, charts, and so on. If there is any referenced value
     * of the worksheet, it will cause a file error when you open it. The
     * excelize only partially updates these references currently.
     * @param sheet The worksheet name
     * @param row The row number
     */
    RemoveRow(sheet: string, row: number): { error: string | null }

    /**
     * SearchSheet provides a function to get cell reference by given worksheet
     * name, cell value, and regular expression. The function doesn't support
     * searching on the calculated result, formatted numbers and conditional
     * lookup currently. If it is a merged cell, it will return the cell
     * reference of the upper left cell of the merged range reference.
     * @param sheet The worksheet name
     * @param value The cell value to search
     * @param reg Specifies if search with regular expression
     */
    SearchSheet(sheet: string, value: string, reg?: boolean): { result: string[], error: string | null }

    /**
     * SetActiveSheet provides a function to set the default active sheet of
     * the workbook by a given index. Note that the active index is different
     * from the ID returned by function GetSheetMap(). It should be greater
     * than or equal to 0 and less than the total worksheet numbers.
     * @param index The sheet index
     */
    SetActiveSheet(index: number): { error: string | null }

    /**
     * SetCellBool provides a function to set bool type value of a cell by
     * given worksheet name, cell reference and cell value.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     */
    SetCellBool(sheet: string, cell: string, value: boolean): { error: string | null }

    /**
     * SetCellDefault provides a function to set string type value of a cell as
     * default format without escaping the cell.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     */
    SetCellDefault(sheet: string, cell: string, value: string): { error: string | null }

    /**
     * SetCellFloat sets a floating point value into a cell. The precision
     * parameter specifies how many places after the decimal will be shown
     * while -1 is a special value that will use as many decimal places as
     * necessary to represent the number. bitSize is 32 or 64 depending on if
     * a float32 or float64 was originally used for the value.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     * @param precision Specifies how many places after the decimal will be
     *  shown
     * @param bitSize BitSize is 32 or 64 depending on if a float32 or float64
     *  was originally used for the value
     */
    SetCellFloat(sheet: string, cell: string, value: number, precision: number, bitSize: number): { error: string | null }

    /**
     * SetCellInt provides a function to set int type value of a cell by given
     * worksheet name, cell reference and cell value.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     */
    SetCellInt(sheet: string, cell: string, value: number): { error: string | null }

    /**
     * SetCellStr provides a function to set string type value of a cell. Total
     * number of characters that a cell can contain 32767 characters.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     */
    SetCellStr(sheet: string, cell: string, value: string): { error: string | null }

    /**
     * SetCellStyle provides a function to add style attribute for cells by
     * given worksheet name, range reference and style ID. Note that
     * diagonalDown and diagonalUp type border should be use same color in the
     * same range. SetCellStyle will overwrite the existing styles for the
     * cell, it won't append or merge style with existing styles.
     * @param sheet The worksheet name
     * @param hCell The top-left cell reference
     * @param vCell The right-bottom cell reference
     * @param styleID The style ID
     */
    SetCellStyle(sheet: string, hCell: string, vCell: string, styleID: number): { error: string | null }

    /**
     * SetCellValue provides a function to set the value of a cell. The
     * specified coordinates should not be in the first row of the table, a
     * complex number can be set with string text.
     *
     * Note that default date format is m/d/yy h:mm of time.Time type value. You
     * can set numbers format by the SetCellStyle function. If you need to set
     * the specialized date in Excel like January 0, 1900 or February 29, 1900,
     * these times can not representation in Go language time.Time data type.
     * Please set the cell value as number 0 or 60, then create and bind the
     * date-time number format style for the cell.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param value The cell value to be write
     */
    SetCellValue(sheet: string, cell: string, value: boolean | number | string ): { error: string | null }

    /**
     * SetColOutlineLevel provides a function to set outline level of a single
     * column by given worksheet name and column name. The value of parameter
     * `level` is 1-7.
     * @param sheet The worksheet name
     * @param col The column name
     * @param level The outline level of the column
     */
    SetColOutlineLevel(sheet: string, col: string, level: number): { error: string | null }

    /**
     * SetColStyle provides a function to set style of columns by given
     * worksheet name, columns range and style ID. Note that this will
     * overwrite the existing styles for the columns, it won't append or merge
     * style with existing styles.
     * @param sheet The worksheet name
     * @param columns The column range
     * @param styleID The style ID
     */
    SetColStyle(sheet: string, columns: string, styleID: number): { error: string | null }

    /**
     * SetColVisible provides a function to set visible columns by given
     * worksheet name, columns range and visibility.
     * @param sheet The worksheet name
     * @param columns The column name
     * @param visible The column's visibility
     */
    SetColVisible(sheet: string, columns: string, visible: boolean): { error: string | null }

    /**
     * SetColWidth provides a function to set the width of a single column or
     * multiple columns.
     * @param sheet The worksheet name
     * @param startCol The start column name
     * @param endCol The end column name
     * @param width The width of the column
     */
    SetColWidth(sheet: string, startCol: string, endCol: string, width: number): { error: string | null }

    /**
     * SetConditionalFormat provides a function to create conditional
     * formatting rule for cell value. Conditional formatting is a feature of
     * Excel which allows you to apply a format to a cell or a range of cells
     * based on certain criteria.
     * @param sheet The worksheet name
     * @param reference The conditional format range reference
     * @param opts The conditional options
     */
    SetConditionalFormat(sheet: string, reference: string, opts: string): { error: string | null }

    /**
     * SetDefaultFont changes the default font in the workbook.
     * @param fontName The font name
     */
    SetDefaultFont(fontName: string): { error: string | null }

    /**
     * SetPanes provides a function to create and remove freeze panes and split
     * panes by given worksheet name and panes format set.
     * @param sheet The worksheet name
     * @param panes The panes format
     */
    SetPanes(sheet: string, panes: string): { error: string | null }

    /**
     * SetRowHeight provides a function to set the height of a single row.
     * @param sheet The worksheet name
     * @param row The row number
     * @param height The height of the row
     */
    SetRowHeight(sheet: string, row: number, height : number): { error: string | null }

    /**
     * SetRowOutlineLevel provides a function to set outline level number of a
     * single row by given worksheet name and Excel row number. The value of
     * parameter `level` is 1-7.
     * @param sheet The worksheet name
     * @param row The row number
     * @param level The outline level of the row
     */
    SetRowOutlineLevel(sheet: string, row: number, level: number): { error: string | null }

    /**
     * SetRowStyle provides a function to set the style of rows by given
     * worksheet name, row range, and style ID. Note that this will overwrite
     * the existing styles for the rows, it won't append or merge style with
     * existing styles.
     * @param sheet The worksheet name
     * @param start The start row number
     * @param end Then end row number
     * @param styleID The style ID
     */
    SetRowStyle(sheet: string, start: number, end: number, styleID: number): { error: string | null }

    /**
     * SetRowStyle provides a function to set the style of rows by given
     * worksheet name, row range, and style ID. Note that this will overwrite
     * the existing styles for the rows, it won't append or merge style with
     * existing styles.
     * @param sheet The worksheet name
     * @param row The row number
     * @param visible The row's visibility
     */
    SetRowVisible(sheet: string, row: number, visible: boolean): { error: string | null }

    /**
     * SetSheetCol writes an array to column by given worksheet name, starting
     * cell reference and a pointer to array type 'slice'.
     * @param sheet The worksheet name
     * @param cell The cell reference
     * @param slice The column cells to be write
     */
    SetSheetCol(sheet: string, cell: string, slice: Array<boolean | number | string>): { error: string | null }

    /**
     * SetSheetName provides a function to set the worksheet name by given
     * source and target worksheet names. Maximum 31 characters are allowed in
     * sheet title and this function only changes the name of the sheet and
     * will not update the sheet name in the formula or reference associated
     * with the cell. So there may be problem formula error or reference
     * missing.
     * @param source The source sheet name
     * @param target The target sheet name
     */
    SetSheetName(source: string, target: string): { error: string | null }

    /**
     * SetSheetRow writes an array to row by given worksheet name, starting
     * cell reference and a pointer to array type 'slice'.
     * @param sheet The worksheet name
     * @param cell The starting cell reference
     * @param slice The array for writes
     */
    SetSheetRow(sheet: string, cell: string, slice: Array<boolean | number | string>): { error: string | null }

    /**
     * SetSheetVisible provides a function to set worksheet visible by given
     * worksheet name. A workbook must contain at least one visible worksheet.
     * If the given worksheet has been activated, this setting will be
     * invalidated.
     * @param sheet The worksheet name
     * @param visible The worksheet visibility
     */
    SetSheetVisible(sheet: string, visible: boolean): { error: string | null }

    /**
     * UngroupSheets provides a function to ungroup worksheets.
     */
    UngroupSheets(): { error: string | null }

    /**
     * UnmergeCell provides a function to unmerge a given range reference.
     * @param sheet The worksheet name
     * @param hCell The top-left cell reference
     * @param vCell The right-bottom cell reference
     */
    UnmergeCell(sheet: string, hCell: string, vCell: string): { error: string | null }

    /**
     * UnprotectSheet provides a function to remove protection for a sheet,
     * specified the second optional password parameter to remove sheet
     * protection with password verification.
     * @param sheet The worksheet name
     * @param password The password for sheet protection
     */
    UnprotectSheet(sheet: string, password?: string): { error: string | null }

    /**
     * UnsetConditionalFormat provides a function to unset the conditional
     * format by given worksheet name and range reference.
     * @param sheet The worksheet name
     * @param reference The conditional format range reference
     */
    UnsetConditionalFormat(sheet: string, reference: string): { error: string | null }

    /**
     * UpdateLinkedValue fix linked values within a spreadsheet are not
     * updating in Office Excel application. This function will be remove
     * value tag when met a cell have a linked value.
     */
    UpdateLinkedValue(): { error: string | null }

    /**
     * WriteToBuffer provides a function to get the contents buffer from the
     * saved file, and it allocates space in memory. Be careful when the file
     * size is large.
     * @param opts The options for save the spreadsheet
     */
    WriteToBuffer(opts?: Options): { buffer: BlobPart, error: string | null };

    /**
     * Error message
     */
    error?: string | null;
  }

  /**
   * init provides a function to compile and instantiate WebAssembly code by a
   * given compressed wasm archive path.
   * @param path The compressed wasm archive path
   */
  export function init(path: string): Promise<{
    CellNameToCoordinates: typeof CellNameToCoordinates,
    ColumnNameToNumber:    typeof ColumnNameToNumber,
    ColumnNumberToName:    typeof ColumnNumberToName,
    CoordinatesToCellName: typeof CoordinatesToCellName,
    HSLToRGB:              typeof HSLToRGB,
    JoinCellName:          typeof JoinCellName,
    RGBToHSL:              typeof RGBToHSL,
    SplitCellName:         typeof SplitCellName,
    ThemeColor:            typeof ThemeColor,
    NewFile:               typeof NewFile;
    OpenReader:            typeof OpenReader;
  }>;
}
