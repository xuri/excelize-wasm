// Type definitions for excelize-wasm

/**
 * AppProperties directly maps the document application properties.
 */
declare type AppProperties = {
  application:        string;
  scale_crop:         boolean;
  doc_security:       number;
  company:            string;
  links_up_to_date:   boolean;
  hyperlinks_changed: boolean;
  app_version:        string;
  error?:             string;
};

/**
 * Options define the options for open and reading spreadsheet.
 */
declare type Options = {
	max_calc_iterations?:  number;
	password?:             string;
	raw_cell_value?:       boolean;
	unzip_size_limit?:     number;
	unzip_xml_size_limit?: number;
};

/**
 * ColumnNumberToName provides a function to convert the integer to Excel sheet
 * column title.
 * @param cell The cell reference
 */
declare function CellNameToCoordinates(cell: string): { col: number, row: number, error: string }

/**
 * ColumnNameToNumber provides a function to convert Excel sheet column name
 * (case-insensitive) to int. The function returns an error if column name
 * incorrect.
 * @param name The column name
 */
declare function ColumnNameToNumber(name: string): { col: number, error: string }

/**
 * ColumnNumberToName provides a function to convert the integer to Excel sheet
 * column title.
 * @param num The column name
 */
declare function ColumnNumberToName(num: number): { col: string, error: string }

/**
 * CoordinatesToCellName converts [X, Y] coordinates to alpha-numeric cell name
 * or returns an error.
 * @param col The column number
 * @param row The row number
 * @param abs Specifies the absolute cell references
 */
declare function CoordinatesToCellName(col: number, row: number, abs?: boolean): { col: string, error: string }

/**
 * HSLToRGB converts an HSL triple to a RGB triple.
 * @param h Hue
 * @param s Saturation
 * @param l Lightness
 */
declare function HSLToRGB(h: number, s: number, l: number): { r: number, g: number, b: number }

/**
 * JoinCellName joins cell name from column name and row number.
 * @param col The column name
 * @param row The row number
 */
declare function JoinCellName(col: string, row: number): { cell: string, error: string }

/**
 * RGBToHSL converts an RGB triple to a HSL triple.
 * @param r Red
 * @param g Green
 * @param b Blue
 */
declare function RGBToHSL(r: number, g: number, b: number): { h: number, s: number, l: number }

/**
 * SplitCellName splits cell name to column name and row number.
 * @param cell The cell reference
 */
declare function SplitCellName(cell: string): { col: string, row: number, error: string }

/**
 * ThemeColor applied the color with tint value.
 * @param baseColor Base color in hex format
 * @param tint A mixture of a color with white
 */
declare function ThemeColor(baseColor: string, tint: number): { color: string }

/**
 * NewFile provides a function to create new file by default template.
 */
declare function NewFile(): NewFile;

/**
 * OpenReader read data stream from buffer and return a populated spreadsheet
 * file.
 * @param r The contents buffer of the file
 * @param opts The options for open and reading spreadsheet
 */
declare function OpenReader(r: Uint8Array[], opts?: Options): NewFile;

/**
 * @constructor
 */
declare class NewFile {
  /**
   * AddChart provides the method to add chart in a sheet by given chart format
   * set (such as offset, scale, aspect ratio setting and print settings) and
   * properties set.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param format The chart format
   * @param combo Specifies the create a chart that combines two or more chart
   *  types in a single chart
   */
  AddChart(sheet: string, cell: string, format: string, combo?: string): { error: string }

  /**
   * AddChartSheet provides the method to create a chartsheet by given chart
   * format set (such as offset, scale, aspect ratio setting and print
   * settings) and properties set. In Excel a chartsheet is a worksheet that
   * only contains a chart.
   * @param sheet The worksheet name
   * @param format The chart format
   * @param combo Specifies the create a chart that combines two or more chart
   *  types in a single chart
   */
  AddChartSheet(sheet: string, format: string, combo?: string): { error: string }

  /**
   * AddComment provides the method to add comment in a sheet by given
   * worksheet index, cell and format set (such as author and text). Note that
   * the max author length is 255 and the max text length is 32512.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param format The comment format
   */
  AddComment(sheet: string, cell: string, format: string): { error: string }

  /**
   * AddPictureFromBytes provides the method to add picture in a sheet by given
   * picture format set (such as offset, scale, aspect ratio setting and print
   * settings), file base name, extension name and file bytes.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param format The picture format
   * @param name The picture name
   * @param extension The extension name
   * @param file The contents buffer of the file
   */
  AddPictureFromBytes(sheet: string, cell: string, format: string, name: string, extension: string, file: Uint8Array[]): { error: string }

  /**
   * AddPivotTable provides the method to add pivot table by given pivot table
   * options. Note that the same fields can not in Columns, Rows and Filter
   * fields at the same time.
   * @param opt The pivot table option
   */
  AddPivotTable(opt: string): { error: string }

  /**
   * AddShape provides the method to add shape in a sheet by given worksheet
   * index, shape format set (such as offset, scale, aspect ratio setting and
   * print settings) and properties set.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param format The shape format
   */
  AddShape(sheet: string, cell: string, format: string): { error: string }

  /**
   * AddTable provides the method to add table in a worksheet by given
   * worksheet name, range reference and format set.
   * @param sheet The worksheet name
   * @param hCell The top-left cell reference
   * @param vCell The right-bottom cell reference
   * @param format The table format
   */
  AddTable(sheet: string, hCell: string, vCell: string, format: string): { error: string }

  /**
   * AutoFilter provides the method to add auto filter in a worksheet by given
   * worksheet name, range reference and settings. An auto filter in Excel is
   * a way of filtering a 2D range of data based on some simple criteria.
   * @param sheet The worksheet name
   * @param hCell The top-left cell reference
   * @param vCell The right-bottom cell reference
   * @param format The auto filter format
   */
  AutoFilter(sheet: string, hCell: string, vCell: string, format: string): { error: string }

  /**
   * CalcCellValue provides a function to get calculated cell value. This
   * feature is currently in working processing. Iterative calculation,
   * implicit intersection, explicit intersection, array formula, table
   * formula and some other formulas are not supported currently.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  CalcCellValue(sheet: string, cell: string): { result: string, error: string }

  /**
   * CopySheet provides a function to duplicate a worksheet by gave source and
   * target worksheet index. Note that currently doesn't support duplicate
   * workbooks that contain tables, charts or pictures.
   * @param from Source sheet name
   * @param to Target sheet name
   */
  CopySheet(from: number, to: number): { error: string }

  /**
   * DeleteChart provides a function to delete chart in spreadsheet by given
   * worksheet name and cell reference.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  DeleteChart(sheet: string, cell: string): { error: string }

  /**
   * DeleteComment provides the method to delete comment in a sheet by given
   * worksheet name.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  DeleteComment(sheet: string, cell: string): { error: string }

  /**
   * DeleteDataValidation delete data validation by given worksheet name and
   * reference sequence. All data validations in the worksheet will be deleted
   * if not specify reference sequence parameter.
   * @param sheet The worksheet name
   * @param sqref The cell reference sequence
   */
  DeleteDataValidation(sheet: string, sqref?: string): { error: string }

  /**
   * DeletePicture provides a function to delete charts in spreadsheet by given
   * worksheet name and cell reference. Note that the image file won't be
   * deleted from the document currently.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  DeletePicture(sheet: string, cell: string): { error: string }

  /**
   * DeleteSheet provides a function to delete worksheet in a workbook by given
   * worksheet name. Use this method with caution, which will affect changes in
   * references such as formulas, charts, and so on. If there is any referenced
   * value of the deleted worksheet, it will cause a file error when you open
   * it. This function will be invalid when only one worksheet is left.
   * @param sheet The worksheet name
   */
  DeleteSheet(sheet: string)

  /**
   * DuplicateRow inserts a copy of specified row (by its Excel row number)
   * below. Use this method with caution, which will affect changes in
   * references such as formulas, charts, and so on. If there is any
   * referenced value of the worksheet, it will cause a file error when you
   * open it. The excelize only partially updates these references currently.
   * @param sheet The worksheet name
   * @param row The row number
   */
  DuplicateRow(sheet: string, row: number): { error: string }

  /**
   * DuplicateRowTo inserts a copy of specified row by it Excel number to
   * specified row position moving down exists rows after target position. Use
   * this method with caution, which will affect changes in references such as
   * formulas, charts, and so on. If there is any referenced value of the
   * worksheet, it will cause a file error when you open it. The excelize only
   * partially updates these references currently.
   * @param sheet The worksheet name
   * @param row The source row number
   * @param row2 The target row number
   */
  DuplicateRowTo(sheet: string, row: number, row2: number): { error: string }

  /**
   * GetActiveSheetIndex provides a function to get active sheet index of the
   * spreadsheet. If not found the active sheet will be return integer 0.
   */
  GetActiveSheetIndex(): { index: number }
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
  GetCellFormula(sheet: string, cell: string): { formula: string, error: string }

  /**
   * GetCellHyperLink gets a cell hyperlink based on the given worksheet name
   * and cell reference. If the cell has a hyperlink, it will return 'true'
   * and the link address, otherwise it will return 'false' and an empty link
   * address.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  GetCellHyperLink(sheet: string, cell: string): { ok: boolean, location: string, error: string }

  /**
   * GetCellStyle provides a function to get cell style index by given
   * worksheet name and cell reference.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  GetCellStyle(sheet: string, cell: string): { style: number, error: string }

  /**
   * GetCellValue provides a function to get formatted value from cell by given
   * worksheet name and cell reference in spreadsheet. The return value is
   * converted to the `string` data type. If the cell format can be applied to
   * the value of a cell, the applied value will be returned, otherwise the
   * original value will be returned. All cells' values will be the same in a
   * merged range.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  GetCellValue(sheet: string, cell: string): { value: string, error: string }

  /**
   * GetColOutlineLevel provides a function to get outline level of a single
   * column by given worksheet name and column name.
   * @param sheet The worksheet name
   * @param col The column name
   */
  GetColOutlineLevel(sheet: string, col: string): { level: number, error: string }

  /**
   * GetColStyle provides a function to get column style ID by given worksheet
   * name and column name.
   * @param sheet The worksheet name
   * @param col The column name
   */
  GetColStyle(sheet: string, col: string): { style: number, error: string }

  /**
   * GetColVisible provides a function to get visible of a single column by
   * given worksheet name and column name.
   * @param sheet The worksheet name
   * @param col The column name
   */
  GetColVisible(sheet: string, col: string): { visible: boolean, error: string }

  /**
   * GetColWidth provides a function to get column width by given worksheet
   * name and column name.
   * @param sheet The worksheet name
   * @param col The column name
   */
  GetColWidth(sheet: string, col: string): { width: number, error: string }

  /**
   * GetCols gets the value of all cells by columns on the worksheet based on
   * the given worksheet name, returned as a two-dimensional array, where the
   * value of the cell is converted to the `string` type. If the cell format
   * can be applied to the value of the cell, the applied value will be used,
   * otherwise the original value will be used.
   * @param sheet The worksheet name
   * @param opts
   */
  GetCols(sheet: string, opts?: Options): { result: string[][], error: string }

  /**
   * GetRowHeight provides a function to get row height by given worksheet name
   * and row number.
   * @param sheet The worksheet name
   * @param row The row number
   */
  GetRowHeight(sheet: string, row: number): { height: number, error: string }

  /**
   * GetRowOutlineLevel provides a function to get outline level number of a
   * single row by given worksheet name and Excel row number.
   * @param sheet The worksheet name
   * @param row The row number
   */
  GetRowOutlineLevel(sheet: string, row: number): { level: number, error: string }

  /**
   * GetRowVisible provides a function to get visible of a single row by given
   * worksheet name and Excel row number.
   * @param sheet The worksheet name
   * @param row The row number
   */
  GetRowVisible(sheet: string, row: number): { visible: boolean, error: string }

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
  GetRows(sheet: string, opts?: Options): { result: string[][], error: string }

  /**
   * GetSheetIndex provides a function to get a sheet index of the workbook by
   * the given sheet name. If the given sheet name is invalid or sheet doesn't
   * exist, it will return an integer type value -1.
   * @param sheet The worksheet name
   */
  GetSheetIndex(sheet: string): number

  /**
   * GetSheetList provides a function to get worksheets, chart sheets, and
   * dialog sheets name list of the workbook.
   */
  GetSheetList(): { list: string[] }

  /**
   * GetSheetMap provides a function to get worksheets, chart sheets, dialog
   * sheets ID and name map of the workbook.
   */
  GetSheetMap(): Map<string,string>

  /**
   * GetSheetName provides a function to get the sheet name of the workbook by
   * the given sheet index. If the given sheet index is invalid, it will
   * return an empty string.
   * @param index The sheet index
   */
  GetSheetName(index: number): { name: string }

  /**
   * GetSheetVisible provides a function to get worksheet visible by given
   * worksheet name.
   * @param sheet The worksheet name
   */
  GetSheetVisible(sheet: string): boolean

  /**
   * GroupSheets provides a function to group worksheets by given worksheets
   * name. Group worksheets must contain an active worksheet.
   * @param sheet The worksheet names
   */
  GroupSheets(sheets: string[]): { error: string }

  /**
   * InsertCols provides a function to insert new columns before the given
   * column name and number of columns.
   *
   * Use this method with caution, which will affect changes in references such
   * as formulas, charts, and so on. If there is any referenced value of the
   * worksheet, it will cause a file error when you open it. The excelize only
   * partially updates these references currently.
   * @param sheet The worksheet name
   * @param col The base column name
   * @param n The instert columns count
   */
  InsertCols(sheet: string, col: string, n: number): { error: string }

  /**
   * InsertPageBreak create a page break to determine where the printed page
   * ends and where begins the next one by given worksheet name and cell, so
   * the content before the page break will be printed on one page and after
   * the page break on another.
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  InsertPageBreak(sheet: string, cell: string): { error: string }

  /**
   * InsertRows provides a function to insert new rows after the given Excel
   * row number starting from 1 and number of rows.
   *
   * Use this method with caution, which will affect changes in references such
   * as formulas, charts, and so on. If there is any referenced value of the
   * worksheet, it will cause a file error when you open it. The excelize only
   * partially updates these references currently.
   * @param sheet The worksheet name
   * @param row The base row number
   * @param n Insert rows count
   */
  InsertRows(sheet: string, row: number, n: number): { error: string }

  /**
   * MergeCell provides a function to merge cells by given range reference and
   * sheet name. Merging cells only keeps the upper-left cell value, and
   * discards the other values.
   *
   * If you create a merged cell that overlaps with another existing merged
   * cell, those merged cells that already exist will be removed. The cell
   * references tuple after merging in the following range will be: A1
   * (x3,y1) D1(x2,y1) A8(x3,y4) D8(x2,y4)
   *
   *  	             B1(x1,y1)      D1(x2,y1)
   *  	           +------------------------+
   *  	           |                        |
   *  	 A4(x3,y3) |    C4(x4,y3)           |
   *  	+------------------------+          |
   *  	|          |             |          |
   *  	|          |B5(x1,y2)    | D5(x2,y2)|
   *  	|          +------------------------+
   *  	|                        |
   *  	|A8(x3,y4)      C8(x4,y4)|
   *  	+------------------------+
   * @param sheet The worksheet name
   * @param hCell The top-left cell reference
   * @param vCell The right-bottom cell reference
   */
  MergeCell(sheet: string, hCell: string, vCell: string): { error: string }

  /**
   * NewConditionalStyle provides a function to create style for conditional
   * format by given style format. The parameters are the same with the
   * NewStyle function. Note that the color field uses RGB color code and only
   * support to set font, fills, alignment and borders currently.
   * @param style
   */
  NewConditionalStyle(style: string): { style: number, error: string }

  /**
   * NewSheet provides the function to create a new sheet by given a worksheet
   * name and returns the index of the sheets in the workbook after it
   * appended. Note that when creating a new workbook, the default worksheet
   * named `Sheet1` will be created.
   * @param sheet The worksheet name
   */
  NewSheet(sheet: string): number

  /**
   * NewStyle provides a function to create the style for cells by given JSON.
   * Note that the color field uses RGB color code.
   * @param style The style format
   */
  NewStyle(style: string): { style: number, error: string }

  /**
   * RemoveCol provides a function to remove single column by given worksheet
   * name and column index.
   *
   * Use this method with caution, which will affect changes in references such
   * as formulas, charts, and so on. If there is any referenced value of the
   * worksheet, it will cause a file error when you open it. The excelize only
   * partially updates these references currently.
   * @param sheet The worksheet name
   * @param col The column name
   */
  RemoveCol(sheet: string, col: string): { error: string }

  /**
   * RemovePageBreak remove a page break by given worksheet name and cell
   * reference
   * @param sheet The worksheet name
   * @param cell The cell reference
   */
  RemovePageBreak(sheet: string, cell: string): { error: string }

  /**
   * RemoveRow provides a function to remove single row by given worksheet name
   * and Excel row number.
   *
   * Use this method with caution, which will affect changes in references such
   * as formulas, charts, and so on. If there is any referenced value of the
   * worksheet, it will cause a file error when you open it. The excelize only
   * partially updates these references currently.
   * @param sheet The worksheet name
   * @param row The row number
   */
  RemoveRow(sheet: string, row: number): { error: string }

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
  SearchSheet(sheet: string, value: string, reg?: boolean): { result: string, error: string }

  /**
   * SetActiveSheet provides a function to set the default active sheet of the
   * workbook by a given index. Note that the active index is different from
   * the ID returned by function GetSheetMap(). It should be greater than or
   * equal to 0 and less than the total worksheet numbers.
   * @param index The sheet index
   */
  SetActiveSheet(index: number)

  /**
   * SetCellBool provides a function to set bool type value of a cell by given
   * worksheet name, cell reference and cell value.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   */
  SetCellBool(sheet: string, cell: string, value: boolean): { error: string }

  /**
   * SetCellDefault provides a function to set string type value of a cell as
   * default format without escaping the cell.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   */
  SetCellDefault(sheet: string, cell: string, value: string): { error: string }

  /**
   * SetCellFloat sets a floating point value into a cell. The precision
   * parameter specifies how many places after the decimal will be shown
   * while -1 is a special value that will use as many decimal places as
   * necessary to represent the number. bitSize is 32 or 64 depending on if a
   * float32 or float64 was originally used for the value. For Example:
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   * @param precision Specifies how many places after the decimal will be shown
   * @param bitSize BitSize is 32 or 64 depending on if a float32 or float64
   *  was originally used for the value
   */
  SetCellFloat(sheet: string, cell: string, value: number, precision: number, bitSize: number): { error: string }

  /**
   * SetCellInt provides a function to set int type value of a cell by given
   * worksheet name, cell reference and cell value.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   */
  SetCellInt(sheet: string, cell: string, value: number): { error: string }

  /**
   * SetCellStr provides a function to set string type value of a cell. Total
   * number of characters that a cell can contain 32767 characters.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   */
  SetCellStr(sheet: string, cell: string, value: string): { error: string }

  /**
   * SetCellStyle provides a function to add style attribute for cells by given
   * worksheet name, range reference and style ID. Note that diagonalDown and
   * diagonalUp type border should be use same color in the same range.
   * SetCellStyle will overwrite the existing styles for the cell, it won't
   * append or merge style with existing styles.
   * @param sheet The worksheet name
   * @param hCell The top-left cell reference
   * @param vCell The right-bottom cell reference
   * @param styleID The style ID
   */
  SetCellStyle(sheet: string, hCell: string, vCell: string, styleID: number): { error: string }

  /**
   * SetCellValue provides a function to set the value of a cell. The specified
   * coordinates should not be in the first row of the table, a complex number
   * can be set with string text.
   *
   * You can set numbers format by the SetCellStyle function. If you need to set
   * the specialized date in Excel like January 0, 1900 or February 29, 1900.
   * Please set the cell value as number 0 or 60, then create and bind the
   * date-time number format style for the cell.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param value The cell value to be write
   */
  SetCellValue(sheet: string, cell: string, value: boolean | number | string ): { error: string }

  /**
   * SetColOutlineLevel provides a function to set outline level of a single
   * column by given worksheet name and column name. The value of parameter
   * `level` is 1-7.
   * @param sheet The worksheet name
   * @param col The column name
   * @param level The outline level of the column
   */
  SetColOutlineLevel(sheet: string, col: string, level: number): { error: string }

  /**
   * SetColStyle provides a function to set style of columns by given worksheet
   * name, columns range and style ID. Note that this will overwrite the
   * existing styles for the columns, it won't append or merge style with
   * existing styles.
   * @param sheet The worksheet name
   * @param columns The column range
   * @param styleID The style ID
   */
  SetColStyle(sheet: string, columns: string, styleID: number): { error: string }

  /**
   * SetColVisible provides a function to set visible columns by given
   * worksheet name, columns range and visibility.
   * @param sheet The worksheet name
   * @param columns The column name
   * @param visible The column's visibility
   */
  SetColVisible(sheet: string, columns: string, visible: boolean): { error: string }

  /**
   * SetColWidth provides a function to set the width of a single column or
   * multiple columns.
   * @param sheet The worksheet name
   * @param startCol The start column name
   * @param endCol The end column name
   * @param width The width of the column
   */
  SetColWidth(sheet: string, startCol: string, endCol: string, width: number): { error: string }

  /**
   * SetDefaultFont changes the default font in the workbook.
   * @param fontName The font name
   */
  SetDefaultFont(fontName: string)

  /**
   * SetPanes provides a function to create and remove freeze panes and split
   * panes by given worksheet name and panes format set.
   * @param sheet The worksheet name
   * @param panes The panes format
   */
  SetPanes(sheet: string, panes: string): { error: string }

  /**
   * SetRowHeight provides a function to set the height of a single row.
   * @param sheet The worksheet name
   * @param row The row number
   * @param height The height of the row
   */
  SetRowHeight(sheet: string, row: number, height : number): { error: string }

  /**
   * SetRowOutlineLevel provides a function to set outline level number of a
   * single row by given worksheet name and Excel row number. The value of
   * parameter `level` is 1-7.
   * @param sheet The worksheet name
   * @param row The row number
   * @param level The outline level of the row
   */
  SetRowOutlineLevel(sheet: string, row: number, level: number): { error: string }

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
  SetRowStyle(sheet: string, start: number, end: number, styleID: number): { error: string }

  /**
   * SetRowStyle provides a function to set the style of rows by given
   * worksheet name, row range, and style ID. Note that this will overwrite
   * the existing styles for the rows, it won't append or merge style with
   * existing styles.
   * @param sheet The worksheet name
   * @param row The row number
   * @param visible The row's visibility
   */
  SetRowVisible(sheet: string, row: number, visible: boolean): { error: string }

  /**
   * SetSheetCol writes an array to column by given worksheet name, starting
   * cell reference and a pointer to array type 'slice'.
   * @param sheet The worksheet name
   * @param cell The cell reference
   * @param slice The column cells to be write
   */
  SetSheetCol(sheet: string, cell: string, slice: Array<boolean | number | string>): { error: string }

  /**
   * SetSheetName provides a function to set the worksheet name by given source
   * and target worksheet names. Maximum 31 characters are allowed in sheet
   * title and this function only changes the name of the sheet and will not
   * update the sheet name in the formula or reference associated with the
   * cell. So there may be problem formula error or reference missing.
   * @param source The source sheet name
   * @param target The target sheet name
   */
  SetSheetName(source: string, target: string)

  /**
   * SetSheetRow writes an array to row by given worksheet name, starting cell
   * reference and a pointer to array type 'slice'.
   * @param sheet The worksheet name
   * @param cell The starting cell reference
   * @param slice The array for writes
   */
  SetSheetRow(sheet: string, cell: string, slice: Array<boolean | number | string>): { error: string }

  /**
   * SetSheetVisible provides a function to set worksheet visible by given
   * worksheet name. A workbook must contain at least one visible worksheet.
   * If the given worksheet has been activated, this setting will be
   * invalidated.
   * @param sheet The worksheet name
   * @param visible The worksheet visibility
   */
  SetSheetVisible(sheet: string, visible: boolean): { error: string }

  /**
   * UngroupSheets provides a function to ungroup worksheets.
   */
  UngroupSheets(): { error: string }

  /**
   * UnmergeCell provides a function to unmerge a given range reference.
   * @param sheet The worksheet name
   * @param hCell The top-left cell reference
   * @param vCell The right-bottom cell reference
   */
  UnmergeCell(sheet: string, hCell: string, vCell: string): { error: string }

  /**
   * UnprotectSheet provides a function to remove protection for a sheet,
   * specified the second optional password parameter to remove sheet
   * protection with password verification.
   * @param sheet The worksheet name
   * @param password The password for sheet protection
   */
  UnprotectSheet(sheet: string, password?: string): { error: string }

  /**
   * UnsetConditionalFormat provides a function to unset the conditional format
   * by given worksheet name and range reference.
   * @param sheet The worksheet name
   * @param reference The conditional format range reference
   */
  UnsetConditionalFormat(sheet: string, reference: string): { error: string }

  /**
   * UpdateLinkedValue fix linked values within a spreadsheet are not updating
   * in Office Excel application. This function will be remove value tag when
   * met a cell have a linked value.
   */
  UpdateLinkedValue(): { error: string }

  /**
   * WriteToBuffer provides a function to get the contents buffer from the
   * saved file, and it allocates space in memory. Be careful when the file
   * size is large.
   * @param opts The options for save the spreadsheet
   */
  WriteToBuffer(opts?: Options): Uint8Array[] | string;

  /**
   * Error message
   */
  error?: string;
}

declare function excelize(path: string): Promise<void>;

/**
 * A window containing a DOM document; the document property points to the DOM
 * document loaded in that window.
 */
interface Window {
  NewFile: NewFile;
  OpenReader: NewFile;
}
