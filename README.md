# excelize-wasm

<p align="center"><img width="500" src="https://github.com/xuri/excelize-wasm/raw/main/excelize-wasm.svg" alt="excelize-wasm logo"></p>

<p align="center">
    <a href="https://github.com/xuri/excelize-wasm/actions/workflows/publish.yml"><img src="https://github.com/xuri/excelize-wasm/actions/workflows/publish.yml/badge.svg" alt="NPM publish"></a>
    <a href="https://github.com/xuri/excelize-wasm/actions/workflows/go.yml"><img src="https://github.com/xuri/excelize-wasm/actions/workflows/go.yml/badge.svg" alt="Build Status"></a>
    <a href="https://codecov.io/gh/xuri/excelize-wasm"><img src="https://codecov.io/gh/xuri/excelize-wasm/branch/master/graph/badge.svg" alt="Code Coverage"></a>
    <a href="https://goreportcard.com/report/github.com/xuri/excelize-wasm"><img src="https://goreportcard.com/badge/github.com/xuri/excelize-wasm" alt="Go Report Card"></a>
    <a href="https://pkg.go.dev/github.com/xuri/excelize/v2"><img src="https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white" alt="go.dev"></a>
    <a href="https://opensource.org/licenses/BSD-3-Clause"><img src="https://img.shields.io/badge/license-bsd-orange.svg" alt="Licenses"></a>
    <a href="https://www.paypal.com/paypalme/xuri"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>
</p>

Excelize-wasm is a pure WebAssembly / Javascript port of Go [Excelize](https://github.com/xuri/excelize) library that allow you to write to and read from XLAM / XLSM / XLSX / XLTM / XLTX files. Supports reading and writing spreadsheet documents generated by Microsoft Excel&trade; 2007 and later. Supports complex components by high compatibility. The full API docs can be found at [docs reference](https://xuri.me/excelize/).

## Environment Compatibility

Browser | Version
---|---
Chrome | &ge;57
Chrome for Android and Android Browser | &ge;105
Edge | &ge;16
Safari on macOS and iOS | &ge;11
Firefox | &ge;52
Firefox for Android | &ge;104
Opera | &ge;44
Opera Mobile | &ge;64
Samsung Internet | &ge;7.2
UC Browser for Android | &ge;13.4
QQ Browser | &ge;10.4
Node.js | &ge;8.0.0
Deno | &ge;1.0

## Basic Usage

### Installation

#### Node.js

```bash
npm install excelize-wasm
```

#### Browser

```html
<script src="excelize-wasm/index.js"></script>
````

### Create spreadsheet

Here is a minimal example usage that will create spreadsheet file.

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = NewFile();
  // Create a new sheet.
  const index = f.NewSheet("Sheet2")
  // Set value of a cell.
  f.SetCellValue("Sheet2", "A2", "Hello world.")
  f.SetCellValue("Sheet1", "B2", 100)
  // Set active sheet of the workbook.
  f.SetActiveSheet(index)
  // Save spreadsheet by the given path.
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

Create spreadsheet in browser:

<details>
  <summary>View code</summary>

```html
<html>
<head>
  <meta charset="utf-8">
  <script src="excelize-wasm/index.js"></script>
</head>
<body>
  <div>
    <button onclick="download()">Download</button>
  </div>
  <script>
  function download() {
    excelize('https://xuri.me/excelize-wasm/v0.0.1/excelize.wasm.gz').then(() => {
      const f = NewFile();
      // Create a new sheet.
      const index = f.NewSheet("Sheet2")
      // Set value of a cell.
      f.SetCellValue("Sheet2", "A2", "Hello world.")
      f.SetCellValue("Sheet1", "B2", 100)
      // Set active sheet of the workbook.
      f.SetActiveSheet(index)
      // Save spreadsheet by the given path.
      const link = document.createElement('a');
      link.download = 'Book1.xlsx';
      link.href = URL.createObjectURL(
        new Blob([f.WriteToBuffer()],
        { type: 'application/vnd.ms-excel' })
      );
      link.click();
    });
  }
  </script>
</body>
```

</details>

### Reading spreadsheet

The following constitutes the bare to read a spreadsheet document.

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = OpenReader(fs.readFileSync('Book1.xlsx'));
  // Create a new sheet.
  const index = f.NewSheet("Sheet2")
  // Set value of a cell.
  var { cell, error } = f.GetCellValue("Sheet1", "B2")
  if (error) {
    console.log(error);
    return;
  }
  console.log(cell)
  // Get all the rows in the Sheet1.
  var { result, error } = f.GetRows("Sheet1");
  if (error) {
    console.log(error);
    return;
  }
  result.forEach(row => {
    row.forEach(colCell => {
      process.stdout.write(`${colCell}\t`)
    })
    console.log();
  });
});
```

### Add chart to spreadsheet file

With excelize-wasm chart generation and management is as easy as a few lines of code. You can build charts based on data in your worksheet or generate charts without any data in your worksheet at all.

<p align="center"><img width="650" src="https://github.com/xuri/excelize-wasm/blob/main/chart.png" alt="Excelize"></p>

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const categories = {
    "A2": "Small", "A3": "Normal", "A4": "Large",
    "B1": "Apple", "C1": "Orange", "D1": "Pear"};
  const values = {"B2": 2, "C2": 3, "D2": 3, "B3": 5,
    "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8};
  const f = NewFile();
  for (const k in categories) {
    f.SetCellValue("Sheet1", k, categories[k]);
  };
  for (const k in values) {
    f.SetCellValue("Sheet1", k, values[k]);
  };
  var { error } = f.AddChart("Sheet1", "E1", `{
    "type": "col3DClustered",
    "series": [
    {
        "name": "Sheet1!$A$2",
        "categories": "Sheet1!$B$1:$D$1",
        "values": "Sheet1!$B$2:$D$2"
    },
    {
        "name": "Sheet1!$A$3",
        "categories": "Sheet1!$B$1:$D$1",
        "values": "Sheet1!$B$3:$D$3"
    },
    {
        "name": "Sheet1!$A$4",
        "categories": "Sheet1!$B$1:$D$1",
        "values": "Sheet1!$B$4:$D$4"
    }],
    "title":
    {
        "name": "Fruit 3D Clustered Column Chart"
    }
  }`);
  if (error) {
    console.log(error);
    return
  }
  // Save spreadsheet by the given path.
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

### Add picture to spreadsheet file

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = OpenReader(fs.readFileSync('Book1.xlsx'));
  if (f.error) {
    console.log(f.error);
    return
  }
  // Insert a picture.
  var { error } = f.AddPictureFromBytes("Sheet1", "A2", "",
    "Picture 1", ".png", fs.readFileSync('image.png'))
  if (error) {
    console.log(error);
    return
  }
  // Insert a picture to worksheet with scaling.
  var { error } = f.AddPictureFromBytes("Sheet1", "D2",
    `{"x_scale": 0.5, "y_scale": 0.5}`, "Picture 2", ".png",
    fs.readFileSync('image.jpg'))
  if (error) {
    console.log(error);
    return
  }
  // Insert a picture offset in the cell with printing support.
  var { error } = f.AddPictureFromBytes("Sheet1", "H2", `{
      "x_offset": 15,
      "y_offset": 10,
      "print_obj": true,
      "lock_aspect_ratio": false,
      "locked": false
  }`, "Picture 3", ".png", fs.readFileSync('image.gif'))
  if (error) {
    console.log(error);
    return
  }
  // Save spreadsheet by the given path.
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

## Contributing

Contributions are welcome! Open a pull request to fix a bug, or open an issue to discuss a new feature or change.

## Licenses

This program is under the terms of the BSD 3-Clause License. See [https://opensource.org/licenses/BSD-3-Clause](https://opensource.org/licenses/BSD-3-Clause).

The Excel logo is a trademark of [Microsoft Corporation](https://aka.ms/trademarks-usage). This artwork is an adaptation.

gopher.{ai,svg,png} was created by [Takuya Ueda](https://twitter.com/tenntenn). Licensed under the [Creative Commons 3.0 Attributions license](http://creativecommons.org/licenses/by/3.0/).
