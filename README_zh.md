# excelize-wasm

<p align="center"><img width="500" src="https://github.com/xuri/excelize-wasm/raw/main/excelize-wasm.svg" alt="excelize-wasm logo"></p>

<p align="center">
    <a href="https://www.npmjs.com/package/excelize-wasm"><img src="https://img.shields.io/npm/v/excelize-wasm.svg" alt="NPM version"></a>
    <a href="https://github.com/xuri/excelize-wasm/actions/workflows/go.yml"><img src="https://github.com/xuri/excelize-wasm/actions/workflows/go.yml/badge.svg" alt="Build Status"></a>
    <a href="https://codecov.io/gh/xuri/excelize-wasm"><img src="https://codecov.io/gh/xuri/excelize-wasm/branch/main/graph/badge.svg" alt="Code Coverage"></a>
    <a href="https://goreportcard.com/report/github.com/xuri/excelize-wasm/cmd"><img src="https://goreportcard.com/badge/github.com/xuri/excelize-wasm/cmd" alt="Go Report Card"></a>
    <a href="https://pkg.go.dev/github.com/xuri/excelize/v2"><img src="https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white" alt="go.dev"></a>
    <a href="https://opensource.org/licenses/BSD-3-Clause"><img src="https://img.shields.io/badge/license-bsd-orange.svg" alt="Licenses"></a>
    <a href="https://www.paypal.com/paypalme/xuri"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>
</p>

Excelize-wasm 是基于 WebAssembly / Javascript 实现的 Go [Excelize](https://github.com/xuri/excelize) 基础库，用于操作 Office Excel 文档基础库，基于 ECMA-376，ISO/IEC 29500 国际标准。可以使用它来读取、写入由 Microsoft Excel&trade; 2007 及以上版本创建的电子表格文档。支持 XLAM / XLSM / XLSX / XLTM / XLTX 等多种文档格式，高度兼容带有样式、图片(表)、透视表、切片器等复杂组件的文档。可应用于各类报表平台、云计算、边缘计算等系统。获取更多信息请访问 [参考文档](https://xuri.me/excelize/)。

## 运行环境兼容性

运行环境 | 版本要求
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

## 快速上手

### 安装

#### Node.js

```bash
npm install --save excelize-wasm
```

#### 浏览器

```html
<script src="excelize-wasm/index.js"></script>
````

### 创建 Excel 文档

下面是一个创建 Excel 文档的简单例子：

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = NewFile();
  // 创建一个工作表
  const index = f.NewSheet("Sheet2")
  // 设置单元格的值
  f.SetCellValue("Sheet2", "A2", "Hello world.")
  f.SetCellValue("Sheet1", "B2", 100)
  // 设置工作簿的默认工作表
  f.SetActiveSheet(index)
  // 根据指定路径保存文件
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

在浏览器中创建 Excel 并下载：

<details>
  <summary>查看代码</summary>

```html
<html>
<head>
  <meta charset="utf-8">
  <script src="excelize-wasm/index.js"></script>
</head>
<body>
  <div>
    <button onclick="download()">下载</button>
  </div>
  <script>
  function download() {
    excelize('https://<服务器地址>/excelize-wasm/excelize.wasm.gz').then(() => {
      const f = NewFile();
      // 创建一个工作表
      const index = f.NewSheet("Sheet2")
      // 设置单元格的值
      f.SetCellValue("Sheet2", "A2", "Hello world.")
      f.SetCellValue("Sheet1", "B2", 100)
      // 设置工作簿的默认工作表
      f.SetActiveSheet(index)
      // 根据指定路径保存文件
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

### 读取 Excel 文档

下面是读取 Excel 文档的例子：

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = OpenReader(fs.readFileSync('Book1.xlsx'));
  // 创建一个工作表
  const index = f.NewSheet("Sheet2")
  // 设置单元格的值
  var { cell, error } = f.GetCellValue("Sheet1", "B2")
  if (error) {
    console.log(error);
    return;
  }
  console.log(cell)
  // 获取 Sheet1 上所有单元格
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

### 在 Excel 文档中创建图表

使用 Excelize 生成图表十分简单，仅需几行代码。您可以根据工作表中的已有数据构建图表，或向工作表中添加数据并创建图表。

<p align="center"><img width="650" src="https://raw.githubusercontent.com/xuri/excelize-wasm/main/chart.png" alt="使用 excelize-wasm 在 Excel 电子表格文档中创建图表"></p>

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
  // 根据指定路径保存文件
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

### 向 Excel 文档中插入图片

```javascript
require('excelize-wasm');
const fs = require('fs');

excelize('excelize.wasm.gz').then(() => {
  const f = OpenReader(fs.readFileSync('Book1.xlsx'));
  if (f.error) {
    console.log(f.error);
    return
  }
  // 插入图片
  var { error } = f.AddPictureFromBytes("Sheet1", "A2", "",
    "Picture 1", ".png", fs.readFileSync('image.png'))
  if (error) {
    console.log(error);
    return
  }
  // 在工作表中插入图片，并设置图片的缩放比例
  var { error } = f.AddPictureFromBytes("Sheet1", "D2",
    `{"x_scale": 0.5, "y_scale": 0.5}`, "Picture 2", ".png",
    fs.readFileSync('image.jpg'))
  if (error) {
    console.log(error);
    return
  }
  // 在工作表中插入图片，并设置图片的打印属性
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
  // 根据指定路径保存文件
  fs.writeFile('Book1.xlsx', f.WriteToBuffer(), 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

## 社区合作

欢迎您为此项目贡献代码，提出建议或问题、修复 Bug 以及参与讨论对新功能的想法。

## 开源许可

本项目遵循 BSD 3-Clause 开源许可协议，访问 [https://opensource.org/licenses/BSD-3-Clause](https://opensource.org/licenses/BSD-3-Clause) 查看许可协议文件。

Excel 徽标是 [Microsoft Corporation](https://aka.ms/trademarks-usage) 的商标，项目的图片是一种改编。

gopher.{ai,svg,png} 由 [Takuya Ueda](https://twitter.com/tenntenn) 创作，遵循 [Creative Commons 3.0 Attributions license](http://creativecommons.org/licenses/by/3.0/) 创作共用授权条款。
