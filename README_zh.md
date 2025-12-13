# excelize-wasm

<p align="center"><img width="500" src="https://github.com/xuri/excelize-wasm/raw/main/excelize-wasm.svg" alt="excelize-wasm logo"></p>

<p align="center">
    <a href="https://www.npmjs.com/package/excelize-wasm"><img src="https://img.shields.io/npm/v/excelize-wasm.svg" alt="NPM version"></a>
    <a href="https://github.com/xuri/excelize-wasm/actions/workflows/publish.yml"><img src="https://github.com/xuri/excelize-wasm/actions/workflows/publish.yml/badge.svg" alt="Build Status"></a>
    <a href="https://codecov.io/gh/xuri/excelize-wasm"><img src="https://codecov.io/gh/xuri/excelize-wasm/branch/main/graph/badge.svg" alt="Code Coverage"></a>
    <a href="https://goreportcard.com/report/github.com/xuri/excelize-wasm/cmd"><img src="https://goreportcard.com/badge/github.com/xuri/excelize-wasm/cmd" alt="Go Report Card"></a>
    <a href="https://pkg.go.dev/github.com/xuri/excelize/v2"><img src="https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white" alt="go.dev"></a>
    <a href="https://opensource.org/licenses/BSD-3-Clause"><img src="https://img.shields.io/badge/license-bsd-orange.svg" alt="Licenses"></a>
    <a href="https://www.paypal.com/paypalme/xuri"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>
</p>

Excelize-wasm 是 [Excelize](https://github.com/xuri/excelize) 基础库的 WebAssembly / Javascript 实现，可用于操作 Office Excel 文档，基于 ECMA-376，ISO/IEC 29500 国际标准。可以使用它来读取、写入由 Microsoft Excel&trade; 2007 及以上版本创建的电子表格文档。支持 XLAM / XLSM / XLSX / XLTM / XLTX 等多种文档格式，高度兼容带有样式、图片(表)、透视表、切片器等复杂组件的文档。可应用于各类报表平台、云计算、边缘计算等系统。获取更多信息请访问 [参考文档](https://xuri.me/excelize/)。

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
Node.js | &ge;12.0.0
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
```

### 创建 Excel 文档

下面是一个创建 Excel 文档的简单例子：

```javascript
const { init } = require('excelize-wasm');
const fs = require('fs');

init('./node_modules/excelize-wasm/excelize.wasm.gz').then((excelize) => {
  const f = excelize.NewFile();
  if (f.error) {
    console.log(f.error);
    return;
  }
  // 新建一张工作表
  const { index } = f.NewSheet('Sheet2');
  // 设置单元格的值
  f.SetCellValue('Sheet2', 'A2', 'Hello world.');
  f.SetCellValue('Sheet1', 'B2', 100);
  // 设置工作簿的默认工作表
  f.SetActiveSheet(index);
  // 根据指定路径保存文件
  const { buffer, error } = f.WriteToBuffer();
  if (error) {
    console.log(error);
    return;
  }
  fs.writeFile('Book1.xlsx', buffer, 'binary', (error) => {
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
  <script src="https://<服务器地址>/excelize-wasm/index.js"></script>
</head>
<body>
  <div>
    <button onclick="download()">下载</button>
  </div>
  <script>
  function download() {
    excelizeWASM
      .init('https://<服务器地址>/excelize-wasm/excelize.wasm.gz')
      .then((excelize) => {
        const f = excelize.NewFile();
        if (f.error) {
          console.log(f.error);
          return;
        }
        // 创建一个工作表
        const { index } = f.NewSheet('Sheet2');
        // 设置单元格的值
        f.SetCellValue('Sheet2', 'A2', 'Hello world.');
        f.SetCellValue('Sheet1', 'B2', 100);
        // 设置工作簿的默认工作表
        f.SetActiveSheet(index);
        // 根据指定路径保存文件
        const { buffer, error } = f.WriteToBuffer();
        if (error) {
          console.log(error);
          return;
        }
        const link = document.createElement('a');
        link.download = 'Book1.xlsx';
        link.href = URL.createObjectURL(
          new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          })
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
const { init } = require('excelize-wasm');
const fs = require('fs');

init('./node_modules/excelize-wasm/excelize.wasm.gz').then((excelize) => {
  const f = excelize.OpenReader(fs.readFileSync('Book1.xlsx'));
  if (f.error) {
    console.log(f.error);
    return;
  }
  // 设置单元格的值
  const ret1 = f.GetCellValue('Sheet1', 'B2');
  if (ret1.error) {
    console.log(ret1.error);
    return;
  }
  console.log(ret1.value);
  // 获取 Sheet1 上所有单元格
  const ret2 = f.GetRows('Sheet1');
  if (ret2.error) {
    console.log(ret2.error);
    return;
  }
  ret2.result.forEach((row) => {
    row.forEach((colCell) => {
      process.stdout.write(`${colCell}\t`);
    });
    console.log();
  });
});
```

### 在 Excel 文档中创建图表

使用 Excelize 生成图表十分简单，仅需几行代码。您可以根据工作表中的已有数据构建图表，或向工作表中添加数据并创建图表。

<p align="center"><img width="650" src="https://raw.githubusercontent.com/xuri/excelize-wasm/main/chart.png" alt="使用 excelize-wasm 在 Excel 电子表格文档中创建图表"></p>

```javascript
const { init } = require('excelize-wasm');
const fs = require('fs');

init('./node_modules/excelize-wasm/excelize.wasm.gz').then((excelize) => {
  const f = excelize.NewFile();
  if (f.error) {
    console.log(f.error);
    return;
  }
  [
    [null, 'Apple', 'Orange', 'Pear'],
    ['Small', 2, 3, 3],
    ['Normal', 5, 2, 4],
    ['Large', 6, 7, 8],
  ].forEach((row, idx) => {
    const ret1 = excelize.CoordinatesToCellName(1, idx + 1);
    if (ret1.error) {
      console.log(ret1.error);
      return;
    }
    const res2 = f.SetSheetRow('Sheet1', ret1.cell, row);
    if (res2.error) {
      console.log(res2.error);
      return;
    }
  });
  const ret3 = f.AddChart('Sheet1', 'E1', {
    Type: excelize.Col3DClustered,
    Series: [
      {
        Name: 'Sheet1!$A$2',
        Categories: 'Sheet1!$B$1:$D$1',
        Values: 'Sheet1!$B$2:$D$2',
      },
      {
        Name: 'Sheet1!$A$3',
        Categories: 'Sheet1!$B$1:$D$1',
        Values: 'Sheet1!$B$3:$D$3',
      },
      {
        Name: 'Sheet1!$A$4',
        Categories: 'Sheet1!$B$1:$D$1',
        Values: 'Sheet1!$B$4:$D$4',
      },
    ],
    Title: [{
      Text: 'Fruit 3D Clustered Column Chart',
    }],
  });
  if (ret3.error) {
    console.log(ret3.error);
    return;
  }
  // 根据指定路径保存文件
  const { buffer, error } = f.WriteToBuffer();
  if (error) {
    console.log(error);
    return;
  }
  fs.writeFile('Book1.xlsx', buffer, 'binary', (error) => {
    if (error) {
      console.log(error);
    }
  });
});
```

### 向 Excel 文档中插入图片

```javascript
const { init } = require('excelize-wasm');
const fs = require('fs');

init('./node_modules/excelize-wasm/excelize.wasm.gz').then((excelize) => {
  const f = excelize.OpenReader(fs.readFileSync('Book1.xlsx'));
  if (f.error) {
    console.log(f.error);
    return;
  }
  // 插入图片
  const ret1 = f.AddPictureFromBytes('Sheet1', 'A2', {
    Extension: '.png',
    File: fs.readFileSync('image.png'),
    Format: { AltText: 'Picture 1' },
  });
  if (ret1.error) {
    console.log(ret1.error);
    return;
  }
  // 在工作表中插入图片，并设置图片的缩放比例
  const ret2 = f.AddPictureFromBytes('Sheet1', 'D2', {
    Extension: '.jpg',
    File: fs.readFileSync('image.jpg'),
    Format: { AltText: 'Picture 2', ScaleX: 0.5, ScaleY: 0.5 },
  });
  if (ret2.error) {
    console.log(ret2.error);
    return;
  }
  // 在工作表中插入图片，并设置图片的打印属性
  const ret3 = f.AddPictureFromBytes('Sheet1', 'H2', {
    Extension: '.gif',
    File: fs.readFileSync('image.gif'),
    Format: {
      AltText: 'Picture 3',
      OffsetX: 15,
      OffsetY: 10,
      PrintObject: true,
      LockAspectRatio: false,
      Locked: false,
    },
  });
  if (ret3.error) {
    console.log(ret3.error);
    return;
  }
  // 根据指定路径保存文件
  const { buffer, error } = f.WriteToBuffer();
  if (error) {
    console.log(error);
    return;
  }
  fs.writeFile('Book1.xlsx', buffer, 'binary', (error) => {
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

Go gopher 由 [Renee French](https://go.dev/doc/gopher/README) 创作，遵循 [Creative Commons 4.0 Attributions license](http://creativecommons.org/licenses/by/4.0/) 创作共用授权条款。
