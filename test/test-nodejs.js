/// <reference path="../src/index.d.ts"/>
import { readFileSync, rmSync, writeFileSync } from "fs";
import assert from "node:assert/strict";
import * as path from "path";
import { fileURLToPath } from "url";
import { init } from "../dist/main.cjs";

/**
 * @typedef { import("excelize-wasm").Init } Init
 * @typedef { import("excelize-wasm").NewFile } NewFile
 */

const currentDir = path.dirname(fileURLToPath(import.meta.url));
const wasmPath = path.join(currentDir, "../dist/excelize.wasm.gz");
const testFilePath = path.join(currentDir, "test.xlsx");

async function start() {
  /** @type {Init} */
  const excelize = await init(wasmPath);

  assert.equal(!!excelize.NewFile, true);

  const workbook = excelize.NewFile();
  const { index } = workbook.NewSheet("Sheet2");

  workbook.SetCellValue("Sheet1", "B2", 100);
  workbook.SetCellValue("Sheet2", "A2", "Hello world.");

  const { buffer, error } = workbook.WriteToBuffer();
  if (error) {
    console.log(error);
    assert.fail(error);
  }

  writeFileSync(testFilePath, buffer, "utf-8");

  const workbook2 = excelize.OpenReader(readFileSync(testFilePath));

  const { value: b2Value, error: b2Error } = workbook2.GetCellValue("Sheet1", "B2");
  const { value: a2Value, error: a2Error } = workbook2.GetCellValue("Sheet2", "A2");

  assert.equal(b2Error, null);
  assert.equal(a2Error, null);

  assert.equal(b2Value, "100");
  assert.equal(a2Value, "Hello world.");

  rmSync(testFilePath);
}

start();
