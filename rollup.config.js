import commonjs from "@rollup/plugin-commonjs";
import { nodeResolve } from "@rollup/plugin-node-resolve";
import terser from "@rollup/plugin-terser";
import dts from "rollup-plugin-dts";
import nodePolyfills from "rollup-plugin-polyfill-node";
import pkg from "./package.json" assert { type: "json" };

const input = "src/index.js";

export default [
  {
    // Plain browser <script>
    input,
    output: {
      file: pkg.exports.script,
      format: "iife",
      generatedCode: "es2015",
      name: "excelizeWASM",
      sourcemap: true,
    },
    plugins: [commonjs(), nodePolyfills(), nodeResolve(), terser()],
  },
  {
    // ES6 module and <script type="module">
    input,
    output: {
      file: pkg.exports.default,
      format: "esm",
      generatedCode: "es2015",
      sourcemap: true,
    },
    plugins: [commonjs(), nodePolyfills(), nodeResolve(), terser()],
  },
  {
    // CommonJS Node module
    input,
    output: {
      file: pkg.exports.require,
      format: "cjs",
      generatedCode: "es2015",
      sourcemap: true,
    },
    external: ["path", "fs"],
    plugins: [commonjs(), nodeResolve(), terser()],
  },
  {
    // Include type definitions
    input: "./src/index.d.ts",
    output: [{ file: "dist/index.d.ts", format: "es" }],
    plugins: [dts()],
  },
];
