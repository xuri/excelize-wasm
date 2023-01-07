import commonjs from '@rollup/plugin-commonjs';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import nodePolyfills from 'rollup-plugin-polyfill-node';
import terser from '@rollup/plugin-terser';
import typescript from '@rollup/plugin-typescript';
import dts from 'rollup-plugin-dts';
import pkg from './package.json' assert { type: 'json' };
import stripCode from 'rollup-plugin-strip-code';

const input = 'src/index.ts';

// Removes import of node modules from browser bundles
const stripCodePlugin = stripCode({
  start_comment: 'START.NODE_ONLY',
  end_comment: 'END.NODE_ONLY',
});

export default [
  // Plain browser <script>
  {
    input,
    output: {
      file: pkg.exports.script,
      file: './dist/index.js',
      format: 'iife',
      generatedCode: 'es2015',
      name: 'excelizeWASM',
      sourcemap: true,
    },
    plugins: [commonjs(), nodeResolve(), stripCodePlugin, typescript()],
  },

  // ES6 module and <script type="module">
  {
    input,
    output: {
      file: pkg.exports.default,
      format: 'esm',
      generatedCode: 'es2015',
      sourcemap: true,
    },
    plugins: [commonjs(), nodeResolve(), stripCodePlugin, typescript()],
  },

  // CommonJS Node module
  {
    input,
    output: {
      file: pkg.exports.require,
      format: 'cjs',
      generatedCode: 'es2015',
      sourcemap: true,
    },
    external: ['path', 'fs', 'crypto'],
    plugins: [commonjs(), nodeResolve(), typescript()],
  },

  // Include type definitions
  {
    input: './src/index.d.ts',
    output: [{ file: 'dist/index.d.ts', format: 'es' }],
    plugins: [dts()],
  },
];
