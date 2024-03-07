import commonjs from '@rollup/plugin-commonjs';
import nodePolyfills from 'rollup-plugin-polyfill-node';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import terser from '@rollup/plugin-terser';
import pkg from './package.json' assert {type: 'json'};
const input = 'src/index.js';

export default [
    {
        // Plain browser <script>
        input,
        output: {
            file: pkg.exports.script,
            format: 'iife',
            generatedCode: 'es2015',
            name: 'excelizeWASM',
            sourcemap: false,
        },
        plugins: [
            commonjs(),
            nodePolyfills(),
            nodeResolve(),
            terser()
        ]
    },
    {
        // ES6 module and <script type="module">
        input,
        output: {
            file: pkg.exports.default,
            format: 'esm',
            generatedCode: 'es2015',
            sourcemap: false,
        },
        plugins: [
            commonjs(),
            nodePolyfills(),
            nodeResolve(),
            terser()
        ]
    },
    {
        // CommonJS Node module
        input,
        output: {
            file: pkg.exports.require,
            format: 'cjs',
            generatedCode: 'es2015',
            sourcemap: false,
            dynamicImportInCjs: false
        },
        external: ['path', 'fs'],
        plugins: [
            commonjs(),
            nodeResolve(),
            terser()
        ]
    }
];
