import commonjs from '@rollup/plugin-commonjs';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import terser from '@rollup/plugin-terser';
const input = 'src/index.js';

export default [
    {
        // Plain browser <script>
        input,
        output: {
            file: './dist/index.js',
            format: 'iife',
            generatedCode: 'es2015',
            name: 'excelizeWASM',
            sourcemap: false,
        },
        plugins: [
            nodeResolve(),
            terser()
        ]
    },
    {
        // ES6 module and <script type="module">
        input,
        output: {
            file: './dist/main.js',
            format: 'esm',
            generatedCode: 'es2015',
            sourcemap: false,
        },
        plugins: [
            nodeResolve(),
            terser()
        ]
    },
    {
        // CommonJS Node module
        input,
        output: {
            file: './dist/main.cjs',
            format: 'cjs',
            generatedCode: 'es2015',
            sourcemap: false,
        },
        external: ['path', 'fs'],
        plugins: [
            commonjs(),
            nodeResolve(),
            terser()
        ]
    }
];
