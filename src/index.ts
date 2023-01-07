import pako from 'pako';
import { Go } from './Go';
import { getCrypto, getFsOrPolyfill } from './utils';

if (!globalThis.TextEncoder) {
  throw new Error('globalThis.TextEncoder is not available, polyfill required');
}

if (!globalThis.TextDecoder) {
  throw new Error('globalThis.TextDecoder is not available, polyfill required');
}

export async function init(wasmPath: string) {
  const encoder = new TextEncoder();
  const decoder = new TextDecoder('utf-8');
  const [fs, crypto] = await Promise.all([getFsOrPolyfill(decoder), getCrypto()]);
  const go = new Go(fs, crypto, encoder, decoder);
  globalThis.go = go; // do we need this to be global?
  globalThis.excelize = {};

  let buffer: Uint8Array;
  if (typeof window === 'undefined') {
    buffer = pako.ungzip(fs.readFileSync(wasmPath));
  } else {
    buffer = pako.ungzip(await (await fetch(wasmPath)).arrayBuffer());
  }
  if (buffer[0] === 0x1f && buffer[1] === 0x8b) {
    buffer = pako.ungzip(buffer);
  }
  const result = await WebAssembly.instantiate(buffer, go.importObject);

  await go.run(result.instance);

  return globalThis.excelize;
}
