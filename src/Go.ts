import { Fs } from './utils';

export class Go {
  exited = false;
  argv = ['js'];
  env = {};
  mem: DataView;
  // Imported frmo node or polyfilled for browser
  fs: Fs;
  // Customized version of crypto to support node and browser
  crypto: any; // { getRandomValues(b: Uint8Array): void };

  _inst: WebAssembly.Instance;
  // JS values that Go currently has references to, indexed by reference id
  _values = [NaN, 0, null, true, false, globalThis, this];
  _goRefCounts = new Array(this._values.length).fill(Infinity);
  // mapping from JS values to reference ids
  private _ids = new Map<any, number>([
    [0, 1],
    [null, 2],
    [true, 3],
    [false, 4],
    [globalThis, 5],
    [this, 6],
  ]);
  _pendingEvent: { id: string; this: Go; args: IArguments; result: unknown };
  _idPool;
  _scheduledTimeouts = new Map();
  _nextCallbackTimeoutID;
  _resolveExitPromise: (value?: unknown) => void;
  _exitPromise = new Promise((resolve) => {
    this._resolveExitPromise = resolve;
  });
  encoder: TextEncoder;
  decoder: TextDecoder;
  timeOrigin = Date.now() - performance.now();

  constructor(fs: any, crypto, encoder: TextEncoder, decoder: TextDecoder) {
    this.fs = fs;
    this.crypto = crypto;
    this.encoder = encoder;
    this.decoder = decoder;
  }

  _resume() {
    console.log('_resume called');
    if (this.exited) {
      throw new Error('Go program has already exited');
    }
    if (typeof this._inst.exports.resume === 'function') {
      this._inst.exports.resume();
    }
    if (this.exited) {
      this._resolveExitPromise();
    }
  }

  _makeFuncWrapper(id: string) {
    console.log('_makeFuncWrapper called');
    const go = this;
    return function () {
      const event = { id, this: this, args: arguments, result: undefined };
      go._pendingEvent = event;
      go._resume();
      return event.result;
    };
  }

  exit(code: number): void {
    console.log('exit called');
    if (code !== 0) {
      console.warn('exit code:', code);
    }
  }

  setInt64(addr: number, v: any): void {
    console.log('setInt64 called');
    this.mem.setUint32(addr + 0, v, true);
    this.mem.setUint32(addr + 4, Math.floor(v / 4294967296), true);
  }

  getInt64(addr: number): number {
    console.log('getInt64 called');
    const low = this.mem.getUint32(addr + 0, true);
    const high = this.mem.getInt32(addr + 4, true);
    return low + high * 4294967296;
  }

  loadValue(addr: number): number | Uint8Array | Uint8ClampedArray {
    console.log('loadValue called');
    const f = this.mem.getFloat64(addr, true);
    if (f === 0) {
      return undefined;
    }
    if (!isNaN(f)) {
      return f;
    }

    const id = this.mem.getUint32(addr, true);
    return this._values[id];
  }

  storeValue(addr: number, v: any): void {
    console.log('storeValue called');
    const nanHead = 0x7ff80000;

    if (typeof v === 'number' && v !== 0) {
      if (isNaN(v)) {
        this.mem.setUint32(addr + 4, nanHead, true);
        this.mem.setUint32(addr, 0, true);
        return;
      }
      this.mem.setFloat64(addr, v, true);
      return;
    }

    if (v === undefined) {
      this.mem.setFloat64(addr, 0, true);
      return;
    }

    let id = this._ids.get(v);
    if (id === undefined) {
      id = this._idPool.pop();
      if (id === undefined) {
        id = this._values.length;
      }
      this._values[id] = v;
      this._goRefCounts[id] = 0;
      this._ids.set(v, id);
    }
    this._goRefCounts[id]++;
    let typeFlag = 0;
    switch (typeof v) {
      case 'object':
        if (v !== null) {
          typeFlag = 1;
        }
        break;
      case 'string':
        typeFlag = 2;
        break;
      case 'symbol':
        typeFlag = 3;
        break;
      case 'function':
        typeFlag = 4;
        break;
    }
    this.mem.setUint32(addr + 4, nanHead | typeFlag, true);
    this.mem.setUint32(addr, id, true);
  }

  loadSlice(addr: number): Uint8Array {
    console.log('loadSlice called');
    const array = this.getInt64(addr + 0);
    const len = this.getInt64(addr + 8);
    return new Uint8Array((this._inst.exports.mem as WebAssembly.Memory).buffer, array, len);
  }

  loadSliceOfValues(addr: number): number[] {
    console.log('loadSliceOfValues called');
    const array = this.getInt64(addr + 0);
    const len = this.getInt64(addr + 8);
    const a = new Array(len);
    for (let i = 0; i < len; i++) {
      a[i] = this.loadValue(array + i * 8);
    }
    return a;
  }

  loadString(addr: number): string {
    console.log('loadString called');
    const saddr = this.getInt64(addr + 0);
    const len = this.getInt64(addr + 8);
    return this.decoder.decode(new DataView((this._inst.exports.mem as WebAssembly.Memory).buffer, saddr, len));
  }

  async run(instance: Go) {
    console.log('run called');
    if (!(instance instanceof WebAssembly.Instance)) {
      throw new Error('Go.run: WebAssembly.Instance expected');
    }
    this._inst = instance;
    this.mem = new DataView((this._inst.exports.mem as WebAssembly.Memory).buffer);
    this._values = [
      // JS values that Go currently has references to, indexed by reference id
      NaN,
      0,
      null,
      true,
      false,
      globalThis,
      this,
    ];
    this._goRefCounts = new Array(this._values.length).fill(Infinity); // number of references that Go has to a JS value, indexed by reference id
    // mapping from JS values to reference ids
    this._ids = new Map<any, number>([
      [0, 1],
      [null, 2],
      [true, 3],
      [false, 4],
      [globalThis, 5],
      [this, 6],
    ]);
    this._idPool = []; // unused ids that have been garbage collected
    this.exited = false; // whether the Go program has exited

    // Pass command line arguments and environment variables to WebAssembly by writing them to the linear memory.
    let offset = 4096;

    const strPtr = (str) => {
      console.log('strPtr called');
      const ptr = offset;
      const bytes = this.encoder.encode(str + '\0');
      new Uint8Array(this.mem.buffer, offset, bytes.length).set(bytes);
      offset += bytes.length;
      if (offset % 8 !== 0) {
        offset += 8 - (offset % 8);
      }
      return ptr;
    };

    const argc = this.argv.length;

    const argvPtrs = [];
    this.argv.forEach((arg) => {
      argvPtrs.push(strPtr(arg));
    });
    argvPtrs.push(0);

    const keys = Object.keys(this.env).sort();
    keys.forEach((key) => {
      argvPtrs.push(strPtr(`${key}=${this.env[key]}`));
    });
    argvPtrs.push(0);

    const argv = offset;
    argvPtrs.forEach((ptr) => {
      this.mem.setUint32(offset, ptr, true);
      this.mem.setUint32(offset + 4, 0, true);
      offset += 8;
    });

    // The linker guarantees global data starts from at least wasmMinDataAddr.
    // Keep in sync with cmd/link/internal/ld/data.go:wasmMinDataAddr.
    const wasmMinDataAddr = 4096 + 8192;
    if (offset >= wasmMinDataAddr) {
      throw new Error('total length of command line and environment variables exceeds limit');
    }

    if (typeof this._inst.exports.run === 'function') {
      this._inst.exports.run(argc, argv);
    }
    if (this.exited) {
      this._resolveExitPromise();
    }
    await this._exitPromise;
  }

  importObject = {
    go: {
      // func wasmExit(code int32)
      'runtime.wasmExit': (sp: number) => {
        console.log('runtime.wasmExit');
        sp >>>= 0;
        const code = this.mem.getInt32(sp + 8, true);
        this.exited = true;
        this._inst = undefined;
        this._values = undefined;
        this._goRefCounts = undefined;
        this._ids = undefined;
        this._idPool = undefined;
        this.exit(code);
      },

      // func wasmWrite(fd uintptr, p unsafe.Pointer, n int32)
      'runtime.wasmWrite': (sp: number) => {
        console.log('runtime.wasmWrite');
        sp >>>= 0;
        const fd = this.getInt64(sp + 8);
        const p = this.getInt64(sp + 16);
        const n = this.mem.getInt32(sp + 24, true);
        this.fs.writeSync(fd, new Uint8Array((this._inst.exports.mem as WebAssembly.Memory).buffer, p, n));
      },

      // func resetMemoryDataView()
      'runtime.resetMemoryDataView': (sp: number) => {
        console.log('runtime.resetMemoryDataView');
        sp >>>= 0;
        this.mem = new DataView((this._inst.exports.mem as WebAssembly.Memory).buffer);
      },

      // func nanotime1() int64
      'runtime.nanotime1': (sp: number) => {
        console.log('runtime.nanotime1');
        sp >>>= 0;
        this.setInt64(sp + 8, (this.timeOrigin + performance.now()) * 1000000);
      },

      // func walltime() (sec int64, nsec int32)
      'runtime.walltime': (sp: number) => {
        console.log('runtime.walltime');
        sp >>>= 0;
        const msec = new Date().getTime();
        this.setInt64(sp + 8, msec / 1000);
        this.mem.setInt32(sp + 16, (msec % 1000) * 1000000, true);
      },

      // func scheduleTimeoutEvent(delay int64) int32
      'runtime.scheduleTimeoutEvent': (sp: number) => {
        console.log('runtime.scheduleTimeoutEvent');
        sp >>>= 0;
        const id = this._nextCallbackTimeoutID;
        this._nextCallbackTimeoutID++;
        this._scheduledTimeouts.set(
          id,
          setTimeout(
            () => {
              this._resume();
              while (this._scheduledTimeouts.has(id)) {
                // for some reason Go failed to register the timeout event, log and try again
                // (temporary workaround for https://github.com/golang/go/issues/28975)
                console.warn('scheduleTimeoutEvent: missed timeout event');
                this._resume();
              }
            },
            this.getInt64(sp + 8) + 1, // setTimeout has been seen to fire up to 1 millisecond early
          ),
        );
        this.mem.setInt32(sp + 16, id, true);
      },

      // func clearTimeoutEvent(id int32)
      'runtime.clearTimeoutEvent': (sp: number) => {
        console.log('runtime.clearTimeoutEvent');
        sp >>>= 0;
        const id = this.mem.getInt32(sp + 8, true);
        clearTimeout(this._scheduledTimeouts.get(id));
        this._scheduledTimeouts.delete(id);
      },

      // func getRandomData(r []byte)
      'runtime.getRandomData': (sp: number) => {
        console.log('runtime.getRandomData');
        sp >>>= 0;
        this.crypto.getRandomValues(this.loadSlice(sp + 8));
      },

      // func finalizeRef(v ref)
      'syscall/js.finalizeRef': (sp: number) => {
        console.log('syscall/js.finalizeRef called');
        sp >>>= 0;
        const id = this.mem.getUint32(sp + 8, true);
        this._goRefCounts[id]--;
        if (this._goRefCounts[id] === 0) {
          const v = this._values[id];
          this._values[id] = null;
          this._ids.delete(v);
          this._idPool.push(id);
        }
      },

      // func stringVal(value string) ref
      'syscall/js.stringVal': (sp: number) => {
        console.log('syscall/js.stringVal called');
        sp >>>= 0;
        this.storeValue(sp + 24, this.loadString(sp + 8));
      },

      // func valueGet(v ref, p string) ref
      'syscall/js.valueGet': (sp: number) => {
        console.log('syscall/js.valueGet called');
        sp >>>= 0;
        const result = Reflect.get(this.loadValue(sp + 8) as any, this.loadString(sp + 16));
        if (typeof this._inst.exports.getsp === 'function') {
          sp = this._inst.exports.getsp() >>> 0; // see comment above
        }
        this.storeValue(sp + 32, result);
      },

      // func valueSet(v ref, p string, x ref)
      'syscall/js.valueSet': (sp: number) => {
        console.log('syscall/js.valueSet called');
        sp >>>= 0;
        Reflect.set(this.loadValue(sp + 8) as any, this.loadString(sp + 16), this.loadValue(sp + 32));
      },

      // func valueDelete(v ref, p string)
      'syscall/js.valueDelete': (sp: number) => {
        console.log('syscall/js.valueDelete called');
        sp >>>= 0;
        Reflect.deleteProperty(this.loadValue(sp + 8) as any, this.loadString(sp + 16));
      },

      // func valueIndex(v ref, i int) ref
      'syscall/js.valueIndex': (sp: number) => {
        console.log('syscall/js.valueIndex called');
        sp >>>= 0;
        this.storeValue(sp + 24, Reflect.get(this.loadValue(sp + 8) as any, this.getInt64(sp + 16)));
      },

      // valueSetIndex(v ref, i int, x ref)
      'syscall/js.valueSetIndex': (sp: number) => {
        console.log('syscall/js.valueSetIndex called');
        sp >>>= 0;
        Reflect.set(this.loadValue(sp + 8) as any, this.getInt64(sp + 16), this.loadValue(sp + 24));
      },

      // func valueCall(v ref, m string, args []ref) (ref, bool)
      'syscall/js.valueCall': (sp: number) => {
        console.log('syscall/js.valueCall called');
        sp >>>= 0;
        try {
          const v = this.loadValue(sp + 8);
          const m = Reflect.get(v as any, this.loadString(sp + 16));
          const args = this.loadSliceOfValues(sp + 32);
          const result = Reflect.apply(m, v, args);
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 56, result);
          this.mem.setUint8(sp + 64, 1);
        } catch (err) {
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 56, err);
          this.mem.setUint8(sp + 64, 0);
        }
      },

      // func valueInvoke(v ref, args []ref) (ref, bool)
      'syscall/js.valueInvoke': (sp: number) => {
        console.log('syscall/js.valueInvoke called');
        sp >>>= 0;
        try {
          const v = this.loadValue(sp + 8);
          const args = this.loadSliceOfValues(sp + 16);
          const result = Reflect.apply(v as any, undefined, args);
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 40, result);
          this.mem.setUint8(sp + 48, 1);
        } catch (err) {
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 40, err);
          this.mem.setUint8(sp + 48, 0);
        }
      },

      // func valueNew(v ref, args []ref) (ref, bool)
      'syscall/js.valueNew': (sp: number) => {
        console.log('syscall/js.valueNew called');
        sp >>>= 0;
        try {
          const v = this.loadValue(sp + 8);
          const args = this.loadSliceOfValues(sp + 16);
          const result = Reflect.construct(v as any, args);
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 40, result);
          this.mem.setUint8(sp + 48, 1);
        } catch (err) {
          if (typeof this._inst.exports.getsp === 'function') {
            sp = this._inst.exports.getsp() >>> 0; // see comment above
          }
          this.storeValue(sp + 40, err);
          this.mem.setUint8(sp + 48, 0);
        }
      },

      // func valueLength(v ref) int
      'syscall/js.valueLength': (sp: number) => {
        console.log('syscall/js.valueLength called');
        sp >>>= 0;
        const v = this.loadValue(sp + 8);
        this.setInt64(sp + 16, parseInt(String(Array.isArray(v) ? v.length : v || 0)));
      },

      // valuePrepareString(v ref) (ref, int)
      'syscall/js.valuePrepareString': (sp: number) => {
        console.log('syscall/js.valuePrepareString called');
        sp >>>= 0;
        const str = this.encoder.encode(String(this.loadValue(sp + 8)));
        this.storeValue(sp + 16, str);
        this.setInt64(sp + 24, str.length);
      },

      // valueLoadString(v ref, b []byte)
      'syscall/js.valueLoadString': (sp: number) => {
        console.log('syscall/js.valueLoadString called');
        sp >>>= 0;
        const str = this.loadValue(sp + 8);
        this.loadSlice(sp + 16).set(str as ArrayLike<number>);
      },

      // func valueInstanceOf(v ref, t ref) bool
      'syscall/js.valueInstanceOf': (sp: number) => {
        console.log('syscall/js.valueInstanceOf called');
        sp >>>= 0;
        this.mem.setUint8(sp + 24, this.loadValue(sp + 8) instanceof (this.loadValue(sp + 16) as any) ? 1 : 0);
      },

      // func copyBytesToGo(dst []byte, src ref) (int, bool)
      'syscall/js.copyBytesToGo': (sp: number) => {
        console.log('syscall/js.copyBytesToGo called');
        sp >>>= 0;
        const dst = this.loadSlice(sp + 8);
        const src = this.loadValue(sp + 32);
        if (!(src instanceof Uint8Array || src instanceof Uint8ClampedArray)) {
          this.mem.setUint8(sp + 48, 0);
          return;
        }
        const toCopy = src.subarray(0, dst.length);
        dst.set(toCopy);
        this.setInt64(sp + 40, toCopy.length);
        this.mem.setUint8(sp + 48, 1);
      },

      // func copyBytesToJS(dst ref, src []byte) (int, bool)
      'syscall/js.copyBytesToJS': (sp: number) => {
        console.log('syscall/js.copyBytesToJS called');
        sp >>>= 0;
        const dst = this.loadValue(sp + 8);
        const src = this.loadSlice(sp + 16);
        if (!(dst instanceof Uint8Array || dst instanceof Uint8ClampedArray)) {
          this.mem.setUint8(sp + 48, 0);
          return;
        }
        const toCopy = src.subarray(0, dst.length);
        dst.set(toCopy);
        this.setInt64(sp + 40, toCopy.length);
        this.mem.setUint8(sp + 48, 1);
      },

      debug: (value: any) => {
        console.log(value);
      },
    },
  };
}
