type Callback = (err: Error | null, value?: any) => void;

export type Fs = Awaited<ReturnType<typeof getFsOrPolyfill>>;

export class EnosysError extends Error {
  code: string;
  constructor() {
    super('not implemented');
    this.code = 'ENOSYS';
  }
  static createEnosysError() {
    return new EnosysError();
  }
}

export function addPolyfills() {
  if (!globalThis.TextEncoder) {
    throw new Error('globalThis.TextEncoder is not available, polyfill required');
  }

  if (!globalThis.TextDecoder) {
    throw new Error('globalThis.TextDecoder is not available, polyfill required');
  }

  if (!globalThis.process) {
    (globalThis as any).process = {
      getuid() {
        return -1;
      },
      getgid() {
        return -1;
      },
      geteuid() {
        return -1;
      },
      getegid() {
        return -1;
      },
      getgroups() {
        throw new EnosysError();
      },
      pid: -1,
      ppid: -1,
      umask() {
        throw new EnosysError();
      },
      cwd() {
        throw new EnosysError();
      },
      chdir() {
        throw new EnosysError();
      },
    };
  }
}

export async function getCrypto() {
  /*START.NODE_ONLY*/
  if (typeof globalThis.window === 'undefined') {
    const nodeCrypto = await import('crypto');
    return {
      getRandomValues: (b: any) => nodeCrypto.randomFillSync(b),
    };
  }
  /*END.NODE_ONLY*/

  return window.crypto;
}

export async function getFsOrPolyfill(decoder: TextDecoder) {
  /*START.NODE_ONLY*/
  if (typeof globalThis.window === 'undefined') {
    return await import('fs');
  }
  /*END.NODE_ONLY*/

  let outputBuf = '';
  return Promise.resolve({
    constants: { O_WRONLY: -1, O_RDWR: -1, O_CREAT: -1, O_TRUNC: -1, O_APPEND: -1, O_EXCL: -1 }, // unused
    writeSync(fd: number, buf: BufferSource, p?: unknown, n?: unknown) {
      outputBuf += decoder.decode(buf);
      const nl = outputBuf.lastIndexOf('\n');
      if (nl != -1) {
        console.log(outputBuf.substring(0, nl));
        outputBuf = outputBuf.substring(nl + 1);
      }
      return buf.byteLength;
    },
    write(fd: unknown, buf, offset: number, length: number, position: number | null, callback: Callback) {
      if (offset !== 0 || length !== buf.length || position !== null) {
        callback(new EnosysError());
        return;
      }
      const n = this.writeSync(fd, buf);
      callback(null, n);
    },
    readFileSync(path: unknown) {
      throw new EnosysError();
    },
    chmod(path: unknown, mode: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    chown(path: unknown, uid: unknown, gid: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    close(fd: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    fchmod(fd: unknown, mode: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    fchown(fd: unknown, uid: unknown, gid: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    fstat(fd: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    fsync(fd: unknown, callback: Callback) {
      callback(null);
    },
    ftruncate(fd: unknown, length: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    lchown(path: unknown, uid: unknown, gid: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    link(path: unknown, link: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    lstat(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    mkdir(path: unknown, perm: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    open(path: unknown, flags: unknown, mode: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    read(fd: unknown, buffer: unknown, offset: unknown, length: unknown, position: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    readdir(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    readlink(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    rename(from: unknown, to: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    rmdir(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    stat(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    symlink(path: unknown, link: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    truncate(path: unknown, length: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    unlink(path: unknown, callback: Callback) {
      callback(new EnosysError());
    },
    utimes(path: unknown, atime: unknown, mtime: unknown, callback: Callback) {
      callback(new EnosysError());
    },
  });
}
