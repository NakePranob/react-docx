"use strict";

// ../../node_modules/.pnpm/tsup@8.5.1_jiti@2.6.1_postcss@8.5.32_typescript@5.9.3/node_modules/tsup/assets/cjs_shims.js
var getImportMetaUrl = () => typeof document === "undefined" ? new URL(`file:${__filename}`).href : document.currentScript && document.currentScript.tagName.toUpperCase() === "SCRIPT" ? document.currentScript.src : new URL("main.js", document.baseURI).href;
var importMetaUrl = /* @__PURE__ */ getImportMetaUrl();

// ../wasm/src/generated/docx_wasm.js
function build_doc_model_from_bytes(bytes, options) {
  try {
    const retptr = wasm.__wbindgen_add_to_stack_pointer(-16);
    const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_export);
    const len0 = WASM_VECTOR_LEN;
    wasm.build_doc_model_from_bytes(retptr, ptr0, len0, isLikeNone(options) ? 0 : addHeapObject(options));
    var r0 = getDataViewMemory0().getInt32(retptr + 4 * 0, true);
    var r1 = getDataViewMemory0().getInt32(retptr + 4 * 1, true);
    var r2 = getDataViewMemory0().getInt32(retptr + 4 * 2, true);
    if (r2) {
      throw takeObject(r1);
    }
    return takeObject(r0);
  } finally {
    wasm.__wbindgen_add_to_stack_pointer(16);
  }
}
function __wbg_get_imports() {
  const import0 = {
    __proto__: null,
    __wbg___wbindgen_debug_string_56c147eb1a51f0c4: function(arg0, arg1) {
      const ret = debugString(getObject(arg1));
      const ptr1 = passStringToWasm0(ret, wasm.__wbindgen_export, wasm.__wbindgen_export2);
      const len1 = WASM_VECTOR_LEN;
      getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
      getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
    },
    __wbg___wbindgen_is_null_ced4761460071341: function(arg0) {
      const ret = getObject(arg0) === null;
      return ret;
    },
    __wbg___wbindgen_is_object_3a2c414391dbf751: function(arg0) {
      const val = getObject(arg0);
      const ret = typeof val === "object" && val !== null;
      return ret;
    },
    __wbg___wbindgen_is_string_6541b0f6ecd4e8e5: function(arg0) {
      const ret = typeof getObject(arg0) === "string";
      return ret;
    },
    __wbg___wbindgen_is_undefined_4410e3c20a99fa97: function(arg0) {
      const ret = getObject(arg0) === void 0;
      return ret;
    },
    __wbg___wbindgen_string_get_fa2687d531ed17a5: function(arg0, arg1) {
      const obj = getObject(arg1);
      const ret = typeof obj === "string" ? obj : void 0;
      var ptr1 = isLikeNone(ret) ? 0 : passStringToWasm0(ret, wasm.__wbindgen_export, wasm.__wbindgen_export2);
      var len1 = WASM_VECTOR_LEN;
      getDataViewMemory0().setInt32(arg0 + 4 * 1, len1, true);
      getDataViewMemory0().setInt32(arg0 + 4 * 0, ptr1, true);
    },
    __wbg___wbindgen_throw_bbadd78c1bac3a77: function(arg0, arg1) {
      throw new Error(getStringFromWasm0(arg0, arg1));
    },
    __wbg__wbg_cb_unref_c2301a3c9b78104b: function(arg0) {
      getObject(arg0)._wbg_cb_unref();
    },
    __wbg_assign_99ad8a391f999526: function(arg0, arg1) {
      const ret = Object.assign(getObject(arg0), getObject(arg1));
      return addHeapObject(ret);
    },
    __wbg_concat_c71bf85e52813325: function(arg0, arg1) {
      const ret = getObject(arg0).concat(getObject(arg1));
      return addHeapObject(ret);
    },
    __wbg_create_00e11a7cefe80760: function(arg0) {
      const ret = Object.create(getObject(arg0));
      return addHeapObject(ret);
    },
    __wbg_defineProperties_8e58c44ab7b28a65: function(arg0, arg1) {
      const ret = Object.defineProperties(getObject(arg0), getObject(arg1));
      return addHeapObject(ret);
    },
    __wbg_defineProperty_b2fe86d27f6dfc3f: function() {
      return handleError(function(arg0, arg1, arg2) {
        const ret = Reflect.defineProperty(getObject(arg0), getObject(arg1), getObject(arg2));
        return ret;
      }, arguments);
    },
    __wbg_from_8a57180716c586ee: function(arg0) {
      const ret = Array.from(getObject(arg0));
      return addHeapObject(ret);
    },
    __wbg_getOwnPropertyDescriptors_7a1e06e1a2828998: function(arg0) {
      const ret = Object.getOwnPropertyDescriptors(getObject(arg0));
      return addHeapObject(ret);
    },
    __wbg_get_52a8a619f7b88df6: function() {
      return handleError(function(arg0, arg1) {
        const ret = Reflect.get(getObject(arg0), getObject(arg1));
        return addHeapObject(ret);
      }, arguments);
    },
    __wbg_get_unchecked_46e778e3cec74b5e: function(arg0, arg1) {
      const ret = getObject(arg0)[arg1 >>> 0];
      return addHeapObject(ret);
    },
    __wbg_instanceof_ArrayBuffer_a581da923203f29f: function(arg0) {
      let result;
      try {
        result = getObject(arg0) instanceof ArrayBuffer;
      } catch (_) {
        result = false;
      }
      const ret = result;
      return ret;
    },
    __wbg_instanceof_Uint8Array_b6fe1ac89eba107e: function(arg0) {
      let result;
      try {
        result = getObject(arg0) instanceof Uint8Array;
      } catch (_) {
        result = false;
      }
      const ret = result;
      return ret;
    },
    __wbg_isArray_139f48e3c057ede8: function(arg0) {
      const ret = Array.isArray(getObject(arg0));
      return ret;
    },
    __wbg_is_edd40bac2fc20257: function(arg0, arg1) {
      const ret = Object.is(getObject(arg0), getObject(arg1));
      return ret;
    },
    __wbg_keys_bd51ff67a9b04698: function(arg0) {
      const ret = Object.keys(getObject(arg0));
      return addHeapObject(ret);
    },
    __wbg_length_68a9d5278d084f4f: function(arg0) {
      const ret = getObject(arg0).length;
      return ret;
    },
    __wbg_length_7abf942ae12c9744: function(arg0) {
      const ret = getObject(arg0).length;
      return ret;
    },
    __wbg_length_fb04d16d7bdf6d4c: function(arg0) {
      const ret = getObject(arg0).length;
      return ret;
    },
    __wbg_new_0b303268aa395a38: function() {
      const ret = new Array();
      return addHeapObject(ret);
    },
    __wbg_new_20b778a4c5c691c3: function() {
      const ret = new Object();
      return addHeapObject(ret);
    },
    __wbg_new_b06772b280cc6e52: function(arg0) {
      const ret = new Uint8Array(getObject(arg0));
      return addHeapObject(ret);
    },
    __wbg_new_from_slice_bb2d1778c0b87eb1: function(arg0, arg1) {
      const ret = new Uint8Array(getArrayU8FromWasm0(arg0, arg1));
      return addHeapObject(ret);
    },
    __wbg_now_5a9864021b76bf46: function() {
      const ret = performance.now();
      return ret;
    },
    __wbg_parse_246201845d0eb98c: function() {
      return handleError(function(arg0, arg1) {
        const ret = JSON.parse(getStringFromWasm0(arg0, arg1));
        return addHeapObject(ret);
      }, arguments);
    },
    __wbg_prototypesetcall_956c7493c68e29b4: function(arg0, arg1, arg2) {
      Uint8Array.prototype.set.call(getArrayU8FromWasm0(arg0, arg1), getObject(arg2));
    },
    __wbg_push_ceb8ef046afb2041: function(arg0, arg1) {
      const ret = getObject(arg0).push(getObject(arg1));
      return ret;
    },
    __wbg_set_a6ba3ac0e634b822: function() {
      return handleError(function(arg0, arg1, arg2) {
        const ret = Reflect.set(getObject(arg0), getObject(arg1), getObject(arg2));
        return ret;
      }, arguments);
    },
    __wbg_stringify_4408512926f9a81c: function() {
      return handleError(function(arg0, arg1) {
        const ret = JSON.stringify(getObject(arg0), getObject(arg1));
        return addHeapObject(ret);
      }, arguments);
    },
    __wbg_stringify_a42c95ea9a7591c9: function() {
      return handleError(function(arg0) {
        const ret = JSON.stringify(getObject(arg0));
        return addHeapObject(ret);
      }, arguments);
    },
    __wbindgen_cast_0000000000000001: function(arg0, arg1) {
      const ret = makeMutClosure(arg0, arg1, __wasm_bindgen_func_elem_8501);
      return addHeapObject(ret);
    },
    __wbindgen_cast_0000000000000002: function(arg0) {
      const ret = arg0;
      return addHeapObject(ret);
    },
    __wbindgen_cast_0000000000000003: function(arg0, arg1) {
      const ret = getStringFromWasm0(arg0, arg1);
      return addHeapObject(ret);
    },
    __wbindgen_object_clone_ref: function(arg0) {
      const ret = getObject(arg0);
      return addHeapObject(ret);
    },
    __wbindgen_object_drop_ref: function(arg0) {
      takeObject(arg0);
    }
  };
  return {
    __proto__: null,
    "./docx_wasm_bg.js": import0
  };
}
function __wasm_bindgen_func_elem_8501(arg0, arg1, arg2, arg3) {
  const ret = wasm.__wasm_bindgen_func_elem_8501(arg0, arg1, addHeapObject(arg2), addHeapObject(arg3));
  return takeObject(ret);
}
function addHeapObject(obj) {
  if (heap_next === heap.length) heap.push(heap.length + 1);
  const idx = heap_next;
  heap_next = heap[idx];
  heap[idx] = obj;
  return idx;
}
var CLOSURE_DTORS = typeof FinalizationRegistry === "undefined" ? { register: () => {
}, unregister: () => {
} } : new FinalizationRegistry((state) => wasm.__wbindgen_export4(state.a, state.b));
function debugString(val) {
  const type = typeof val;
  if (type == "number" || type == "boolean" || val == null) {
    return `${val}`;
  }
  if (type == "string") {
    return `"${val}"`;
  }
  if (type == "symbol") {
    const description = val.description;
    if (description == null) {
      return "Symbol";
    } else {
      return `Symbol(${description})`;
    }
  }
  if (type == "function") {
    const name = val.name;
    if (typeof name == "string" && name.length > 0) {
      return `Function(${name})`;
    } else {
      return "Function";
    }
  }
  if (Array.isArray(val)) {
    const length = val.length;
    let debug = "[";
    if (length > 0) {
      debug += debugString(val[0]);
    }
    for (let i = 1; i < length; i++) {
      debug += ", " + debugString(val[i]);
    }
    debug += "]";
    return debug;
  }
  const builtInMatches = /\[object ([^\]]+)\]/.exec(toString.call(val));
  let className;
  if (builtInMatches && builtInMatches.length > 1) {
    className = builtInMatches[1];
  } else {
    return toString.call(val);
  }
  if (className == "Object") {
    try {
      return "Object(" + JSON.stringify(val) + ")";
    } catch (_) {
      return "Object";
    }
  }
  if (val instanceof Error) {
    return `${val.name}: ${val.message}
${val.stack}`;
  }
  return className;
}
function dropObject(idx) {
  if (idx < 1028) return;
  heap[idx] = heap_next;
  heap_next = idx;
}
function getArrayU8FromWasm0(ptr, len) {
  ptr = ptr >>> 0;
  return getUint8ArrayMemory0().subarray(ptr / 1, ptr / 1 + len);
}
var cachedDataViewMemory0 = null;
function getDataViewMemory0() {
  if (cachedDataViewMemory0 === null || cachedDataViewMemory0.buffer.detached === true || cachedDataViewMemory0.buffer.detached === void 0 && cachedDataViewMemory0.buffer !== wasm.memory.buffer) {
    cachedDataViewMemory0 = new DataView(wasm.memory.buffer);
  }
  return cachedDataViewMemory0;
}
function getStringFromWasm0(ptr, len) {
  return decodeText(ptr >>> 0, len);
}
var cachedUint8ArrayMemory0 = null;
function getUint8ArrayMemory0() {
  if (cachedUint8ArrayMemory0 === null || cachedUint8ArrayMemory0.byteLength === 0) {
    cachedUint8ArrayMemory0 = new Uint8Array(wasm.memory.buffer);
  }
  return cachedUint8ArrayMemory0;
}
function getObject(idx) {
  return heap[idx];
}
function handleError(f, args) {
  try {
    return f.apply(this, args);
  } catch (e) {
    wasm.__wbindgen_export3(addHeapObject(e));
  }
}
var heap = new Array(1024).fill(void 0);
heap.push(void 0, null, true, false);
var heap_next = heap.length;
function isLikeNone(x) {
  return x === void 0 || x === null;
}
function makeMutClosure(arg0, arg1, f) {
  const state = { a: arg0, b: arg1, cnt: 1 };
  const real = (...args) => {
    state.cnt++;
    const a = state.a;
    state.a = 0;
    try {
      return f(a, state.b, ...args);
    } finally {
      state.a = a;
      real._wbg_cb_unref();
    }
  };
  real._wbg_cb_unref = () => {
    if (--state.cnt === 0) {
      wasm.__wbindgen_export4(state.a, state.b);
      state.a = 0;
      CLOSURE_DTORS.unregister(state);
    }
  };
  CLOSURE_DTORS.register(real, state, state);
  return real;
}
function passArray8ToWasm0(arg, malloc) {
  const ptr = malloc(arg.length * 1, 1) >>> 0;
  getUint8ArrayMemory0().set(arg, ptr / 1);
  WASM_VECTOR_LEN = arg.length;
  return ptr;
}
function passStringToWasm0(arg, malloc, realloc) {
  if (realloc === void 0) {
    const buf = cachedTextEncoder.encode(arg);
    const ptr2 = malloc(buf.length, 1) >>> 0;
    getUint8ArrayMemory0().subarray(ptr2, ptr2 + buf.length).set(buf);
    WASM_VECTOR_LEN = buf.length;
    return ptr2;
  }
  let len = arg.length;
  let ptr = malloc(len, 1) >>> 0;
  const mem = getUint8ArrayMemory0();
  let offset = 0;
  for (; offset < len; offset++) {
    const code = arg.charCodeAt(offset);
    if (code > 127) break;
    mem[ptr + offset] = code;
  }
  if (offset !== len) {
    if (offset !== 0) {
      arg = arg.slice(offset);
    }
    ptr = realloc(ptr, len, len = offset + arg.length * 3, 1) >>> 0;
    const view = getUint8ArrayMemory0().subarray(ptr + offset, ptr + len);
    const ret = cachedTextEncoder.encodeInto(arg, view);
    offset += ret.written;
    ptr = realloc(ptr, len, offset, 1) >>> 0;
  }
  WASM_VECTOR_LEN = offset;
  return ptr;
}
function takeObject(idx) {
  const ret = getObject(idx);
  dropObject(idx);
  return ret;
}
var cachedTextDecoder = new TextDecoder("utf-8", { ignoreBOM: true, fatal: true });
cachedTextDecoder.decode();
var MAX_SAFARI_DECODE_BYTES = 2146435072;
var numBytesDecoded = 0;
function decodeText(ptr, len) {
  numBytesDecoded += len;
  if (numBytesDecoded >= MAX_SAFARI_DECODE_BYTES) {
    cachedTextDecoder = new TextDecoder("utf-8", { ignoreBOM: true, fatal: true });
    cachedTextDecoder.decode();
    numBytesDecoded = len;
  }
  return cachedTextDecoder.decode(getUint8ArrayMemory0().subarray(ptr, ptr + len));
}
var cachedTextEncoder = new TextEncoder();
if (!("encodeInto" in cachedTextEncoder)) {
  cachedTextEncoder.encodeInto = function(arg, view) {
    const buf = cachedTextEncoder.encode(arg);
    view.set(buf);
    return {
      read: arg.length,
      written: buf.length
    };
  };
}
var WASM_VECTOR_LEN = 0;
var wasmModule;
var wasmInstance;
var wasm;
function __wbg_finalize_init(instance, module2) {
  wasmInstance = instance;
  wasm = instance.exports;
  wasmModule = module2;
  cachedDataViewMemory0 = null;
  cachedUint8ArrayMemory0 = null;
  return wasm;
}
async function __wbg_load(module2, imports) {
  if (typeof Response === "function" && module2 instanceof Response) {
    if (typeof WebAssembly.instantiateStreaming === "function") {
      try {
        return await WebAssembly.instantiateStreaming(module2, imports);
      } catch (e) {
        const validResponse = module2.ok && expectedResponseType(module2.type);
        if (validResponse && module2.headers.get("Content-Type") !== "application/wasm") {
          console.warn("`WebAssembly.instantiateStreaming` failed because your server does not serve Wasm with `application/wasm` MIME type. Falling back to `WebAssembly.instantiate` which is slower. Original error:\n", e);
        } else {
          throw e;
        }
      }
    }
    const bytes = await module2.arrayBuffer();
    return await WebAssembly.instantiate(bytes, imports);
  } else {
    const instance = await WebAssembly.instantiate(module2, imports);
    if (instance instanceof WebAssembly.Instance) {
      return { instance, module: module2 };
    } else {
      return instance;
    }
  }
  function expectedResponseType(type) {
    switch (type) {
      case "basic":
      case "cors":
      case "default":
        return true;
    }
    return false;
  }
}
async function __wbg_init(module_or_path) {
  if (wasm !== void 0) return wasm;
  if (module_or_path !== void 0) {
    if (Object.getPrototypeOf(module_or_path) === Object.prototype) {
      ({ module_or_path } = module_or_path);
    } else {
      console.warn("using deprecated parameters for the initialization function; pass a single object instead");
    }
  }
  if (module_or_path === void 0) {
    module_or_path = new URL("docx_wasm_bg.wasm", importMetaUrl);
  }
  const imports = __wbg_get_imports();
  if (typeof module_or_path === "string" || typeof Request === "function" && module_or_path instanceof Request || typeof URL === "function" && module_or_path instanceof URL) {
    module_or_path = fetch(module_or_path);
  }
  const { instance, module: module2 } = await __wbg_load(await module_or_path, imports);
  return __wbg_finalize_init(instance, module2);
}

// ../wasm/src/index.ts
var initPromise;
var overrideSource;
var initialized;
function setWasmSource(source) {
  if (initPromise) {
    throw new Error(
      "react-docx: setWasmSource must be called before the first parse/serialize call initializes WASM"
    );
  }
  overrideSource = source;
}
async function nodeFsPromises() {
  const proc = globalThis.process;
  if (typeof proc?.getBuiltinModule === "function") {
    return proc.getBuiltinModule("node:fs/promises");
  }
  const dynamicImport = new Function("specifier", "return import(specifier)");
  return dynamicImport("node:fs/promises");
}
async function defaultWasmSource() {
  const wasmUrl = new URL("./docx_wasm_bg.wasm", importMetaUrl);
  if (wasmUrl.protocol === "file:") {
    const fs = await nodeFsPromises();
    return fs.readFile(wasmUrl);
  }
  const response = await fetch(wasmUrl);
  if (!response.ok) {
    throw new Error(
      `react-docx: failed to load the bundled WebAssembly binary from ${wasmUrl.href} (${response.status} ${response.statusText}). If your bundler did not emit this asset, import @extend-ai/react-docx/docx_wasm_bg.wasm?url and pass it to setWasmSource() before importing a document.`
    );
  }
  return response;
}
async function initWasm(source) {
  if (!initPromise) {
    const chosen = source ?? overrideSource;
    initPromise = (chosen !== void 0 ? Promise.resolve(chosen) : defaultWasmSource()).then((module_or_path) => __wbg_init({ module_or_path })).then((output) => initialized = output).catch((error) => {
      if (error instanceof WebAssembly.CompileError) {
        throw new Error(
          `react-docx: the bundled WebAssembly binary failed to compile. It requires WebAssembly SIMD support (Chrome 91+, Firefox 89+, Safari 16.4+, Node 16.4+). Original error: ${error.message}`,
          { cause: error }
        );
      }
      throw error;
    });
  }
  return initPromise;
}
function toUint8Array(value) {
  return value instanceof Uint8Array ? value : Uint8Array.from(value);
}
function wasmPackageToMaps(raw) {
  const parts = /* @__PURE__ */ new Map();
  for (const [name, part] of Object.entries(raw.parts ?? {})) {
    parts.set(name, {
      name: part.name ?? name,
      content: part.content
    });
  }
  const binaryAssets = /* @__PURE__ */ new Map();
  for (const [name, asset] of Object.entries(raw.binaryAssets ?? {})) {
    binaryAssets.set(name, toUint8Array(asset));
  }
  return {
    parts,
    binaryAssets,
    warnings: Array.isArray(raw.warnings) ? [...raw.warnings] : []
  };
}
async function wasmBuildDocModelFromBytes(bytes, options) {
  await initWasm();
  const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return build_doc_model_from_bytes(payload, options);
}

// ../doc-model/src/form-field.ts
function resolveFormFieldSourcePreview(field) {
  const preview = field.sourcePreview;
  if (field.fieldType !== "checkbox" || typeof field.sourceXml !== "string" || !field.sourceXml.trim() || !preview || typeof preview.relationshipId !== "string" || !preview.relationshipId.trim() || typeof preview.shapeId !== "string" || !preview.shapeId.trim() || typeof preview.checked !== "boolean" || preview.checked !== Boolean(field.checked ?? field.widget?.checkbox?.defaultChecked)) {
    return void 0;
  }
  const image = preview.image;
  if (!image || image.type !== "image") {
    return void 0;
  }
  const hasSource = typeof image.src === "string" && image.src.trim().length > 0;
  const hasData = (image.data instanceof Uint8Array || Array.isArray(image.data)) && image.data.length > 0;
  return hasSource || hasData ? image : void 0;
}
function invalidateFormFieldSource(field) {
  return { ...field, sourceXml: void 0, sourcePreview: void 0 };
}

// ../doc-model/src/opendocument.ts
var OPEN_DOCUMENT_BODY_PART = "content.xml";
function isOpenDocumentModel(model) {
  return model.metadata?.sourcePartNames?.document === OPEN_DOCUMENT_BODY_PART;
}
function bytesToBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString("base64");
  }
  if (typeof btoa !== "function") {
    return void 0;
  }
  let binary = "";
  const chunkSize = 32768;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}
function imageDataSource(image) {
  const contentType = image.contentType?.trim().toLowerCase();
  if (!image.data?.byteLength || !contentType || !/^image\/[a-z0-9.+-]+$/.test(contentType)) {
    return void 0;
  }
  const base64 = bytesToBase64(image.data);
  return base64 ? `data:${contentType};base64,${base64}` : void 0;
}

// ../doc-model/src/normalize.ts
function normalizeUint8Array(value) {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (Array.isArray(value)) {
    return Uint8Array.from(value);
  }
  return void 0;
}
function normalizeParagraphChild(child, context) {
  const { invalidated } = context;
  if (child.type === "form-field" && child.sourcePreview) {
    const image2 = resolveFormFieldSourcePreview(child);
    if (!image2) {
      const stateChanged = typeof child.sourcePreview.checked === "boolean" && (child.fieldType !== "checkbox" || child.sourcePreview.checked !== Boolean(child.checked ?? child.widget?.checkbox?.defaultChecked));
      const normalized = stateChanged ? invalidateFormFieldSource(child) : { ...child, sourcePreview: void 0 };
      if (stateChanged) invalidated.add(normalized);
      return normalized;
    }
    const normalizedImage = normalizeParagraphChild(image2, context);
    return normalizedImage === image2 ? child : { ...child, sourcePreview: { ...child.sourcePreview, image: normalizedImage } };
  }
  if (child.type !== "image") {
    return child;
  }
  const image = child;
  const data = normalizeUint8Array(image.data);
  const modelAsset = image.model3d?.asset;
  const modelData = normalizeUint8Array(modelAsset?.data);
  const model3d = modelAsset && modelData !== modelAsset.data ? { ...image.model3d, asset: { ...modelAsset, data: modelData } } : image.model3d;
  const src = context.fillImageSources && !image.src && !image.contentNodes ? imageDataSource({ data, contentType: image.contentType }) : image.src;
  if (data === image.data && model3d === image.model3d && src === image.src && !image.contentNodes) {
    return child;
  }
  return {
    ...image,
    ...src === image.src ? {} : { src },
    data,
    model3d,
    contentNodes: image.contentNodes?.map((node) => normalizeDocNode(node, context))
  };
}
function normalizeTableCellContent(node, context) {
  return normalizeDocNode(node, context);
}
function normalizeDocNode(node, context) {
  const { invalidated } = context;
  if (node.type === "paragraph") {
    const children = node.children.map((child) => normalizeParagraphChild(child, context));
    const stateChanged2 = invalidated.size > 0 && children.some((child) => invalidated.has(child));
    const normalized2 = {
      ...node,
      reviewContent: node.reviewContent?.map((fragment) => ({
        ...fragment,
        children: fragment.children.map((child) => normalizeParagraphChild(child, context))
      })),
      ...stateChanged2 ? { sourceXml: void 0, sourceTextPatch: void 0 } : {},
      children
    };
    if (stateChanged2) invalidated.add(normalized2);
    return normalized2;
  }
  const rows = node.rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => ({
      ...cell,
      nodes: cell.nodes.map((content) => normalizeTableCellContent(content, context))
    }))
  }));
  const stateChanged = invalidated.size > 0 && rows.some((row) => row.cells.some((cell) => cell.nodes.some((content) => invalidated.has(content))));
  const normalized = {
    ...node,
    reviewRows: node.reviewRows?.map((fragment) => ({
      ...fragment,
      row: normalizeDocNode({ type: "table", rows: [fragment.row] }, context).rows[0]
    })),
    ...stateChanged ? { sourceXml: void 0, sourceTextPatches: void 0 } : {},
    rows
  };
  if (stateChanged) invalidated.add(normalized);
  return normalized;
}
function normalizeDocModel(model) {
  const context = {
    invalidated: /* @__PURE__ */ new Set(),
    fillImageSources: isOpenDocumentModel(model)
  };
  const normalize = (node) => normalizeDocNode(node, context);
  return {
    ...model,
    nodes: model.nodes.map(normalize),
    metadata: {
      ...model.metadata,
      headerSections: model.metadata.headerSections.map((section) => ({
        ...section,
        nodes: section.nodes.map(normalize)
      })),
      footerSections: model.metadata.footerSections.map((section) => ({
        ...section,
        nodes: section.nodes.map(normalize)
      })),
      sections: model.metadata.sections?.map((section) => ({
        ...section,
        headerSections: section.headerSections.map((header) => ({
          ...header,
          nodes: header.nodes.map(normalize)
        })),
        footerSections: section.footerSections.map((footer) => ({
          ...footer,
          nodes: footer.nodes.map(normalize)
        }))
      })),
      footnotes: model.metadata.footnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(normalize)
      })),
      comments: model.metadata.comments?.map((comment) => ({
        ...comment,
        nodes: comment.nodes?.map(normalize)
      })),
      endnotes: model.metadata.endnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(normalize)
      }))
    }
  };
}

// ../doc-model/src/index.ts
async function buildDocModelFromBytes(bytes, options) {
  const startedAt = performance.now();
  const result = await wasmBuildDocModelFromBytes(bytes, options);
  const builtAt = performance.now();
  const { parts, binaryAssets, warnings } = wasmPackageToMaps(result.package);
  const pkg = { parts, binaryAssets, ...warnings.length > 0 ? { warnings } : {} };
  const model = normalizeDocModel(result.model);
  const finishedAt = performance.now();
  const buildModelMs = result.timings.buildModelMs + finishedAt - builtAt;
  return {
    package: pkg,
    model,
    timings: {
      parseMs: Math.max(0, builtAt - startedAt - result.timings.buildModelMs),
      buildModelMs,
      totalMs: finishedAt - startedAt
    }
  };
}

// src/import-transferables.ts
function collectImportTransferables(value) {
  const buffers = /* @__PURE__ */ new Set();
  const visited = /* @__PURE__ */ new Set();
  const pending = [value];
  while (pending.length > 0) {
    const current = pending.pop();
    if (!current || typeof current !== "object" || visited.has(current)) {
      continue;
    }
    visited.add(current);
    if (current instanceof ArrayBuffer) {
      buffers.add(current);
    } else if (ArrayBuffer.isView(current)) {
      if (current.buffer instanceof ArrayBuffer) {
        buffers.add(current.buffer);
      }
    } else if (current instanceof Map) {
      for (const [key, entry] of current) {
        pending.push(key, entry);
      }
    } else {
      for (const entry of Object.values(current)) {
        pending.push(entry);
      }
    }
  }
  return [...buffers];
}

// src/docx-import-worker.ts
function performanceNow() {
  return typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();
}
function serializeError(error) {
  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
      stack: error.stack
    };
  }
  return {
    message: String(error)
  };
}
self.addEventListener(
  "message",
  async (event) => {
    const request = event.data;
    if (!request || request.type !== "import-docx") {
      return;
    }
    try {
      if (request.wasmSource !== void 0) {
        setWasmSource(request.wasmSource);
      }
      const startedAt = performanceNow();
      const result = await buildDocModelFromBytes(request.buffer, request.parseOptions);
      const finishedAt = performanceNow();
      const timings = {
        totalMs: finishedAt - startedAt,
        parseMs: result.timings.parseMs,
        buildModelMs: result.timings.buildModelMs
      };
      const response = {
        id: request.id,
        type: "success",
        package: result.package,
        model: result.model,
        timings
      };
      self.postMessage(response, {
        transfer: collectImportTransferables(response)
      });
    } catch (error) {
      const response = {
        id: request.id,
        type: "error",
        error: serializeError(error)
      };
      self.postMessage(response);
    }
  }
);
var ready = { type: "ready" };
self.postMessage(ready);
//# sourceMappingURL=docx-import-worker.cjs.map