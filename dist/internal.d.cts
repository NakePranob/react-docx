/* tslint:disable */
/* eslint-disable */

declare function build_doc_model_from_bytes(bytes: Uint8Array, options?: any | null): any;

declare function build_doc_model_from_package(_package: any, options?: any | null): string;

declare function build_doc_model_from_package_js(_package: any, options?: any | null): any;

declare function model_to_document_xml_from_json_wasm(model_json: string, base_package?: any | null): string;

declare function package_to_array_buffer_wasm(_package: any): Uint8Array;

declare function parse_docx_wasm(bytes: Uint8Array): any;

declare function serialize_docx_from_json_wasm(model_json: string, base_package?: any | null): Uint8Array;

declare function serialize_docx_wasm(model: any, base_package?: any | null): Uint8Array;

type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly build_doc_model_from_bytes: (a: number, b: number, c: number, d: number) => void;
    readonly build_doc_model_from_package: (a: number, b: number, c: number) => void;
    readonly build_doc_model_from_package_js: (a: number, b: number, c: number) => void;
    readonly can_preserve_image_paint_source_wasm: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly export_document_wasm: (a: number, b: number, c: number) => void;
    readonly extract_document_blocks_wasm: (a: number, b: number, c: number) => void;
    readonly extract_document_blocks_with_options_wasm: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly extract_document_blocks_with_revisions_wasm: (a: number, b: number, c: number) => void;
    readonly extract_document_review_wasm: (a: number, b: number, c: number) => void;
    readonly inspect_document_source_wasm: (a: number, b: number, c: number) => void;
    readonly model_to_document_xml_from_json_wasm: (a: number, b: number, c: number, d: number) => void;
    readonly model_to_document_xml_wasm: (a: number, b: number, c: number) => void;
    readonly package_to_array_buffer_wasm: (a: number, b: number) => void;
    readonly parse_document_bundle_wasm: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly parse_docx_wasm: (a: number, b: number, c: number) => void;
    readonly serialize_docx_from_json_wasm: (a: number, b: number, c: number, d: number) => void;
    readonly serialize_docx_wasm: (a: number, b: number, c: number) => void;
    readonly __wasm_bindgen_func_elem_8501: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_export: (a: number, b: number) => number;
    readonly __wbindgen_export2: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_export3: (a: number) => void;
    readonly __wbindgen_export4: (a: number, b: number) => void;
    readonly __wbindgen_add_to_stack_pointer: (a: number) => number;
    readonly __wbindgen_export5: (a: number, b: number, c: number) => void;
}

type WasmSource = InitInput;
interface DocxParseOptions {
    include3dModels?: boolean;
}
declare function initWasm$1(source?: WasmSource): Promise<InitOutput>;
declare function wasmReviewReady(): boolean;
declare function wasmExtractDocumentReviewSync(model: unknown): unknown;
declare function wasmExtractDocumentReview(model: unknown): Promise<unknown>;
declare function wasmExtractDocumentBlocksWithRevisions(model: unknown): Promise<unknown>;
interface WasmOoxmlPart {
    name: string;
    content: string;
}
interface WasmOoxmlPackage {
    parts: Record<string, WasmOoxmlPart>;
    binaryAssets: Record<string, Uint8Array>;
    warnings?: string[];
}
/** Package shape produced by pre-Uint8Array versions of this library. */
interface LegacyWasmOoxmlPackage {
    parts: Record<string, WasmOoxmlPart>;
    binaryAssets: Record<string, number[]>;
    warnings?: string[];
}
declare function docModelToWasmJson(model: unknown): string;
declare function wasmPackageToMaps(raw: WasmOoxmlPackage | LegacyWasmOoxmlPackage): {
    parts: Map<string, WasmOoxmlPart>;
    binaryAssets: Map<string, Uint8Array>;
    warnings: string[];
};
declare function mapsToWasmPackage(input: {
    parts: Map<string, WasmOoxmlPart>;
    binaryAssets: Map<string, Uint8Array>;
    warnings?: string[];
}): WasmOoxmlPackage;
declare function wasmParseDocx(bytes: ArrayBuffer | Uint8Array): Promise<WasmOoxmlPackage>;
declare function wasmBuildDocModelFromPackage(pkg: WasmOoxmlPackage, options?: DocxParseOptions): Promise<unknown>;
declare function wasmInspectDocumentSource(bytes: ArrayBuffer | Uint8Array): Promise<unknown>;
declare function wasmExtractDocumentBlocks(model: unknown, options?: {
    includeWords?: boolean;
    include3dModels?: boolean;
}): Promise<unknown>;
declare function wasmBuildDocModelFromBytes(bytes: ArrayBuffer | Uint8Array, options?: DocxParseOptions): Promise<{
    package: WasmOoxmlPackage;
    model: unknown;
    timings: {
        parseMs: number;
        buildModelMs: number;
    };
}>;
declare function wasmCanPreserveImagePaintSource(image: unknown, basePackage?: WasmOoxmlPackage, ownerPartName?: string): Promise<boolean>;
declare function wasmSerializeDocx(model: unknown, basePackage?: WasmOoxmlPackage): Promise<ArrayBuffer>;
declare function wasmModelToDocumentXml(model: unknown, basePackage?: WasmOoxmlPackage): Promise<string>;
declare function wasmPackageToArrayBuffer(pkg: WasmOoxmlPackage): Promise<ArrayBuffer>;

type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;
declare function setWasmSource(source: WasmSource): void;
declare function initWasm(source?: WasmSource): ReturnType<typeof initWasm$1>;
declare function canUseConfiguredWasmSourceInWorker(): boolean;
declare function getConfiguredWorkerWasmSource(): WorkerWasmSource | undefined;

export { type DocxParseOptions, type LegacyWasmOoxmlPackage, type WasmOoxmlPackage, type WasmOoxmlPart, type WasmSource, type WorkerWasmSource, build_doc_model_from_bytes, build_doc_model_from_package, build_doc_model_from_package_js, canUseConfiguredWasmSourceInWorker, docModelToWasmJson, getConfiguredWorkerWasmSource, initWasm, mapsToWasmPackage, model_to_document_xml_from_json_wasm, package_to_array_buffer_wasm, parse_docx_wasm, serialize_docx_from_json_wasm, serialize_docx_wasm, setWasmSource, wasmBuildDocModelFromBytes, wasmBuildDocModelFromPackage, wasmCanPreserveImagePaintSource, wasmExtractDocumentBlocks, wasmExtractDocumentBlocksWithRevisions, wasmExtractDocumentReview, wasmExtractDocumentReviewSync, wasmInspectDocumentSource, wasmModelToDocumentXml, wasmPackageToArrayBuffer, wasmPackageToMaps, wasmParseDocx, wasmReviewReady, wasmSerializeDocx };
