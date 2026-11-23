import * as React from 'react';
import { Properties } from 'csstype';

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

interface OoxmlPart {
    name: string;
    content: string;
}
interface OoxmlPackage {
    parts: Map<string, OoxmlPart>;
    binaryAssets: Map<string, Uint8Array>;
    /** Diagnostics produced while adapting a non-OOXML source such as legacy DOC. */
    warnings?: string[];
}
declare function parseDocx(input: ArrayBuffer): Promise<OoxmlPackage>;
declare function packageToArrayBuffer(pkg: OoxmlPackage): Promise<ArrayBuffer>;
declare function createMinimalDocxPackage(documentXml?: string): OoxmlPackage;
declare function getPart(pkg: OoxmlPackage, partName: string): OoxmlPart | undefined;
declare function withPart(pkg: OoxmlPackage, part: OoxmlPart): OoxmlPackage;

type ParagraphAlignment = "left" | "center" | "right" | "justify";
type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
interface TextRunBorderStyle {
    type: string;
    color?: string;
    sizeEighthPt?: number;
    spacePt?: number;
    frame?: boolean;
    shadow?: boolean;
}
type TextLigatures = "none" | "standard" | "contextual" | "historical" | "discretional" | "standardContextual" | "standardHistorical" | "contextualHistorical" | "standardDiscretional" | "contextualDiscretional" | "historicalDiscretional" | "standardContextualHistorical" | "standardContextualDiscretional" | "standardHistoricalDiscretional" | "contextualHistoricalDiscretional" | "all";
interface TextStyle {
    textShadow?: TextShadow;
    textOutline?: TextOutline;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    allCaps?: boolean;
    smallCaps?: boolean;
    color?: string;
    highlight?: string;
    backgroundColor?: string;
    fontSizePt?: number;
    fontSizeCsPt?: number;
    /** Resolved single-family fallback retained for backwards compatibility. */
    fontFamily?: string;
    /** @internal Import-time fallback used to detect fontFamily edits. */
    sourceFontFamily?: string;
    /** Explicit OOXML w:rFonts slots. */
    fontFamilyAscii?: string;
    fontFamilyHAnsi?: string;
    fontFamilyEastAsia?: string;
    fontFamilyCs?: string;
    fontThemeAscii?: string;
    fontThemeHAnsi?: string;
    fontThemeEastAsia?: string;
    fontThemeCs?: string;
    /** @internal Concrete families resolved from the document theme per slot. */
    resolvedFontFamilyAscii?: string;
    /** @internal Concrete families resolved from the document theme per slot. */
    resolvedFontFamilyHAnsi?: string;
    /** @internal Concrete families resolved from the document theme per slot. */
    resolvedFontFamilyEastAsia?: string;
    /** @internal Concrete families resolved from the document theme per slot. */
    resolvedFontFamilyCs?: string;
    fontHint?: string;
    /** OOXML w:lang slots and complex-script direction controls. */
    language?: string;
    languageEastAsia?: string;
    languageBidi?: string;
    rightToLeft?: boolean;
    complexScript?: boolean;
    characterSpacingTwips?: number;
    /** Horizontal character width as an integer percentage; 100 is unscaled. */
    characterScalePercent?: number;
    sourceCharacterScale?: string;
    /** Minimum font size in points at which OOXML kerning is enabled. */
    kerningMinFontSizePt?: number;
    ligatures?: TextLigatures;
    contextualAlternates?: boolean;
    verticalAlign?: "superscript" | "subscript";
    runBorder?: TextRunBorderStyle;
}
interface TextShadow {
    color: string;
    opacity: number;
    offsetXPx: number;
    offsetYPx: number;
    blurRadiusPx: number;
    sourceXml?: string;
    sourceColor?: string;
}
interface TextOutline {
    color: string;
    opacity: number;
    widthPx: number;
    lineJoin: "round" | "bevel" | "miter";
    sourceXml?: string;
    sourceColor?: string;
}
interface TextRunNode {
    type: "text";
    /**
     * Run text. For a run with `fieldInstruction`, Word files leave this empty
     * and OpenDocument files store the field's saved result (e.g. "7").
     */
    text: string;
    /**
     * Instruction of a dynamic field such as `PAGE` or `NUMPAGES`. The viewer
     * shows the live value in place of `text` when it can resolve one.
     */
    fieldInstruction?: string;
    style?: TextStyle;
    link?: string;
    noteReference?: {
        kind: "footnote" | "endnote";
        id: number;
    };
}
interface ImagePaintBounds {
    xPx: number;
    yPx: number;
    widthPx: number;
    heightPx: number;
    layoutWidthPx: number;
    layoutHeightPx: number;
}
interface ChartPoint {
    index: number;
    categories?: {
        level: number;
        label: string;
    }[];
    value?: string;
    x?: string;
    y?: string;
    bubbleSize?: string;
}
interface ChartData {
    title?: string | null;
    series: {
        name?: string | null;
        chartType: string;
        points: ChartPoint[];
    }[];
}
interface ImagePropertyAttribute {
    name: string;
    value: string;
}
interface ImagePropertyNode {
    name: string;
    attributes?: ImagePropertyAttribute[];
    children?: ImagePropertyNode[];
    text?: string;
}
interface ImageAssetReference {
    element: string;
    kind: string;
    relationshipId?: string;
    target?: string;
    embeddedAssetAvailable: boolean;
}
interface ImageTransform {
    rotationDegrees?: number;
    flipHorizontal?: boolean;
    flipVertical?: boolean;
    offsetXEmu?: number;
    offsetYEmu?: number;
    widthEmu?: number;
    heightEmu?: number;
}
interface ImageTile {
    offsetXEmu?: number;
    offsetYEmu?: number;
    scaleX?: number;
    scaleY?: number;
    alignment?: string;
    flip?: string;
}
interface ImageCrop {
    leftFraction?: number;
    topFraction?: number;
    rightFraction?: number;
    bottomFraction?: number;
}
interface ImagePresentation {
    transformApplied?: boolean;
    assets?: ImageAssetReference[];
    transform?: ImageTransform;
    sourceRectangle?: ImageCrop;
    fillRectangle?: ImageCrop;
    tile?: ImageTile;
    properties?: ImagePropertyNode[];
}
interface OfficeModel3D {
    modelRelationshipId?: string;
    linkedModelRelationshipId?: string;
    asset?: {
        partName: string;
        contentType: string;
        data?: Uint8Array;
    };
    preview?: {
        modelRelationshipId?: string | null;
        rasterRelationshipId: string;
        renderer?: string | null;
        rendererVersion?: string | null;
    };
    sourceXml: string;
}
interface ImageRunNode {
    model3d?: OfficeModel3D;
    presentation?: ImagePresentation;
    chartData?: ChartData;
    chartPartName?: string;
    paintBounds?: ImagePaintBounds;
    /** MathML of a formula; `semanticKind` is then "formula". */
    mathMl?: string;
    blockId?: string;
    /**
     * Nested story of a text box or text frame. An OpenDocument text frame has
     * no `src` picture and is drawn from these nodes.
     */
    contentNodes?: DocNode[];
    /** Parser classification, e.g. "text_box" or "formula". */
    semanticKind?: string;
    type: "image";
    /**
     * Picture to draw. Images of OpenDocument files arrive with `data` and
     * `contentType` only; `normalizeDocModel` fills this from them.
     */
    src?: string;
    alt?: string;
    widthPx?: number;
    heightPx?: number;
    partName?: string;
    contentType?: string;
    data?: Uint8Array;
    sourceXml?: string;
    crop?: ImageCrop;
    cssFilter?: string;
    cssOpacity?: number;
    floating?: {
        xPx?: number;
        yPx?: number;
        horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
        verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
        horizontalRelativeTo?: string;
        verticalRelativeTo?: string;
        distLPx?: number;
        distRPx?: number;
        distTPx?: number;
        distBPx?: number;
        wrapType?: "none" | "square" | "tight" | "through" | "topAndBottom";
        wrapText?: "bothSides" | "left" | "right" | "largest";
        behindDocument?: boolean;
        zIndex?: number;
    };
    syntheticTextBox?: boolean;
    textBoxText?: string;
}
type FormFieldType = "checkbox" | "text" | "date" | "dropdown";
type FormFieldSourceKind = "sdt" | "legacy";
interface FormFieldOption {
    displayText: string;
    value?: string;
}
interface FormFieldTextWidgetSettings {
    inputType?: "regular" | "number" | "date" | "currentDate" | "currentTime" | "calculated" | (string & {});
    defaultText?: string;
    maxLength?: number;
    textFormat?: string;
}
interface FormFieldCheckboxWidgetSettings {
    defaultChecked?: boolean;
    sizeMode?: "auto" | "exact";
    sizePt?: number;
}
interface FormFieldDropdownWidgetSettings {
    defaultValue?: string;
}
interface FormFieldWidgetSettings {
    name?: string;
    enabled?: boolean;
    calcOnExit?: boolean;
    text?: FormFieldTextWidgetSettings;
    checkbox?: FormFieldCheckboxWidgetSettings;
    dropdown?: FormFieldDropdownWidgetSettings;
}
interface FormFieldSourcePreview {
    relationshipId: string;
    shapeId: string;
    checked: boolean;
    image: ImageRunNode;
}
interface FormFieldRunNode {
    type: "form-field";
    fieldType: FormFieldType;
    sourceKind?: FormFieldSourceKind;
    id?: number;
    tag?: string;
    title?: string;
    placeholder?: string;
    checked?: boolean;
    value?: string;
    options?: FormFieldOption[];
    widget?: FormFieldWidgetSettings;
    checkedSymbol?: string;
    uncheckedSymbol?: string;
    style?: TextStyle;
    link?: string;
    sourceXml?: string;
    sourcePreview?: FormFieldSourcePreview;
}
type ParagraphChildNode = TextRunNode | ImageRunNode | FormFieldRunNode;
interface ParagraphNumbering {
    numId: number;
    ilvl: number;
}
interface ParagraphSpacing {
    beforeTwips?: number;
    afterTwips?: number;
    beforeAutoSpacing?: boolean;
    afterAutoSpacing?: boolean;
    lineTwips?: number;
    lineRule?: "auto" | "exact" | "atLeast";
}
interface ParagraphIndent {
    leftTwips?: number;
    rightTwips?: number;
    firstLineTwips?: number;
    hangingTwips?: number;
}
interface ParagraphBorderStyle {
    type: string;
    color?: string;
    sizeEighthPt?: number;
    spacePt?: number;
    frame?: boolean;
    shadow?: boolean;
}
interface ParagraphBorderSet {
    top?: ParagraphBorderStyle;
    right?: ParagraphBorderStyle;
    bottom?: ParagraphBorderStyle;
    left?: ParagraphBorderStyle;
    between?: ParagraphBorderStyle;
    bar?: ParagraphBorderStyle;
}
interface ParagraphTabStop {
    alignment?: "clear" | "left" | "center" | "right" | "decimal" | "bar";
    leader?: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
    positionTwips?: number;
}
type ParagraphTextAlignment = "auto" | "top" | "center" | "baseline" | "bottom";
interface ParagraphStyle {
    /** @internal Imported effective automatic before-spacing. */
    sourceBeforeAutoSpacing?: boolean;
    /** @internal Imported effective automatic after-spacing. */
    sourceAfterAutoSpacing?: boolean;
    align?: ParagraphAlignment;
    snapToGrid?: boolean;
    /** @internal Imported effective document-grid snapping. */
    sourceSnapToGrid?: boolean;
    /** @internal Whether the imported paragraph had a direct grid property. */
    sourceHasSnapToGrid?: boolean;
    /** @internal Inherited grid setting exposed when a direct override is cleared. */
    sourceInheritedSnapToGrid?: boolean;
    /** @internal Table-style grid setting before paragraph-style overrides. */
    sourceTableSnapToGrid?: boolean;
    textAlignment?: ParagraphTextAlignment;
    /** @internal Imported effective vertical character alignment. */
    sourceTextAlignment?: ParagraphTextAlignment;
    /** @internal Whether the imported paragraph had a direct alignment property. */
    sourceHasTextAlignment?: boolean;
    /** @internal Inherited alignment exposed when a direct override is cleared. */
    sourceInheritedTextAlignment?: ParagraphTextAlignment;
    /** @internal Table-style alignment before paragraph-style overrides. */
    sourceTableTextAlignment?: ParagraphTextAlignment;
    headingLevel?: HeadingLevel;
    styleId?: string;
    styleName?: string;
    numbering?: ParagraphNumbering;
    spacing?: ParagraphSpacing;
    indent?: ParagraphIndent;
    backgroundColor?: string;
    borders?: ParagraphBorderSet;
    tabStops?: ParagraphTabStop[];
    contextualSpacing?: boolean;
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    pageBreakBefore?: boolean;
    dropCap?: {
        type: "drop" | "margin";
        lines?: number;
        wrap?: string;
        horizontalAnchor?: string;
        verticalAnchor?: string;
        xTwips?: number;
        yTwips?: number;
        horizontalSpaceTwips?: number;
        verticalSpaceTwips?: number;
    };
}
/** @internal Imported mark layers before paragraph-style inheritance. */
interface ParagraphMarkFormattingSource {
    defaultStyle?: TextStyle;
    characterStyle?: TextStyle;
    directStyle?: TextStyle;
}
/**
 * Part of a paragraph that a comment is anchored to, as a range of the
 * paragraph's children. Present on paragraphs that carry no Word source markup
 * (OpenDocument text); for Word files the anchors are read from `sourceXml`.
 */
interface ParagraphCommentRange {
    commentId: number;
    /** Index of the first child inside the range. */
    childStart: number;
    /** Index after the last child inside the range. */
    childEnd: number;
    /** True when the comment's anchor ends in this paragraph. */
    endsHere: boolean;
}
interface ParagraphNode {
    reviewContent?: {
        sourceXml: string;
        children: ParagraphChildNode[];
    }[];
    /** Comment anchors of a paragraph without Word source markup. */
    commentRanges?: ParagraphCommentRange[];
    type: "paragraph";
    /** Stable identity for measurement caches; survives edits via cloneParagraph. */
    blockId?: string;
    children: ParagraphChildNode[];
    style?: ParagraphStyle;
    /** Effective formatting of the paragraph-ending mark. */
    paragraphMarkStyle?: TextStyle;
    /** @internal Imported mark formatting used to detect changes. */
    sourceParagraphMarkStyle?: TextStyle;
    /** @internal Presence distinguishes imported mark layers from fresh marks. */
    sourceParagraphMarkFormatting?: ParagraphMarkFormattingSource;
    /** @internal Original paragraph mark properties retained during regeneration. */
    sourceParagraphMarkPropertiesXml?: string;
    paragraphMarkDeleted?: boolean;
    sourceXml?: string;
    /** @internal Validated provenance for a surgical plain-text sourceXml patch. */
    sourceTextPatch?: {
        runs: Array<{
            style?: TextStyle;
            link?: string;
            noteReference?: TextRunNode["noteReference"];
        }>;
    };
    /** @internal Immutable import-time run metadata used by surgical editors. */
    sourceRunProvenance?: {
        runs: Array<{
            style?: TextStyle;
            link?: string;
            noteReference?: TextRunNode["noteReference"];
        }>;
    };
}
interface TableBoxSpacing {
    topTwips?: number;
    rightTwips?: number;
    bottomTwips?: number;
    leftTwips?: number;
}
interface TableBorderStyle {
    type: string;
    color?: string;
    sizeEighthPt?: number;
}
interface TableBorderSet {
    top?: TableBorderStyle;
    right?: TableBorderStyle;
    bottom?: TableBorderStyle;
    left?: TableBorderStyle;
    insideH?: TableBorderStyle;
    insideV?: TableBorderStyle;
    tl2br?: TableBorderStyle;
    tr2bl?: TableBorderStyle;
}
/** dxa values are twips; pct values are ordinary percentages. */
type TablePreferredWidth = {
    type: "dxa";
    value: number;
} | {
    type: "pct";
    value: number;
} | {
    type: "auto";
} | {
    type: "nil";
};
/** @internal Imported baseline for edits through either width API. */
interface TableWidthSource {
    preferredWidth?: TablePreferredWidth;
    widthTwips?: number;
    inheritedPreferredWidth?: TablePreferredWidth;
}
type TableAlignment = "left" | "center" | "right";
interface TableCellStyle {
    backgroundColor?: string;
    gridSpan?: number;
    rowSpan?: number;
    vMergeContinuation?: boolean;
    widthTwips?: number;
    preferredWidth?: TablePreferredWidth;
    /** @internal */
    sourceWidth?: TableWidthSource;
    marginTwips?: TableBoxSpacing;
    verticalAlign?: "top" | "center" | "bottom";
    textDirection?: "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV" | "tbLrV";
    /** @internal */
    sourceTextDirection?: TableCellStyle["textDirection"];
    /** Protects the preferred cell width during automatic table layout. */
    noWrap?: boolean;
    /** @internal */
    sourceNoWrap?: boolean;
    borders?: TableBorderSet;
}
type TableCellContentNode = ParagraphNode | TableNode;
interface TableCellNode {
    type: "table-cell";
    nodes: TableCellContentNode[];
    style?: TableCellStyle;
}
/** @internal Imported baseline for row geometry edits and clears. */
interface TableRowGeometrySource {
    gridBefore?: number;
    gridAfter?: number;
    widthBefore?: TablePreferredWidth;
    widthAfter?: TablePreferredWidth;
}
interface TableRowStyle {
    gridBefore?: number;
    gridAfter?: number;
    widthBefore?: TablePreferredWidth;
    widthAfter?: TablePreferredWidth;
    /** @internal */
    sourceRowGeometry?: TableRowGeometrySource;
    backgroundColor?: string;
    heightTwips?: number;
    heightRule?: "auto" | "atLeast" | "exact";
    cantSplit?: boolean;
    isHeader?: boolean;
}
interface TableRowNode {
    type: "table-row";
    cells: TableCellNode[];
    style?: TableRowStyle;
}
interface TableStyle {
    styleId?: string;
    /** @internal Imported table-style identity used to detect edits and clears. */
    sourceStyleId?: string;
    widthTwips?: number;
    preferredWidth?: TablePreferredWidth;
    /** @internal */
    sourceWidth?: TableWidthSource;
    alignment?: TableAlignment;
    /** @internal Imported effective alignment before editing. */
    sourceAlignment?: TableAlignment;
    /** @internal Inherited alignment exposed when a direct override is cleared. */
    sourceInheritedAlignment?: TableAlignment;
    bidiVisual?: boolean;
    /** @internal Imported effective table direction before editing. */
    sourceBidiVisual?: boolean;
    /** @internal Inherited direction exposed when a direct override is cleared. */
    sourceInheritedBidiVisual?: boolean;
    indentTwips?: number;
    layout?: "fixed" | "autofit";
    cellSpacingTwips?: number;
    cellMarginTwips?: TableBoxSpacing;
    columnWidthsTwips?: number[];
    /** @internal Validated zero-width logical columns excluded from physical layout. */
    collapsedGridColumns?: number[];
    borders?: TableBorderSet;
    floating?: {
        xTwips?: number;
        yTwips?: number;
        leftFromTextTwips?: number;
        rightFromTextTwips?: number;
        topFromTextTwips?: number;
        bottomFromTextTwips?: number;
        horizontalAnchor?: string;
        verticalAnchor?: string;
        horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
        verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
    };
}
interface TableNode {
    reviewRows?: {
        sourceIndex: number;
        sourceXml: string;
        row: TableRowNode;
    }[];
    type: "table";
    /** Stable identity for measurement caches; survives edits via cloneTable. */
    blockId?: string;
    rows: TableRowNode[];
    style?: TableStyle;
    sourceXml?: string;
    /** @internal Validated text-only descendant replacements for sourceXml. */
    sourceTextPatches?: Array<{
        sourceParagraphXml: string;
        paragraph: ParagraphNode;
    }>;
}
type DocNode = ParagraphNode | TableNode;
interface HeaderSection {
    partName: string;
    referenceType?: string;
    nodes: DocNode[];
}
interface FooterSection {
    partName: string;
    referenceType?: string;
    nodes: DocNode[];
}
interface DocumentSection {
    startNodeIndex: number;
    sectionPropertiesXml?: string;
    headerSections: HeaderSection[];
    footerSections: FooterSection[];
}
interface ParagraphStyleDefinition {
    id: string;
    name: string;
    basedOnId?: string;
    nextStyleId?: string;
    align?: ParagraphAlignment;
    snapToGrid?: boolean;
    /** @internal Whether the style chain owns grid snapping before document defaults. */
    sourceHasSnapToGrid?: boolean;
    textAlignment?: ParagraphTextAlignment;
    /** @internal Whether the style chain owns alignment before document defaults. */
    sourceHasTextAlignment?: boolean;
    headingLevel?: HeadingLevel;
    numbering?: ParagraphNumbering;
    spacing?: ParagraphSpacing;
    indent?: ParagraphIndent;
    backgroundColor?: string;
    borders?: ParagraphBorderSet;
    tabStops?: ParagraphTabStop[];
    contextualSpacing?: boolean;
    keepNext?: boolean;
    keepLines?: boolean;
    widowControl?: boolean;
    pageBreakBefore?: boolean;
    runStyle?: TextStyle;
    uiPriority?: number;
    isDefault?: boolean;
    isPrimary?: boolean;
}
interface NumberingLevelDefinition {
    ilvl: number;
    start?: number;
    format?: string;
    text?: string;
    suffix?: "tab" | "space" | "nothing";
    /**
     * `w:lvlRestart`: one-based level whose use restarts this level; 0 means the
     * level never restarts. Absent means the OOXML default (restart when any
     * shallower level is used).
     */
    lvlRestart?: number;
    /**
     * `w:isLgl`: legal numbering — every placeholder in this level's text
     * renders as decimal regardless of the referenced levels' formats.
     */
    isLgl?: boolean;
    indent?: ParagraphIndent;
    runStyle?: TextStyle;
    bulletFontFamily?: string;
    bulletColor?: string;
    pictureBulletId?: number;
    pictureBullet?: NumberingPictureBulletDefinition;
}
interface NumberingAbstractDefinition {
    abstractNumId: number;
    levels: NumberingLevelDefinition[];
}
interface NumberingInstanceDefinition {
    numId: number;
    abstractNumId: number;
    levelStartOverrides?: Record<string, number>;
    levelOverrides?: NumberingLevelDefinition[];
}
interface NumberingDefinitionSet {
    abstracts: NumberingAbstractDefinition[];
    instances: NumberingInstanceDefinition[];
}
interface NumberingPictureBulletDefinition {
    numPicBulletId: number;
    src?: string;
    widthPx?: number;
    heightPx?: number;
    partName?: string;
    contentType?: string;
}
interface DocumentNoteDefinition {
    id: number;
    text: string;
    nodes?: DocNode[];
}
interface DocumentCommentDefinition {
    nodes?: DocNode[];
    id: number;
    author?: string;
    initials?: string;
    date?: string;
    text: string;
    /** Comment id this comment replies to (from commentsExtended threading). */
    parentId?: number;
    /** True when the comment thread is marked done in commentsExtended. */
    resolved?: boolean;
    /** @internal Original `w:comment` block retained for lossless serialization. */
    sourceXml?: string;
    /** @internal Final comment paragraph id used by `commentsExtended.xml`. */
    extendedParagraphId?: string;
    /** @internal Resolution state read from `commentsExtended.xml`. */
    sourceResolved?: boolean;
    /** @internal True when `resolved` differs from the imported state. */
    resolutionDirty?: boolean;
    /** @internal True for a comment created in the editor rather than imported. */
    isNew?: boolean;
}
interface DocumentCompatibilitySettings {
    adjustLineHeightInTable?: boolean | null;
    noLeading?: boolean;
    suppressTopSpacingWP?: boolean;
    truncateFontHeightsLikeWP6?: boolean;
    noExtraLineSpacing?: boolean;
    spaceForUL?: boolean;
    suppressBottomSpacing?: boolean;
    compatibilityMode?: number;
    enableOpenTypeFeatures?: boolean;
    suppressSpacingBeforeAfterPageBreak?: boolean;
    usePrinterMetrics?: boolean;
    useFixedHtmlParagraphSpacing?: boolean;
    doNotBreakWrappedTables?: boolean;
    doNotBreakConstrainedForcedTable?: boolean;
    evenAndOddHeaders?: boolean;
}
interface DocModel {
    nodes: DocNode[];
    metadata: {
        documentId?: string;
        sourcePartNames?: Record<string, string>;
        sourceParts: number;
        warnings: string[];
        documentPageCount?: number;
        documentOpenTag?: string;
        documentBackgroundColor?: string;
        sectionPropertiesXml?: string;
        sections?: DocumentSection[];
        headerSections: HeaderSection[];
        footerSections: FooterSection[];
        paragraphStyles: ParagraphStyleDefinition[];
        defaultParagraphStyleId?: string;
        defaultTabStopTwips?: number;
        /** @internal Imported settings are authoritative when their values are cleared. */
        documentSettingsImported?: boolean;
        /** @internal Import consumed the line-spacing compatibility fields. */
        lineSpacingCompatibilityImported?: boolean;
        numberingDefinitions?: NumberingDefinitionSet;
        compatibility?: DocumentCompatibilitySettings;
        footnotes?: DocumentNoteDefinition[];
        endnotes?: DocumentNoteDefinition[];
        comments?: DocumentCommentDefinition[];
    };
}

interface ImagePaintRect {
    xPx: number;
    yPx: number;
    widthPx: number;
    heightPx: number;
}
declare function resolveImagePaintBounds(image: Pick<ImageRunNode, "paintBounds" | "widthPx" | "heightPx">, widthPx?: number | undefined, heightPx?: number | undefined): ImagePaintRect | undefined;

declare const DOCUMENT_BLOCK_TYPES: readonly ["text", "heading", "section_heading", "figure", "table", "table_head", "table_cell", "key_value", "page_number", "barcode", "formula", "header", "footer", "watermark", "legend"];
type DocumentBlockType = (typeof DOCUMENT_BLOCK_TYPES)[number];
interface DocxSourceRef {
    documentId: string;
    nodeId: string;
    partName: string;
    story: string;
    start?: number;
    end?: number;
    offsetEncoding: "utf-16";
}
interface ExtractDocumentBlocksOptions {
    includeWords?: boolean;
    include3dModels?: boolean;
}
interface DocumentWord {
    object: "word";
    id: string;
    blockId: string;
    content: string;
    sourceRefs: DocxSourceRef[];
}
interface DocumentBlock {
    object: "block";
    id: string;
    type: DocumentBlockType;
    content: string;
    details: Record<string, unknown>;
    sourceRefs: DocxSourceRef[];
    parentBlockId?: string;
    children: DocumentBlock[];
    words?: DocumentWord[];
}
interface StructuredDocument {
    schemaVersion: 1;
    parserVersion: string;
    documentId: string;
    blocks: DocumentBlock[];
    warnings: string[];
    geometryStatus: "requires_layout";
}
declare function extractDocumentBlocks(model: DocModel, options?: ExtractDocumentBlocksOptions): Promise<StructuredDocument>;
declare function flattenDocumentBlocks(blocks: readonly DocumentBlock[]): DocumentBlock[];
interface SourceElement {
    id: string;
    kind: "document" | "element" | "text" | "comment" | "processing_instruction";
    namespace?: string;
    name?: string;
    text?: string;
    byteStart: number;
    byteEnd: number;
    attributes: Array<{
        namespace: string | null;
        name: string;
        value: string;
    }>;
    children: SourceElement[];
}
interface DocumentSourceGraph {
    schemaVersion: 1;
    documentId: string;
    parts: Array<{
        name: string;
        byteLength: number;
        root: SourceElement | null;
        error: string | null;
    }>;
    assets: Array<{
        name: string;
        byteLength: number;
        root: null;
        error: null;
    }>;
    warnings: string[];
    preservedFeatures: Array<{
        partName: string;
        feature: string;
        occurrences: number;
        support: "preserved_only" | "approximate_rendering";
    }>;
}
declare function inspectDocumentSource(bytes: ArrayBuffer | Uint8Array): Promise<DocumentSourceGraph>;

interface WordProperty {
    name: string;
    attributes?: Record<string, string>;
    children?: WordProperty[];
}
interface DocumentRevision {
    id: string;
    sourceId?: string;
    kind: "insertion" | "deletion" | "move_from" | "move_to" | "formatting" | "cell_merge";
    scope: "text" | "paragraph_mark" | "paragraph" | "section" | "table" | "table_grid" | "table_row" | "table_cell";
    author?: string;
    date?: string;
    moveName?: string;
    text?: string;
    properties?: {
        before: WordProperty[];
        after: WordProperty[];
    };
    sourceRefs: DocxSourceRef[];
    content?: DocNode[];
    childStart?: number;
    childEnd?: number;
    sourceIndex?: number;
}
interface CommentTarget extends DocxSourceRef {
    childStart: number;
    childEnd: number;
}
interface DocumentComment {
    id: number;
    author?: string | null;
    initials?: string | null;
    date?: string | null;
    text: string;
    parentId?: number | null;
    resolved?: boolean | null;
    blockIds: string[];
    targets: CommentTarget[];
    reference?: DocxSourceRef | null;
}
interface DocumentReview {
    schemaVersion: 1;
    revisions: DocumentRevision[];
    comments: DocumentComment[];
}
declare function extractDocumentReview(model: DocModel): Promise<DocumentReview>;
declare function extractDocumentReviewSync(model: DocModel): DocumentReview;

/** The parent grid bounds skipped-column counts without changing their source values. */
declare function tableGridColumnBound(table: Pick<TableNode, "style" | "rows">): number;
/** Ignore skipped-grid counts that exceed the parent grid. */
declare function resolveTableGridSkipCount(count: number | undefined, gridColumnCount: number): number;
declare function tableRowGridSkipCount(table: Pick<TableNode, "style" | "rows">, row: TableRowNode, edge: "before" | "after"): number;

interface GridCell {
    style?: {
        gridSpan?: number;
    };
}
interface GridRow {
    cells: GridCell[];
    style?: {
        gridBefore?: number;
        gridAfter?: number;
    };
}
interface GridTable {
    rows: GridRow[];
    sourceXml?: string;
    sourceTextPatches?: unknown;
    style?: {
        layout?: string;
        columnWidthsTwips?: number[];
        collapsedGridColumns?: number[];
        cellSpacingTwips?: number;
    };
}
interface GridProjectionOptions {
    /** Effective border spacing after the consumer has resolved inherited styles. */
    resolvedBorderSpacingPx?: number;
}
interface LogicalBoundaryRange {
    start: number;
    end: number;
}
type GridProjectionResult<T extends GridTable> = {
    kind: 'identity';
    sourceTable: T;
    table: T;
}
/** A marked grid rejected here must not be passed to a physical layout consumer. */
 | {
    kind: 'rejected';
    sourceTable: T;
    reason: string;
} | {
    kind: 'projected';
    sourceTable: T;
    table: T;
    logicalToPhysicalBoundary: number[];
    physicalToLogicalBoundaryRanges: LogicalBoundaryRange[];
    physicalToLogicalColumn: number[];
    collapsedGridColumns: number[];
};
/** Produces a physical view while retaining the original logical model and source provenance. */
declare function projectCollapsedTableGrid<T extends GridTable>(sourceTable: T, options?: GridProjectionOptions): GridProjectionResult<T>;
/** Invalidates only the projected table's source provenance at a committed structural edit. */
declare function materializeCollapsedTableGrid<T extends GridTable>(sourceTable: T, options?: GridProjectionOptions): GridProjectionResult<T>;

/** Returns a fresh block id, unique for the lifetime of this module instance. */
declare function allocateBlockId(): string;
/**
 * Assigns a fresh block id to every paragraph/table node (body, headers,
 * footers, footnotes, endnotes, nested table content) that does not already
 * have one. Mutates the model in place and returns it.
 */
declare function ensureDocModelBlockIds(model: DocModel): DocModel;
/**
 * Returns block ids that appear on more than one node within the same
 * container (body, a header/footer part, a note). Nodes without ids are
 * ignored, and the same id may legitimately recur across containers (section
 * views reference the same header/footer parts as the flat metadata lists).
 */
declare function collectDuplicateDocModelBlockIds(model: DocModel): string[];

declare function cloneParagraphMarkFormatting(formatting?: ParagraphMarkFormattingSource): ParagraphMarkFormattingSource | undefined;
declare function cloneParagraphNode(paragraph: ParagraphNode): ParagraphNode;
declare function cloneTableNode(table: TableNode): TableNode;
declare function cloneFormFieldSourcePreview(preview?: FormFieldSourcePreview): FormFieldSourcePreview | undefined;
declare function cloneDocModel(model: DocModel): DocModel;

declare function resolveFormFieldSourcePreview(field: FormFieldRunNode): ImageRunNode | undefined;
declare function invalidateFormFieldSource(field: FormFieldRunNode): FormFieldRunNode;

/**
 * Dev-only commit guard: freezes a committed model so in-place mutation of
 * structure shared with history snapshots throws instead of corrupting them.
 * Already-frozen subtrees are skipped, so once a model is frozen, freezing a
 * structurally shared successor only pays for the changed nodes.
 */
declare function deepFreezeDocModel(model: DocModel): DocModel;

declare function normalizeDocModel(model: DocModel): DocModel;

/**
 * True for a model read from OpenDocument text (`.odt`, `.ott`, `.odm`,
 * `.otm`, `.oth` or flat `.fodt`).
 *
 * Such a model has no Word package or Word source markup behind it, so it can
 * be displayed but not serialized to DOCX. The check reads the source part the
 * parser recorded, not a file name.
 */
declare function isOpenDocumentModel(model: Pick<DocModel, "metadata">): boolean;
/**
 * True for a block id the parser gave to OpenDocument content. Such ids name
 * an OpenDocument part after the document hash: `docx:<sha256>/content.xml/…`
 * for the body, notes and comments, `docx:<sha256>/styles.xml/…` for headers
 * and footers. Lets code that sees one block, not the model, tell that its
 * formatting came resolved from the file.
 */
declare function isOpenDocumentBlockId(blockId: string | undefined): boolean;

declare function buildDocModel(pkg: OoxmlPackage, options?: DocxParseOptions): Promise<DocModel>;
declare function buildDocModelFromBytes(bytes: ArrayBuffer | Uint8Array, options?: DocxParseOptions): Promise<{
    package: OoxmlPackage;
    model: DocModel;
    timings: {
        parseMs: number;
        buildModelMs: number;
        totalMs: number;
    };
}>;

interface LineElementMetrics {
    ascent: number;
    descent: number;
    suppressibleBottomSpace?: number;
}
interface LineVerticalMetrics {
    ascent: number;
    descent: number;
    height: number;
    suppressibleBottomSpace: number;
}
/** Signed 32-bit multiply/divide, including saturation and half-away rounding. */
declare function multiplyDivideRound(a: number, b: number, denominator: number): number;
/** All measurements and limits must use the same unit. */
declare function aggregateLineMetrics(elements: readonly LineElementMetrics[]): LineVerticalMetrics;
declare function checkLineVerticalFit(metrics: LineVerticalMetrics, available: number, secondaryLimit: number): {
    fits: boolean;
    appliedSuppression: number;
    rejectedSuppression: number;
};
declare function constrainParagraphLineSplit(linesRemaining: number, fittingLines: number, minimumBefore: number, minimumAfter: number): number;

/** All amounts must use the same layout unit. */
declare function collapseParagraphMargins(amounts: readonly number[]): number;
/** The preceding amount has already been included in the flow position. */
declare function paragraphMarginContribution(previousAmount: number, nextAmount: number): number;

type TableRowHeightRule = "auto" | "atLeast" | "exact";
interface TableRowHeightConstraint {
    /** Intrinsic outer height already includes cell padding. */
    intrinsicHeight: number;
    declaredHeight?: number;
    heightRule?: TableRowHeightRule;
    maximumBottomPadding?: number;
}
interface ResolvedTableRowHeight {
    height: number;
    exact: boolean;
    contentHeightLimit?: number;
}
declare function normalizeTableRowHeightTwips(height?: number): number | undefined;
/** Measurements must use the same unit; border geometry is resolved separately. */
declare function resolveTableRowHeight(constraint: TableRowHeightConstraint): ResolvedTableRowHeight;
declare function resolveTableCellContentClipHeight(rowHeightLimit: number, cellTopPadding?: number, cellBottomPadding?: number, outerRowHeight?: number): number;

type TableWidthDefinition = Pick<TableStyle, "preferredWidth" | "widthTwips" | "sourceWidth">;
declare function resolveEffectiveTablePreferredWidth(style: TableWidthDefinition | undefined): TablePreferredWidth | undefined;
/** The reference is page text width for tables and resolved table width for cells. */
declare function resolveTablePreferredWidthTwips(style: TableWidthDefinition | undefined, percentageReferenceWidthTwips?: number): number | undefined;
/** Convert table source widths to CSS pixels without a device-pixel rounding step. */
declare function tableWidthTwipsToPixels(widthTwips?: number): number | undefined;
/** Reduce fixed-layout grid widths proportionally when they exceed the table width. */
declare function reduceFixedTableColumnWidths(columnWidths: readonly number[], tableWidth: number): number[];
interface FixedTableCellWidth {
    gridSpan?: number;
    preferredWidth?: number;
}
interface FixedTableRowWidths {
    gridBefore?: number;
    gridAfter?: number;
    beforeWidth?: number;
    afterWidth?: number;
    cells: readonly FixedTableCellWidth[];
}
/** Resolve ordered fixed-layout grid requests in a single measurement unit. */
declare function resolveFixedTableGridWidths(params: {
    initialGridWidths?: readonly number[];
    rows: readonly FixedTableRowWidths[];
    preferredTableWidth?: number;
}): number[];
declare function resolveEffectiveTableAlignment(style: Pick<TableStyle, "alignment" | "sourceAlignment" | "sourceInheritedAlignment"> | undefined): TableAlignment | undefined;
declare function resolveEffectiveTableBidiVisual(style: Pick<TableStyle, "bidiVisual" | "sourceBidiVisual" | "sourceInheritedBidiVisual"> | undefined): boolean | undefined;
declare function resolveTableLeadingIndent(alignment: TableAlignment | undefined, indent?: number): number;
declare function resolveTableAlignmentOffset(params: {
    alignment?: TableAlignment;
    bidiVisual?: boolean;
    availableWidth: number;
    tableWidth: number;
    indent?: number;
}): number;

interface LayoutOptions {
    pageWidth?: number;
    pageHeight?: number;
    margin?: number;
    minLineHeight?: number;
    paragraphSpacing?: number;
    tableCellPadding?: number;
}
interface LayoutFormField {
    checked: boolean;
    label: string;
    fallbackText: string;
    childIndex: number;
    style?: TextStyle;
    disabled?: boolean;
}
interface LayoutTextRun {
    kind: "text";
    id: string;
    text: string;
    style?: TextStyle;
    link?: string;
    formField?: LayoutFormField;
}
interface LayoutImageRun {
    chartData?: ImageRunNode["chartData"];
    kind: "image";
    id: string;
    src?: string;
    alt?: string;
    widthPx?: number;
    heightPx?: number;
    contentType?: string;
    data?: Uint8Array;
    floating?: boolean;
    floatingPosition?: ImageRunNode["floating"];
    crop?: ImageRunNode["crop"];
    paintBounds?: ImageRunNode["paintBounds"];
    cssFilter?: string;
    cssOpacity?: number;
    formField?: LayoutFormField;
}
type LayoutRun = LayoutTextRun | LayoutImageRun;
interface LayoutParagraphBlock {
    kind: "paragraph";
    id: string;
    runs: LayoutRun[];
    align: ParagraphAlignment;
    headingLevel?: HeadingLevel;
    x: number;
    y: number;
    width: number;
    height: number;
    formFieldParagraphIndex?: number;
}
interface LayoutTableCell {
    id: string;
    colSpan?: number;
    backgroundColor?: string;
    paragraphs: LayoutParagraphBlock[];
}
interface LayoutTableRow {
    id: string;
    backgroundColor?: string;
    cells: LayoutTableCell[];
}
interface LayoutTableBlock {
    kind: "table";
    id: string;
    x: number;
    y: number;
    width: number;
    height: number;
    rows: LayoutTableRow[];
}
type LayoutBlock = LayoutParagraphBlock | LayoutTableBlock;
interface LayoutPage {
    number: number;
    blocks: LayoutBlock[];
}
declare function layoutDocument(model: DocModel, options?: LayoutOptions): LayoutPage[];

declare function setWasmSource(source: WasmSource): void;
declare function initWasm(source?: WasmSource): ReturnType<typeof initWasm$1>;

interface DocxImportResult {
    package: OoxmlPackage;
    model: DocModel;
    source: "worker" | "main-thread";
    timings?: DocxImportWorkerTimings;
}
interface DocxImportOptions extends DocxParseOptions {
    signal?: AbortSignal;
    transferBuffer?: boolean;
    useWorker?: boolean | "required";
}
interface DocxImportWorkerTimings {
    totalMs: number;
    parseMs: number;
    buildModelMs: number;
}

type DocxViewerSource = ArrayBuffer | ArrayBufferView | Blob;
interface ParseDocxForViewerOptions extends DocxImportOptions {
    loadEmbeddedFonts?: boolean | "defer";
    fileName?: string;
}
interface ParsedDocxPerformanceTimings {
    fileReadMs: number;
    parseMs: number;
    modelConstructionMs: number;
    fontLoadingMs: number;
    totalMs: number;
}
interface ParsedDocxDocument {
    readonly model: DocModel;
    readonly package: OoxmlPackage;
    readonly source: DocxImportResult["source"];
    readonly fileName?: string;
    readonly timings: ParsedDocxPerformanceTimings;
    readonly embeddedFontsLoaded: boolean;
    loadEmbeddedFonts(): Promise<void>;
    dispose(): void;
}
declare function parseDocxForViewer(source: DocxViewerSource, options?: ParseDocxForViewerOptions): Promise<ParsedDocxDocument>;

type ViewerZoomMode = "automatic" | "fit-page" | "fit-width";
type ViewerZoomLevel = number | ViewerZoomMode;
type ViewerZoomState = {
    level: ViewerZoomLevel;
    resolvedZoom: number;
};

type TableRow = TableNode["rows"][number];
interface TableCellPhysicalGridRange {
    cellIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
}
type TableGridColumnInsertionRowPlan = {
    rowIndex: number;
} & ({
    kind: "unchanged";
} | {
    kind: "expand-before";
    gridBefore: number;
} | {
    kind: "expand-after";
    gridAfter: number;
} | {
    kind: "extend-cell";
    cellIndex: number;
    gridSpan: number;
} | {
    kind: "insert-cell";
    cellIndex: number;
    referenceCellIndex?: number;
});
type TableGridColumnDeletionRowPlan = {
    rowIndex: number;
} & ({
    kind: "unchanged";
} | {
    kind: "reduce-before";
    gridBefore: number;
} | {
    kind: "reduce-after";
    gridAfter: number;
} | {
    kind: "shrink-cell";
    cellIndex: number;
    gridSpan: number;
} | {
    kind: "remove-cell";
    cellIndex: number;
    removesOnlyCell: boolean;
});
interface TableGridColumnInsertionPlan {
    columnIndex: number;
    columnCount: number;
    rows: TableGridColumnInsertionRowPlan[];
}
interface TableGridColumnDeletionPlan {
    columnIndex: number;
    columnCount: number;
    deletesLastPhysicalColumn: boolean;
    deletesAllActualCells: boolean;
    rows: TableGridColumnDeletionRowPlan[];
}
declare function tableCellPhysicalGridRange(row: TableRow, cellIndex: number, gridColumnBound: number): TableCellPhysicalGridRange | undefined;
declare function tablePhysicalGridColumnCount(table: TableNode): number;
declare function planTableGridColumnInsertion(table: TableNode, columnIndex: number): TableGridColumnInsertionPlan | undefined;
declare function planTableGridColumnDeletion(table: TableNode, columnIndex: number): TableGridColumnDeletionPlan | undefined;
declare function tableCellsIntersectingGridRange(table: TableNode, startRowIndex: number, endRowIndex: number, startColumnIndex: number, endColumnIndex: number): Array<{
    rowIndex: number;
    cellIndex: number;
}>;

type AnnotationMutationFailureReason = "stale" | "unsupported" | "unsafe-xml";
type AnnotationMutationResult = {
    ok: true;
    model: DocModel;
} | {
    ok: false;
    reason: AnnotationMutationFailureReason;
};
type CommentCreationResult = {
    ok: true;
    model: DocModel;
    commentId: number;
} | {
    ok: false;
    reason: AnnotationMutationFailureReason;
};
type ParagraphRevisionKind = "insertion" | "deletion" | "move-from" | "move-to" | "format-change" | "paragraph-format-change";
interface ParagraphRevisionTarget {
    /** Top-level body node containing the revision. Table locations are unsupported. */
    nodeIndex: number;
    /** OOXML `w:id` from the `w:ins` or `w:del` wrapper. */
    revisionId: string | number;
    kind: ParagraphRevisionKind;
}
interface CreateParagraphCommentInput {
    /** Top-level body paragraph index. */
    nodeIndex: number;
    /** UTF-16 offsets into the paragraph's plain text, matching DOM selection offsets. */
    startOffset: number;
    endOffset: number;
    text: string;
    author?: string;
    initials?: string;
    date?: string;
}
declare function acceptParagraphRevision(model: DocModel, target: ParagraphRevisionTarget): AnnotationMutationResult;
declare function rejectParagraphRevision(model: DocModel, target: ParagraphRevisionTarget): AnnotationMutationResult;
declare function setCommentResolved(model: DocModel, commentId: number, resolved: boolean): AnnotationMutationResult;
/**
 * Creates a comment over a plain-text range in one top-level paragraph. The
 * operation splits only the boundary runs and preserves their original rPr.
 */
declare function createParagraphComment(model: DocModel, input: CreateParagraphCommentInput): CommentCreationResult;

interface InsertParagraphOptions {
    paragraphStyle?: ParagraphStyle;
    runStyle?: TextStyle;
}
interface UpdateTextOptions {
    insertedStyle?: TextStyle;
}
declare function splitParagraphChildrenAtTextOffsets(paragraph: ParagraphNode, text: string, startOffset: number, endOffset: number, options?: {
    beforeInsertedStyle?: TextStyle;
    afterInsertedStyle?: TextStyle;
}): {
    beforeChildren: ParagraphNode["children"];
    afterChildren: ParagraphNode["children"];
};
declare function insertParagraph(model: DocModel, text: string, index?: number, options?: InsertParagraphOptions): DocModel;
declare function removeParagraph(model: DocModel, index: number): DocModel;
declare function duplicateParagraph(model: DocModel, index: number): DocModel;
declare function updateParagraphText(model: DocModel, index: number, text: string, options?: UpdateTextOptions): DocModel;
declare function updateTableCellText(model: DocModel, tableIndex: number, rowIndex: number, cellIndex: number, text: string, options?: UpdateTextOptions): DocModel;
declare function updateTableCellParagraphText(model: DocModel, tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string, options?: UpdateTextOptions): DocModel;
declare function updateTableCellParagraphTextRecursive(model: DocModel, tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string, options?: UpdateTextOptions): DocModel;
declare function replaceText(model: DocModel, searchValue: string | RegExp, replacement: string): DocModel;
declare function setParagraphHeading(model: DocModel, nodeIndex: number, headingLevel?: HeadingLevel): DocModel;
declare function setParagraphAlignment(model: DocModel, nodeIndex: number, align?: ParagraphAlignment): DocModel;
declare function applyRunStyle(model: DocModel, nodeIndex: number, runIndex: number, style: Partial<TextStyle>): DocModel;
declare function toggleRunStyleFlag(model: DocModel, nodeIndex: number, runIndex: number, key: "bold" | "italic" | "underline" | "strike"): DocModel;
declare function setRunHighlight(model: DocModel, nodeIndex: number, runIndex: number, highlight?: string): DocModel;
declare function setRunColor(model: DocModel, nodeIndex: number, runIndex: number, color?: string): DocModel;
declare function copyParagraphs(model: DocModel, startIndex: number, endIndex?: number): ParagraphNode[];
declare function pasteParagraphs(model: DocModel, index: number, paragraphs: ParagraphNode[]): DocModel;
declare function serializeParagraphsForClipboard(paragraphs: ParagraphNode[]): string;
declare function parseParagraphsFromClipboard(input: string): ParagraphNode[] | undefined;

interface DocumentLayoutMetrics {
    pageWidthPx: number;
    pageHeightPx: number;
    marginsPx: {
        top: number;
        right: number;
        bottom: number;
        left: number;
    };
    headerDistancePx: number;
    footerDistancePx: number;
    docGridLinePitchPx?: number;
}
declare function parseSectionLayout(sectionPropertiesXml?: string, missingSectionLayout?: DocumentLayoutMetrics): DocumentLayoutMetrics;
declare function resolveDocumentLayout(model: DocModel): DocumentLayoutMetrics;

type DocxDocumentTheme = "light" | "dark";
type DocxHeadingStyleMap = Partial<Record<HeadingLevel, Properties<string | number>>>;
type LetterheadFloatSide = "left" | "right";
declare function paragraphLetterheadFloatSideAtNodeIndex(nodes: DocModel["nodes"], nodeIndex: number): LetterheadFloatSide | undefined;
type DocxImageWrapMode = "inline" | "square" | "tight" | "through" | "topAndBottom" | "behindText" | "inFrontOfText";
interface DocxImageWrapState {
    mode: DocxImageWrapMode;
    moveWithText: boolean;
    fixedPositionOnPage: boolean;
}
interface DocxImageWrapMenuOption {
    actionId: DocxContextMenuActionId | (string & {});
    label: string;
    checked: boolean;
    disabled?: boolean;
    onSelect: () => void;
}
interface UseDocxImageWrapMenuResult {
    state: DocxImageWrapState;
    wrapOptions: DocxImageWrapMenuOption[];
    positioningOptions: DocxImageWrapMenuOption[];
    editWrapBoundaryOption: DocxImageWrapMenuOption;
    moreLayoutOptionsOption: DocxImageWrapMenuOption;
    setMode: (mode: DocxImageWrapMode) => void;
    setMoveWithText: (moveWithText: boolean) => void;
}
type DocxListType = "unordered" | "ordered";
type DocxEditorSelection = {
    kind: "paragraph";
    nodeIndex: number;
} | {
    kind: "table-cell";
    tableIndex: number;
    rowIndex: number;
    cellIndex: number;
};
interface DocxTableCellSelectionRange {
    tableIndex: number;
    anchorRowIndex: number;
    anchorCellIndex: number;
    focusRowIndex: number;
    focusCellIndex: number;
    gridRange?: {
        startColumnIndex: number;
        endColumnIndex: number;
    };
}
interface ParagraphLocationInBody {
    kind: "paragraph";
    nodeIndex: number;
}
interface ParagraphLocationInCell {
    kind: "table-cell";
    tableIndex: number;
    rowIndex: number;
    cellIndex: number;
    paragraphIndex: number;
}
type DocxTextRangeLocation = ParagraphLocationInBody | ParagraphLocationInCell;
interface DocxTextRangeBoundary {
    location: DocxTextRangeLocation;
    offset: number;
}
interface DocxTextRange {
    start: DocxTextRangeBoundary;
    end: DocxTextRangeBoundary;
}
interface DocxHistoryRestoreRequest {
    nonce: number;
    selection: DocxEditorSelection;
    activeTextRange?: DocxTextRange;
    textInput?: boolean;
}
type ParagraphLocation = DocxTextRangeLocation;
type DocxImageLocation = ParagraphLocation & {
    childIndex: number;
};
type DocxSectionRegion = "header" | "footer";
interface DocxSectionParagraphLocation {
    region: DocxSectionRegion;
    partName: string;
    nodeIndex: number;
    rowIndex?: number;
    cellIndex?: number;
    paragraphIndex?: number;
}
type DocxSectionImageLocation = DocxSectionParagraphLocation & {
    childIndex: number;
};
type DocxFormFieldLocation = ParagraphLocation & {
    childIndex: number;
};
interface DocxSelectedFormField {
    location: DocxFormFieldLocation;
    field: FormFieldRunNode;
}
type DocxImageDropTarget = ParagraphLocation & {
    childIndex: number;
    /** UTF-16 offset inside the target text run. */
    textOffset?: number;
};
type DocxTrackedChangeKind = "insertion" | "deletion" | "move-from" | "move-to" | "format-change" | "paragraph-format-change";
interface DocxTrackedChange {
    revision?: DocumentRevision;
    id: string;
    /** Document generation that produced this handle. */
    documentLoadNonce: number;
    /** Stable block identity used to reject shifted/stale handles. */
    blockId?: string;
    /** Opaque source-state signature used to fail closed after mutation. */
    sourceFingerprint: string;
    inlineAnchorId?: string;
    /** OOXML `w:id` used by safe accept/reject operations. */
    revisionId?: string;
    kind: DocxTrackedChangeKind;
    author?: string;
    date?: string;
    text?: string;
    nodeIndex: number;
    location: DocxTextRangeLocation;
}
/**
 * A document comment anchored to a text range.
 *
 * Comment definitions come from `word/comments.xml` (threading and resolution
 * state from `word/commentsExtended.xml`); the anchor location is resolved
 * from the paragraph that carries the `commentReference` run.
 */
interface DocxComment {
    /** Stable identifier (unique per rendered anchor). */
    id: string;
    /** Document generation that produced this handle. */
    documentLoadNonce: number;
    /** Stable block identity used to reject shifted/stale handles. */
    blockId?: string;
    /** Opaque source/comment signature used to fail closed after mutation. */
    sourceFingerprint: string;
    /** Comment id from `word/comments.xml` (`w:id`). */
    commentId: number;
    author?: string;
    initials?: string;
    date?: string;
    /** Plain-text comment body. */
    text: string;
    /** Comment id this comment replies to, when part of a thread. */
    parentId?: number;
    /** True when the comment thread is marked done. */
    resolved?: boolean;
    /** Plain-text excerpt of the commented document range. */
    anchorText?: string;
    nodeIndex: number;
    location: DocxTextRangeLocation;
}
type DocxAnnotationCommandFailureReason = AnnotationMutationFailureReason;
type DocxAnnotationCommandResult = {
    ok: true;
} | {
    ok: false;
    reason: DocxAnnotationCommandFailureReason;
};
type DocxCommentCreationCommandResult = {
    ok: true;
    commentId: number;
} | {
    ok: false;
    reason: DocxAnnotationCommandFailureReason;
};
interface DocxCreateCommentOptions {
    author?: string;
    initials?: string;
    date?: string;
}
type DocxLineSpacingRule = "auto" | "exact" | "atLeast";
type DocxSelectionSessionKind = "idle" | "pointer" | "keyboard" | "composition" | "command" | "history-restore";
interface DocxLineSpacingInfo {
    lineRule: DocxLineSpacingRule;
    lineTwips?: number;
    multiple: number;
}
type DocxBorderContext = "paragraph" | "table";
type DocxBorderPreset = "bottom" | "top" | "left" | "right" | "none" | "all" | "outside" | "inside" | "inside-horizontal" | "inside-vertical" | "diagonal-down" | "diagonal-up" | "horizontal-line";
type DocxBorderPresetState = Record<DocxBorderPreset, boolean>;
/**
 * Initial state for `useDocxEditor`.
 *
 * These values are read when the editor controller is created. Later document
 * imports update the controller through `editor.importDocxFile`.
 */
interface UseDocxEditorOptions {
    /**
     * Parsed document to use as the initial editor state without reading or
     * parsing the source again.
     */
    document?: ParsedDocxDocument;
    /**
     * Loads embedded fonts when an initial parsed document deferred them.
     *
     * @defaultValue `true`
     */
    loadEmbeddedFonts?: boolean;
    /**
     * Worker mode for `editor.importDocxFile`, with the same values as
     * `parseDocxForViewer`: `"required"` reports an import error when the import
     * worker cannot start, `true` parses on the main thread when the worker
     * cannot start, and `false` always parses on the main thread.
     *
     * Browsers refuse to run a worker script from another origin, so pass `false`
     * when the package's assets are served from a different origin than the page.
     * Unlike the other options, this one is read on every import.
     *
     * @defaultValue `"required"`, or `false` when the configured WASM source
     * cannot be forwarded to a worker
     */
    useWorker?: boolean | "required";
    /**
     * THUNDER: rewrites each imported document's model before the editor holds
     * it — part of the load, not an edit, so it leaves the history empty. Read on
     * every import, like `useWorker`.
     */
    transformImportedModel?: (model: DocModel) => DocModel;
    /**
     * Model used for `editor.newDocument()` and for the initial empty editor.
     *
     * @defaultValue `defaultStarterModel`
     */
    starterModel?: DocModel;
    /**
     * Initial display name before a user imports a file.
     *
     * @defaultValue `"(new document)"`
     */
    initialFileName?: string;
    /**
     * Initial status text exposed on `editor.status`.
     *
     * @defaultValue `"Ready"`
     */
    initialStatus?: string;
    /**
     * Initial document surface theme.
     *
     * @defaultValue `"light"`
     */
    initialDocumentTheme?: DocxDocumentTheme;
    /**
     * Whether tracked changes are visible when the editor first mounts.
     *
     * @defaultValue `false`
     */
    initialShowTrackedChanges?: boolean;
    /**
     * Whether document comments are visible when the editor first mounts.
     *
     * @defaultValue `false`
     */
    initialShowComments?: boolean;
}
/**
 * Controller returned by `useDocxEditor`.
 *
 * Pass this object to `DocxEditorViewer`, then wire toolbar buttons and custom
 * controls to the methods exposed here.
 *
 * @example
 * ```tsx
 * const editor = useDocxEditor();
 *
 * return (
 *   <>
 *     <button onClick={() => editor.toggleBold()}>Bold</button>
 *     <DocxEditorViewer editor={editor} />
 *   </>
 * );
 * ```
 */
interface DocxEditorController {
    model: DocModel;
    /**
     * THUNDER: runs an edit of the whole model as one undoable step, the way the
     * editor's own commands do. Return `current` unchanged to make no step.
     */
    applyModelChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void;
    /**
     * THUNDER: the block of table cells selected in the viewer — by dragging
     * across cells or selecting the whole table — or undefined when none is.
     * Indices are the table's node index and each row's cell index.
     */
    tableCellSelection?: DocxTableCellSelectionRange;
    /** THUNDER: where `DocxEditorViewer` reports its cell selection. */
    reportTableCellSelection: (range: DocxTableCellSelectionRange | undefined) => void;
    documentLoadNonce: number;
    historyRevision: number;
    fileName: string;
    status: string;
    importError?: Error;
    isImporting: boolean;
    /**
     * Set when the loaded document can only be viewed, with the reason to show
     * the user. OpenDocument text (`.odt` and its variants) has no Word package
     * to save into, so `DocxEditorViewer` renders it read-only whatever its
     * `mode`, edits leave the model unchanged, `exportDocx()` reports this
     * reason in `status` and `getDocxBytes()` rejects with it.
     */
    viewOnlyReason?: string;
    documentTheme: DocxDocumentTheme;
    selection: DocxEditorSelection;
    activeTextRange?: DocxTextRange;
    historyRestoreRequest?: DocxHistoryRestoreRequest;
    selectedFormField?: DocxSelectedFormField;
    selectedParagraph?: ParagraphNode;
    selectedRunStyle?: TextRunNode["style"];
    selectedLink?: string;
    pendingRunStyle?: TextRunNode["style"];
    selectionSessionKind: DocxSelectionSessionKind;
    suppressNextDomSelectionRestore: () => void;
    beginSelectionSession: (kind: Exclude<DocxSelectionSessionKind, "idle">, options?: {
        settleAfterMs?: number;
    }) => void;
    clearSelectionSession: (expectedKind?: DocxSelectionSessionKind) => void;
    selectedParagraphStyleId?: string;
    selectedLineSpacing: DocxLineSpacingInfo;
    selectedBorderContext: DocxBorderContext;
    activeBorderPresets: DocxBorderPresetState;
    availableParagraphStyles: ParagraphStyleDefinition[];
    trackedChanges: DocxTrackedChange[];
    showTrackedChanges: boolean;
    comments: DocxComment[];
    showComments: boolean;
    currentPage: number;
    totalPages: number;
    hasUnorderedList: boolean;
    hasOrderedList: boolean;
    canUndo: boolean;
    canRedo: boolean;
    registerPendingExportModelTransformer: (transformer?: (model: DocModel, documentLoadNonce: number) => DocModel) => void;
    setStatus: (value: string | ((previous: string) => string)) => void;
    setDocumentTheme: (theme: DocxDocumentTheme) => void;
    setShowTrackedChanges: (showTrackedChanges: boolean) => void;
    setShowComments: (showComments: boolean) => void;
    syncPaginationInfo: (pagination: DocxPaginationInfo) => void;
    /** Selects a numeric or responsive viewer zoom level. */
    setZoom: (level: ViewerZoomLevel) => void;
    /** Returns the selected zoom level, including an active responsive mode. */
    getZoom: () => ViewerZoomLevel;
    /** Returns the currently resolved zoom percentage. */
    getResolvedZoom: () => number;
    toggleShowTrackedChanges: () => void;
    toggleShowComments: () => void;
    acceptTrackedChange: (change: DocxTrackedChange) => DocxAnnotationCommandResult;
    rejectTrackedChange: (change: DocxTrackedChange) => DocxAnnotationCommandResult;
    createComment: (text: string, options?: DocxCreateCommentOptions) => DocxCommentCreationCommandResult;
    setCommentResolved: (comment: DocxComment, resolved: boolean) => DocxAnnotationCommandResult;
    importDocxFile: (file: File) => Promise<void>;
    newDocument: () => void;
    exportDocx: () => void;
    /** The current document as DOCX bytes, the same bytes `exportDocx()` downloads. */
    getDocxBytes: () => Promise<ArrayBuffer>;
    undo: () => void;
    redo: () => void;
    setHeading: (heading?: HeadingLevel) => void;
    setParagraphStyle: (styleId?: string) => void;
    setLineSpacing: (lineMultiple: number) => void;
    setFontFamily: (fontFamily: string) => void;
    setFontSize: (fontSizePt: number) => void;
    toggleBold: () => void;
    toggleItalic: () => void;
    toggleUnderline: () => void;
    toggleStrike: () => void;
    toggleSuperscript: () => void;
    toggleSubscript: () => void;
    setHighlight: (highlight?: string) => void;
    setTextColor: (color?: string) => void;
    setLink: (link?: string) => void;
    selectFormField: (location?: DocxFormFieldLocation) => void;
    toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
    setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
    updateFormFieldWidget: (location: DocxFormFieldLocation, patch: Partial<NonNullable<FormFieldRunNode["widget"]>>) => void;
    applyBorderPreset: (preset: DocxBorderPreset) => void;
    setAlignment: (align?: ParagraphAlignment) => void;
    toggleList: (listType: DocxListType) => void;
    adjustSelectedListDepth: (levelDelta: number, draftText?: string) => boolean;
    insertListItemAfterSelection: (draftText: string, startOffset: number, endOffset?: number, targetLocation?: ParagraphLocation) => {
        paragraphIndex: number;
        caretOffset: number;
    } | undefined;
    splitParagraphAtSelection: (draftText: string, startOffset: number, endOffset?: number, targetLocation?: ParagraphLocation) => {
        paragraphIndex: number;
        caretOffset: number;
    } | undefined;
    insertTable: () => void;
    insertImageFile: (file: File) => Promise<void>;
    appendParagraph: (text?: string) => number;
    resizeImage: (location: DocxImageLocation, widthPx: number, heightPx: number) => void;
    setSyntheticTextBoxText: (location: DocxImageLocation, text: string, storyIndex?: number) => boolean;
    setImageWrapMode: (location: DocxImageLocation, mode: DocxImageWrapMode, seedFloating?: Partial<NonNullable<ImageRunNode["floating"]>>) => void;
    moveFloatingImage: (location: DocxImageLocation, patch: Partial<NonNullable<ImageRunNode["floating"]>>, options?: {
        reparentToParagraphNodeIndex?: number;
        anchorTextOffset?: number;
    }) => void;
    moveSectionFloatingImage: (location: DocxSectionImageLocation, patch: Partial<NonNullable<ImageRunNode["floating"]>>) => void;
    moveParagraphDropCap: (nodeIndex: number, patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>) => void;
    setParagraphDropCapFontSizePt: (nodeIndex: number, fontSizePt: number) => void;
    setParagraphDropCapText: (nodeIndex: number, text: string) => void;
    moveImage: (source: DocxImageLocation, target: DocxImageDropTarget) => void;
    setActiveTextRange: (range?: DocxTextRange) => void;
    selectParagraph: (nodeIndex: number) => void;
    selectTableCell: (tableIndex: number, rowIndex: number, cellIndex: number) => void;
    clearTableCellContents: (tableIndex: number, cells: Array<{
        rowIndex: number;
        cellIndex: number;
    }>) => void;
    insertTableRow: (tableIndex: number, rowIndex: number, direction: "above" | "below") => void;
    insertTableColumn: (tableIndex: number, cellIndex: number, direction: "left" | "right", rowIndex?: number) => void;
    deleteTableRow: (tableIndex: number, rowIndex: number) => void;
    deleteTableColumn: (tableIndex: number, cellIndex: number, rowIndex?: number) => void;
    deleteTable: (tableIndex: number) => void;
    moveTable: (tableIndex: number, targetNodeIndex: number) => void;
    setTableFloating: (tableIndex: number, patch: Partial<NonNullable<NonNullable<TableNode["style"]>["floating"]>>, options?: {
        moveToNodeIndex?: number;
        anchorTextOffset?: number;
    }) => void;
    moveEmbeddedTableToBody: (tableRuntimeKey: string, targetNodeIndex: number) => void;
    replaceExpandedSelection: (text: string, range?: DocxTextRange) => DocxTextRange | undefined;
    deleteExpandedSelection: (range?: DocxTextRange) => DocxTextRange | undefined;
    commitParagraphText: (nodeIndex: number, text: string) => void;
    commitTableCellText: (tableIndex: number, rowIndex: number, cellIndex: number, text: string) => void;
    commitTableCellParagraphTextRecursive: (tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string) => void;
    commitParagraphTextAtRange: (location: ParagraphLocation, text: string, caret: {
        start: number;
        end: number;
    }, options?: {
        insertionOffset?: number;
        coalesceKey?: string;
    }) => boolean;
    commitTableCellTextAtRange: (tableIndex: number, rowIndex: number, cellIndex: number, text: string, caret: {
        paragraphIndex: number;
        start: number;
        end: number;
    }) => boolean;
    /**
     * THUNDER: replaces a paragraph's runs as a text edit — one undo step, or
     * part of the last one when `coalesceKey` continues a burst of typing.
     */
    commitParagraphChildrenAtRange: (location: ParagraphLocation, children: ParagraphNode["children"], caret: number, options?: {
        insertionOffset?: number;
        coalesceKey?: string;
    }) => boolean;
    commitSectionParagraphText: (location: DocxSectionParagraphLocation, text: string) => void;
}
type DocxTableContextMenuActionId = "insert-row-above" | "insert-row-below" | "insert-column-left" | "insert-column-right" | "delete-row" | "delete-column" | "delete-table";
type DocxContextMenuActionId = DocxTableContextMenuActionId | "cut" | "copy" | "paste" | "image-wrap-inline" | "image-wrap-square" | "image-wrap-tight" | "image-wrap-through" | "image-wrap-top-and-bottom" | "image-wrap-behind-text" | "image-wrap-in-front-of-text" | "image-edit-wrap-boundary" | "image-move-with-text" | "image-fix-position-on-page" | "image-more-layout-options" | "image-bring-to-front" | "image-bring-forward" | "image-in-front-of-text" | "image-send-to-back" | "image-send-backward" | "image-behind-text";
interface DocxTableContextMenuContext {
    /** Zero-based table index in the document body. */
    tableIndex: number;
    /** Zero-based row index under the pointer. */
    rowIndex: number;
    /** Zero-based cell index under the pointer. */
    cellIndex: number;
}
interface DocxTableContextMenuAction {
    /** Built-in table command id. */
    id: DocxTableContextMenuActionId;
    /** Human-readable menu label. */
    label: string;
    /** True when the action removes content or structure. */
    destructive?: boolean;
}
interface DocxContextMenuAction {
    /** Built-in command id, or a custom id from your menu renderer. */
    id: DocxContextMenuActionId | (string & {});
    /** Human-readable menu label. */
    label: string;
    /** Optional keyboard shortcut text to display next to the label. */
    shortcut?: string;
    /** True when the action removes content or structure. */
    destructive?: boolean;
    /** Prevents the action from being selected. */
    disabled?: boolean;
    /** Marks toggle-style menu items as active. */
    checked?: boolean;
    /** Adds a visual separator before this item in default-style menus. */
    separatorBefore?: boolean;
    /** Nested submenu actions. */
    children?: DocxContextMenuAction[];
}
/**
 * Arguments passed to `renderTableContextMenu`.
 */
interface DocxTableContextMenuRenderProps {
    /** Table cell that opened the menu. */
    context: DocxTableContextMenuContext;
    /** Built-in table actions available for this context. */
    actions: DocxTableContextMenuAction[];
    /** Runs a built-in table action. */
    runAction: (actionId: DocxTableContextMenuActionId) => void;
    /** Closes the menu without running an action. */
    closeMenu: () => void;
    /** Viewport coordinates where the menu should be placed. */
    position: {
        x: number;
        y: number;
    };
    /** Current document theme, useful for custom menu styling. */
    documentTheme: DocxDocumentTheme;
}
interface DocxContextMenuContext {
    /** Type of target that opened the menu. */
    kind: "text" | "table" | "image";
    /** Active text range when the menu was opened from text. */
    activeTextRange?: DocxTextRange;
    /** Text location under the pointer when available. */
    location?: DocxTextRangeLocation;
    /** Table context when `kind` is `"table"`. */
    tableContext?: DocxTableContextMenuContext;
    /** Image context when `kind` is `"image"`. */
    image?: {
        location: DocxImageLocation;
        floating?: NonNullable<ImageRunNode["floating"]>;
        wrap?: DocxImageWrapState;
    } | undefined;
}
/**
 * Arguments passed to `renderContextMenu`.
 */
interface DocxContextMenuRenderProps {
    /** Target context that opened the menu. */
    context: DocxContextMenuContext;
    /** Built-in actions available for this context. */
    actions: DocxContextMenuAction[];
    /** Runs a built-in or custom action id. */
    runAction: (actionId: DocxContextMenuActionId | (string & {})) => void;
    /** Closes the menu without running an action. */
    closeMenu: () => void;
    /** Viewport coordinates where the menu should be placed. */
    position: {
        x: number;
        y: number;
    };
    /** Current document theme, useful for custom menu styling. */
    documentTheme: DocxDocumentTheme;
}
/**
 * Controls how `DocxEditorViewer` windows page DOM nodes.
 *
 * Virtualization is enabled by default for multi-page documents. It keeps
 * offscreen pages as lightweight placeholders, which is important for large
 * DOCX files with many tables or complex paragraphs.
 */
interface DocxPageVirtualizationOptions {
    /**
     * Enables or disables internal page virtualization.
     *
     * Disable only when you need every page mounted at once, for example when
     * generating thumbnails for all pages from the live DOM.
     *
     * @defaultValue `true`
     */
    enabled?: boolean;
    /**
     * Number of pages to keep mounted before and after the visible viewport.
     *
     * Higher values reduce mount/unmount churn while scrolling but increase
     * initial DOM work.
     *
     * @defaultValue `2`
     */
    overscan?: number;
    /**
     * Compatibility option from earlier releases.
     *
     * Page virtualization now starts immediately, so this option is accepted but
     * does not delay the initial page window.
     *
     * @defaultValue `0`
     */
    settleDelayMs?: number;
    /**
     * Explicit scroll container for internal page virtualization.
     *
     * Omit this to let the viewer discover the nearest scrollable ancestor.
     * Provide it when the host app owns scrolling, such as when the viewer is
     * mounted inside a custom scroll-area viewport.
     */
    scrollElement?: HTMLElement | null;
    /**
     * Effective visual scale applied around the viewer.
     *
     * Omit this to let the viewer infer CSS `zoom` from its ancestor chain.
     * Provide it when toolbar zoom is controlled outside the viewer so virtual
     * page offsets update synchronously with the selected zoom.
     */
    zoomScale?: number;
}
/**
 * Externally controlled visible page window.
 *
 * Provide this when your app owns scrolling or virtualization and wants
 * `DocxEditorViewer` to mount only a specific page range.
 */
interface DocxVisiblePageRange {
    /**
     * Zero-based first page index to render.
     *
     * @example
     * ```tsx
     * <DocxEditorViewer
     *   editor={editor}
     *   visiblePageRange={{ startPageIndex: 0, endPageIndex: 2 }}
     * />
     * ```
     */
    startPageIndex: number;
    /** Zero-based last page index to render. */
    endPageIndex: number;
}
/**
 * Props for the full DOCX editor/viewer surface.
 *
 * Use this component with a controller from `useDocxEditor`.
 *
 * @example
 * ```tsx
 * const editor = useDocxEditor();
 *
 * return (
 *   <DocxEditorViewer
 *     editor={editor}
 *     mode="read-only"
 *     pageVirtualization={{ overscan: 3 }}
 *   />
 * );
 * ```
 */
interface DocxEditorViewerProps {
    /**
     * Editor controller returned by `useDocxEditor`.
     */
    editor: DocxEditorController;
    /**
     * CSS class applied to the outer viewer root.
     */
    className?: string;
    /**
     * Inline styles applied to the outer viewer root.
     */
    style?: React.CSSProperties;
    /** Controlled numeric percentage or responsive zoom mode. */
    zoom?: ViewerZoomLevel;
    /** Initial zoom level when `zoom` is uncontrolled. */
    defaultZoom?: ViewerZoomLevel;
    /** Called when the selected level or its resolved percentage changes. */
    onZoomChange?: (state: ViewerZoomState) => void;
    /**
     * Background color of each rendered page surface.
     *
     * @defaultValue Viewer theme page color.
     */
    pageBackgroundColor?: string;
    /**
     * Background color shown between pages.
     *
     * @defaultValue `"transparent"`
     */
    pageGapBackgroundColor?: string;
    /**
     * Hides the document behind `loadingState` until initial pagination settles.
     *
     * Disable to show approximate layout before pagination settles.
     *
     * @defaultValue `true`
     */
    deferInitialPaginationPaint?: boolean;
    /**
     * Custom content shown while initial pagination is settling.
     *
     * Used only when `deferInitialPaginationPaint` is true.
     *
     * @defaultValue A compact `"Loading..."` pill.
     */
    loadingState?: React.ReactNode;
    /**
     * Configures internal page virtualization.
     *
     * @defaultValue `{ enabled: true, overscan: 2 }`
     */
    pageVirtualization?: DocxPageVirtualizationOptions;
    /**
     * Controlled page range to mount.
     *
     * When provided, internal page virtualization is bypassed and only this
     * range is rendered.
     */
    visiblePageRange?: DocxVisiblePageRange;
    /**
     * Hard set of zero-based pages to mount. Unlike `visiblePageRange`, omitted
     * pages remain metadata-only even when they fall between requested pages.
     */
    pageIndexes?: readonly number[];
    /**
     * Called whenever the viewer's resolved page count changes.
     */
    onPageCountChange?: (pageCount: number) => void;
    /**
     * Called when the viewer needs a page to become visible, such as after a
     * bookmark or cross-reference navigation.
     *
     * If omitted, the internal virtualizer scrolls the page into view.
     */
    onRequestPageReveal?: (pageIndex: number) => void;
    /**
     * Overrides visual styles for heading levels.
     */
    headingStyles?: DocxHeadingStyleMap;
    /**
     * Interaction mode for the viewer.
     *
     * @defaultValue `"edit"`
     */
    mode?: DocxEditorViewerMode;
    /**
     * Overrides whether tracked changes are shown.
     *
     * If omitted, the value from `useDocxTrackChanges(editor)` or
     * `editor.showTrackedChanges` is used.
     */
    showTrackedChanges?: boolean;
    /** Places removed content in margin balloons by default. */
    revisionDisplay?: "balloons" | "inline";
    /**
     * Custom renderer for tracked-change cards in the page gutter.
     */
    renderTrackedChangeCard?: (props: DocxTrackedChangeCardRenderProps) => React.ReactNode;
    /**
     * Overrides whether document comments are shown.
     *
     * If omitted, the value from `useDocxComments(editor)` or
     * `editor.showComments` is used.
     */
    showComments?: boolean;
    /**
     * Custom renderer for comment cards in the page gutter.
     */
    renderCommentCard?: (props: DocxCommentCardRenderProps) => React.ReactNode;
    /**
     * Custom renderer for table context menus.
     *
     * Call `props.runAction(action.id)` to execute built-in actions.
     */
    renderTableContextMenu?: (props: DocxTableContextMenuRenderProps) => React.ReactNode;
    /**
     * THUNDER: turns typed or pasted text into inline nodes — the template
     * editor's `{{…}}` into locked fields. Return undefined to insert the text
     * as it is.
     */
    inlineNodesFromText?: (text: string) => ParagraphNode["children"] | undefined;
    /**
     * Custom renderer for text, image, and table context menus.
     *
     * Call `props.closeMenu()` after handling a custom action.
     */
    renderContextMenu?: (props: DocxContextMenuRenderProps) => React.ReactNode;
    /**
     * Called when a form field is double-clicked.
     *
     * Use this to open a custom field settings panel.
     */
    onFormFieldDoubleClick?: (location: DocxFormFieldLocation) => void;
}
interface DocxTrackedChangeCardRenderProps {
    /** Tracked-change data represented by the card. */
    change: DocxTrackedChange;
    /** Short display label for the change kind. */
    kindLabel: string;
    /** Plain-text excerpt for the changed content. */
    snippet: string;
    /** Rendered removed content, preserving inline formatting and images. */
    content?: React.ReactNode;
    /** Formatted change date, if the source document provided one. */
    formattedDate?: string;
    /** Accent color chosen for this change kind. */
    accentColor: string;
    /** Current document theme. */
    documentTheme: DocxDocumentTheme;
    /** Zero-based page index that owns the card. */
    pageIndex: number;
    /** Positioning style computed by the viewer. Apply this to the card root. */
    style: Properties<string | number>;
    /** Accept this change. Unsupported or stale markup fails closed. */
    accept: () => DocxAnnotationCommandResult;
    /** Reject this change. Unsupported or stale markup fails closed. */
    reject: () => DocxAnnotationCommandResult;
}
interface DocxCommentCardRenderProps {
    /** Comment data represented by the card. */
    comment: DocxComment;
    /** Plain-text comment body (already normalized for display). */
    snippet: string;
    /** Formatted comment date, if the source document provided one. */
    formattedDate?: string;
    /** Accent color chosen for comments. */
    accentColor: string;
    /** Current document theme. */
    documentTheme: DocxDocumentTheme;
    /** Zero-based page index that owns the card. */
    pageIndex: number;
    /** Positioning style computed by the viewer. Apply this to the card root. */
    style: Properties<string | number>;
    /** Mark the comment resolved or reopen it. */
    setResolved: (resolved: boolean) => DocxAnnotationCommandResult;
}
/**
 * Viewer interaction mode.
 *
 * - `"edit"` allows editing when the document state permits it.
 * - `"read-only"` disables editing affordances while preserving navigation.
 */
type DocxEditorViewerMode = "edit" | "read-only";
interface UseDocxDocumentThemeResult {
    documentTheme: DocxDocumentTheme;
    isDarkDocument: boolean;
    setDocumentTheme: (theme: DocxDocumentTheme) => void;
    toggleDocumentTheme: () => void;
}
interface UseDocxParagraphStylesResult {
    paragraphStyles: ParagraphStyleDefinition[];
    selectedParagraphStyleId?: string;
    setParagraphStyle: (styleId?: string) => void;
}
interface UseDocxLineSpacingResult {
    lineSpacing: DocxLineSpacingInfo;
    setLineSpacing: (lineMultiple: number) => void;
}
interface UseDocxBordersResult {
    borderContext: DocxBorderContext;
    activeBorderPresets: DocxBorderPresetState;
    applyBorderPreset: (preset: DocxBorderPreset) => void;
}
interface UseDocxFormFieldsResult {
    formFields: DocxSelectedFormField[];
    selectedFormField?: DocxSelectedFormField;
    selectFormField: (location?: DocxFormFieldLocation) => void;
    setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
    toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
    updateFormFieldWidget: (location: DocxFormFieldLocation, patch: Partial<NonNullable<FormFieldRunNode["widget"]>>) => void;
    updateSelectedFormFieldWidget: (patch: Partial<NonNullable<FormFieldRunNode["widget"]>>) => void;
}
interface UseDocxTrackChangesResult {
    trackedChanges: DocxTrackedChange[];
    showTrackedChanges: boolean;
    setShowTrackedChanges: (showTrackedChanges: boolean) => void;
    toggleShowTrackedChanges: () => void;
    changesByLocation: Map<string, DocxTrackedChange[]>;
    getChangesForLocation: (location: DocxTextRangeLocation) => DocxTrackedChange[];
    acceptChange: (change: DocxTrackedChange) => DocxAnnotationCommandResult;
    rejectChange: (change: DocxTrackedChange) => DocxAnnotationCommandResult;
}
interface UseDocxCommentsResult {
    comments: DocxComment[];
    showComments: boolean;
    setShowComments: (showComments: boolean) => void;
    toggleShowComments: () => void;
    commentsByLocation: Map<string, DocxComment[]>;
    getCommentsForLocation: (location: DocxTextRangeLocation) => DocxComment[];
    createComment: (text: string, options?: DocxCreateCommentOptions) => DocxCommentCreationCommandResult;
    setCommentResolved: (comment: DocxComment, resolved: boolean) => DocxAnnotationCommandResult;
}
interface DocxSectionColumnLayout {
    count: number;
    gapPx: number;
}
interface DocxPageLayoutInfo {
    pageWidthPx: number;
    pageHeightPx: number;
    contentWidthPx: number;
    contentHeightPx: number;
    marginsPx: {
        top: number;
        right: number;
        bottom: number;
        left: number;
    };
    headerDistancePx: number;
    footerDistancePx: number;
    pageNumberStart: number;
    columns?: DocxSectionColumnLayout;
    viewportDefaults: {
        zoomPercent: number;
        pageGapPx: number;
    };
}
interface UseDocxPageLayoutResult {
    layout: DocxPageLayoutInfo;
}
interface DocxPaginationInfo {
    currentPage: number;
    totalPages: number;
}
interface UseDocxPaginationResult {
    pagination: DocxPaginationInfo;
}
interface DocxPageThumbnailResolutionOptions {
    /** Source page width in CSS pixels. */
    sourceWidthPx: number;
    /** Source page height in CSS pixels. */
    sourceHeightPx: number;
    /**
     * Desired thumbnail bounds.
     *
     * A number means "fit within this max width and max height". An object lets
     * you constrain width and height independently.
     *
     * @defaultValue `160`
     */
    resolution?: DocxPageThumbnailBounds;
    /**
     * Maximum CSS pixel width for the thumbnail.
     *
     * Overrides `resolution.maxWidth` when provided.
     */
    maxWidthPx?: number;
    /**
     * Maximum CSS pixel height for the thumbnail.
     *
     * Overrides `resolution.maxHeight` when provided.
     */
    maxHeightPx?: number;
    /**
     * Canvas backing-store pixel ratio.
     *
     * Increase for sharper thumbnails on high-density displays.
     *
     * @defaultValue `window.devicePixelRatio`, capped internally.
     */
    pixelRatio?: number;
}
/**
 * Thumbnail size constraint.
 *
 * @example
 * ```tsx
 * useDocxPageThumbnails(editor, { resolution: 180 });
 * useDocxPageThumbnails(editor, { resolution: { maxWidth: 120 } });
 * ```
 */
type DocxPageThumbnailBounds = number | {
    /** Maximum thumbnail CSS height. */
    maxHeight?: number;
    /** Maximum thumbnail CSS width. */
    maxWidth?: number;
};
interface DocxPageThumbnailResolution {
    /** Thumbnail CSS width. */
    widthPx: number;
    /** Thumbnail CSS height. */
    heightPx: number;
    /** Canvas backing-store width. */
    pixelWidthPx: number;
    /** Canvas backing-store height. */
    pixelHeightPx: number;
    /** Scale from source page pixels to thumbnail CSS pixels. */
    scale: number;
}
interface DocxPageThumbnailRenderWindow {
    /**
     * Page indexes whose attached thumbnail canvases should render first.
     *
     * Use this for the thumbnails currently visible in a virtualized sidebar.
     */
    visiblePageIndexes?: readonly number[];
    /**
     * Page indexes to rasterize into the thumbnail surface cache after visible
     * thumbnails. Prefetched pages paint quickly once their canvases mount.
     */
    prefetchPageIndexes?: readonly number[];
}
/**
 * Options for `useDocxPageThumbnails`.
 */
interface UseDocxPageThumbnailsOptions {
    /**
     * Desired thumbnail bounds.
     *
     * @defaultValue `160`
     */
    resolution?: DocxPageThumbnailBounds;
    /** Maximum thumbnail CSS width. */
    maxWidthPx?: number;
    /** Maximum thumbnail CSS height. */
    maxHeightPx?: number;
    /**
     * Canvas backing-store pixel ratio.
     *
     * @defaultValue `window.devicePixelRatio`, capped internally.
     */
    pixelRatio?: number;
    /**
     * Minimum interval between repeat raster jobs for the same thumbnail canvas.
     *
     * Lower this when the consumer already limits work to a small visible
     * thumbnail window.
     *
     * @defaultValue `200`
     */
    minRasterIntervalMs?: number;
    /**
     * Prioritizes thumbnails for consumer-owned virtualized thumbnail rails.
     */
    renderWindow?: DocxPageThumbnailRenderWindow;
    /**
     * Hard set of zero-based pages exposed and eligible for rasterization.
     * Omit to preserve the existing all-pages behavior.
     */
    pageIndexes?: readonly number[];
    /**
     * Default scheduling policy for thumbnail paint requests.
     *
     * @defaultValue `"idle"`
     */
    scheduling?: DocxThumbnailScheduling;
    /**
     * Prevents thumbnail rendering while keeping stable item metadata.
     *
     * @defaultValue `false`
     */
    disabled?: boolean;
}
type DocxPageThumbnailStatus = "idle" | "rendering" | "ready" | "unavailable" | "error";
type DocxThumbnailScheduling = "idle" | "immediate";
interface DocxRenderToCanvasOptions {
    /** Bypasses idle scheduling for a visible imperative request. */
    scheduling?: DocxThumbnailScheduling;
    /** Ignores unchanged-content and surface-cache fast paths. */
    force?: boolean;
}
interface DocxPageThumbnailItem extends DocxPageThumbnailResolution {
    /** Source page aspect ratio. */
    aspectRatio: number;
    /** Alias for `sourceHeightPx` for compatibility with other preview APIs. */
    contentHeight: number;
    /** Alias for `sourceWidthPx` for compatibility with other preview APIs. */
    contentWidth: number;
    /** Alias for `heightPx`. */
    height: number;
    /** Zero-based page index. */
    pageIndex: number;
    /** One-based page number for display. */
    pageNumber: number;
    /** Source page width in CSS pixels. */
    sourceWidthPx: number;
    /** Source page height in CSS pixels. */
    sourceHeightPx: number;
    /** True when the source page DOM is currently mounted. */
    isMounted: boolean;
    /** Current thumbnail render status. */
    status: DocxPageThumbnailStatus;
    /** Last thumbnail rendering error, if any. */
    error?: Error;
    /** Paints this thumbnail into a canvas. Returns false when unavailable. */
    paint: (canvas: HTMLCanvasElement | null) => boolean;
    /** Ref callback that keeps an attached canvas rendered as the page changes. */
    canvasRef: (canvas: HTMLCanvasElement | null) => void;
    /** Asynchronously renders this thumbnail into a canvas. */
    renderToCanvas: (canvas: HTMLCanvasElement, options?: DocxRenderToCanvasOptions) => Promise<void>;
    /** Alias for `widthPx`. */
    width: number;
}
interface UseDocxPageThumbnailsResult {
    /** Paints the requested page thumbnail into a canvas. */
    paintThumbnail: (pageIndex: number, canvas: HTMLCanvasElement | null) => boolean;
    /** Thumbnail metadata and paint helpers for each known page. */
    thumbnails: DocxPageThumbnailItem[];
    /** Re-renders every thumbnail canvas currently attached through `canvasRef`. */
    rerenderAttachedThumbnails: () => Promise<void>;
}
type UseDocxViewerThumbnailsOptions = UseDocxPageThumbnailsOptions;
type DocxViewerThumbnails = UseDocxPageThumbnailsResult;
declare const defaultStarterModel: DocModel;
interface DocxCommentLink {
    id: number;
    parentId?: number;
    resolved?: boolean;
    blockIds: string[];
    targets: {
        blockId: string;
        start?: number;
        end?: number;
    }[];
}
declare function resolveDocxCommentLinks(model: DocModel): DocxCommentLink[];
declare function useDocxEditor(options?: UseDocxEditorOptions): DocxEditorController;
declare function resolveDocxPageThumbnailResolution(options: DocxPageThumbnailResolutionOptions): DocxPageThumbnailResolution;
declare function useDocxPageThumbnails(editor: DocxEditorController, options?: UseDocxPageThumbnailsOptions): UseDocxPageThumbnailsResult;
declare const useDocxViewerThumbnails: typeof useDocxPageThumbnails;
declare function useDocxDocumentTheme(editor: Pick<DocxEditorController, "documentTheme" | "setDocumentTheme">): UseDocxDocumentThemeResult;
declare function useDocxParagraphStyles(editor: Pick<DocxEditorController, "availableParagraphStyles" | "selectedParagraphStyleId" | "setParagraphStyle">): UseDocxParagraphStylesResult;
declare function useDocxImageWrapMenu(menu: Pick<DocxContextMenuRenderProps, "context" | "runAction">): UseDocxImageWrapMenuResult | undefined;
declare function useDocxLineSpacing(editor: Pick<DocxEditorController, "selectedLineSpacing" | "setLineSpacing">): UseDocxLineSpacingResult;
declare function useDocxBorders(editor: Pick<DocxEditorController, "selectedBorderContext" | "activeBorderPresets" | "applyBorderPreset">): UseDocxBordersResult;
declare function useDocxFormFields(editor: Pick<DocxEditorController, "model" | "selectedFormField" | "selectFormField" | "setFormFieldValue" | "toggleFormCheckbox" | "updateFormFieldWidget">): UseDocxFormFieldsResult;
declare function useDocxTrackChanges(editor: Pick<DocxEditorController, "trackedChanges" | "showTrackedChanges" | "setShowTrackedChanges" | "toggleShowTrackedChanges" | "acceptTrackedChange" | "rejectTrackedChange">): UseDocxTrackChangesResult;
/**
 * Exposes document comments and the show/hide state, mirroring
 * `useDocxTrackChanges`. Pair with `DocxEditorViewer`'s `renderCommentCard`
 * to fully customize how comments render.
 */
declare function useDocxComments(editor: Pick<DocxEditorController, "comments" | "showComments" | "setShowComments" | "toggleShowComments" | "createComment" | "setCommentResolved">): UseDocxCommentsResult;
declare function useDocxPageLayout(editor: Pick<DocxEditorController, "model">): UseDocxPageLayoutResult;
declare function useDocxPagination(editor: Pick<DocxEditorController, "currentPage" | "totalPages">): UseDocxPaginationResult;
declare function DocxEditorViewer({ editor, className, style, zoom, defaultZoom, onZoomChange, pageBackgroundColor, pageGapBackgroundColor, deferInitialPaginationPaint, loadingState, pageVirtualization, visiblePageRange, pageIndexes, onPageCountChange, onRequestPageReveal, headingStyles, showTrackedChanges, revisionDisplay, renderTrackedChangeCard, showComments, renderCommentCard, renderTableContextMenu, inlineNodesFromText, renderContextMenu, onFormFieldDoubleClick, mode, }: DocxEditorViewerProps): React.JSX.Element;

declare function invalidateFontMetrics(): void;

interface BlockBoundingBox {
    left: number;
    top: number;
    right: number;
    bottom: number;
}
interface DocxBlockFragment {
    nodeId: string;
    pageIndex: number;
    pageWidth: number;
    pageHeight: number;
    boundingBox: BlockBoundingBox;
    polygon: Array<{
        x: number;
        y: number;
    }>;
    coordinateSpace: "page-css-px";
}
interface DocxBlockGeometry {
    status: "rendered" | "partially_rendered" | "not_mounted" | "not_visible" | "unresolved";
    scope: "mounted_pages";
    issues: Array<{
        nodeId: string;
        reason: "unsupported_range" | "document_mismatch" | "not_on_page" | "range_text_mismatch";
    }>;
    fragments: DocxBlockFragment[];
    unresolvedSourceIds: string[];
}
declare function resolveDocxSourceGeometry(root: HTMLElement, refs: readonly DocxSourceRef[]): DocxBlockGeometry;
declare function resolveDocxBlockGeometries(root: HTMLElement, blocks: readonly DocumentBlock[]): Map<string, DocxBlockGeometry>;
declare function resolveDocxWordGeometries(root: HTMLElement, words: readonly DocumentWord[]): Map<string, DocxBlockGeometry>;
declare function DocxWordHighlights({ word, words, ...props }: {
    rootRef: React.RefObject<HTMLElement | null>;
    word?: DocumentWord;
    words?: readonly DocumentWord[];
    color?: string;
    onGeometry?: (geometry: DocxBlockGeometry) => void;
}): React.JSX.Element;
declare function revealDocxSource(root: HTMLElement, source: string | DocxSourceRef): void;
interface DocxBlockHighlightsProps {
    rootRef: React.RefObject<HTMLElement | null>;
    block?: DocumentBlock;
    blocks?: readonly DocumentBlock[];
    color?: string;
    colorForBlock?: (block: DocumentBlock) => string;
    onGeometry?: (geometry: DocxBlockGeometry) => void;
}
declare function DocxBlockHighlights({ rootRef, block, blocks, color, colorForBlock, onGeometry, }: DocxBlockHighlightsProps): React.JSX.Element;
declare function normalizedDocxBoundingBox(fragment: DocxBlockFragment): {
    pageIndex: number;
    left: number;
    top: number;
    width: number;
    height: number;
};
interface DocxPageBlock extends Omit<DocumentBlock, "id" | "parentBlockId" | "children"> {
    id: string;
    sourceBlockId: string;
    parentBlockId?: string;
    metadata: {
        page: {
            number: number;
            width: number;
            height: number;
        };
    };
    boundingBox: BlockBoundingBox;
    polygon: Array<{
        x: number;
        y: number;
    }>;
    fragments: DocxBlockFragment[];
}
declare function toDocxPageBlocks(block: DocumentBlock, geometry: DocxBlockGeometry): DocxPageBlock[];
type ExtendParserBlockDetails = {
    type: "text_details";
} | {
    type: "table_details";
    rowCount: number;
    columnCount: number;
} | {
    type: "table_cell_details";
    rowIndex: number;
    columnIndex: number;
} | {
    type: "figure_details";
    imageUrl?: string;
    figureType?: string;
} | {
    type: "barcode_details";
    imageUrl?: string;
    decodedValue?: string;
    format?: string;
} | {
    type: "key_value_details";
    ocrText?: string;
} | {
    type: "formula_details";
    latex: string;
} | Record<string, never>;
interface ExtendParserBlock extends Omit<DocxPageBlock, "details" | "fragments"> {
    details: ExtendParserBlockDetails;
    sourceDetails: DocumentBlock["details"];
    coordinateSpace: "page-points";
    children: ExtendParserBlock[];
}
declare function toExtendParserBlocks(blocks: readonly DocumentBlock[], geometries: ReadonlyMap<string, DocxBlockGeometry>): ExtendParserBlock[];

interface ResolvedModelSection {
    startNodeIndex: number;
    sectionPropertiesXml?: string;
    headerSections: HeaderSection[];
    footerSections: FooterSection[];
}
interface PaginationSectionMetrics {
    startNodeIndex: number;
    pageContentWidthPx: number;
    pageContentHeightPx: number;
    docGridLinePitchPx?: number;
}
interface TableExplicitPageBreakInfo {
    startRowIndexes: number[];
    breakAfterTable: boolean;
}
interface DocumentPageRange {
    startNodeIndex: number;
    endNodeIndex: number;
}
declare function sectionBreakPropertiesStartNewPage(sectionPropertiesXml: string): boolean;
declare function paragraphHasExplicitPageBreak(paragraph: ParagraphNode): boolean;
declare function paragraphHasPageBreakBefore(paragraph: ParagraphNode): boolean;
declare function sectionBreakAfterParagraphStartsNewPage(paragraph: ParagraphNode): boolean;
declare function paragraphHasLastRenderedPageBreak(paragraph: ParagraphNode): boolean;
declare function paragraphStartsWithLastRenderedPageBreak(paragraph: ParagraphNode): boolean;
declare function paragraphBeforeSpacingPx(paragraph: ParagraphNode): number;
declare function paragraphAfterSpacingPx(paragraph: ParagraphNode): number;
declare function resolveParagraphBeforeSpacingPx(model: DocModel, nodeIndex: number, paragraph: ParagraphNode, pageConsumedHeightPx: number, suppressSpacingBeforeAfterPageBreak: boolean): number;
declare function resolveDocumentSectionsFromMetadata(metadata: DocModel["metadata"]): ResolvedModelSection[];
declare function sectionTitlePageEnabled(sectionPropertiesXml?: string): boolean;
declare function selectSectionVariantForPage<T extends HeaderSection | FooterSection>(sections: T[], sectionPropertiesXml: string | undefined, pageIndex: number, options?: {
    evenAndOddHeaders?: boolean;
}): T | undefined;
declare function resolveSectionIndexForNodeIndex(sections: Pick<ResolvedModelSection, "startNodeIndex">[], nodeIndex: number, previousSectionIndex: number): number;
declare function resolveSectionPropertiesXmlForNodeIndex(sections: ResolvedModelSection[], nodeIndex: number, fallbackSectionPropertiesXml?: string): string | undefined;
declare function resolvePaginationSectionMetricsIndexForNodeIndex(metricsBySection: PaginationSectionMetrics[], nodeIndex: number, previousSectionIndex: number): number;
declare function scalePaginationSectionMetricsHeights(metricsBySection: PaginationSectionMetrics[], heightScale: number): PaginationSectionMetrics[];
declare function collectTableExplicitPageBreakInfo(table: TableNode): TableExplicitPageBreakInfo;
declare function collectTopLevelExplicitPageBreakStartNodeIndexes(nodes: DocModel["nodes"]): Set<number>;
declare function collectDocxHardPageBreakStartNodeIndexes(model: DocModel): Set<number>;
declare function collectDocxLastRenderedPageBreakStartNodeIndexes(model: DocModel): number[];
declare function buildDocumentPageRanges(nodeCount: number, pageBreakStartNodeIndexes: Iterable<number>): DocumentPageRange[];

declare const DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX = 2;
declare const DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX = 14;
interface TableRowRange {
    startRowIndex: number;
    endRowIndex: number;
}
interface ParagraphLineRange {
    startLineIndex: number;
    endLineIndex: number;
    totalLineCount: number;
    lineHeightPx: number;
}
interface DocumentPageNodeSegment {
    nodeIndex: number;
    tableRowRange?: TableRowRange;
    paragraphLineRange?: ParagraphLineRange;
    paragraphBeforeSpacingPx?: number;
}
type PaginationPageStartReason = "initial" | "overflow" | "paragraph-break" | "manual-break" | "section";
declare function suppressParagraphBeforeSpacingAtPageStart(reason: PaginationPageStartReason, compatibilityMode?: number): boolean;
interface LetterheadColumnSegmentGroup {
    startOffset: number;
    endOffset: number;
    leftSegments: DocumentPageNodeSegment[];
    rightSegments: DocumentPageNodeSegment[];
}
interface ParagraphSplitControlOptions {
    allowKeepLinesOverflow?: boolean;
    allowKeepNextOverflow?: boolean;
}
interface PageSegmentationCallbacks {
    estimateDocNodeHeightPx: (node: DocModel["nodes"][number], availableWidthPx?: number, numberingDefinitions?: NumberingDefinitionSet, docGridLinePitchPx?: number) => number;
    paragraphHasVisibleText: (paragraph: ParagraphNode) => boolean;
    paragraphIsStructuralSectionBreakSpacer: (paragraph: ParagraphNode) => boolean;
    estimateParagraphHeightPx: (paragraph: ParagraphNode, availableWidthPx?: number, numberingDefinitions?: NumberingDefinitionSet, docGridLinePitchPx?: number) => number;
    estimateParagraphLineHeightPx: (paragraph: ParagraphNode, docGridLinePitchPx?: number) => number;
    paragraphLineCountWithinWidth: (paragraph: ParagraphNode, availableWidthPx?: number, numberingDefinitions?: NumberingDefinitionSet) => number;
    paragraphWidowControlEnabled: (paragraph: ParagraphNode) => boolean;
    paragraphCanSplitAcrossPages: (paragraph: ParagraphNode, lineCount: number, options?: ParagraphSplitControlOptions) => boolean;
    estimateTableRowHeightsPx: (table: TableNode, maxAvailableWidthPx?: number, numberingDefinitions?: NumberingDefinitionSet, docGridLinePitchPx?: number) => number[];
}
interface OverflowBreakCollectionOptions {
    suppressSpacingBeforeAfterPageBreak?: boolean;
    pageOverflowTolerancePx?: number;
}
interface DocumentPageSegmentationOptions extends OverflowBreakCollectionOptions {
    allowParagraphLineSplitting?: boolean;
    measuredTableRowHeightsByNodeIndex?: Record<number, number[]>;
    measuredPageContentHeightsPxByPageIndex?: number[];
    minParagraphLineHeightPx?: number;
    preferLastRenderedParagraphStartBreaks?: boolean;
}
declare function paragraphLetterheadColumnGroupAtSegmentOffset(nodeSegments: DocumentPageNodeSegment[], startOffset: number, resolveFloatSideAtNodeIndex: (nodeIndex: number) => "left" | "right" | undefined): LetterheadColumnSegmentGroup | undefined;
declare function collectDocxEstimatedOverflowBreakStartNodeIndexes(model: DocModel, hardBreakStartNodeIndexes: Set<number>, pageContentHeightPx: number, pageContentWidthPx: number, callbacks: PageSegmentationCallbacks, numberingDefinitions?: NumberingDefinitionSet, paginationMetricsBySection?: PaginationSectionMetrics[], options?: OverflowBreakCollectionOptions): Set<number>;
declare function buildDocumentPageNodeSegments(model: DocModel, pageContentHeightPx: number, pageContentWidthPx: number, callbacks: PageSegmentationCallbacks, numberingDefinitions?: NumberingDefinitionSet, paginationMetricsBySection?: PaginationSectionMetrics[], options?: DocumentPageSegmentationOptions): DocumentPageNodeSegment[][];
declare function resolveDocumentPageSegmentStartNodeIndex(pageSegments: DocumentPageNodeSegment[]): number | undefined;
declare function scorePaginationAgainstStoredPageBreaks(pages: DocumentPageNodeSegment[][], storedBreakStartNodeIndexes: number[]): number;

declare function normalizeSectionPageNumber(pageNumber: number, properties?: string): number;
declare function createSectionParityBlankPage<T>(): T[];
declare function isSectionParityBlankPage(page: readonly unknown[]): boolean;
declare function createSectionPageNumberTracker(metadata: DocModel["metadata"]): {
    advancePage: () => void;
    startSection: (nodeIndex: number, atPageTop: boolean, hasPreviousPages: boolean) => boolean;
};

type LayoutSnapshotOptions = LayoutOptions;
interface LayoutSnapshotRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
interface LayoutSnapshotPageSize {
    width: number;
    height: number;
}
interface LayoutSnapshotMargins {
    top: number;
    right: number;
    bottom: number;
    left: number;
}
interface ResolvedLayoutMetrics {
    pageSizePx: LayoutSnapshotPageSize;
    marginsPx: LayoutSnapshotMargins;
    contentBoxPx: LayoutSnapshotRect;
}
interface ResolvedDocumentSection {
    index: number;
    startNodeIndex: number;
    synthetic: boolean;
    layout: ResolvedLayoutMetrics;
}
interface ResolvedDocument {
    source: "layout-engine-adapter";
    nodeCount: number;
    pageCountHint?: number;
    layout: ResolvedLayoutMetrics;
    sections: ResolvedDocumentSection[];
    metadata: {
        sourceParts: number;
        warningCount: number;
        warnings: string[];
        headerSectionCount: number;
        footerSectionCount: number;
    };
}
interface LayoutFragmentSource {
    kind: "paragraph" | "table" | "table-cell" | "table-cell-paragraph";
    nodeIndex: number;
    rowIndex?: number;
    cellIndex?: number;
    paragraphIndex?: number;
}
interface LayoutSnapshotTextRun {
    kind: "text";
    id: string;
    text: string;
    style?: TextStyle;
    link?: string;
}
interface LayoutSnapshotImageRun {
    kind: "image";
    id: string;
    src?: string;
    alt?: string;
    widthPx?: number;
    heightPx?: number;
    floating?: boolean;
}
type LayoutSnapshotRun = LayoutSnapshotTextRun | LayoutSnapshotImageRun;
interface LayoutSnapshotParagraphFragment {
    kind: "paragraph";
    id: string;
    framePx: LayoutSnapshotRect;
    align: ParagraphAlignment;
    headingLevel?: HeadingLevel;
    runs: LayoutSnapshotRun[];
    source?: LayoutFragmentSource;
}
interface LayoutSnapshotTableCellFragment {
    id: string;
    colSpan?: number;
    backgroundColor?: string;
    paragraphs: LayoutSnapshotParagraphFragment[];
    source?: LayoutFragmentSource;
}
interface LayoutSnapshotTableRowFragment {
    id: string;
    backgroundColor?: string;
    cells: LayoutSnapshotTableCellFragment[];
}
interface LayoutSnapshotTableFragment {
    kind: "table";
    id: string;
    framePx: LayoutSnapshotRect;
    rows: LayoutSnapshotTableRowFragment[];
    source?: LayoutFragmentSource;
}
type LayoutSnapshotBlockFragment = LayoutSnapshotParagraphFragment | LayoutSnapshotTableFragment;
interface LayoutSnapshotPage {
    index: number;
    number: number;
    pageSizePx: LayoutSnapshotPageSize;
    marginsPx: LayoutSnapshotMargins;
    contentBoxPx: LayoutSnapshotRect;
    blocks: LayoutSnapshotBlockFragment[];
}
interface LayoutSnapshot {
    version: 1;
    source: "layout-engine-adapter";
    resolvedDocument: ResolvedDocument;
    pages: LayoutSnapshotPage[];
}
declare function resolveDocumentForLayout(model: DocModel, options?: LayoutSnapshotOptions): ResolvedDocument;
declare function buildLayoutSnapshot(model: DocModel, options?: LayoutSnapshotOptions): LayoutSnapshot;

declare function modelToDocumentXml(model: DocModel, basePackage?: OoxmlPackage): Promise<string>;
declare function serializeDocModel(model: DocModel, basePackage?: OoxmlPackage): Promise<OoxmlPackage>;
declare function serializeDocx(model: DocModel, basePackage?: OoxmlPackage): Promise<ArrayBuffer>;
declare function canPreserveImagePaintSource(image: ImageRunNode, basePackage?: OoxmlPackage, ownerPartName?: string): Promise<boolean>;

/**
 * @deprecated Since v0.10.0, use DocxEditorViewerProps with mode="read-only"
 * and a controller from useDocxEditor.
 */
interface ReactDocxViewerProps {
    /**
     * Raw `.docx` file contents to parse and render.
     *
     * Pass either `file` or `model`. When both are provided, `model` wins and
     * the file buffer is ignored.
     *
     * @example
     * ```tsx
     * const buffer = await file.arrayBuffer();
     * <ReactDocxViewer file={buffer} />
     * ```
     */
    file?: ArrayBuffer;
    /**
     * Prebuilt document model to render without parsing a `.docx` buffer.
     *
     * Useful when you already parsed the document with `parseDocx` and
     * `buildDocModel`, or when you are rendering a model produced by your own
     * pipeline.
     */
    model?: DocModel;
    /** Parsed input returned by `parseDocxForViewer`. */
    document?: ParsedDocxDocument;
    /** Loads deferred embedded fonts when a parsed document is mounted. */
    loadEmbeddedFonts?: boolean;
    /** Hard set of zero-based pages to mount. */
    pageIndexes?: readonly number[];
    /**
     * CSS class applied to the outer viewer container.
     *
     * @example
     * ```tsx
     * <ReactDocxViewer file={buffer} className="docx-preview" />
     * ```
     */
    className?: string;
    /** Controlled numeric percentage or responsive zoom mode. */
    zoom?: ViewerZoomLevel;
    /** Initial zoom level when `zoom` is uncontrolled. */
    defaultZoom?: ViewerZoomLevel;
    /** Called when the selected level or its resolved percentage changes. */
    onZoomChange?: (state: ViewerZoomState) => void;
    /**
     * Layout overrides for the simple read-only renderer.
     *
     * If omitted, page width and height are derived from the document section
     * properties when available.
     *
     * @example
     * ```tsx
     * <ReactDocxViewer
     *   file={buffer}
     *   layoutOptions={{ pageWidth: 816, pageHeight: 1056, margin: 72 }}
     * />
     * ```
     */
    layoutOptions?: LayoutOptions;
    /**
     * Content shown when neither `file` nor `model` is provided.
     *
     * @defaultValue `"No DOCX loaded."`
     */
    emptyState?: React.ReactNode;
}
interface UseDocxModelOptions {
    /**
     * Worker mode for the import, with the same values as `parseDocxForViewer`:
     * `"required"` reports an error when the import worker cannot start, `true`
     * parses on the main thread when the worker cannot start, and `false` always
     * parses on the main thread.
     *
     * Browsers refuse to run a worker script from another origin, so pass `false`
     * when the package's assets are served from a different origin than the page.
     * Changing this value parses the current file again.
     *
     * @defaultValue `"required"`
     */
    useWorker?: boolean | "required";
}
interface UseDocxModelState {
    /** Parsed document model, available once loading succeeds. */
    model?: DocModel;
    /** True while the hook is parsing the current file buffer. */
    isLoading: boolean;
    /** Parse or model-build failure for the current file buffer. */
    error?: Error;
}
declare function useDocxModel(file?: ArrayBuffer, options?: UseDocxModelOptions): UseDocxModelState;
/**
 * @deprecated Since v0.10.0, use DocxEditorViewer with mode="read-only"
 * and a controller from useDocxEditor.
 */
declare function ReactDocxViewer({ file, model, document, loadEmbeddedFonts, pageIndexes, className, zoom, defaultZoom, onZoomChange, layoutOptions, emptyState, }: ReactDocxViewerProps): React.JSX.Element;
type DocxThumbnailOutput = "canvas" | "blob" | "imageBitmap";
interface CreateDocxThumbnailRendererOptions {
    scheduling?: DocxThumbnailScheduling;
    pixelRatio?: number;
    resolution?: DocxPageThumbnailBounds;
    pageIndexes?: readonly number[];
}
interface RenderDocxThumbnailPageOptions {
    scheduling?: DocxThumbnailScheduling;
    pixelRatio?: number;
    resolution?: DocxPageThumbnailBounds;
    maxWidth?: number;
    maxHeight?: number;
    output?: DocxThumbnailOutput;
    canvas?: HTMLCanvasElement;
    mimeType?: string;
    quality?: number;
}
interface DocxThumbnailRenderTimings {
    paginationMs: number;
    pageMountMs: number;
    rasterizationMs: number;
    encodingMs: number;
    totalMs: number;
}
interface DocxThumbnailRenderResult {
    pageIndex: number;
    pageCount: number;
    width: number;
    height: number;
    pixelWidth: number;
    pixelHeight: number;
    output: DocxThumbnailOutput;
    canvas?: HTMLCanvasElement;
    blob?: Blob;
    imageBitmap?: ImageBitmap;
    timings: DocxThumbnailRenderTimings;
}
type DocxThumbnailDocument = ParsedDocxDocument | DocModel;
interface DocxThumbnailRenderer {
    readonly document: DocxThumbnailDocument;
    readonly pageCount: number;
    renderPage(pageIndex: number, options?: RenderDocxThumbnailPageOptions): Promise<DocxThumbnailRenderResult>;
    dispose(): void;
}
declare function createDocxThumbnailRenderer(documentModel: DocxThumbnailDocument, options?: CreateDocxThumbnailRendererOptions): DocxThumbnailRenderer;

export { type AnnotationMutationFailureReason, type AnnotationMutationResult, type BlockBoundingBox, type ChartData, type ChartPoint, type CommentCreationResult, type CommentTarget, type CreateDocxThumbnailRendererOptions, type CreateParagraphCommentInput, DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX, DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX, DOCUMENT_BLOCK_TYPES, type DocModel, type DocNode, type DocumentBlock, type DocumentBlockType, type DocumentComment, type DocumentCommentDefinition, type DocumentCompatibilitySettings, type DocumentNoteDefinition, type DocumentPageNodeSegment, type DocumentPageRange, type DocumentPageSegmentationOptions, type DocumentReview, type DocumentRevision, type DocumentSection, type DocumentSourceGraph, type DocumentWord, type DocxAnnotationCommandFailureReason, type DocxAnnotationCommandResult, type DocxBlockFragment, type DocxBlockGeometry, DocxBlockHighlights, type DocxBlockHighlightsProps, type DocxBorderContext, type DocxBorderPreset, type DocxBorderPresetState, type DocxComment, type DocxCommentCardRenderProps, type DocxCommentCreationCommandResult, type DocxCommentLink, type DocxContextMenuAction, type DocxContextMenuActionId, type DocxContextMenuContext, type DocxContextMenuRenderProps, type DocxCreateCommentOptions, type DocxDocumentTheme, type DocxEditorController, type DocxEditorSelection, DocxEditorViewer, type DocxEditorViewerMode, type DocxEditorViewerProps, type DocxFormFieldLocation, type DocxHeadingStyleMap, type DocxImageDropTarget, type DocxImageLocation, type DocxImageWrapMenuOption, type DocxImageWrapMode, type DocxImageWrapState, type DocxLineSpacingInfo, type DocxListType, type DocxPageBlock, type DocxPageLayoutInfo, type DocxPageThumbnailBounds, type DocxPageThumbnailItem, type DocxPageThumbnailRenderWindow, type DocxPageThumbnailResolution, type DocxPageThumbnailResolutionOptions, type DocxPageThumbnailStatus, type DocxPageVirtualizationOptions, type DocxPaginationInfo, type DocxParseOptions, type DocxRenderToCanvasOptions, type DocxSectionColumnLayout, type DocxSelectedFormField, type DocxSourceRef, type DocxTableCellSelectionRange, type DocxTableContextMenuAction, type DocxTableContextMenuActionId, type DocxTableContextMenuContext, type DocxTableContextMenuRenderProps, type DocxTextRange, type DocxTextRangeLocation, type DocxThumbnailDocument, type DocxThumbnailOutput, type DocxThumbnailRenderResult, type DocxThumbnailRenderTimings, type DocxThumbnailRenderer, type DocxThumbnailScheduling, type DocxTrackedChange, type DocxTrackedChangeCardRenderProps, type DocxTrackedChangeKind, type DocxViewerSource, type DocxViewerThumbnails, type DocxVisiblePageRange, DocxWordHighlights, type ExtendParserBlock, type ExtendParserBlockDetails, type ExtractDocumentBlocksOptions, type FixedTableCellWidth, type FixedTableRowWidths, type FooterSection, type FormFieldCheckboxWidgetSettings, type FormFieldDropdownWidgetSettings, type FormFieldOption, type FormFieldRunNode, type FormFieldSourceKind, type FormFieldSourcePreview, type FormFieldTextWidgetSettings, type FormFieldType, type FormFieldWidgetSettings, type GridProjectionOptions, type GridProjectionResult, type GridTable, type HeaderSection, type HeadingLevel, type ImageAssetReference, type ImageCrop, type ImagePaintBounds, type ImagePaintRect, type ImagePresentation, type ImagePropertyAttribute, type ImagePropertyNode, type ImageRunNode, type ImageTile, type ImageTransform, type InsertParagraphOptions, type LayoutBlock, type LayoutFormField, type LayoutFragmentSource, type LayoutImageRun, type LayoutOptions, type LayoutPage, type LayoutParagraphBlock, type LayoutRun, type LayoutSnapshot, type LayoutSnapshotBlockFragment, type LayoutSnapshotImageRun, type LayoutSnapshotMargins, type LayoutSnapshotOptions, type LayoutSnapshotPage, type LayoutSnapshotPageSize, type LayoutSnapshotParagraphFragment, type LayoutSnapshotRect, type LayoutSnapshotRun, type LayoutSnapshotTableCellFragment, type LayoutSnapshotTableFragment, type LayoutSnapshotTableRowFragment, type LayoutSnapshotTextRun, type LayoutTableBlock, type LayoutTableCell, type LayoutTableRow, type LayoutTextRun, type LegacyWasmOoxmlPackage, type LetterheadColumnSegmentGroup, type LineElementMetrics, type LineVerticalMetrics, type LogicalBoundaryRange, type NumberingAbstractDefinition, type NumberingDefinitionSet, type NumberingInstanceDefinition, type NumberingLevelDefinition, type NumberingPictureBulletDefinition, type OfficeModel3D, type OoxmlPackage, type OoxmlPart, type OverflowBreakCollectionOptions, type PageSegmentationCallbacks, type PaginationPageStartReason, type PaginationSectionMetrics, type ParagraphAlignment, type ParagraphBorderSet, type ParagraphBorderStyle, type ParagraphChildNode, type ParagraphCommentRange, type ParagraphIndent, type ParagraphLineRange, type ParagraphMarkFormattingSource, type ParagraphNode, type ParagraphNumbering, type ParagraphRevisionKind, type ParagraphRevisionTarget, type ParagraphSpacing, type ParagraphSplitControlOptions, type ParagraphStyle, type ParagraphStyleDefinition, type ParagraphTabStop, type ParagraphTextAlignment, type ParseDocxForViewerOptions, type ParsedDocxDocument, type ParsedDocxPerformanceTimings, ReactDocxViewer, type ReactDocxViewerProps, type RenderDocxThumbnailPageOptions, type ResolvedDocument, type ResolvedDocumentSection, type ResolvedLayoutMetrics, type ResolvedModelSection, type ResolvedTableRowHeight, type SourceElement, type StructuredDocument, type TableAlignment, type TableBorderSet, type TableBorderStyle, type TableBoxSpacing, type TableCellContentNode, type TableCellNode, type TableCellPhysicalGridRange, type TableCellStyle, type TableExplicitPageBreakInfo, type TableGridColumnDeletionPlan, type TableGridColumnDeletionRowPlan, type TableGridColumnInsertionPlan, type TableGridColumnInsertionRowPlan, type TableNode, type TablePreferredWidth, type TableRowGeometrySource, type TableRowHeightConstraint, type TableRowHeightRule, type TableRowNode, type TableRowRange, type TableRowStyle, type TableStyle, type TableWidthDefinition, type TableWidthSource, type TextLigatures, type TextOutline, type TextRunBorderStyle, type TextRunNode, type TextShadow, type TextStyle, type UpdateTextOptions, type UseDocxBordersResult, type UseDocxCommentsResult, type UseDocxDocumentThemeResult, type UseDocxEditorOptions, type UseDocxFormFieldsResult, type UseDocxImageWrapMenuResult, type UseDocxLineSpacingResult, type UseDocxModelOptions, type UseDocxModelState, type UseDocxPageLayoutResult, type UseDocxPageThumbnailsOptions, type UseDocxPageThumbnailsResult, type UseDocxPaginationResult, type UseDocxParagraphStylesResult, type UseDocxTrackChangesResult, type UseDocxViewerThumbnailsOptions, type ViewerZoomLevel, type ViewerZoomMode, type ViewerZoomState, type WasmOoxmlPackage, type WasmSource, type WordProperty, acceptParagraphRevision, aggregateLineMetrics, allocateBlockId, applyRunStyle, buildDocModel, buildDocModelFromBytes, buildDocumentPageNodeSegments, buildDocumentPageRanges, buildLayoutSnapshot, canPreserveImagePaintSource, checkLineVerticalFit, cloneDocModel, cloneFormFieldSourcePreview, cloneParagraphMarkFormatting, cloneParagraphNode, cloneTableNode, collapseParagraphMargins, collectDocxEstimatedOverflowBreakStartNodeIndexes, collectDocxHardPageBreakStartNodeIndexes, collectDocxLastRenderedPageBreakStartNodeIndexes, collectDuplicateDocModelBlockIds, collectTableExplicitPageBreakInfo, collectTopLevelExplicitPageBreakStartNodeIndexes, constrainParagraphLineSplit, copyParagraphs, createDocxThumbnailRenderer, createMinimalDocxPackage, createParagraphComment, createSectionPageNumberTracker, createSectionParityBlankPage, deepFreezeDocModel, defaultStarterModel, duplicateParagraph, ensureDocModelBlockIds, extractDocumentBlocks, extractDocumentReview, extractDocumentReviewSync, flattenDocumentBlocks, getPart, initWasm, insertParagraph, inspectDocumentSource, invalidateFormFieldSource, isOpenDocumentBlockId, isOpenDocumentModel, isSectionParityBlankPage, layoutDocument, materializeCollapsedTableGrid, modelToDocumentXml, multiplyDivideRound, normalizeDocModel, normalizeSectionPageNumber, normalizeTableRowHeightTwips, normalizedDocxBoundingBox, packageToArrayBuffer, paragraphAfterSpacingPx, paragraphBeforeSpacingPx, paragraphHasExplicitPageBreak, paragraphHasLastRenderedPageBreak, paragraphHasPageBreakBefore, paragraphLetterheadColumnGroupAtSegmentOffset, paragraphLetterheadFloatSideAtNodeIndex, paragraphMarginContribution, paragraphStartsWithLastRenderedPageBreak, parseDocx, parseDocxForViewer, parseParagraphsFromClipboard, parseSectionLayout, pasteParagraphs, planTableGridColumnDeletion, planTableGridColumnInsertion, projectCollapsedTableGrid, reduceFixedTableColumnWidths, invalidateFontMetrics as refreshDocxFontMetrics, rejectParagraphRevision, removeParagraph, replaceText, resolveDocumentForLayout, resolveDocumentLayout, resolveDocumentPageSegmentStartNodeIndex, resolveDocumentSectionsFromMetadata, resolveDocxBlockGeometries, resolveDocxCommentLinks, resolveDocxPageThumbnailResolution, resolveDocxSourceGeometry, resolveDocxWordGeometries, resolveEffectiveTableAlignment, resolveEffectiveTableBidiVisual, resolveEffectiveTablePreferredWidth, resolveFixedTableGridWidths, resolveFormFieldSourcePreview, resolveImagePaintBounds, resolvePaginationSectionMetricsIndexForNodeIndex, resolveParagraphBeforeSpacingPx, resolveSectionIndexForNodeIndex, resolveSectionPropertiesXmlForNodeIndex, resolveTableAlignmentOffset, resolveTableCellContentClipHeight, resolveTableGridSkipCount, resolveTableLeadingIndent, resolveTablePreferredWidthTwips, resolveTableRowHeight, revealDocxSource, scalePaginationSectionMetricsHeights, scorePaginationAgainstStoredPageBreaks, sectionBreakAfterParagraphStartsNewPage, sectionBreakPropertiesStartNewPage, sectionTitlePageEnabled, selectSectionVariantForPage, serializeDocModel, serializeDocx, serializeParagraphsForClipboard, setCommentResolved, setParagraphAlignment, setParagraphHeading, setRunColor, setRunHighlight, setWasmSource, splitParagraphChildrenAtTextOffsets, suppressParagraphBeforeSpacingAtPageStart, tableCellPhysicalGridRange, tableCellsIntersectingGridRange, tableGridColumnBound, tablePhysicalGridColumnCount, tableRowGridSkipCount, tableWidthTwipsToPixels, toDocxPageBlocks, toExtendParserBlocks, toggleRunStyleFlag, updateParagraphText, updateTableCellParagraphText, updateTableCellParagraphTextRecursive, updateTableCellText, useDocxBorders, useDocxComments, useDocxDocumentTheme, useDocxEditor, useDocxFormFields, useDocxImageWrapMenu, useDocxLineSpacing, useDocxModel, useDocxPageLayout, useDocxPageThumbnails, useDocxPagination, useDocxParagraphStyles, useDocxTrackChanges, useDocxViewerThumbnails, withPart };
