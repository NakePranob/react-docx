import { Properties } from 'csstype';
import * as React from 'react';

type ViewerZoomMode = "automatic" | "fit-page" | "fit-width";
type ViewerZoomLevel = number | ViewerZoomMode;
type ViewerZoomState = {
    level: ViewerZoomLevel;
    resolvedZoom: number;
};

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

interface DocxSourceRef {
    documentId: string;
    nodeId: string;
    partName: string;
    story: string;
    start?: number;
    end?: number;
    offsetEncoding: "utf-16";
}

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

type AnnotationMutationFailureReason = "stale" | "unsupported" | "unsafe-xml";

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

type DocxDocumentTheme = "light" | "dark";
type DocxHeadingStyleMap = Partial<Record<HeadingLevel, Properties<string | number>>>;
type DocxImageWrapMode = "inline" | "square" | "tight" | "through" | "topAndBottom" | "behindText" | "inFrontOfText";
interface DocxImageWrapState {
    mode: DocxImageWrapMode;
    moveWithText: boolean;
    fixedPositionOnPage: boolean;
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
interface DocxPaginationInfo {
    currentPage: number;
    totalPages: number;
}
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

type DocxVanillaContent = Node | string | number | null;
interface DocxVanillaViewerOptions extends UseDocxEditorOptions {
    file?: ArrayBuffer | Blob;
    src?: string;
    fileName?: string;
    className?: string;
    style?: Properties<string | number>;
    height?: Properties<string | number>["height"];
    zoom?: ViewerZoomLevel;
    defaultZoom?: ViewerZoomLevel;
    onZoomChange?: (state: ViewerZoomState) => void;
    pageBackgroundColor?: string;
    pageGapBackgroundColor?: string;
    deferInitialPaginationPaint?: boolean;
    pageVirtualization?: DocxPageVirtualizationOptions;
    visiblePageRange?: DocxVisiblePageRange;
    pageIndexes?: readonly number[];
    onPageCountChange?: (pageCount: number) => void;
    onRequestPageReveal?: (pageIndex: number) => void;
    headingStyles?: DocxHeadingStyleMap;
    mode?: DocxEditorViewerMode;
    showTrackedChanges?: boolean;
    showComments?: boolean;
    onFormFieldDoubleClick?: (location: DocxFormFieldLocation) => void;
    emptyState?: DocxVanillaContent;
    loadingState?: DocxVanillaContent;
    errorState?: DocxVanillaContent | ((error: Error) => DocxVanillaContent);
    renderTrackedChangeCard?: (props: DocxTrackedChangeCardRenderProps) => DocxVanillaContent;
    renderCommentCard?: (props: DocxCommentCardRenderProps) => DocxVanillaContent;
    renderTableContextMenu?: (props: DocxTableContextMenuRenderProps) => DocxVanillaContent;
    renderContextMenu?: (props: DocxContextMenuRenderProps) => DocxVanillaContent;
}
interface DocxVanillaViewerController extends DocxEditorController {
    isLoading: boolean;
    error?: Error;
    hasDocument: boolean;
}
interface DocxVanillaViewer {
    readonly element: HTMLElement;
    getController(): DocxVanillaViewerController | null;
    reload(): void;
    setOptions(options: Partial<DocxVanillaViewerOptions>): void;
    subscribe(listener: (controller: DocxVanillaViewerController) => void): () => void;
    destroy(): void;
}
declare class DocxViewer implements DocxVanillaViewer {
    readonly element: HTMLElement;
    private readonly viewer;
    constructor(container: Element | DocumentFragment, options?: DocxVanillaViewerOptions);
    getController(): DocxVanillaViewerController | null;
    setOptions(options: Partial<DocxVanillaViewerOptions>): void;
    reload(): void;
    subscribe(listener: (controller: DocxVanillaViewerController) => void): () => void;
    destroy(): void;
}
declare function createDocxViewer(container: Element | DocumentFragment, options?: DocxVanillaViewerOptions): DocxViewer;

export { type DocModel, type DocxCommentCardRenderProps, type DocxCommentLink, type DocxContextMenuRenderProps, type DocxEditorController, type DocxEditorViewerMode, type DocxFormFieldLocation, type DocxHeadingStyleMap, type DocxPageVirtualizationOptions, type DocxTableContextMenuRenderProps, type DocxTrackedChangeCardRenderProps, type DocxVanillaContent, type DocxVanillaViewer, type DocxVanillaViewerController, type DocxVanillaViewerOptions, DocxViewer, type DocxViewerSource, type DocxVisiblePageRange, type ParseDocxForViewerOptions, type ParsedDocxDocument, type UseDocxEditorOptions, type ViewerZoomLevel, type ViewerZoomState, type WasmSource, createDocxViewer, initWasm, parseDocxForViewer, resolveDocxCommentLinks, setWasmSource };
