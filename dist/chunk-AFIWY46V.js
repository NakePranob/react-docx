// ../doc-model/src/index.ts
import {
  mapsToWasmPackage,
  wasmBuildDocModelFromBytes,
  wasmBuildDocModelFromPackage,
  wasmPackageToMaps
} from "@extend-ai/react-docx/internal";

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
function isOpenDocumentBlockId(blockId) {
  return blockId !== void 0 && /^docx:[0-9a-f]+\/(?:content|styles)\.xml\//.test(blockId);
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

// ../doc-model/src/image-paint.ts
function resolveImagePaintBounds(image, widthPx = image.widthPx, heightPx = image.heightPx) {
  const bounds = image.paintBounds;
  if (!bounds || !Number.isFinite(bounds.xPx) || !Number.isFinite(bounds.yPx)) return void 0;
  const positive = (value) => typeof value === "number" && Number.isFinite(value) && value > 0;
  if (![bounds.widthPx, bounds.heightPx, bounds.layoutWidthPx, bounds.layoutHeightPx].every(positive) || !positive(widthPx) || !positive(heightPx)) return void 0;
  const scaleX = widthPx / bounds.layoutWidthPx;
  const scaleY = heightPx / bounds.layoutHeightPx;
  const paint = {
    xPx: bounds.xPx * scaleX,
    yPx: bounds.yPx * scaleY,
    widthPx: bounds.widthPx * scaleX,
    heightPx: bounds.heightPx * scaleY
  };
  return Object.values(paint).every(Number.isFinite) && positive(paint.widthPx) && positive(paint.heightPx) ? paint : void 0;
}

// ../doc-model/src/blocks.ts
import {
  wasmExtractDocumentBlocks,
  wasmInspectDocumentSource
} from "@extend-ai/react-docx/internal";
var DOCUMENT_BLOCK_TYPES = [
  "text",
  "heading",
  "section_heading",
  "figure",
  "table",
  "table_head",
  "table_cell",
  "key_value",
  "page_number",
  "barcode",
  "formula",
  "header",
  "footer",
  "watermark",
  "legend"
];
async function extractDocumentBlocks(model, options = {}) {
  return await wasmExtractDocumentBlocks(model, options);
}
function flattenDocumentBlocks(blocks) {
  return blocks.flatMap((block) => [
    block,
    ...flattenDocumentBlocks(block.children)
  ]);
}
async function inspectDocumentSource(bytes) {
  return await wasmInspectDocumentSource(bytes);
}

// ../doc-model/src/review.ts
import { wasmExtractDocumentReview, wasmExtractDocumentReviewSync } from "@extend-ai/react-docx/internal";
async function extractDocumentReview(model) {
  return await wasmExtractDocumentReview(model);
}
function extractDocumentReviewSync(model) {
  return wasmExtractDocumentReviewSync(model);
}

// ../doc-model/src/table-grid.ts
function tableGridColumnBound(table) {
  if (table.style?.columnWidthsTwips !== void 0) return table.style.columnWidthsTwips.length;
  return Math.max(0, ...table.rows.map((row) => row.cells.reduce((total, cell) => {
    const span = cell.style?.gridSpan;
    return total + (Number.isFinite(span) && span > 0 ? Math.max(1, Math.floor(span)) : 1);
  }, 0)));
}
function resolveTableGridSkipCount(count, gridColumnCount) {
  const value = Number.isFinite(count) && count >= 0 ? Math.floor(count) : 0;
  return Number.isFinite(gridColumnCount) && gridColumnCount >= value ? value : 0;
}
function tableRowGridSkipCount(table, row, edge) {
  return resolveTableGridSkipCount(
    edge === "before" ? row.style?.gridBefore : row.style?.gridAfter,
    tableGridColumnBound(table)
  );
}

// ../doc-model/src/table-grid-projection.ts
function projectCollapsedTableGrid(sourceTable, options = {}) {
  const marker = sourceTable.style?.collapsedGridColumns;
  if (marker === void 0 || Array.isArray(marker) && marker.length === 0) {
    return { kind: "identity", sourceTable, table: sourceTable };
  }
  const reject = (reason) => ({ kind: "rejected", sourceTable, reason });
  if (!Array.isArray(marker)) return reject("invalid-collapsed-marker");
  if (sourceTable.style?.layout !== "fixed") return reject("fixed-layout-required");
  if (options.resolvedBorderSpacingPx !== 0) return reject("zero-resolved-spacing-not-confirmed");
  if (sourceTable.style.cellSpacingTwips !== void 0 && sourceTable.style.cellSpacingTwips !== 0) {
    return reject("nonzero-source-cell-spacing");
  }
  const widths = sourceTable.style.columnWidthsTwips;
  if (!Array.isArray(widths) || widths.length === 0 || widths.some((value) => !Number.isFinite(value) || value < 0)) {
    return reject("invalid-logical-grid");
  }
  const collapsed = /* @__PURE__ */ new Set();
  for (const index of marker) {
    if (!Number.isInteger(index) || index < 0 || index >= widths.length) return reject("invalid-collapsed-index");
    if (collapsed.has(index)) return reject("duplicate-collapsed-index");
    if (widths[index] !== 0) return reject("marked-column-has-positive-width");
    collapsed.add(index);
  }
  const logicalToPhysicalBoundary = [0];
  const physicalToLogicalColumn = [];
  const physicalWidths = [];
  for (let index = 0; index < widths.length; index += 1) {
    if (!collapsed.has(index)) {
      if (widths[index] === 0) return reject("unmarked-zero-column");
      physicalToLogicalColumn.push(index);
      physicalWidths.push(widths[index]);
    }
    logicalToPhysicalBoundary.push(physicalWidths.length);
  }
  if (physicalWidths.length === 0) return reject("empty-physical-grid");
  const physicalToLogicalBoundaryRanges = [];
  for (let logical = 0; logical < logicalToPhysicalBoundary.length; logical += 1) {
    const physical = logicalToPhysicalBoundary[logical];
    const range = physicalToLogicalBoundaryRanges[physical];
    if (range) range.end = logical;
    else physicalToLogicalBoundaryRanges[physical] = { start: logical, end: logical };
  }
  const physicalSums = [0];
  for (const value of physicalWidths) {
    const total = physicalSums.at(-1) + value;
    if (!Number.isFinite(total)) return reject("invalid-physical-grid-width");
    physicalSums.push(total);
  }
  if (!(physicalSums.at(-1) > 0)) return reject("invalid-physical-grid-width");
  const rows = [];
  for (const row of sourceTable.rows) {
    if ((row.style?.gridBefore ?? 0) !== 0 || (row.style?.gridAfter ?? 0) !== 0) {
      return reject("row-grid-gaps-not-supported");
    }
    let logicalStart = 0;
    let changed = false;
    const cells = [];
    for (const cell of row.cells) {
      const sourceSpan = cell.style?.gridSpan ?? 1;
      if (!Number.isInteger(sourceSpan) || sourceSpan <= 0) return reject("invalid-cell-span");
      const logicalEnd = logicalStart + sourceSpan;
      if (logicalEnd > widths.length) return reject("cell-exceeds-logical-grid");
      const physicalStart = logicalToPhysicalBoundary[logicalStart];
      const physicalEnd = logicalToPhysicalBoundary[logicalEnd];
      const span = physicalEnd - physicalStart;
      if (span <= 0) return reject("cell-collapses-to-empty-span");
      if (!(physicalSums[physicalEnd] - physicalSums[physicalStart] > 0)) {
        return reject("cell-has-nonpositive-physical-width");
      }
      if (span === sourceSpan) {
        cells.push(cell);
      } else {
        changed = true;
        const style2 = { ...cell.style };
        if (span === 1) delete style2.gridSpan;
        else style2.gridSpan = span;
        cells.push({ ...cell, style: style2 });
      }
      logicalStart = logicalEnd;
    }
    rows.push(changed ? { ...row, cells } : row);
  }
  const { collapsedGridColumns: _marker, ...style } = sourceTable.style;
  const table = { ...sourceTable, style: { ...style, columnWidthsTwips: physicalWidths }, rows };
  return {
    kind: "projected",
    sourceTable,
    table,
    logicalToPhysicalBoundary,
    physicalToLogicalBoundaryRanges,
    physicalToLogicalColumn,
    collapsedGridColumns: [...collapsed].sort((left, right) => left - right)
  };
}
function materializeCollapsedTableGrid(sourceTable, options = {}) {
  const projection = projectCollapsedTableGrid(sourceTable, options);
  if (projection.kind !== "projected") return projection;
  const table = { ...projection.table };
  delete table.sourceXml;
  delete table.sourceTextPatches;
  return { ...projection, table };
}

// ../doc-model/src/block-id.ts
var nextBlockIdValue = 1;
function allocateBlockId() {
  const value = nextBlockIdValue;
  nextBlockIdValue += 1;
  return `b${value}`;
}
function ensureNodeBlockIds(node) {
  if (!node.blockId) {
    node.blockId = allocateBlockId();
  }
  if (node.type !== "table") {
    node.children.forEach((child, index) => {
      if (child.type !== "image") return;
      child.blockId ??= `${node.blockId}/child/${index}`;
      child.contentNodes?.forEach(ensureNodeBlockIds);
    });
    return;
  }
  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const child of cell.nodes) {
        ensureNodeBlockIds(child);
      }
    }
  }
}
function ensureDocModelBlockIds(model) {
  model.nodes.forEach(ensureNodeBlockIds);
  for (const section of model.metadata.headerSections ?? []) {
    section.nodes.forEach(ensureNodeBlockIds);
  }
  for (const section of model.metadata.footerSections ?? []) {
    section.nodes.forEach(ensureNodeBlockIds);
  }
  for (const section of model.metadata.sections ?? []) {
    for (const headerSection of section.headerSections ?? []) {
      headerSection.nodes.forEach(ensureNodeBlockIds);
    }
    for (const footerSection of section.footerSections ?? []) {
      footerSection.nodes.forEach(ensureNodeBlockIds);
    }
  }
  for (const note of model.metadata.footnotes ?? []) {
    note.nodes?.forEach(ensureNodeBlockIds);
  }
  for (const note of model.metadata.endnotes ?? []) {
    note.nodes?.forEach(ensureNodeBlockIds);
  }
  for (const comment of model.metadata.comments ?? []) {
    comment.nodes?.forEach(ensureNodeBlockIds);
  }
  return model;
}
function collectNodeBlockIds(node, seen, duplicates) {
  if (node.blockId) {
    if (seen.has(node.blockId)) {
      duplicates.add(node.blockId);
    }
    seen.add(node.blockId);
  }
  if (node.type !== "table") {
    return;
  }
  for (const row of node.rows) {
    for (const cell of row.cells) {
      for (const child of cell.nodes) {
        collectNodeBlockIds(child, seen, duplicates);
      }
    }
  }
}
function collectDuplicateDocModelBlockIds(model) {
  const duplicates = /* @__PURE__ */ new Set();
  const visitContainer = (nodes) => {
    if (!nodes) {
      return;
    }
    const seen = /* @__PURE__ */ new Set();
    nodes.forEach((node) => collectNodeBlockIds(node, seen, duplicates));
  };
  visitContainer(model.nodes);
  for (const section of model.metadata.headerSections ?? []) {
    visitContainer(section.nodes);
  }
  for (const section of model.metadata.footerSections ?? []) {
    visitContainer(section.nodes);
  }
  for (const note of model.metadata.footnotes ?? []) {
    visitContainer(note.nodes);
  }
  for (const note of model.metadata.endnotes ?? []) {
    visitContainer(note.nodes);
  }
  return Array.from(duplicates);
}

// ../doc-model/src/clone.ts
function isParagraphCellContent(node) {
  return node.type === "paragraph";
}
function isTableCellContentTable(node) {
  return node.type === "table";
}
function cloneTableCellContent(nodes) {
  return nodes.map((node) => {
    if (isParagraphCellContent(node)) {
      return cloneParagraph(node);
    }
    if (isTableCellContentTable(node)) {
      return cloneTable(node);
    }
    return node;
  });
}
function cloneParagraphNumbering(numbering) {
  return numbering ? { ...numbering } : void 0;
}
function cloneParagraphSpacing(spacing) {
  return spacing ? { ...spacing } : void 0;
}
function cloneParagraphIndent(indent) {
  return indent ? { ...indent } : void 0;
}
function cloneParagraphBorderStyle(border) {
  return border ? { ...border } : void 0;
}
function cloneParagraphBorderSet(borders) {
  if (!borders) {
    return void 0;
  }
  return {
    top: cloneParagraphBorderStyle(borders.top),
    right: cloneParagraphBorderStyle(borders.right),
    bottom: cloneParagraphBorderStyle(borders.bottom),
    left: cloneParagraphBorderStyle(borders.left),
    between: cloneParagraphBorderStyle(borders.between),
    bar: cloneParagraphBorderStyle(borders.bar)
  };
}
function cloneParagraphStyle(style) {
  if (!style) {
    return void 0;
  }
  return {
    ...style,
    numbering: cloneParagraphNumbering(style.numbering),
    spacing: cloneParagraphSpacing(style.spacing),
    indent: cloneParagraphIndent(style.indent),
    borders: cloneParagraphBorderSet(style.borders),
    tabStops: style.tabStops?.map((stop) => ({ ...stop })),
    dropCap: style.dropCap ? {
      ...style.dropCap
    } : void 0
  };
}
function cloneTextStyle(style) {
  return style ? {
    ...style,
    runBorder: style.runBorder ? { ...style.runBorder } : void 0,
    textShadow: style.textShadow ? { ...style.textShadow } : void 0,
    textOutline: style.textOutline ? { ...style.textOutline } : void 0
  } : void 0;
}
function cloneParagraphMarkFormatting(formatting) {
  return formatting ? {
    defaultStyle: cloneTextStyle(formatting.defaultStyle),
    characterStyle: cloneTextStyle(formatting.characterStyle),
    directStyle: cloneTextStyle(formatting.directStyle)
  } : void 0;
}
function cloneParagraphNode(paragraph) {
  return cloneParagraph(paragraph);
}
function cloneTableNode(table) {
  return cloneTable(table);
}
function cloneImagePropertyNode(node) {
  return {
    ...node,
    attributes: node.attributes?.map((attribute) => ({ ...attribute })),
    children: node.children?.map(cloneImagePropertyNode)
  };
}
function cloneImagePresentation(presentation) {
  return {
    ...presentation,
    assets: presentation.assets?.map((asset) => ({ ...asset })),
    transform: presentation.transform ? { ...presentation.transform } : void 0,
    sourceRectangle: presentation.sourceRectangle ? { ...presentation.sourceRectangle } : void 0,
    fillRectangle: presentation.fillRectangle ? { ...presentation.fillRectangle } : void 0,
    tile: presentation.tile ? { ...presentation.tile } : void 0,
    properties: presentation.properties?.map(cloneImagePropertyNode)
  };
}
function cloneImageRun(image) {
  return {
    ...image,
    chartData: image.chartData && structuredClone(image.chartData),
    model3d: image.model3d && structuredClone(image.model3d),
    presentation: image.presentation ? cloneImagePresentation(image.presentation) : void 0,
    contentNodes: image.contentNodes?.map(cloneDocNode),
    data: image.data ? new Uint8Array(image.data) : void 0,
    crop: image.crop ? { ...image.crop } : void 0,
    paintBounds: image.paintBounds ? { ...image.paintBounds } : void 0,
    floating: image.floating ? { ...image.floating } : void 0
  };
}
function cloneFormFieldSourcePreview(preview) {
  return preview ? { ...preview, image: cloneImageRun(preview.image) } : void 0;
}
function cloneParagraph(paragraph) {
  return {
    type: "paragraph",
    blockId: paragraph.blockId,
    style: cloneParagraphStyle(paragraph.style),
    paragraphMarkStyle: cloneTextStyle(paragraph.paragraphMarkStyle),
    sourceParagraphMarkStyle: cloneTextStyle(paragraph.sourceParagraphMarkStyle),
    sourceParagraphMarkFormatting: cloneParagraphMarkFormatting(
      paragraph.sourceParagraphMarkFormatting
    ),
    sourceParagraphMarkPropertiesXml: paragraph.sourceParagraphMarkPropertiesXml,
    paragraphMarkDeleted: paragraph.paragraphMarkDeleted,
    sourceXml: paragraph.sourceXml,
    commentRanges: paragraph.commentRanges?.map((range) => ({ ...range })),
    reviewContent: paragraph.reviewContent?.map((fragment) => ({
      sourceXml: fragment.sourceXml,
      children: cloneParagraph({ type: "paragraph", children: fragment.children }).children
    })),
    sourceTextPatch: paragraph.sourceTextPatch ? {
      runs: paragraph.sourceTextPatch.runs.map((run) => ({
        style: cloneTextStyle(run.style),
        link: run.link,
        noteReference: run.noteReference ? { ...run.noteReference } : void 0
      }))
    } : void 0,
    sourceRunProvenance: paragraph.sourceRunProvenance ? {
      runs: paragraph.sourceRunProvenance.runs.map((run) => ({
        style: cloneTextStyle(run.style),
        link: run.link,
        noteReference: run.noteReference ? { ...run.noteReference } : void 0
      }))
    } : void 0,
    children: paragraph.children.map((child) => {
      if (child.type === "text") {
        return {
          type: "text",
          text: child.text,
          fieldInstruction: child.fieldInstruction,
          style: cloneTextStyle(child.style),
          link: child.link,
          noteReference: child.noteReference ? { ...child.noteReference } : void 0
        };
      }
      if (child.type === "form-field") {
        return {
          type: "form-field",
          fieldType: child.fieldType,
          sourceKind: child.sourceKind,
          id: child.id,
          tag: child.tag,
          title: child.title,
          placeholder: child.placeholder,
          checked: child.checked,
          value: child.value,
          options: child.options?.map((option) => ({
            displayText: option.displayText,
            value: option.value
          })),
          widget: child.widget ? {
            name: child.widget.name,
            enabled: child.widget.enabled,
            calcOnExit: child.widget.calcOnExit,
            text: child.widget.text ? {
              inputType: child.widget.text.inputType,
              defaultText: child.widget.text.defaultText,
              maxLength: child.widget.text.maxLength,
              textFormat: child.widget.text.textFormat
            } : void 0,
            checkbox: child.widget.checkbox ? {
              defaultChecked: child.widget.checkbox.defaultChecked,
              sizeMode: child.widget.checkbox.sizeMode,
              sizePt: child.widget.checkbox.sizePt
            } : void 0,
            dropdown: child.widget.dropdown ? {
              defaultValue: child.widget.dropdown.defaultValue
            } : void 0
          } : void 0,
          checkedSymbol: child.checkedSymbol,
          uncheckedSymbol: child.uncheckedSymbol,
          style: child.style ? { ...child.style } : void 0,
          link: child.link,
          sourceXml: child.sourceXml,
          sourcePreview: cloneFormFieldSourcePreview(child.sourcePreview)
        };
      }
      return cloneImageRun(child);
    })
  };
}
function cloneTableBoxSpacing(spacing) {
  if (!spacing) {
    return void 0;
  }
  return {
    topTwips: spacing.topTwips,
    rightTwips: spacing.rightTwips,
    bottomTwips: spacing.bottomTwips,
    leftTwips: spacing.leftTwips
  };
}
function cloneTableBorderStyle(border) {
  if (!border) {
    return void 0;
  }
  return {
    type: border.type,
    color: border.color,
    sizeEighthPt: border.sizeEighthPt
  };
}
function cloneTableBorderSet(borders) {
  if (!borders) {
    return void 0;
  }
  return {
    top: cloneTableBorderStyle(borders.top),
    right: cloneTableBorderStyle(borders.right),
    bottom: cloneTableBorderStyle(borders.bottom),
    left: cloneTableBorderStyle(borders.left),
    insideH: cloneTableBorderStyle(borders.insideH),
    insideV: cloneTableBorderStyle(borders.insideV),
    tl2br: cloneTableBorderStyle(borders.tl2br),
    tr2bl: cloneTableBorderStyle(borders.tr2bl)
  };
}
function cloneTableFloatingStyle(floating) {
  if (!floating) {
    return void 0;
  }
  return {
    xTwips: floating.xTwips,
    yTwips: floating.yTwips,
    leftFromTextTwips: floating.leftFromTextTwips,
    rightFromTextTwips: floating.rightFromTextTwips,
    topFromTextTwips: floating.topFromTextTwips,
    bottomFromTextTwips: floating.bottomFromTextTwips,
    horizontalAnchor: floating.horizontalAnchor,
    verticalAnchor: floating.verticalAnchor,
    horizontalAlign: floating.horizontalAlign,
    verticalAlign: floating.verticalAlign
  };
}
function cloneTablePreferredWidth(width) {
  return width ? { ...width } : void 0;
}
function cloneTableWidthSource(source) {
  return source ? {
    ...source,
    preferredWidth: cloneTablePreferredWidth(source.preferredWidth),
    inheritedPreferredWidth: cloneTablePreferredWidth(source.inheritedPreferredWidth)
  } : void 0;
}
function cloneTableRowStyle(style) {
  return style ? {
    ...style,
    widthBefore: cloneTablePreferredWidth(style.widthBefore),
    widthAfter: cloneTablePreferredWidth(style.widthAfter),
    sourceRowGeometry: style.sourceRowGeometry ? {
      ...style.sourceRowGeometry,
      widthBefore: cloneTablePreferredWidth(style.sourceRowGeometry.widthBefore),
      widthAfter: cloneTablePreferredWidth(style.sourceRowGeometry.widthAfter)
    } : void 0
  } : void 0;
}
function cloneTable(table) {
  return {
    type: "table",
    blockId: table.blockId,
    sourceXml: table.sourceXml,
    reviewRows: table.reviewRows?.map((fragment) => ({
      ...fragment,
      row: cloneTable({ type: "table", rows: [fragment.row] }).rows[0]
    })),
    sourceTextPatches: table.sourceTextPatches?.map((patch) => ({
      sourceParagraphXml: patch.sourceParagraphXml,
      paragraph: cloneParagraph(patch.paragraph)
    })),
    style: table.style ? {
      styleId: table.style.styleId,
      sourceStyleId: table.style.sourceStyleId,
      widthTwips: table.style.widthTwips,
      preferredWidth: cloneTablePreferredWidth(table.style.preferredWidth),
      sourceWidth: cloneTableWidthSource(table.style.sourceWidth),
      alignment: table.style.alignment,
      sourceAlignment: table.style.sourceAlignment,
      sourceInheritedAlignment: table.style.sourceInheritedAlignment,
      bidiVisual: table.style.bidiVisual,
      sourceBidiVisual: table.style.sourceBidiVisual,
      sourceInheritedBidiVisual: table.style.sourceInheritedBidiVisual,
      indentTwips: table.style.indentTwips,
      layout: table.style.layout,
      cellSpacingTwips: table.style.cellSpacingTwips,
      floating: cloneTableFloatingStyle(table.style.floating),
      cellMarginTwips: cloneTableBoxSpacing(table.style.cellMarginTwips),
      columnWidthsTwips: table.style.columnWidthsTwips ? [...table.style.columnWidthsTwips] : void 0,
      collapsedGridColumns: table.style.collapsedGridColumns ? [...table.style.collapsedGridColumns] : void 0,
      borders: cloneTableBorderSet(table.style.borders)
    } : void 0,
    rows: table.rows.map((row) => ({
      type: "table-row",
      style: cloneTableRowStyle(row.style),
      cells: row.cells.map((cell) => ({
        type: "table-cell",
        style: cell.style ? {
          ...cell.style,
          preferredWidth: cloneTablePreferredWidth(cell.style.preferredWidth),
          sourceWidth: cloneTableWidthSource(cell.style.sourceWidth),
          marginTwips: cloneTableBoxSpacing(cell.style.marginTwips),
          borders: cloneTableBorderSet(cell.style.borders)
        } : void 0,
        nodes: cloneTableCellContent(cell.nodes)
      }))
    }))
  };
}
function cloneDocNode(node) {
  return node.type === "paragraph" ? cloneParagraph(node) : cloneTable(node);
}
function cloneNumberingDefinitions(numberingDefinitions) {
  if (!numberingDefinitions) {
    return void 0;
  }
  return {
    abstracts: numberingDefinitions.abstracts.map((abstractDefinition) => ({
      abstractNumId: abstractDefinition.abstractNumId,
      levels: abstractDefinition.levels.map((level) => ({
        ...level,
        runStyle: level.runStyle ? { ...level.runStyle } : void 0,
        pictureBullet: level.pictureBullet ? { ...level.pictureBullet } : void 0
      }))
    })),
    instances: numberingDefinitions.instances.map((instanceDefinition) => ({
      numId: instanceDefinition.numId,
      abstractNumId: instanceDefinition.abstractNumId,
      levelStartOverrides: instanceDefinition.levelStartOverrides ? { ...instanceDefinition.levelStartOverrides } : void 0,
      levelOverrides: instanceDefinition.levelOverrides ? instanceDefinition.levelOverrides.map((level) => ({
        ...level,
        runStyle: level.runStyle ? { ...level.runStyle } : void 0,
        pictureBullet: level.pictureBullet ? { ...level.pictureBullet } : void 0
      })) : void 0
    }))
  };
}
function cloneDocModel(model) {
  return {
    nodes: model.nodes.map(cloneDocNode),
    metadata: {
      documentId: model.metadata.documentId,
      sourcePartNames: model.metadata.sourcePartNames ? { ...model.metadata.sourcePartNames } : void 0,
      sourceParts: model.metadata.sourceParts,
      warnings: [...model.metadata.warnings],
      documentPageCount: model.metadata.documentPageCount,
      documentOpenTag: model.metadata.documentOpenTag,
      documentBackgroundColor: model.metadata.documentBackgroundColor,
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      sections: model.metadata.sections?.map((section) => ({
        startNodeIndex: section.startNodeIndex,
        sectionPropertiesXml: section.sectionPropertiesXml,
        headerSections: (section.headerSections ?? []).map((headerSection) => ({
          partName: headerSection.partName,
          referenceType: headerSection.referenceType,
          nodes: headerSection.nodes.map(cloneDocNode)
        })),
        footerSections: (section.footerSections ?? []).map((footerSection) => ({
          partName: footerSection.partName,
          referenceType: footerSection.referenceType,
          nodes: footerSection.nodes.map(cloneDocNode)
        }))
      })),
      headerSections: (model.metadata.headerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map(cloneDocNode)
      })),
      footerSections: (model.metadata.footerSections ?? []).map((section) => ({
        partName: section.partName,
        referenceType: section.referenceType,
        nodes: section.nodes.map(cloneDocNode)
      })),
      paragraphStyles: (model.metadata.paragraphStyles ?? []).map((style) => ({
        ...style,
        runStyle: style.runStyle ? { ...style.runStyle } : void 0,
        numbering: cloneParagraphNumbering(style.numbering),
        spacing: cloneParagraphSpacing(style.spacing),
        indent: cloneParagraphIndent(style.indent),
        borders: cloneParagraphBorderSet(style.borders)
      })),
      defaultParagraphStyleId: model.metadata.defaultParagraphStyleId,
      defaultTabStopTwips: model.metadata.defaultTabStopTwips,
      documentSettingsImported: model.metadata.documentSettingsImported,
      lineSpacingCompatibilityImported: model.metadata.lineSpacingCompatibilityImported,
      numberingDefinitions: cloneNumberingDefinitions(
        model.metadata.numberingDefinitions
      ),
      compatibility: model.metadata.compatibility ? { ...model.metadata.compatibility } : void 0,
      footnotes: model.metadata.footnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(cloneDocNode)
      })),
      endnotes: model.metadata.endnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(cloneDocNode)
      })),
      comments: model.metadata.comments?.map((comment) => ({ ...comment, nodes: comment.nodes?.map(cloneDocNode) }))
    }
  };
}

// ../doc-model/src/freeze.ts
function deepFreezeValue(value) {
  if (typeof value !== "object" || value === null) {
    return;
  }
  if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) {
    return;
  }
  if (Object.isFrozen(value)) {
    return;
  }
  Object.freeze(value);
  if (Array.isArray(value)) {
    for (const entry of value) {
      deepFreezeValue(entry);
    }
    return;
  }
  for (const key of Object.keys(value)) {
    deepFreezeValue(value[key]);
  }
}
function deepFreezeDocModel(model) {
  deepFreezeValue(model);
  return model;
}

// ../doc-model/src/index.ts
async function buildDocModel(pkg, options) {
  const wasmPackage = mapsToWasmPackage({
    parts: pkg.parts,
    binaryAssets: pkg.binaryAssets,
    warnings: pkg.warnings
  });
  const model = await wasmBuildDocModelFromPackage(wasmPackage, options);
  return normalizeDocModel(model);
}
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

export {
  resolveFormFieldSourcePreview,
  invalidateFormFieldSource,
  isOpenDocumentModel,
  isOpenDocumentBlockId,
  normalizeDocModel,
  resolveImagePaintBounds,
  DOCUMENT_BLOCK_TYPES,
  extractDocumentBlocks,
  flattenDocumentBlocks,
  inspectDocumentSource,
  extractDocumentReview,
  extractDocumentReviewSync,
  tableGridColumnBound,
  resolveTableGridSkipCount,
  tableRowGridSkipCount,
  projectCollapsedTableGrid,
  materializeCollapsedTableGrid,
  allocateBlockId,
  ensureDocModelBlockIds,
  collectDuplicateDocModelBlockIds,
  cloneParagraphMarkFormatting,
  cloneParagraphNode,
  cloneTableNode,
  cloneFormFieldSourcePreview,
  cloneDocModel,
  deepFreezeDocModel,
  buildDocModel,
  buildDocModelFromBytes
};
//# sourceMappingURL=chunk-AFIWY46V.js.map