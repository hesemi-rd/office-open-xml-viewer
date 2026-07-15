import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { ParagraphLayoutContext } from './layout-context.js';
import { acquireShapeTextBoxLayout } from './layout/paragraph.js';
import { paintTextBoxLayout } from './paint/canvas-text.js';
import { createLayoutServices, type DecodedImage, type RenderState } from './renderer.js';
import { textBoxAcquisitionInput } from './parser-model.js';
import type { DocxDocumentModel, ShapeRun } from './types.js';

function servicesFor(
  ctx: CanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string>,
): ReturnType<typeof createLayoutServices> {
  return createLayoutServices({
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses,
  } as DocxDocumentModel, { measureContext: ctx });
}

/** Test-only adapter over the production retained APIs. It intentionally has no
 * production caller: tests use it to acquire once, then exercise measure-free
 * point-geometry paint at an arbitrary viewport scale. */
export function acquireAndPaintShapeTextBox(
  shape: ShapeRun,
  x: number, y: number, w: number, h: number,
  ctx: CanvasRenderingContext2D,
  scale: number,
  fontFamilyClasses: Record<string, string> = {},
  images: Map<string, DecodedImage> = new Map(),
  state?: RenderState,
): void {
  const services = state?.layoutServices ?? servicesFor(ctx, fontFamilyClasses);
  const grid = state?.docGrid;
  const lineGridActive = grid?.linePitchPt != null && grid.linePitchPt > 0
    && (grid.type === 'lines' || grid.type === 'linesAndChars' || grid.type === 'snapToChars');
  const characterGridActive = grid?.charSpacePt != null
    && (grid.type === 'linesAndChars' || grid.type === 'snapToChars');
  const context: ParagraphLayoutContext = {
    lineGrid: { active: lineGridActive, pitchPt: lineGridActive ? grid?.linePitchPt ?? null : null },
    characterGrid: {
      active: characterGridActive,
      deltaPt: characterGridActive ? grid.charSpacePt ?? 0 : 0,
    },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
    baseRtl: false, isJustified: false, stretchLastLine: false,
    tabStops: [], hasRuby: false, hasEastAsianText: false,
    kinsoku: state?.kinsoku ?? DEFAULT_KINSOKU_RULES,
    defaultTabPt: state?.defaultTabPt ?? 36,
    mathDefJc: state?.mathDefJc,
  };
  const source = { story: 'textbox' as const, storyInstance: 'test-shape', path: [] as number[] };
  const layout = acquireShapeTextBoxLayout(shape, {
    xPt: x / scale, yPt: y / scale, widthPt: w / scale, heightPt: h / scale,
  }, {
    id: 'test-shape-textbox',
    source,
    flowDomainId: 'test-shape', context,
    measurer: { context: ctx, fontFamilyClasses },
    environment: {
      pageIndex: state?.pageIndex ?? 0,
      totalPages: state?.totalPages ?? 1,
      displayPageNumber: state?.displayPageNumber,
      pageNumberFormat: state?.pageNumberFormat,
      currentDateMs: state?.currentDateMs,
      noteNumbers: state?.noteNumbers,
      currentNoteNumber: state?.currentNoteNumber,
      verticalCJK: state?.verticalCJK,
      documentHasEastAsianText: state?.docEastAsian
        ?? shape.textBlocks?.some((block) => /[\u3000-\u9fff\uf900-\ufaff]/u.test(block.text))
        ?? false,
      resolvedLocalFonts: state?.resolvedLocalFonts ?? services.text.localMetrics,
      layoutServices: services,
    },
    input: textBoxAcquisitionInput(shape, source),
  });
  if (!layout) return;
  const viewportContext = scale === 1 ? ctx : new Proxy(ctx, {
    get(target, property) {
      if (property === 'fillText') return (text: string, px: number, py: number) =>
        target.fillText(text, px * scale, py * scale);
      if (property === 'fillRect') return (px: number, py: number, pw: number, ph: number) =>
        target.fillRect(px * scale, py * scale, pw * scale, ph * scale);
      if (property === 'strokeRect') return (px: number, py: number, pw: number, ph: number) =>
        target.strokeRect(px * scale, py * scale, pw * scale, ph * scale);
      if (property === 'translate') return (px: number, py: number) =>
        target.translate(px * scale, py * scale);
      if (property === 'moveTo') return (px: number, py: number) =>
        target.moveTo(px * scale, py * scale);
      if (property === 'lineTo') return (px: number, py: number) =>
        target.lineTo(px * scale, py * scale);
      if (property === 'drawImage') return (
        image: CanvasImageSource, px: number, py: number, pw: number, ph: number,
      ) => target.drawImage(image, px * scale, py * scale, pw * scale, ph * scale);
      const value = Reflect.get(target, property, target);
      return typeof value === 'function' ? value.bind(target) : value;
    },
    set(target, property, value) {
      if (property === 'font' && typeof value === 'string') {
        target.font = value.replace(/([\d.]+)px/u, (_, size: string) => `${Number(size) * scale}px`);
        return true;
      }
      if (property === 'letterSpacing' && typeof value === 'string') {
        target.letterSpacing = value.replace(/([\d.-]+)px/u, (_, size: string) => `${Number(size) * scale}px`);
        return true;
      }
      return Reflect.set(target, property, value, target);
    },
  });
  const imageByResourceKey = new Map<string, DecodedImage>();
  layout.paragraphs.forEach((paragraph, blockIndex) => {
    const block = shape.textBlocks?.[blockIndex];
    const image = block?.imagePath ? images.get(block.imagePath) : undefined;
    if (!image) return;
    for (const resource of paragraph.resources) {
      if (resource.kind === 'image') imageByResourceKey.set(resource.resourceKey, image);
    }
  });
  paintTextBoxLayout(layout, {
    ctx: viewportContext, scale, dpr: state?.dpr ?? 1,
    defaultTextColor: shape.defaultTextColor
      ? `#${shape.defaultTextColor}` : state?.defaultColor ?? '#000000',
    showTrackChanges: state?.showTrackChanges,
    resources: {
      paint(resourceKey, kind, bounds, paintContext) {
        if (kind !== 'image') return;
        const image = imageByResourceKey.get(resourceKey);
        if (!image) return;
        paintContext.drawImage(
          image,
          bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt,
        );
      },
    },
    onTextRun: state?.onTextRun,
  });
}
