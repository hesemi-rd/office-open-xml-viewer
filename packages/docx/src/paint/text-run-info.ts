import type { HyperlinkTarget } from '@silurus/ooxml-core';

/** Shared overlay payload emitted by both retained and adapter text painters. */
export interface TextRunPaintInfo {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly fontSize: number;
  readonly font: string;
  readonly letterSpacingPx?: number;
  readonly transform?: string;
  readonly hyperlink?: HyperlinkTarget;
  readonly eastAsianVert?: boolean;
}

/** Keeps optional overlay fields absent instead of materializing `undefined`. */
export function textRunPaintInfo(info: TextRunPaintInfo): TextRunPaintInfo {
  return {
    text: info.text,
    x: info.x,
    y: info.y,
    w: info.w,
    h: info.h,
    fontSize: info.fontSize,
    font: info.font,
    ...(info.letterSpacingPx !== undefined ? { letterSpacingPx: info.letterSpacingPx } : {}),
    ...(info.transform !== undefined ? { transform: info.transform } : {}),
    ...(info.hyperlink !== undefined ? { hyperlink: info.hyperlink } : {}),
    ...(info.eastAsianVert !== undefined ? { eastAsianVert: info.eastAsianVert } : {}),
  };
}
