export interface PaintPageOptions {
  readonly scale: number;
  readonly dpr: number;
}

export interface CanvasPaintContext {
  readonly ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  readonly scale: number;
  readonly dpr: number;
}
