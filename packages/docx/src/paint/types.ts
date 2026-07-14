export interface PaintPageOptions {
  readonly scale: number;
  readonly dpr: number;
}

export interface PaintCanvas2D {
  fillStyle: string | CanvasGradient | CanvasPattern;
  fillRect(x: number, y: number, width: number, height: number): void;
}

export interface CanvasPaintContext {
  readonly ctx: PaintCanvas2D;
  readonly scale: number;
  readonly dpr: number;
}
