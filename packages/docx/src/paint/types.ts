export interface PaintPageOptions {
  readonly scale: number;
  readonly dpr: number;
}

export interface PaintCanvas2D {
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  fillRect(x: number, y: number, width: number, height: number): void;
  strokeRect(x: number, y: number, width: number, height: number): void;
  setLineDash(segments: number[]): void;
  fillText(text: string, x: number, y: number): void;
}

export interface CanvasPaintContext {
  readonly ctx: PaintCanvas2D;
  readonly scale: number;
  readonly dpr: number;
}
