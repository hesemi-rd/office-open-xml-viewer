import type { DocComment, DocNote, RenderPageOptions, WorkerResponse } from './types';

/** Lightweight summary returned by the render worker's `parse` — everything
 *  the main-thread proxy needs for its synchronous getters. The full model
 *  stays in the worker. */
export interface DocumentMeta {
  pageCount: number;
  comments: DocComment[];
  footnotes: DocNote[];
  endnotes: DocNote[];
  /** ECMA-376 §17.6.13 / §17.6.11 — per-page page size (pt), one entry per page,
   *  index-aligned with `pageCount`. Built worker-side from the paginated pages'
   *  `sectionGeom` so the main thread can lay out (e.g. a scroll viewer's spacer)
   *  without the full model. Genuinely per-page for a mixed-geometry document. */
  pageSizes: { widthPt: number; heightPt: number }[];
}

/** Serializable subset of RenderPageOptions (callbacks cannot cross the wire). */
export type WireRenderPageOptions = Omit<RenderPageOptions, 'onTextRun'>;

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; useGoogleFonts?: boolean }
  | { type: 'renderPage'; id: number; pageIndex: number; opts: WireRenderPageOptions }
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerResponse =
  | Exclude<WorkerResponse, { type: 'parsed' }>
  | { type: 'parsedMeta'; id: number; meta: DocumentMeta }
  | { type: 'pageRendered'; id: number; bitmap: ImageBitmap };
