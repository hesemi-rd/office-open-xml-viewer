/**
 * Indirection module for the render worker. `?worker&inline` embeds the whole
 * worker bundle (renderer + core) as a base64 string, so this module is only
 * reachable via dynamic import — main-mode users never download it.
 */
import RenderWorker from './render-worker.ts?worker&inline';

export function createRenderWorker(): Worker {
  return new RenderWorker();
}
