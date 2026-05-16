export { parsePptx, extractMedia as extractPptxMedia } from './pptx';
export { parseDocx } from './docx';
export { parseXlsx, parseSheet as parseXlsxSheet, parseXlsxAllSheets } from './xlsx';
export {
  renderSlideNode,
  installImageBitmapShim,
  type NodeCanvasLike,
  type NodeCanvasFactory,
  type NodeImageLike,
} from './render';
