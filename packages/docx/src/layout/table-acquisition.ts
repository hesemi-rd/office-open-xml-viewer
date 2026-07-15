import type { BorderSpec, CellBorders, DocParagraph, DocTable, TableBorders } from '../types.js';
import {
  acquireTableCellBlocks,
  isStructuralTrailingParagraph,
} from './table-cell-blocks.js';
import type { ParagraphBorderEdges } from './paragraph-border-adjacency.js';
import { layoutTable } from './table.js';
import { tableCellHorizontalSpacingInsets } from './table-columns.js';
import type {
  LayoutServices,
  ParagraphLayout,
  TableBorderInput,
  TableEdgeInputs,
  TableFormatInput,
  TableLayout,
  TableLayoutInput,
} from './types.js';

export interface RetainedTableAcquisitionDependencies<State> {
  layoutServices(state: State): LayoutServices | undefined;
  tableFormat(table: Readonly<DocTable>): TableFormatInput;
  resolveColumns(table: DocTable, contentWidthPt: number, state: State): readonly number[];
  createCellState(state: State, contentWidthPt: number, cell: DocTable['rows'][number]['cells'][number]): State;
  acquireParagraph(
    state: State,
    paragraph: DocParagraph,
    contentWidthPt: number,
    sourcePath: readonly number[],
    flowDomainId: string,
    paragraphBorderEdges: ParagraphBorderEdges,
  ): ParagraphLayout;
  advanceState(state: State, advancePt: number): void;
}

function retainedBorder(border: BorderSpec | null): TableBorderInput | null {
  if (!border) return null;
  const authored = border.color ?? '000000';
  return Object.freeze({
    widthPt: border.width,
    color: authored.startsWith('#') ? authored : `#${authored}`,
    authoredStyle: border.style,
  });
}

function retainedEdges(edges: CellBorders | TableBorders): TableEdgeInputs {
  return Object.freeze({
    top: retainedBorder(edges.top),
    right: retainedBorder(edges.right),
    bottom: retainedBorder(edges.bottom),
    left: retainedBorder(edges.left),
    insideH: retainedBorder(edges.insideH),
    insideV: retainedBorder(edges.insideV),
  });
}

function physicalAlignment(
  value: string | null | undefined,
  bidiVisual: boolean,
): TableLayoutInput['alignment'] {
  if (value === 'center') return 'center';
  const trailing = value === 'right' || value === 'end';
  return (bidiVisual ? !trailing : trailing) ? 'right' : 'left';
}

/**
 * Acquire an ordinary or nested table from final-width retained children.
 * Parser-private authored-presence and lexical facts arrive only through the
 * immutable TableFormatInput; this fold owns recursive table geometry and never
 * reaches back into parser/model metadata.
 */
export function acquireRetainedTableLayout<State>(
  table: DocTable,
  columnWidthsPt: readonly number[],
  contentWidthPt: number,
  outerState: State,
  sourcePath: readonly number[],
  dependencies: RetainedTableAcquisitionDependencies<State>,
): TableLayout {
  const services = dependencies.layoutServices(outerState);
  if (!services) throw new Error('Retained table acquisition requires layout services');
  const flowDomainId = `table:${sourcePath.join('.')}`;
  const format = dependencies.tableFormat(table);
  const bidiVisual = table.bidiVisual === true;
  const firstRowException = format.firstRowException;
  const tableIndentPt = firstRowException?.indentAuthored
    ? (firstRowException.indentPt ?? 0)
    : (table.tblInd ?? 0);
  const rows: TableLayoutInput['rows'] = table.rows.map((row, rowIndex) => {
    const rowFormat = format.rows[rowIndex];
    let columnStart = Math.max(0, Math.min(columnWidthsPt.length, row.gridBefore ?? 0));
    const cells = row.cells.map((cell, cellIndex) => {
      const formatMargins = rowFormat?.cells[cellIndex]?.marginsPt ?? {
        top: cell.marginTop ?? table.cellMarginTop,
        right: cell.marginRight ?? table.cellMarginRight,
        bottom: cell.marginBottom ?? table.cellMarginBottom,
        left: cell.marginLeft ?? table.cellMarginLeft,
      };
      const currentColumnStart = columnStart;
      const columnSpan = Math.min(
        Math.max(1, cell.colSpan),
        Math.max(0, columnWidthsPt.length - currentColumnStart),
      );
      columnStart += columnSpan;
      const cellTotalWidthPt = columnWidthsPt
        .slice(currentColumnStart, currentColumnStart + columnSpan)
        .reduce((sum, width) => sum + width, 0);
      const spacingInsets = tableCellHorizontalSpacingInsets(
        rowFormat?.cellSpacingPt ?? 0,
        currentColumnStart,
        columnSpan,
        columnWidthsPt.length,
      );
      const cellPath = [...sourcePath, rowIndex, cellIndex];
      const acquired = cell.vMerge === false
        ? []
        : acquireTableCellBlocks({
            cell,
            table,
            cellTotalWidthPt,
            outerState,
            sourcePath: cellPath,
          }, {
            resolveContentWidthPt: (_cell, _table, totalWidthPt) => Math.max(
              0,
              totalWidthPt
                - spacingInsets.startPt
                - spacingInsets.endPt
                - formatMargins.left
                - formatMargins.right,
            ),
            createCellState: dependencies.createCellState,
            acquireParagraph: (
              cellState,
              paragraph,
              paragraphWidthPt,
              paragraphPath,
              paragraphBorderEdges,
            ) => dependencies.acquireParagraph(
              cellState,
              paragraph,
              paragraphWidthPt,
              paragraphPath,
              `${flowDomainId}:cell:${rowIndex}.${cellIndex}`,
              paragraphBorderEdges,
            ),
            acquireNestedTable: (cellState, nestedTable, nestedContentWidthPt, nestedPath) => {
              const nestedColumns = dependencies.resolveColumns(
                nestedTable,
                nestedContentWidthPt,
                cellState,
              );
              return acquireRetainedTableLayout(
                nestedTable,
                nestedColumns,
                nestedContentWidthPt,
                cellState,
                nestedPath,
                dependencies,
              );
            },
            advanceState: dependencies.advanceState,
          });
      return {
        id: `${flowDomainId}:cell:${rowIndex}.${cellIndex}`,
        source: { story: 'body' as const, storyInstance: 'body', path: cellPath },
        columnStart: currentColumnStart,
        columnSpan,
        verticalMerge: cell.vMerge === true
          ? 'restart' as const
          : cell.vMerge === false ? 'continue' as const : 'none' as const,
        margins: {
          topPt: formatMargins.top,
          rightPt: formatMargins.right,
          bottomPt: formatMargins.bottom,
          leftPt: formatMargins.left,
        },
        vAlign: cell.vAlign,
        ...(cell.background ? {
          background: {
            color: cell.background.startsWith('#') ? cell.background : `#${cell.background}`,
          },
        } : {}),
        borders: retainedEdges(cell.borders),
        blocks: acquired.map((layout, blockIndex) => {
          if (layout.kind === 'table' && !('flowBounds' in layout)) {
            throw new Error('Ordinary table acquisition cannot retain a page-slice table fragment');
          }
          return {
            layout: layout as ParagraphLayout | TableLayout,
            ...(isStructuralTrailingParagraph(cell.content, blockIndex)
              ? { structuralTrailing: true }
              : {}),
          };
        }),
      };
    });
    const heightRule = rowFormat?.height?.rule ?? 'auto';
    return {
      id: `${flowDomainId}:row:${rowIndex}`,
      source: { story: 'body' as const, storyInstance: 'body', path: [...sourcePath, rowIndex] },
      heightPt: rowFormat?.height?.valuePt ?? null,
      heightRule,
      cellSpacingPt: rowFormat?.cellSpacingPt ?? 0,
      exceptionBorders: rowFormat?.exception?.borders
        ? retainedEdges(rowFormat.exception.borders)
        : null,
      alignment: physicalAlignment(rowFormat?.justification ?? table.jc, bidiVisual),
      indentPt: tableIndentPt,
      cells,
      ...(row.isHeader ? { repeatedHeader: true } : {}),
    };
  });
  const input: TableLayoutInput = {
    kind: 'table',
    id: flowDomainId,
    source: { story: 'body', storyInstance: 'body', path: [...sourcePath] },
    flowDomainId,
    ordinaryFlow: true,
    alignment: physicalAlignment(table.jc, bidiVisual),
    indentPt: tableIndentPt,
    bidiVisual,
    columnWidthsPt,
    borders: retainedEdges(table.borders),
    rows,
  };
  const bounds = {
    xPt: 0,
    yPt: 0,
    widthPt: contentWidthPt,
    heightPt: 1,
  };
  return layoutTable(input, {
    container: { id: flowDomainId, kind: 'tableCell', bounds },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: bounds,
  }, services).layout;
}
