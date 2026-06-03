// OMML AST shared by docx/xlsx/pptx parsers. Mirrors ECMA-376 §22.1.2.
// Each Rust parser emits JSON matching these types; the core math engine consumes them.
export type MathStyle = 'roman' | 'italic' | 'bold' | 'boldItalic';

export interface MathRun {
  kind: 'run';
  text: string;
  style: MathStyle;
}
export interface MathFraction {
  kind: 'fraction';
  num: MathNode[];
  den: MathNode[];
  /** false = no rule (e.g. binomial); defaults to true. */
  bar?: boolean;
}
export interface MathScript {
  kind: 'sup' | 'sub' | 'subSup';
  base: MathNode[];
  sup?: MathNode[];
  sub?: MathNode[];
}
export interface MathNary {
  kind: 'nary';
  /** operator char, e.g. '∑', '∫', '∏'. */
  op: string;
  sub?: MathNode[];
  sup?: MathNode[];
  body: MathNode[];
}
export interface MathDelimiter {
  kind: 'delimiter';
  /** opening char (default '('). */
  begChar: string;
  /** closing char (default ')'). */
  endChar: string;
  /** separated groups (e.g. for cases / multiple args). */
  items: MathNode[][];
}
export interface MathRadical {
  kind: 'radical';
  /** optional index (e.g. cube root); empty/omitted = square root. */
  index?: MathNode[];
  radicand: MathNode[];
}
/** Lower/upper limit (`m:limLow` / `m:limUpp`), e.g. lim under n→∞. */
export interface MathLimit {
  kind: 'limit';
  base: MathNode[];
  lower?: MathNode[];
  upper?: MathNode[];
}
/** Matrix (`m:m`) or aligned equation array (`m:eqArr`). rows → cells → nodes. */
export interface MathArray {
  kind: 'array';
  rows: MathNode[][][];
  /** 'eq' = alternating right/left (eqArr); 'center' = matrix; 'left'. */
  align: 'eq' | 'center' | 'left';
}
/** Group character (`m:groupChr`), e.g. under/over brace. */
export interface MathGroupChr {
  kind: 'groupChr';
  char: string;
  pos: 'top' | 'bot';
  base: MathNode[];
}
/** Over/under bar (`m:bar`). */
export interface MathBar {
  kind: 'bar';
  pos: 'top' | 'bot';
  base: MathNode[];
}
/** Accent (`m:acc`), e.g. hat, bar, vector arrow over the base. */
export interface MathAccent {
  kind: 'accent';
  char: string;
  base: MathNode[];
}
export interface MathFunc {
  kind: 'func';
  name: MathNode[];
  arg: MathNode[];
}
export interface MathGroup {
  kind: 'group';
  items: MathNode[];
}

export type MathNode =
  | MathRun
  | MathFraction
  | MathScript
  | MathNary
  | MathDelimiter
  | MathRadical
  | MathLimit
  | MathArray
  | MathGroupChr
  | MathBar
  | MathAccent
  | MathFunc
  | MathGroup;

const KINDS = new Set([
  'run',
  'fraction',
  'sup',
  'sub',
  'subSup',
  'nary',
  'delimiter',
  'radical',
  'limit',
  'array',
  'groupChr',
  'bar',
  'accent',
  'func',
  'group',
]);

export function isMathNode(v: unknown): v is MathNode {
  return (
    !!v &&
    typeof v === 'object' &&
    KINDS.has((v as { kind?: string }).kind ?? '')
  );
}

/** Top-level math container as emitted by parsers. `display` = block (`m:oMathPara`). */
export interface MathFormula {
  nodes: MathNode[];
  display: boolean;
}
