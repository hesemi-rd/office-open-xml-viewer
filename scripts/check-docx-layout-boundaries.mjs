#!/usr/bin/env node

import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { basename, dirname, join, relative, resolve, sep } from 'node:path';
import { spawnSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const require = createRequire(new URL('../packages/docx/package.json', import.meta.url));
const ts = require('typescript');

const BASELINE_PATH = 'scripts/docx-layout-boundary-baseline.json';
const DOCX_SOURCE = 'packages/docx/src';
const PAINT_SOURCE = `${DOCX_SOURCE}/paint`;

const LEGACY_SYMBOLS = [
  'fitMeasureReuseEnabled',
  'fragmentPaintEnabled',
  'lineReuseEnabled',
  'isFragmentPaintableParagraph',
  'layoutLinesInputs',
  'stampParagraphLines',
  'renderBodyParagraphLines',
  'renderShapeText',
  'tableRequiresLegacyPaint',
  'isFragmentPaintableTable',
  'tableReuseEnabled',
  'renderTableFragment',
  'computePages',
  'computeTableLayout',
  'calculateRowHeight',
  'measureCellContentHeightPx',
  'buildTableCellBlocks',
  'renderHeaderFooter',
  'measureFootnoteHeight',
  'deferFront',
  'sectionBreakSpacer',
  'collapsedSpacer',
  'leadsCollapsedRun',
  'hiddenCollapsed',
  'tableColWidthsPt',
  'tableRowHeightsPt',
  'tableLayoutInputs',
];

const LEGACY_RENDERER_IMPORTS = new Set([
  'fragment-paint.ts',
  'layout-context.ts',
  'layout-fragments.ts',
  'line-layout.ts',
  'paragraph-measure.ts',
  'table-fragments.ts',
  'table-geometry.ts',
]);

function fail(code, detail) {
  throw new Error(`${code}: ${detail}`);
}

function posixPath(path) {
  return path.split(sep).join('/');
}

function isProductionTypeScript(path) {
  return path.endsWith('.ts')
    && !path.endsWith('.d.ts')
    && !/\.(test|spec|stories)\.tsx?$/.test(path)
    && !path.includes('/wasm/');
}

function listFiles(root) {
  if (!existsSync(root)) return [];
  const files = [];
  for (const entry of readdirSync(root)) {
    const path = join(root, entry);
    if (statSync(path).isDirectory()) files.push(...listFiles(path));
    else files.push(path);
  }
  return files;
}

function sourceFile(path) {
  return ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
}

function importIsTypeOnly(statement) {
  const clause = statement.importClause;
  if (!clause) return false;
  if (clause.isTypeOnly) return true;
  if (clause.name || !clause.namedBindings || !ts.isNamedImports(clause.namedBindings)) return false;
  return clause.namedBindings.elements.length > 0
    && clause.namedBindings.elements.every((element) => element.isTypeOnly);
}

function exportIsTypeOnly(statement) {
  if (statement.isTypeOnly) return true;
  return statement.exportClause
    && ts.isNamedExports(statement.exportClause)
    && statement.exportClause.elements.length > 0
    && statement.exportClause.elements.every((element) => element.isTypeOnly);
}

function moduleSpecifiers(path) {
  const source = sourceFile(path);
  const specifiers = [];
  for (const statement of source.statements) {
    if (ts.isImportDeclaration(statement)
      && ts.isStringLiteral(statement.moduleSpecifier)
      && !importIsTypeOnly(statement)) {
      specifiers.push(statement.moduleSpecifier.text);
    }
    if (ts.isExportDeclaration(statement)
      && statement.moduleSpecifier
      && ts.isStringLiteral(statement.moduleSpecifier)
      && !exportIsTypeOnly(statement)) {
      specifiers.push(statement.moduleSpecifier.text);
    }
  }
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && node.expression.kind === ts.SyntaxKind.ImportKeyword
      && node.arguments.length === 1
      && ts.isStringLiteral(node.arguments[0])) {
      specifiers.push(node.arguments[0].text);
    }
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'require'
      && node.arguments.length === 1
      && ts.isStringLiteral(node.arguments[0])) {
      specifiers.push(node.arguments[0].text);
    }
    ts.forEachChild(node, visit);
  };
  ts.forEachChild(source, visit);
  return specifiers;
}

function resolveLocalImport(importer, specifier) {
  if (!specifier.startsWith('.')) return null;
  const clean = specifier.split('?')[0].split('#')[0];
  const unresolved = resolve(dirname(importer), clean);
  const withoutJs = unresolved.replace(/\.(mjs|cjs|js)$/, '');
  const candidates = [
    unresolved,
    `${withoutJs}.ts`,
    `${withoutJs}.tsx`,
    join(unresolved, 'index.ts'),
  ];
  return candidates.find((candidate) => existsSync(candidate) && statSync(candidate).isFile()) ?? null;
}

function dependencyGraph(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const files = listFiles(sourceRoot).filter(isProductionTypeScript);
  const graph = new Map();
  for (const file of files) {
    const dependencies = moduleSpecifiers(file)
      .map((specifier) => resolveLocalImport(file, specifier))
      .filter((candidate) => candidate && candidate.startsWith(sourceRoot));
    graph.set(file, [...new Set(dependencies)]);
  }
  return graph;
}

function forbiddenPaintDependency(root, path) {
  const rel = posixPath(relative(root, path));
  if (rel === `${DOCX_SOURCE}/types.ts`) return true;
  if (LEGACY_RENDERER_IMPORTS.has(basename(path))) return true;
  return rel.startsWith(`${DOCX_SOURCE}/layout/`) && basename(path) !== 'types.ts';
}

function transitivePaintViolations(root) {
  const graph = dependencyGraph(root);
  const paintRoot = resolve(root, PAINT_SOURCE);
  const entries = [...graph.keys()].filter((path) => path.startsWith(`${paintRoot}${sep}`));
  const violations = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      for (const dependency of graph.get(current.path) ?? []) {
        const chain = [...current.chain, dependency];
        if (forbiddenPaintDependency(root, dependency)) {
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (!visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return violations;
}

function identifierCounts(root) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const fileCounts = Object.fromEntries(LEGACY_SYMBOLS.map((symbol) => [symbol, 0]));
    const source = sourceFile(path);
    const visit = (node) => {
      if (ts.isIdentifier(node) && Object.hasOwn(fileCounts, node.text)) fileCounts[node.text] += 1;
      ts.forEachChild(node, visit);
    };
    visit(source);
    const file = posixPath(relative(root, path));
    for (const [symbol, count] of Object.entries(fileCounts)) {
      if (count > 0) counts[`${file}#${symbol}`] = count;
    }
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function rendererImportEdges(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return [];
  return [...new Set(moduleSpecifiers(renderer)
    .map((specifier) => resolveLocalImport(renderer, specifier))
    .filter((path) => path && LEGACY_RENDERER_IMPORTS.has(basename(path)))
    .map((path) => `${DOCX_SOURCE}/renderer.ts -> ${posixPath(relative(root, path))}`))]
    .sort();
}

function currentAllowances(root) {
  return {
    version: 1,
    legacySymbolCounts: identifierCounts(root),
    rendererImportEdges: rendererImportEdges(root),
  };
}

function readBaseline(path) {
  const value = JSON.parse(readFileSync(path, 'utf8'));
  if (value.version !== 1 || typeof value.legacySymbolCounts !== 'object' || !Array.isArray(value.rendererImportEdges)) {
    fail('INVALID_BASELINE', path);
  }
  return value;
}

function git(root, args, allowFailure = false) {
  const result = spawnSync('git', args, { cwd: root, encoding: 'utf8' });
  if (result.status !== 0 && !allowFailure) fail('GIT_ERROR', `${args.join(' ')}: ${result.stderr.trim()}`);
  return result;
}

function mergeBaseBaseline(root, baseRef) {
  const mergeBase = git(root, ['merge-base', baseRef, 'HEAD']).stdout.trim();
  const shown = git(root, ['show', `${mergeBase}:${BASELINE_PATH}`], true);
  if (shown.status !== 0) return null;
  const value = JSON.parse(shown.stdout);
  if (value.version !== 1) fail('INVALID_BASELINE', `${mergeBase}:${BASELINE_PATH}`);
  return value;
}

function assertNoExpansion(head, base) {
  for (const [symbol, count] of Object.entries(head.legacySymbolCounts)) {
    const baseCount = base.legacySymbolCounts[symbol] ?? 0;
    if (count > baseCount) fail('BASELINE_EXPANSION', `${symbol}: ${count} > ${baseCount}`);
  }
  const baseEdges = new Set(base.rendererImportEdges);
  for (const edge of head.rendererImportEdges) {
    if (!baseEdges.has(edge)) fail('BASELINE_EXPANSION', edge);
  }
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function assertExactBaseline(baseline, actual) {
  if (stableJson(baseline) !== stableJson(actual)) {
    fail('BASELINE_MISMATCH', 'baseline must exactly describe current legacy symbols and renderer import edges');
  }
}

function assertNoTransitivePaintViolation(root) {
  const violations = transitivePaintViolations(root);
  if (violations.length > 0) {
    fail('TRANSITIVE_FORBIDDEN_EDGE', violations.map((chain) => chain.join(' -> ')).join('\n'));
  }
}

function parseArguments(argv) {
  const options = {
    root: process.cwd(),
    baseRef: 'origin/main',
    write: false,
    final: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--root') options.root = resolve(argv[++index]);
    else if (arg === '--base-ref') options.baseRef = argv[++index];
    else if (arg === '--write-transitional-baseline') options.write = true;
    else if (arg === '--final') options.final = true;
    else fail('UNKNOWN_ARGUMENT', arg);
  }
  return options;
}

export function checkDocxLayoutBoundaries(options) {
  const root = resolve(options.root);
  const baselinePath = resolve(root, BASELINE_PATH);
  const baselineExists = existsSync(baselinePath);
  assertNoTransitivePaintViolation(root);

  if (options.write) {
    const baseBaseline = mergeBaseBaseline(root, options.baseRef);
    if (baseBaseline) fail('TRANSITIONAL_BASELINE_EXISTS', `${options.baseRef} already contains ${BASELINE_PATH}`);
    const actual = currentAllowances(root);
    mkdirSync(dirname(baselinePath), { recursive: true });
    writeFileSync(baselinePath, stableJson(actual));
    return;
  }

  if (options.final || !baselineExists) {
    if (options.final && baselineExists) fail('FINAL_BASELINE_PRESENT', BASELINE_PATH);
    const actual = currentAllowances(root);
    if (Object.keys(actual.legacySymbolCounts).length > 0 || actual.rendererImportEdges.length > 0) {
      fail('FINAL_LEGACY_BOUNDARY', stableJson(actual).trim());
    }
    return;
  }

  const baseBaseline = mergeBaseBaseline(root, options.baseRef);
  const headBaseline = readBaseline(baselinePath);
  if (baseBaseline) assertNoExpansion(headBaseline, baseBaseline);
  assertExactBaseline(headBaseline, currentAllowances(root));
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    checkDocxLayoutBoundaries(parseArguments(process.argv.slice(2)));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
