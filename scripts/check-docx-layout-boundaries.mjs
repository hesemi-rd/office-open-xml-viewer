#!/usr/bin/env node

import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { basename, dirname, join, relative, resolve, sep } from 'node:path';
import { spawnSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const require = createRequire(new URL('../packages/docx/package.json', import.meta.url));
const ts = require('typescript');

const BASELINE_PATH = 'scripts/docx-layout-boundary-baseline.json';
const DOCX_SOURCE = 'packages/docx/src';
const PAINT_SOURCE = `${DOCX_SOURCE}/paint`;
const LAYOUT_SOURCE = `${DOCX_SOURCE}/layout`;

const FINAL_RENDERER_EXPORTS = new Set([
  'DocxTextRunInfo',
  'RenderDocumentOptions',
  'clearResolvedLocalFonts',
  'documentHasMath',
  'dropColorReplacedCache',
  'paginateDocument',
  'physicalPageSizeForPage',
  'prepareMathRuns',
  'renderDocumentToCanvas',
  'setResolvedLocalFonts',
]);

const FINAL_RENDERER_DECLARATIONS = new Set([
  ...FINAL_RENDERER_EXPORTS,
  'createLayoutServices',
  'normalizeRenderOptions',
]);

const PLANNED_NON_LAYOUT_MODULES = new Set([
  `${DOCX_SOURCE}/parser-model.ts`,
]);

const SHARED_PAINT_IMPORTS = new Map([
  ['@silurus/ooxml-core', new Set(['renderChart'])],
]);

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
  return /\.tsx?$/.test(path)
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
  const kind = path.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS;
  return ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true, kind);
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

function moduleEdges(path) {
  const source = sourceFile(path);
  const edges = [];
  for (const statement of source.statements) {
    if (ts.isImportDeclaration(statement)
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      const bindings = statement.importClause?.namedBindings;
      const importedNames = statement.importClause
        ? [
            ...(statement.importClause.name ? ['default'] : []),
            ...(bindings && ts.isNamespaceImport(bindings) ? ['*'] : []),
            ...(bindings && ts.isNamedImports(bindings)
              ? bindings.elements.map((element) => element.propertyName?.text ?? element.name.text)
              : []),
          ]
        : [];
      edges.push({
        specifier: statement.moduleSpecifier.text,
        typeOnly: importIsTypeOnly(statement),
        literal: true,
        importedNames,
        bare: !statement.importClause,
      });
    }
    if (ts.isExportDeclaration(statement)
      && statement.moduleSpecifier
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      edges.push({
        specifier: statement.moduleSpecifier.text,
        typeOnly: exportIsTypeOnly(statement),
        literal: true,
        importedNames: statement.exportClause && ts.isNamedExports(statement.exportClause)
          ? statement.exportClause.elements.map((element) => element.propertyName?.text ?? element.name.text)
          : ['*'],
      });
    }
  }
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { specifier: argument.text, typeOnly: false, literal: true }
        : { specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'require') {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { specifier: argument.text, typeOnly: false, literal: true }
        : { specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    ts.forEachChild(node, visit);
  };
  ts.forEachChild(source, visit);
  return edges;
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
    graph.set(file, moduleEdges(file));
  }
  return graph;
}

function paintBoundaryViolations(root) {
  const graph = dependencyGraph(root);
  const paintRoot = resolve(root, PAINT_SOURCE);
  const layoutTypes = resolve(root, LAYOUT_SOURCE, 'types.ts');
  const pageGraph = resolve(root, LAYOUT_SOURCE, 'page-graph.ts');
  const entries = [...graph.keys()].filter((path) => path.startsWith(`${paintRoot}${sep}`));
  const violations = [];
  const nonLiteral = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      for (const edge of graph.get(current.path) ?? []) {
        if (!edge.literal) {
          nonLiteral.push(posixPath(relative(root, current.path)));
          continue;
        }
        if (edge.bare) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        if (!edge.specifier.startsWith('.')) {
          const allowedNames = SHARED_PAINT_IMPORTS.get(edge.specifier);
          const allowed = !edge.typeOnly
            && allowedNames
            && edge.importedNames?.length > 0
            && edge.importedNames?.every((name) => allowedNames.has(name));
          if (!allowed) {
            violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          }
          continue;
        }
        const dependency = resolveLocalImport(current.path, edge.specifier);
        if (!dependency) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        const chain = [...current.chain, dependency];
        const insidePaint = dependency.startsWith(`${paintRoot}${sep}`);
        const allowedContract = (edge.typeOnly && dependency === layoutTypes) || dependency === pageGraph;
        if (!insidePaint && !allowedContract) {
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (insidePaint && !visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return { violations, nonLiteral };
}

function identifierText(node) {
  if (ts.isIdentifier(node) || ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return node.text;
  return null;
}

function assertCapabilityBoundaries(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const rel = posixPath(relative(root, path));
    const inPaint = rel.startsWith(`${PAINT_SOURCE}/`);
    const inLayout = rel.startsWith(`${LAYOUT_SOURCE}/`);
    if (!inPaint && !inLayout) continue;
    const source = sourceFile(path);
    const visit = (node) => {
      const text = identifierText(node);
      if (inPaint && text === 'measureText') fail('PAINT_CAPABILITY', `${rel} uses measureText`);
      if (inPaint && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('PAINT_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('LAYOUT_STYLE_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:dpr|displayScale|devicePixelRatio|CanvasRenderingContext2D|OffscreenCanvasRenderingContext2D)$/.test(text)) {
        fail('LAYOUT_DISPLAY_CAPABILITY', `${rel} uses ${text}`);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
}

const MIGRATION_IDENTIFIER = /(?:legacy|(?:use|enable|prefer|require)[a-z0-9]*(?:old|previous|alternate)[a-z0-9]*(?:engine|layout|path|algorithm)|(?:reuse|paint)enabled|requireslegacy|dryrun)/i;

function matchingIdentifierCounts(root, predicate) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    const visit = (node) => {
      if (ts.isIdentifier(node) && predicate(node.text)) {
        const key = `${file}#${node.text}`;
        counts[key] = (counts[key] ?? 0) + 1;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
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

function bindingNames(name, names = []) {
  if (ts.isIdentifier(name)) names.push(name.text);
  else for (const element of name.elements) {
    if (!ts.isOmittedExpression(element)) bindingNames(element.name, names);
  }
  return names;
}

function declarationNames(statement) {
  if ((ts.isFunctionDeclaration(statement)
      || ts.isClassDeclaration(statement)
      || ts.isInterfaceDeclaration(statement)
      || ts.isTypeAliasDeclaration(statement)
      || ts.isEnumDeclaration(statement)
      || ts.isModuleDeclaration(statement))
    && statement.name) {
    return [statement.name.text];
  }
  if (ts.isVariableStatement(statement)) {
    return statement.declarationList.declarations.flatMap((declaration) => bindingNames(declaration.name));
  }
  return [];
}

function declarationKind(statement) {
  if (ts.isVariableStatement(statement)) return 'variable';
  return ts.SyntaxKind[statement.kind];
}

function normalizedNodeHash(node, source) {
  const normalized = node.getText(source).replace(/\s+/g, ' ').trim();
  return createHash('sha256').update(normalized).digest('hex');
}

function declarationInventory(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const nonLayoutDeclarationKeys = [];
  const legacyDeclarationHashes = {};
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const file = posixPath(relative(root, path));
    const migrationOwner = file.startsWith(`${LAYOUT_SOURCE}/`)
      || file.startsWith(`${PAINT_SOURCE}/`)
      || file.startsWith(`${DOCX_SOURCE}/conformance/`)
      || PLANNED_NON_LAYOUT_MODULES.has(file);
    const source = sourceFile(path);
    for (const statement of source.statements) {
      for (const name of declarationNames(statement)) {
        const key = `${file}#${declarationKind(statement)}#${name}`;
        if (!migrationOwner) nonLayoutDeclarationKeys.push(key);
        if (LEGACY_SYMBOLS.includes(name)) {
          legacyDeclarationHashes[key] = normalizedNodeHash(statement, source);
        }
      }
    }
  }
  return {
    nonLayoutDeclarationKeys: [...new Set(nonLayoutDeclarationKeys)].sort(),
    legacyDeclarationHashes: Object.fromEntries(
      Object.entries(legacyDeclarationHashes).sort(([left], [right]) => left.localeCompare(right)),
    ),
  };
}

function rendererImportEdges(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return [];
  return [...new Set(moduleEdges(renderer)
    .filter((edge) => edge.literal)
    .map((edge) => resolveLocalImport(renderer, edge.specifier))
    .filter((path) => path && LEGACY_RENDERER_IMPORTS.has(basename(path)))
    .map((path) => `${DOCX_SOURCE}/renderer.ts -> ${posixPath(relative(root, path))}`))]
    .sort();
}

function currentAllowances(root) {
  const declarations = declarationInventory(root);
  return {
    version: 2,
    legacySymbolCounts: identifierCounts(root),
    migrationIdentifierCounts: matchingIdentifierCounts(root, (name) => MIGRATION_IDENTIFIER.test(name)),
    nonLayoutDeclarationKeys: declarations.nonLayoutDeclarationKeys,
    legacyDeclarationHashes: declarations.legacyDeclarationHashes,
    rendererImportEdges: rendererImportEdges(root),
  };
}

function readBaseline(path) {
  const value = JSON.parse(readFileSync(path, 'utf8'));
  if (value.version !== 2
    || typeof value.legacySymbolCounts !== 'object'
    || typeof value.migrationIdentifierCounts !== 'object'
    || !Array.isArray(value.nonLayoutDeclarationKeys)
    || typeof value.legacyDeclarationHashes !== 'object'
    || !Array.isArray(value.rendererImportEdges)) {
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
  if (value.version !== 2) fail('INVALID_BASELINE', `${mergeBase}:${BASELINE_PATH}`);
  return value;
}

function assertNoExpansion(head, base) {
  for (const [symbol, count] of Object.entries(head.legacySymbolCounts)) {
    const baseCount = base.legacySymbolCounts[symbol] ?? 0;
    if (count > baseCount) fail('BASELINE_EXPANSION', `${symbol}: ${count} > ${baseCount}`);
  }
  for (const [identifier, count] of Object.entries(head.migrationIdentifierCounts)) {
    const baseCount = base.migrationIdentifierCounts[identifier] ?? 0;
    if (count > baseCount) fail('BASELINE_EXPANSION', `${identifier}: ${count} > ${baseCount}`);
  }
  const baseDeclarations = new Set(base.nonLayoutDeclarationKeys);
  for (const declaration of head.nonLayoutDeclarationKeys) {
    if (!baseDeclarations.has(declaration)) fail('BASELINE_EXPANSION', declaration);
  }
  for (const [declaration, hash] of Object.entries(head.legacyDeclarationHashes)) {
    const baseHash = base.legacyDeclarationHashes[declaration];
    if (!baseHash) fail('BASELINE_EXPANSION', declaration);
    if (hash !== baseHash) fail('LEGACY_DECLARATION_CHANGED', declaration);
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

function assertPaintBoundaries(root) {
  const { violations, nonLiteral } = paintBoundaryViolations(root);
  if (nonLiteral.length > 0) {
    fail('NON_LITERAL_MODULE_EDGE', nonLiteral.join('\n'));
  }
  if (violations.length > 0) {
    fail('FORBIDDEN_PAINT_EDGE', violations.map((chain) => chain.join(' -> ')).join('\n'));
  }
}

function hasExportModifier(statement) {
  return statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword) ?? false;
}

function rendererExports(path) {
  const names = [];
  const source = sourceFile(path);
  for (const statement of source.statements) {
    if (hasExportModifier(statement)) {
      const declared = declarationNames(statement);
      names.push(...(declared.length > 0 ? declared : ['default']));
    }
    if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
      names.push(...statement.exportClause.elements.map((element) => element.name.text));
    }
  }
  return [...new Set(names)].sort();
}

function rendererImportBindings(source) {
  const bindings = new Set();
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause) continue;
    if (statement.importClause.name) bindings.add(statement.importClause.name.text);
    const named = statement.importClause.namedBindings;
    if (named && ts.isNamespaceImport(named)) bindings.add(named.name.text);
    if (named && ts.isNamedImports(named)) {
      for (const element of named.elements) bindings.add(element.name.text);
    }
  }
  return bindings;
}

function unwrapAdapterExpression(expression) {
  if (ts.isAwaitExpression(expression)
    || ts.isParenthesizedExpression(expression)
    || ts.isAsExpression(expression)
    || ts.isSatisfiesExpression(expression)) {
    return unwrapAdapterExpression(expression.expression);
  }
  return expression;
}

function adapterValueIsAllowed(expression, callable) {
  const value = unwrapAdapterExpression(expression);
  if (ts.isIdentifier(value)
    || ts.isStringLiteralLike(value)
    || ts.isNumericLiteral(value)
    || value.kind === ts.SyntaxKind.TrueKeyword
    || value.kind === ts.SyntaxKind.FalseKeyword
    || value.kind === ts.SyntaxKind.NullKeyword) return true;
  if (ts.isPropertyAccessExpression(value)) return adapterValueIsAllowed(value.expression, callable);
  if (ts.isElementAccessExpression(value)) {
    return adapterValueIsAllowed(value.expression, callable)
      && (!value.argumentExpression || adapterValueIsAllowed(value.argumentExpression, callable));
  }
  if (ts.isArrayLiteralExpression(value)) {
    return value.elements.every((element) => !ts.isSpreadElement(element)
      ? adapterValueIsAllowed(element, callable)
      : adapterValueIsAllowed(element.expression, callable));
  }
  if (ts.isObjectLiteralExpression(value)) {
    return value.properties.every((property) => {
      if (ts.isPropertyAssignment(property)) return adapterValueIsAllowed(property.initializer, callable);
      if (ts.isShorthandPropertyAssignment(property)) return true;
      if (ts.isSpreadAssignment(property)) return adapterValueIsAllowed(property.expression, callable);
      return false;
    });
  }
  if (ts.isCallExpression(value)) {
    const callee = unwrapAdapterExpression(value.expression);
    const callableName = ts.isIdentifier(callee)
      ? callee.text
      : ts.isPropertyAccessExpression(callee) && ts.isIdentifier(callee.expression)
        ? callee.expression.text
        : null;
    return callableName != null
      && callable.has(callableName)
      && value.arguments.every((argument) => adapterValueIsAllowed(argument, callable));
  }
  return false;
}

function adapterBodyIsAllowed(body, callable) {
  return body.statements.every((statement) => {
    if (ts.isVariableStatement(statement)) {
      return statement.declarationList.declarations.every((declaration) => (
        ts.isIdentifier(declaration.name)
        && (!declaration.initializer || adapterValueIsAllowed(declaration.initializer, callable))
      ));
    }
    if (ts.isExpressionStatement(statement)) return adapterValueIsAllowed(statement.expression, callable);
    if (ts.isReturnStatement(statement)) {
      return !statement.expression || adapterValueIsAllowed(statement.expression, callable);
    }
    return false;
  });
}

function assertFinalRendererAdapter(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) fail('FINAL_ADAPTER_MISSING', `${DOCX_SOURCE}/renderer.ts`);
  const source = sourceFile(renderer);
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement)
      && (!statement.exportClause || !ts.isNamedExports(statement.exportClause))) {
      fail('FINAL_ADAPTER_EXPORT', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      if (!FINAL_RENDERER_DECLARATIONS.has(name)) fail('FINAL_ADAPTER_DECLARATION', name);
    }
  }
  for (const name of rendererExports(renderer)) {
    if (!FINAL_RENDERER_EXPORTS.has(name)) fail('FINAL_ADAPTER_EXPORT', name);
  }
  const callable = rendererImportBindings(source);
  callable.add('createLayoutServices');
  callable.add('normalizeRenderOptions');
  for (const statement of source.statements) {
    if (ts.isFunctionDeclaration(statement)
      && statement.name
      && (statement.name.text === 'paginateDocument' || statement.name.text === 'renderDocumentToCanvas')
      && statement.body
      && !adapterBodyIsAllowed(statement.body, callable)) {
      fail('FINAL_ADAPTER_BODY', statement.name.text);
    }
  }
  for (const edge of moduleEdges(renderer)) {
    if (!edge.literal) fail('FINAL_ADAPTER_IMPORT', '<dynamic>');
    if (edge.bare) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    if (!edge.specifier.startsWith('.')) continue;
    const target = resolveLocalImport(renderer, edge.specifier);
    if (!target) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    const rel = posixPath(relative(root, target));
    const allowed = rel.startsWith(`${LAYOUT_SOURCE}/`)
      || rel.startsWith(`${PAINT_SOURCE}/`)
      || (edge.typeOnly && rel === `${DOCX_SOURCE}/types.ts`);
    if (!allowed) fail('FINAL_ADAPTER_IMPORT', `${DOCX_SOURCE}/renderer.ts -> ${rel}`);
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
  assertPaintBoundaries(root);
  assertCapabilityBoundaries(root);

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
    assertFinalRendererAdapter(root);
    const actual = currentAllowances(root);
    if (Object.keys(actual.legacySymbolCounts).length > 0
      || Object.keys(actual.migrationIdentifierCounts).length > 0
      || Object.keys(actual.legacyDeclarationHashes).length > 0
      || actual.rendererImportEdges.length > 0) {
      fail('FINAL_LEGACY_BOUNDARY', stableJson(actual).trim());
    }
    return;
  }

  const baseBaseline = mergeBaseBaseline(root, options.baseRef);
  const headBaseline = readBaseline(baselinePath);
  if (baseBaseline) assertNoExpansion(headBaseline, baseBaseline);
  const actual = currentAllowances(root);
  if (baseBaseline) assertNoExpansion(actual, baseBaseline);
  assertExactBaseline(headBaseline, actual);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    checkDocxLayoutBoundaries(parseArguments(process.argv.slice(2)));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
