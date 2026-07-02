# プロジェクト総合レビューと改善計画（2026-07）

対象: monorepo 全体（core / docx / pptx / xlsx の TS レンダラー、3 つの Rust/WASM パーサー、ooxml-common、ビルド・CI・周辺パッケージ）。
5 領域を並列レビューし、主要な指摘はコード上で裏取り済み。

## 総評

コードベースの健全性は全体として高い。spec 引用コメントの規律、`as any` の少なさ、lazy image パイプライン、zip-bomb ガード、WorkerBridge・SVG キャッシュ等の共有設計は良く維持されている。一方で、負債は明確に **4 つの構造的パターン** に集中している:

1. **God file 化** — docx `renderer.ts` 8,875 行、pptx parser `lib.rs` 13,645 行、docx parser `parser.rs` 10,255 行、pptx/xlsx `renderer.ts` 各 ~4,400 行、core `chart/renderer.ts` 2,437 行。並列 worktree セッションの衝突源であり、テストの `__test_*` バックドア輸出の根本原因。
2. **同一概念の 3 重実装** — CLAUDE.md の横断統合原則に反する重複が Rust 側（theme / rels / zip helpers / fill・color / txBody パース）と TS 側（bidi-line ×3、google-fonts ×3、画像デコードゲーティング、dash テーブル ×2）に残存。
3. **WASM 境界とホットパスの無駄なコピー・再計算** — 全文書 JSON 文字列の 3〜4 重シリアライズ、xlsx のシート切替毎の全ファイルコピー、docx の毎ページ全画像再デコード、pptx マスター/レイアウト XML の 10 回超 DOM 再パース。
4. **CI が描画を一切検証していない** — VRT はローカル専用のまま、既存の smoke suite（`tests/smoke/layouts.spec.ts`）は CI に未接続。大規模リファクタを安全に進める網が無い。

---

## 領域別の主要指摘（裏取り済み）

### A. core（TS）

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| A1 | `chart/renderer.ts` の 7 ファミリーがレイアウト（タイトル帯・凡例・軸ガター・スケール）を各自再実装し、定数がドリフト済み（bar `h*0.02` vs line `h*0.045` 等）。measured-gutter 化も bar 系のみ | `core/src/chart/renderer.ts` | L |
| A2 | プリセット図形エンジンが 2 系統併存（手書き 2,147 行 switch の `preset.ts` と spec 駆動の `preset-geometry/`）。両レンダラーが両方 import | `core/src/shape/preset.ts` / `preset-geometry/` | L |
| A3 | `applyExtrusion` / `applyBevelShading` が全キャンバス CPU ピクセルループ（O(w·h·steps)、メインスレッド同期） | `core/src/shape/bevel-shading.ts:869-1009` | M |
| A4 | effect ヘルパーが shape 毎・render 毎にフルデバイスサイズの aux canvas を確保（softEdge は 3 枚） | `core/src/shape/effects.ts` | M |
| A5 | preset-geometry evaluator が図形毎・render 毎に formula 文字列を再パース | `core/src/shape/preset-geometry/evaluator.ts:96-101` | M |
| A6 | pptx dash テーブルが core 内に 2 つあり、キー語彙が不一致（`lgDash` が片方で solid になる） | `core/src/shape/paint.ts:121` vs `core/src/draw/dash.ts:130` | S |
| A7 | チャートラベルを文字数 `slice(0,N)` で切断（未計測・省略記号なし・CJK 非対応） | `core/src/chart/renderer.ts:93` 他 | S |
| A8 | 画像 sniff のために Blob 全体をコピー（4 バイトで足りる）。raster/metafile 用ビットマップキャッシュが core に無い（SVG のみ存在） | `core/src/image/wmf.ts:711-727` | S+M |
| A9 | WMF/EMF/DIB が OffscreenCanvas を無条件要求（他は fallback あり） | `core/src/image/{wmf,emf,dib}.ts` | S |
| A10 | `WorkerBridge.request` に timeout / AbortSignal が無く、wedged worker で promise が永久ハング | `core/src/worker/bridge.ts:85-91` | S–M |
| A11 | barrel がテスト専用の bevel 内部関数等を全パッケージに輸出 | `core/src/index.ts:140-159` | S |

### B. docx（TS）

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| B1 | `renderer.ts` 8,875 行。`computePages` 単体 ~950 行、`layoutLines` ~595 行、`drawParagraphLine` ~578 行。自然な分割線（images / paginate / line-layout / paint / tables / shape-text / fonts / notes）は既に存在 | `docx/src/renderer.ts` | L |
| B2 | **measure と paint が手動同期の二重実装**（「Mirror renderParagraph's vertical advancement EXACTLY」コメント、テーブルも pt 系と px 系の 2 コピーで bug #523 は両方に当てた実績）。最大の設計リスク | `renderer.ts:2284-2648` vs `3936-4232`, `2649` vs `7423` | L |
| B3 | 全画像がページ描画毎に再デコード（createImageBitmap + WMF ラスタライズ + colorReplacement の getImageData を全ページ分・毎回） | `renderer.ts:575-619`, `document.ts:48` | S/M |
| B4 | 同一ページ再訪時も buildSegments + layoutLines を全段落で再計算（pagination 時に計算した行を破棄）。header/footer reserve で computePages 自体も 2 回走る | `renderer.ts:4009+`, `2197-2218` | M |
| B5 | textbox 用に第 2 の弱いテキストレイアウトエンジン（kinsoku/bidi/justify/tab なし、O(n²) 計測） | `renderer.ts:6553-7117` | M |
| B6 | font 文字列構築（regex 群）が measureText 毎に再実行、memo なし | `renderer.ts:8293-8539` | S |
| B7 | main モードで JSON parse 後のモデルを structured clone で二重シリアライズ | `docx/src/worker.ts:31-35` | S |
| B8 | paginator がパース済みモデルへ直接フィールドを stamp（公開 API に実行時フィールドが漏れ、2-pass で stale stamp が残存） | `renderer.ts:2502-2535` | M |
| B9 | `as unknown as` ~45 箇所（discriminated union で narrowing 可能な箇所が多い）、56 個の test が src/ 直下にフラット | `renderer.ts` 各所 | S |

### C. pptx / xlsx（TS）

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| C1 | [xlsx] `parseSheet` 毎にワークブック全体を `rawData.slice(0)` でコピーし worker へ structured clone（worker は既に `currentBuffer` を保持）| `xlsx/src/workbook.ts:191` | S |
| C2 | [共通] WASM 境界の 3 重シリアライズ: serde_json 文字列 → worker で JSON.parse → main へ structured clone | 各 `worker.ts` + parser `lib.rs` | M |
| C3 | [xlsx] pointermove ホットパスの `getCellAt` / `getHeaderHit` が最大 1,048,576 行の線形走査（O(log n) の `AxisMetrics` が既にあるのに未使用） | `xlsx/src/viewer.ts:746-775, 924-956` | S/M |
| C4 | [xlsx] scroll イベント毎にフル再描画（rAF coalescing なし）、worker モードに stale-frame dropping なし（pptx の `__pptxRenderToken` 相当が無い）。canvas サイズ再代入で backing store を毎回破棄 | `xlsx/src/viewer.ts:496-516` | M |
| C5 | [xlsx] `destroy()` が wrapper DOM を残し、`<style>` をインスタンス毎に `document.head` へ蓄積 | `xlsx/src/viewer.ts:470-483, 2228-2236` | S |
| C6 | [pptx] `destroy()` が呼び出し側所有の canvas ごと wrapper を除去（canvas が DOM から消える） | `pptx/src/viewer.ts:83-95, 324-330` | S |
| C7 | [共通] bidi-line ×3、google-fonts ×3、画像デコードゲーティング（svgBlip/srcRect/metafile 判定）の再実装、pptx worker 内の `decodeDataUrl` 私的コピー — core へ集約すべき | 各パッケージ | M |
| C8 | [xlsx] `drawShapeText` が経験的定数（inset 7px/4px、line `×1.2` / ascent `×0.85`）— `bodyPr@*Ins` をパーサーが落としている。pptx は実測 ascent 済みで同一 DrawingML が別描画になる | `xlsx/src/renderer.ts:3444-3451, 3758-3794` | M |
| C9 | [pptx] デッキ全体を先頭で一括パース（xlsx 型の per-slide lazy なし）、再訪スライドのレイアウト再計算（bitmap LRU なし） | `pptx/parser/src/lib.rs:16-21`, `pptx/src/viewer.ts` | L（LRU のみなら M） |
| C10 | [pptx] `ChartElement → ChartModel` の 60 フィールド手動コピー + xlsx 側 adapter — Rust が nested `chart` を emit すれば 4 箇所同期が消える | `pptx/src/renderer.ts:4316-4373` | S/M |
| C11 | [pptx] render token を canvas への monkey-patch（`as unknown as`）で実装 — WeakMap へ | `pptx/src/renderer.ts:4178-4180` | S |
| C12 | [共通] renderer.ts 4,400 行級の god file（xlsx viewer は renderer のジオメトリを手動ミラー: 「Mirror renderCurrentSheet's startCol」） | 両 renderer.ts, xlsx viewer.ts | M/L |

### D. Rust パーサー / ooxml-common

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| D1 | 存在確認のためだけにメディアを全 inflate（3 パーサー共通パターン） | `docx/parser/src/parser.rs:1270`, `xlsx/parser/src/drawing.rs:1443`, `pptx/parser/src/lib.rs:9328` | S |
| D2 | `extract_image` / `parse_sheet` 等が呼び出し毎に全ファイルを JS→WASM コピー + central directory 再スキャン。stateful `#[wasm_bindgen]` ハンドルで解消 | 各 parser の WASM 面 | M |
| D3 | [xlsx] `parse_sheet` 毎に workbook.xml / theme / **sharedStrings 全体**を再パース。同一 drawing XML をシート毎に 3 回以上 DOM 化 | `xlsx/parser/src/lib.rs:51-99`, `drawing.rs:1361-1500` | M |
| D4 | [pptx] マスター XML を ~10-12 回 DOM 再パース、レイアウトキャッシュ無し（50 スライド 1 レイアウトで 50 回パース + 50 コピー保持）、スライド XML も 2 回パース | `pptx/parser/src/lib.rs:8976, 9160-9290` | M–L |
| D5 | theme パース 3 重実装 → `ooxml_common::theme` へ | pptx `lib.rs:1926`, xlsx `lib.rs:242`, docx `parser.rs:506` | M |
| D6 | OPC rels + パス解決 3 重実装（leading-slash バグは pptx のみ修正済み = 他 2 つに同じ穴の可能性） → `ooxml_common::rels` へ | 各 parser | S/M |
| D7 | `read_zip_entry`/`read_zip_bytes` 3 重実装（pptx 版のみ cap エラーを Option で握り潰す差異） → `ooxml_common::zip` へ | 各 parser | S |
| D8 | DrawingML color-node / fill / txBody パース 3 重実装 → ooxml-common へ（型+パース+純述語の範囲で） | 各 parser | M |
| D9 | WASM 返却がモノリシック JSON `String`（UTF-8→UTF-16 変換 + postMessage コピー + JSON.parse で 3-4 回走査）。`Vec<u8>` 返却 + transferable + 受信側 1 回 parse へ。xlsx の `to_string(&wb).unwrap()` は panic で WASM インスタンスを殺す | 各 parser `lib.rs` | S–M |
| D10 | release プロファイル未最適化（`codegen-units=1` / `panic="abort"` / `wasm-opt=-Oz` なし）。3 バイナリが zip/roxmltree/serde を各自静的リンク | ルート `Cargo.toml` | S（フラグ）/ L（統合） |
| D11 | `parse_docx` のエラーが JSON エスケープなしの `format!` — `"` を含むエラーで invalid JSON。pptx/xlsx は `Result<String, JsValue>` で不統一 | `docx/parser/src/lib.rs:12-21` | S |
| D12 | pptx `lib.rs` 13,645 行 / docx `parser.rs` 10,255 行の god module（`parse_slide` 10 引数、`parse_layout_placeholders` 15+ 引数）。xlsx のモジュール分割が house pattern | 両 parser | M/L |

### E. ビルド / CI / パッケージング

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| E1 | **CI で 1 ピクセルも描画されない**。既存の font 非依存 smoke suite（`tests/smoke/layouts.spec.ts`）が CI 未接続。worker プロトコル不整合や WASM init 失敗が green でマージされ得る | `.github/workflows/ci.yml` | M |
| E2 | `exports` が全エントリで `"import"` を `"types"` より先に列挙 — TS の条件解決順で型解決が壊れる可能性。`"default"` 条件も無し。attw/publint チェック無し | `package.json:37-56` | S |
| E3 | typecheck ジョブが毎 PR で 3 パーサーをコールドビルド（rust ジョブと publish には cache があるのに typecheck だけ無い） | `ci.yml` typecheck job | S |
| E4 | WASM が base64 データ URL として JS にインライン（+33% サイズ、compileStreaming 不可、worker で byte 毎 atob）。core の worker-host チャンクが tarball に 3 重複（~700 KB） | `vite.config.ts`, 公開 tarball | M/L |
| E5 | `sideEffects` フィールド無し + root barrel が 3 フォーマット全部を pull（docx だけ flat re-export される非対称もあり） | `package.json`, `src/index.ts` | S |
| E6 | publish workflow が typecheck のみゲート — テスト・成果物 sanity（pack して import してみる）無し | `publish.yml` | S |
| E7 | ast-grep ルールが 1 個のみ、汎用 TS linter 無し（`@ts-ignore` 禁止等の追加余地） | `rules/` | M |
| E8 | `packages/node` の skia-canvas 描画 probe テスト 9 件が CI で 100% skip（唯一の headless 描画テストなのに） | `packages/node` | S |
| E9 | VRT がメンテナーの macOS フォントスタックに機械固定（committed demo reference が他環境で無用）。Noto CJK 同梱 + bundled chromium 化で CI-VRT への道が開く | `packages/*/playwright.config.ts` | L |
| E10 | site 0.43.0 / mcp-server crate 0.1.0 のバージョンドリフト（リリース手順の bump 対象外）。3.1 MB の生成物 mathjax bundle がコミット済み | `site/package.json` 他 | S |

---

## 改善計画

原則: **(1) 安全網 → (2) 即効の S 課題 → (3) 境界・キャッシュの perf 改善 → (4) 共有層への集約 → (5) 大型構造リファクタ** の順。大型分割（Phase 4）は CI 描画テスト（Phase 1）が入るまで着手しない。

### Phase 1 — 安全網と即効修正（1〜2 週間、すべて S 中心）

リファクタの前提となる CI 強化と、独立して即マージできる小粒修正。

- **CI**: smoke suite を CI 接続（WASM ビルド → Storybook or VRT fixture 起動 → Playwright chromium で `tests/smoke/layouts.spec.ts`）[E1]。typecheck ジョブに rust-cache [E3]。node の skia probe を CI で実行 [E8]。publish に `pnpm test` + `npm pack` → import smoke + attw/publint [E6]。
- **パッケージング**: exports の `types` 先頭化 + `default` 追加 [E2]。`sideEffects: false`（副作用監査後）+ docx flat re-export の整理 [E5]。
- **正しさ/堅牢性（ユーザー可視バグ）**: pptx `destroy()` の canvas 返還 [C6]。xlsx `destroy()` の wrapper/style 除去 [C5]。dash テーブル統合（`lgDash` が solid になる実バグ）[A6]。`parse_docx` エラー JSON エスケープ + `Result` 統一 [D11]。xlsx serializer の `unwrap` 除去 [D9 の一部]。
- **即効 perf**: zip 存在確認の inflate 廃止 [D1]。xlsx `parseSheet` の全バッファ再送廃止（worker の `currentBuffer` 使用）[C1]。xlsx ヒットテストを `AxisMetrics` 経由に [C3]。docx font 文字列 memo [B6]。Blob sniff を `slice(0,8)` に [A8 前半]。Cargo release プロファイル + wasm-opt フラグ [D10 前半]。
- **小粒衛生**: `read_zip_*` の ooxml-common 集約 [D7]。core barrel のテスト専用輸出削除 [A11]。WMF/EMF/DIB の OffscreenCanvas fallback [A9]。pptx render token の WeakMap 化 [C11]。

### Phase 2 — WASM 境界とキャッシュの再設計（2〜4 週間、M 中心）

ロード時間とナビゲーション体感に最も効くまとまり。着手前に代表ファイル（大型 docx / 200 スライド pptx / 多シート xlsx）でベンチマークを取り、各項目の前後比較を残す。

1. **境界プロトコル統一** [C2, B7, D9]: パーサーは `Vec<u8>`（JSON bytes）返却 → worker は transferable で素通し → 受信側で 1 回だけ `JSON.parse`。3 フォーマット同時に（横断統合原則）。
2. **stateful アーカイブハンドル** [D2, D3]: `#[wasm_bindgen]` の per-document ハンドル（bytes + central directory + sharedStrings/theme キャッシュ）。`extract_image` / `parse_sheet` の毎回全コピーを廃止。
3. **pptx パーサーの DOM 再パース削減** [D4]: マスター/レイアウト/スライドの parse-once 化、レイアウトキャッシュ追加。
4. **docx 画像デコードキャッシュ** [B3]: `Map<string, DecodedImage>` を `DocxDocument` に保持（`destroy()` で解放）。core に `getCachedBitmapByPath`（SVG キャッシュの sibling）を追加して 3 フォーマットで共用 [A8 後半]。
5. **xlsx スクロール描画** [C4]: rAF coalescing + サイズ不変時の canvas 再確保スキップ + worker モードの stale-frame dropping（pptx 方式の移植）。
6. **core effect 系** [A3, A4, A5]: aux canvas の bbox サイズ化 + プール、extrusion の GPU 合成化、preset formula のプリコンパイル。
7. **WorkerBridge** [A10]: timeout / AbortSignal 対応。

### Phase 3 — 共有層への集約（2〜4 週間、CLAUDE.md 横断統合原則の実行）

「型 + パース + 純述語」の範囲で ooxml-common / core へ寄せる。1 関心 = 1 commit、各集約で 3 パッケージの挙動差（= 潜在バグ）を洗い出して同時修正。

- **Rust**: `ooxml_common::theme`（D5、xlsx `parse_theme_ln_widths` の重複も吸収）→ `ooxml_common::rels`（D6、leading-slash 処理の差異を必ず突き合わせ）→ color-node / fill / txBody（D8、color node から段階的に）。
- **TS core**: bidi-line、google-fonts テーブル、画像デコードゲーティング述語（svgBlip/srcRect/metafile 判定）、`decodeDataUrl` の私的コピー削除 [C7]。
- **チャートモデル**: Rust が nested `chart: ChartModel` を emit し、pptx の 60 フィールド手動コピーと xlsx adapter を削除 [C10]。
- **spec 忠実化**: xlsx `drawShapeText` — `bodyPr@lIns/tIns/rIns/bIns` のパース追加（§21.1.2.1.1）+ pptx の実測 ascent 方式を core へ lift して共用 [C8]。チャートラベルの `elideToWidth` 化 [A7]。
- **プロセス**: docx の既知ヒューリスティック（float page-fit §17.4.57、frame wrap-band §17.3.1.11 等）を issue 化して追跡 [B12 相当]。

### Phase 4 — 大型構造リファクタ（順次、L）

CI 描画テストが回っていることを前提に、衝突コストの高いものから。

1. **docx renderer 分割 + measure/paint 統一** [B1, B2, B4, B5, B8]: まず機械的なモジュール分割（images / paginate / line-layout / paint-paragraph / tables / shape-text / fonts / notes）→ `__test_*` バックドア廃止・テスト再配置 [B9, B10] → レイアウトを pt 空間の単一成果物（`LayoutLine[]` + row heights）に一本化し paginator と painter が同じデータを消費 → textbox を主エンジンへ統合。**この統一が docx の回帰クラス（paginate/paint 乖離）を構造的に消す本丸。**
2. **pptx parser 分割** [D12, D4]: xlsx のモジュール構成を範型に theme / fill / text / shape / chart / master / markdown へ。`ParsedMaster` / `ParsedLayout` コンテキスト構造体で 10〜15 引数関数を解消。
3. **core chart layout 抽出** [A1, A7, A11 系]: `computeChartFrame` + 向きパラメタライズされた axis/gridline painter を共有し、ファミリー別コードはマーク描画のみに。凡例は実測サイズ化。
4. **プリセット図形エンジン一本化** [A2]: legacy `buildShapePath` の呼び出し面を棚卸し → spec 駆動エンジンへバッチ移行（バッチ毎に VRT）→ 2,147 行 switch を削除。
5. **pptx lazy slide parsing + bitmap LRU** [C9]: parser API を meta + `parse_slide(i)` に分割（Phase 2-2 のハンドル上に自然に載る）。先行して bitmap LRU だけ入れるのも可。
6. **配布形態** [E4, E9, D10 後半]: `.wasm` の実アセット化（`wasmUrl` オプション、data-URL は fallback）、worker チャンク 3 重複の解消検討、VRT の決定論フォント化 → CI-VRT。

### 取り組み方の注意

- **各 Phase 2 項目は必ずベンチ前後比較を PR に記載**（体感目標: 大型ファイルの初回表示、シート/スライド切替、スクロール中のフレーム落ち）。
- 横断修正（Phase 3）は CLAUDE.md の規定どおり core / ooxml-common を触る 1 本の協調 PR にまとめてよいが、commit は 1 関心ずつ。
- Phase 4-1（docx 統一）だけは他と独立してブランチ寿命が長くなるため、分割（機械的移動）と統一（挙動変更）を別 PR に分けること。
- VRT reference はユーザー指示なしに更新しない（既存ルールどおり）。Phase 4 の各バッチは demo サンプルの VRT をローカルで回してから push。
