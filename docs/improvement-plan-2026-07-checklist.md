# 改善計画 実行チェックリスト（2026-07）

`docs/improvement-plan-2026-07.md` の実行トラッカー。項目 ID（A1, B2, C1, D5, E2 …）は計画本文の表を参照。
ローカルセッションで進める際は、完了した項目にチェックを入れてこのファイルを更新していく。

## 前提・引き継ぎ事項

- レビューはリモートセッションで実施。**主要指摘（god file 行数、`rawData.slice(0)`、pptx/xlsx `destroy()` の挙動、`parse_docx` のエラー整形、exports の条件順、CI typecheck ジョブの rust-cache 欠如）はコード上で裏取り済み。** それ以外の細部（行番号・定数値）は着手時に現物を再確認すること。
- VRT はローカル専用（`private/sample-*` が必要）。**レンダリングに触る修正は必ずローカルで `pnpm build:wasm && pnpm vrt` を回してから push。** リモートでは実行不可だった点に注意。
- 参照画像の更新（`UPDATE_REFS=1`）はユーザー明示指示のみ。
- 横断修正（Phase 3）は CLAUDE.md の規定どおり core / ooxml-common を含む 1 本の協調 PR にまとめてよい。commit は 1 関心ずつ。
- Phase 2 の各項目は **PR にベンチ前後比較を記載**（大型 docx / 200 スライド級 pptx / 多シート xlsx で計測）。
- **レビュー実施後に `DocxScrollViewer` / `PptxScrollViewer`（#650 系〜#656）が main に追加された。** 本計画の viewer 系指摘（C5/C6 の destroy ライフサイクル、C4 のスクロール描画、C7 の共有化）は着手時に ScrollViewer 系クラスも同じ観点で点検し、同種の穴があれば同一 PR で修正する。

## 検証コマンド early reference

```bash
pnpm build:wasm                          # parser を触ったら必須
pnpm test                                # vitest（unit）
pnpm typecheck
pnpm vrt                                 # ローカルのみ・全パッケージ
pnpm --filter @silurus/ooxml-docx vrt    # 単一パッケージ
cargo fmt --all --check && cargo clippy --all-targets -- -D warnings
cargo test
```

---

## Phase 1 — 安全網と即効修正

### CI / パッケージング

- [ ] E1: smoke suite（`tests/smoke/layouts.spec.ts`）を CI に接続（WASM ビルド → Storybook 起動 → Playwright chromium）
- [ ] E3: typecheck ジョブに `Swatinem/rust-cache@v2`（3 parser workspace 指定。rust ジョブ / publish.yml の記述を流用）
- [ ] E8: `packages/node` の skia-canvas probe テストを CI で実行（skia-canvas を devDep 化 or 明示 install、skip されていないことをログで確認）
- [ ] E6: publish.yml に `pnpm test` + `npm pack` → 一時 dir で `import('@silurus/ooxml/docx')` smoke + `attw --pack` / `publint`
- [ ] E2: ルート package.json の全 exports で `types` を先頭に、`default` 条件を追加。**公開済み 0.69.0 の tarball で現状の型解決を先に確認**
- [ ] E5: `sideEffects: false` 追加（module スコープ副作用を監査してから）。`src/index.ts` の docx flat re-export（pptx/xlsx と非対称）を整理

### 正しさ / 堅牢性（ユーザー可視）

- [ ] C6: pptx `viewer.ts` `destroy()` — `wrapper.remove()` 前に caller の canvas を `insertBefore` で返還。検証: destroy → 同一 canvas で再生成
- [ ] C5: xlsx `viewer.ts` `destroy()` — wrapper subtree と注入 `<style>` を除去（style は module-wide 1 回注入への変更でも可）。検証: Storybook で mount/unmount 繰り返し
- [ ] C5/C6 追補: `DocxScrollViewer` / `PptxScrollViewer` の `destroy()` を同観点（caller 所有 DOM の返還・注入 style/listener の除去）で点検し、同種の穴があれば同一 PR で修正
- [ ] A6: dash テーブル統合 — `core/src/draw/dash.ts` を正とし `paint.ts` の `DASH_PATTERNS` を消す。ST_PresetLineDashVal（§20.1.10.49）の全キーで両者の出力差を先に一覧化（`lgDash`→solid バグの確認）
- [ ] D11: `parse_docx` を `Result<String, JsValue>` に統一（エラー JSON の手組み format! 廃止）。TS 側の catch パス確認
- [ ] D9(部分): xlsx `lib.rs` の `serde_json::to_string(&wb).unwrap()` を `map_err` に

### 即効 perf

- [ ] D1: zip 存在確認の全 inflate 廃止（docx `load_media_map` / xlsx `build_drawing_rid_urls` / pptx master bg）→ `index_for_name` 系へ。`ooxml_common::zip::entry_exists` として共有
- [ ] C1: xlsx `workbook.ts` `parseSheet` の `rawData.slice(0)` 廃止 — worker の `currentBuffer` を使用（worker 再起動時の fallback として optional data は残す）
- [ ] C3: xlsx `getCellAt` / `getHeaderHit` を `AxisMetrics.indexAt()` 経由に（frozen-pane 分岐に注意）。検証: 50 万行スクロール後の pointermove
- [ ] B6: docx `normalizeFontFamily` / `buildFont` の memo 化（`Map<family, chain>`、`fontFamilyClasses` は per-doc なので identity キー or render 毎リセット）+ `ctx.font` 不変時の再代入スキップ
- [ ] A8(前半): `decodeRasterOrMetafile` の sniff を `data.slice(0, 8).arrayBuffer()` に
- [ ] D10(前半): ルート Cargo.toml に `codegen-units=1` / `panic="abort"` / `strip="debuginfo"`、各 parser に `wasm-opt = ["-Oz"]`。**前後の .wasm サイズを記録**

### 小粒衛生

- [ ] D7: `read_zip_entry` / `read_zip_bytes` を `ooxml_common::zip` に集約（pptx 版だけ cap エラーを握り潰す差異があるので挙動統一に注意）
- [ ] A11: core barrel からテスト専用 bevel 内部（`edt1d` 等）・未消費 export を削除（テストは deep import に変更）
- [ ] A9: WMF/EMF/DIB の `new OffscreenCanvas` 直呼びを `createAuxCanvas` 系 fallback 付き factory に統一
- [ ] C11: pptx render token（`__pptxRenderToken` monkey-patch）を module-level WeakMap に

## Phase 2 — WASM 境界とキャッシュ再設計

- [ ] ベンチ基盤: 代表ファイル 3 種で初回表示 / シート・スライド切替 / スクロール FPS の計測スクリプトを用意（以降の各項目で前後比較）
- [ ] C2+B7+D9: 境界プロトコル統一 — parser は `Vec<u8>`（JSON bytes）返却 → worker は transferable 素通し → 受信側で 1 回 `JSON.parse`。**3 フォーマット同時に**
- [ ] D2: stateful `#[wasm_bindgen]` アーカイブハンドル（bytes + central directory index 保持、`extract(path)` / `parse_sheet(i)`）。worker の `currentBuffer` ライフサイクルに載せる
- [ ] D3: xlsx sharedStrings / theme / workbook.xml のハンドル内キャッシュ。sheet rels + drawing XML の parse-once 化（現状 1 つの drawing を 3 回以上 DOM 化）
- [ ] D4: pptx マスター DOM の parse-once 化（~10-12 回 → 1 回）、layout cache 追加（`master_cache` の鏡映）、slide XML の 2 回パース解消
- [ ] B3: docx 画像デコードキャッシュ（`Map<string, DecodedImage>` を DocxDocument に、`destroy()` で解放）
- [ ] A8(後半): core に `getCachedBitmapByPath`（SVG キャッシュの sibling、ImageBitmap の close 管理付き）を追加し 3 フォーマットで共用
- [ ] C4: xlsx スクロール — rAF coalescing（1 フレーム 1 render・latest wins）+ サイズ不変時の canvas 再確保スキップ + worker モードの世代カウンタで stale bitmap 破棄
- [ ] A4: effect 系 aux canvas を bbox+ぼかし半径サイズに縮小 + module-level プール
- [ ] A3: `applyExtrusion` を silhouette bbox 限定 + drawImage 反復の GPU 合成に置換、`applyBevelShading` も bbox 限定
- [ ] A5: preset-geometry formula のプリコンパイル（`Map<presetName, CompiledDef>`）。出力バイト一致を確認
- [ ] A10: `WorkerBridge.request` に `{ timeoutMs, signal }` + worker `error` イベントで pending 一括 reject
- [ ] B4: docx 同一ページ再描画時の layout 再計算削減（pt 空間 `LayoutLine[]` のキャッシュ。丸め差に注意 — 本格統一は Phase 4-1）

## Phase 3 — 共有層への集約（横断 1 PR / commit は関心毎）

- [ ] D5: `ooxml_common::theme`（clrScheme / fontScheme / lnStyleLst / objectDefaults）— 3 実装の挙動差を先に一覧化
- [ ] D6: `ooxml_common::rels`（parse / find-by-type / resolve target）— **leading-slash（OPC 絶対 Target）の扱いを 3 パーサーで突き合わせ**（pptx のみ修正済みの穴）
- [ ] D8: DrawingML color-node → fill → txBody の順で段階的に ooxml-common へ（型+パース+純述語の範囲厳守）
- [ ] C7: TS 側 — bidi-line ×3 / google-fonts ×3 を core へ、画像デコードゲーティング述語（svgBlip / srcRect / metafile）を core に、pptx worker の `decodeDataUrl` 私的コピー削除
- [ ] C10: Rust parser が nested `chart: ChartModel` を emit → pptx 60 フィールド手動コピーと xlsx adapter を削除
- [ ] C8: xlsx `bodyPr@lIns/tIns/rIns/bIns` パース追加（§21.1.2.1.1）+ pptx の実測 ascent 方式を core へ lift（7px/4px・×1.2・×0.85 の経験的定数を撤去）
- [ ] A7: チャートラベルの `slice(0,N)` を `elideToWidth(ctx, text, maxPx)` に統一
- [ ] B12: docx の既知ヒューリスティック（float page-fit §17.4.57 / frame wrap-band §17.3.1.11 / inside-outside margin 近似等）を GitHub issue 化

## Phase 4 — 大型構造リファクタ（CI 描画テスト稼働が前提）

- [ ] B1: docx renderer 機械的分割（images / paginate / line-layout / paint-paragraph / tables / shape-text / fonts / notes）— **挙動変更なしの移動のみで 1 PR**
- [ ] B9/B10: `__test_*` バックドア廃止 + 56 テストのモジュール別再配置（分割の直後に）
- [ ] B2: measure/paint 統一 — レイアウトを pt 空間の単一成果物に一本化（**分割とは別 PR**。element type 毎に段階施行、各段で VRT）
- [ ] B5: textbox テキストを主エンジン（buildSegments/layoutLines）に統合（kinsoku/bidi/justify が textbox でも効くようになる）
- [ ] B8: paginator の model 直接 stamp をラッパーレコードに置換
- [ ] D12: pptx parser 分割（xlsx のモジュール構成を範型に theme / fill / text / shape / chart / master / markdown + `ParsedMaster`/`ParsedLayout` 構造体）
- [ ] A1: core chart layout 抽出（`computeChartFrame` + 向きパラメタライズ axis painter、凡例実測サイズ化）— ファミリー毎に段階移行
- [ ] A2: プリセット図形エンジン一本化 — legacy `buildShapePath` 呼び出し面の棚卸し → spec 駆動エンジンへバッチ移行（バッチ毎 VRT）→ 2,147 行 switch 削除
- [ ] C9: pptx parser API を meta + `parse_slide(i)` に分割（D2 ハンドル上に実装）。先行で bitmap LRU のみも可
- [ ] E4: `.wasm` 実アセット化（`wasmUrl` オプション、data-URL fallback 維持）/ worker チャンク 3 重複の解消検討
- [ ] E9: VRT の決定論フォント化（Noto CJK を fixture に `@font-face` 同梱 + bundled chromium）→ CI-VRT・demo reference の可搬化
- [ ] E10: リリース手順の bump 対象に site / mcp-server Cargo.toml を追加、mathjax 生成物のコミット運用見直し

## 推奨 PR 分割（Phase 1 の目安）

1. `ci/rendering-smoke` — E1 + E3 + E8（安全網を最優先で 1 本）
2. `fix/viewer-destroy` — C5 + C6（ユーザー可視のライフサイクルバグ）
3. `fix/core-dash-table` — A6（描画差の出る実バグ、VRT 必須）
4. `fix/parser-error-api` — D11 + D9(unwrap) + D7（Rust エラー/zip 衛生をまとめて）
5. `perf/zip-and-sheet-transfer` — D1 + C1 + D10(フラグ)（ベンチ添付）
6. `perf/hot-path-small` — C3 + B6 + A8(前半) + C11
7. `chore/packaging` — E2 + E5 + E6（npm 消費側の検証を伴うため独立）
