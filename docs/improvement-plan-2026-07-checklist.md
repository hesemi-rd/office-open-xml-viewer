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

## Phase 1 — 安全網と即効修正 ✅ 完了（2026-07-02、PR #658–#664）

全 22 項目完了。各項目の注記は実装時の再定義・追加発見を含む（計画の指摘が現物と異なった項目: A6 / D7 / A8 / A9 / E5 barrel）。CI 運用知見: smoke ジョブの `playwright install chrome --with-deps` が apt/CDN 起因で 30 分超ハングする事象が 2 回連続発生 → プリインストール Chrome があれば skip + 全ジョブ timeout-minutes で恒久対策済み（#662 内）。

### CI / パッケージング

- [x] E1 (#658): smoke suite（`tests/smoke/layouts.spec.ts`）を CI に接続（WASM ビルド → Storybook 起動 → Playwright chromium）
- [x] E3 (#658): typecheck ジョブに `Swatinem/rust-cache@v2`（3 parser workspace 指定。rust ジョブ / publish.yml の記述を流用）
- [x] E8 (#658): devDep 化 + `OOXML_REQUIRE_SKIA=1` で silent skip を hard fail 化。pnpm 10 の build-script 承認（`onlyBuiltDependencies: [skia-canvas]`）が必須と判明し追加。初実行でテスト自体の実バグ 2 件（skia@2 API 不一致・到達不能アサーション）と pptx unhandled-rejection リークを検出・修正
- [x] E6 (#664): publish.yml に test + publint + attw `--profile esm-only`（fail-closed）+ tarball smoke（temp dir に npm i して 3 サブパス import）を追加
- [x] E2 (#664): types 先頭化 + default 追加。実測: tsc 4 象限は before も pass だが、attw が TS bug #50762 の「fallback condition 誤用」🐛 を bundler/node16-ESM 全 5 エントリで検出 → 修正後 🟢、publint 5 errors → clean
- [x] E5 (#664): sideEffects: false を root+4 パッケージに追加（grep 監査 + 独立監査 + DOM なし Node import 成功の三重確認）。「docx flat re-export 非対称」は計画の誤読と判明（root は対称、flat barrel は各パッケージ自身のもの）— API 変更なし

### 正しさ / 堅牢性（ユーザー可視）

- [x] C6 (#659): pptx `viewer.ts` `destroy()` — `wrapper.remove()` 前に caller の canvas を `insertBefore` で返還。検証: destroy → 同一 canvas で再生成
- [x] C5 (#659): xlsx `viewer.ts` `destroy()` — wrapper subtree と注入 `<style>` を除去（style は module-wide 1 回注入への変更でも可）。検証: Storybook で mount/unmount 繰り返し
- [x] C5/C6 追補 (#659): ScrollViewer ×2 は点検の結果、完全実装で修正不要。代わりに **DocxViewer に pptx と同一の canvas 喪失バグを発見し同 PR で修正**。レビューで stale-nextSibling の DOM 仕様違反（NotFoundError）も検出しガード追加、テスト DOM の二重所属欠陥も修正
- [x] A6 (#660) **再定義**: 「lgDash→solid」は誤診（dash.ts のテーブルは実は §20.1.10.82 ST_TextUnderlineType の下線用で、prstDash 値は流れない）。真バグ = paint.ts の `sysDashDotDot` 欠落（solid に化ける）を修正。preset テーブルを dash.ts へ `pptxPresetDashArray` として移設（byte-equivalence テスト付き）、下線側を `pptxUnderlineDashArray` に改名し誤 spec 引用を修正
- [x] D11 (#661): Result 化 + TS 受け口 2 箇所（worker.ts / render-worker.ts）の error フィールド probe 削除。壊れた入力で `docx-parser error: ...` throw を end-to-end 確認
- [x] D9(部分) (#661): map_err 化（panic による WASM インスタンス死を排除）

### 即効 perf

- [x] D1 (#662): 計画の 3 箇所に加え grep で 10 箇所追加発見、**計 13 箇所**を `index_for_name` に置換（bytes を実際に使う 3 箇所は正しく除外）。共有ヘルパーは不要と判断（archive ハンドル既保持のため inline が自然）
- [x] C1 (#662): `data` フィールドをプロトコルから**削除**（optional で残すより契約が明確 — parse 前の parseSheet は明示エラー）。worker-vs-main VRT 0.000% diff で挙動同一を証明
- [x] C3 (#663): `scrollableIndexAt` 新設で O(log n) 化。ベンチ: 50 万行 ×1000 ヒットで 9296ms → 0.378ms（~24,600×）、旧実装 verbatim コピーを oracle にした全数パリティ 0 差
- [x] B6 (#663): `WeakMap<fontFamilyClasses, Map<family, css>>`（identity キー、呼び出し面変更ゼロ）。font 再代入スキップは測定ループの 2 writer（measureText/strAdvance）を単一トラッカーで統一（draw パス 12 箇所は据え置き）
- [x] A8(前半) (#663): slice は **44 バイト**が正（EMF 判定に offset 40-43 が必要。計画の 8 バイトでは不足）。raster は Blob 直渡しで全体コピー消滅
- [x] D10(前半) (#662): 3 バイナリ計 1,967,032 → 1,903,734 bytes（**−3.22%**）。ビルド 23.8s → 33.4s。panic=abort でも console_error_panic_hook のスタック出力は保持（hook は abort 前に走る）

### 小粒衛生

- [x] D7 (#661): 実像は「docx にも私的ヘルパーあり」で**計 5 実装・~45 呼び出し面**を generic `read_zip_bytes/read_zip_string<R: Read+Seek>` に集約。cap 超過は Err 返却とし、画像パスの従来挙動（スキップ）は呼び出し側の明示的 `.ok()` で維持
- [x] A11 (#663): 4 export（edt1d/shadePixel/shadeParamsFor/fillDirFromKey）削除。materialClass/lightDirFromRig は pptx が実使用のため保持。テストは既に deep import で変更不要。dead export はゼロ（全数調査）
- [x] A9 (#663): 無条件直呼びは wmf.ts の 1 箇所のみだった（emf/dib はチェック済み）。createAuxCanvas を `core/canvas/aux-canvas.ts` へ移設し 3 ファイル統一 — EMF/DIB は main-thread fallback を獲得（strict superset）。worker 専用の OffscreenCanvas+transferToImageBitmap は正当につき対象外。pattern-bitmaps の重複 factory も統合
- [x] C11 (#663): pptx に加え **docx にも同型 monkey-patch を発見**（横断原則）— 両方 WeakMap 化。xlsx は同期描画で該当なし

## Phase 2 — WASM 境界とキャッシュ再設計 ✅ 完了（2026-07-03、PR #666–#672）

体感成果: 15MB pptx パース 30→11ms（#666+#669 複合）、xlsx シート切替 1.61×+転送ゼロ、docx ページ再訪 ~9×、xlsx スクロール 100 イベント→1 render、wedged worker の永久ハング解消。全 PR にベンチ/等価性証明（sha256 golden / oracle / byte-identity）添付。

- [x] ベンチ基盤 (#666/#671): `packages/node/src/bench-parse.mjs`（境界込み parse 時間、string/bytes 自動判別）+ `bench-handle.mjs`（parse 後の反復 work: 全シート切替/全画像抽出、--wasm-dir で before 計測）
- [x] C2+B7+D9 (#666): 4 parse 関数を Vec<u8>(JSON bytes) 化、worker は transferable 素通し、main で 1 回 decode+parse（render-worker は in-worker 消費なので現行維持）。15MB pptx median 30→15ms(~2×)。レビューの「WASM memory 全体転送」指摘は生成コードの .slice() 確認で却下
- [x] D2 (#671): stateful `#[wasm_bindgen]` アーカイブハンドル（`DocxArchive`/`PptxArchive`/`XlsxArchive` = 所有 `ZipArchive<Cursor<Vec<u8>>>` + `max` 保持）。bytes を WASM へ 1 回コピー・central directory を 1 回スキャンし `parse()`/`extract_image(path)`/`parse_sheet(i,name)` を retained archive 上で提供。フリー関数（`parse_*`/`extract_*`）は所有 archive を張る thin wrapper 化で温存（node/markdown/stories/MCP 無変更）。各 worker は `currentBuffer: Uint8Array` を `archive: *Archive|null` に置換し、構築後は JS 側 bytes を保持しない（メモリ二重化解消）。再 parse 時 `disposeArchive()` で明示 free（二重 free/UAF ガード）。ベンチ: pptx sample-4 (15MB, parse+5 media) **1.49×**、docx sample-9 (13 images) 1.10×
- [x] D3 (#671): xlsx `WorkbookShared`（workbook.xml/rels source + sheet list + theme palette + sharedStrings）をハンドル内で 1 回パースし全シート切替で再利用。`parse_sheet` 毎の sharedStrings/theme 全再パースを解消（`parse_sheet_with`/`parse_xlsx_inner_with` に分離、フリー関数は毎回 fresh build で従来コスト維持）。ベンチ: sample-12 (8 sheets) **1.61×**、sample-9 (3 sheets) 1.30×。roxmltree `Document` は借用のためキャッシュ不可 → cached string から都度再パース（inflate なしで安価）。drawing XML の parse-once 化は sheet 単位再パース解消を主目的とし本コミット範囲では見送り（別項）
- [x] D4 (#669): master 11 extractors+bg が単一 Document 共有、ParsedLayout+layout_cache で 4×S→1/distinct+1/slide。「slide XML 2 回パース」は誤認（1 回）。layout decorative は slide 固有 smartart 依存で意図的非キャッシュ（unsound 回避）。出力 sha256 byte-identical、-20%/-26% parse 回数、cfg(test) カウンタで 2+2N を regression 固定
- [x] B3 (#668): base bitmap は core キャッシュ、a:clrChange recolor は第 2 層（per-fetchImage WeakMap）でメモ化。ページ再訪の画像コスト ~9×（0.27→0.03ms）。destroy() で 3 キャッシュ全 drop
- [x] A8(後半) (#668): pptx 実装を core/image/bitmap-image-by-path.ts へ verbatim lift（LRU 256、同期 peek、eviction close、rejection-handler 規律をヘッダに明文化）。pptx は thin re-export。**xlsx は意図的非移行**（raster+SVG 同居の workbook 所有 Map で構造が異なる — false-abstraction 回避、二重 LRU と close 所有権競合を防ぐ）
- [x] C4 (#667): scheduleRender で scroll/resize/drag を rAF coalesce（100 イベント→1 render）、明示 API は同期維持。size guard は orchestrator 側にも（+setTransform 冪等化）。_renderSeq 世代で stale bitmap close。worker-vs-main VRT 0.000%
- [x] A4 (#672): innerShadow/softEdge(×3)/reflection の aux を bbox⊕margin に縮小（blur は 3σ=3·blur — pixel テストが margin バグを捕捉）。ピクセル −~99%（79–87×）。**プールは不採用**（bbox 化後はサイズ不一致でヒットせず）。**reflection は full-canvas に revert** — mirror blit のリサンプリングが crop 端で skia プラットフォーム依存（Linux CI で δ≤7 検出）。tripwire テストで再 crop を防止
- [x] A3 (#672) **縮小案採用**: region 限定（getImageData/loop/putImageData）のみ、GPU 合成は VRT 割れリスクで見送り。byte-identical 証明済み。正直な注記: pptx は shape 専用オフスクリーンのため現行パスの実利は pad 分のみ — seam は大 canvas 呼び出し向け
- [x] A5 (#672): {op, argTokens} へ lazy プリコンパイル（WeakMap、evaluatorForDef 単一チョークポイント）。186 プリセット全数 oracle 一致。split() 呼び出し −99.6%（12.9×）
- [x] A10 (#670): per-request timeout + AbortSignal + worker error/messageerror の pending 一括 reject（常時有効）。timeout は opt-in（LoadOptions.workerTimeoutMs、既定無制限 — 巨大ファイルの正当な長時間パースを壊さない）。viewer 5 箇所の明示 forwarding
- [~] B4: **Phase 4-1 に統合**。B3 で再訪コストの主犯（画像再デコード）は解消済み。LayoutLine[] キャッシュは B2 統一（compute-once 単一成果物）の副産物として実装するのが正道 — 中間キャッシュを別実装すると Phase 4-1 で捨てることになるため

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
