# 改善計画 実行チェックリスト(2026-07 第2次)

`docs/improvement-plan-2026-07b.md` の実行トラッカー。項目 ID(CH1, XF2, WD3, IX1, SC5, RB1, PD1, QA6, AR1 …)は計画本文の表を参照。
ローカルセッションで進める際は、完了した項目にチェックを入れてこのファイルを更新していく。

## 前提・引き継ぎ事項

- レビューはリモートセッションで実施(main = 1610253 時点)。**全指摘は独立検証エージェントがコード上で裏取り済み**(152 claim 中 REFUTED 0)だが、行番号・定数は着手時に現物を再確認すること。
- VRT はローカル専用。**レンダリングに触る修正は必ず `pnpm build:wasm && pnpm vrt` を回してから push。**
- 参照画像の更新(`UPDATE_REFS=1`)はユーザー明示指示のみ。**新規参照が必要な項目(CH12 docx チャート、CH15 chartEx、XF9 縦書き、WD4 kern 等)は着手前にユーザーと参照更新の段取りを合意する。**
- perf 項目(SC 系)は PR にベンチ前後比較を記載(bench-parse / bench-handle + 対象別の合成 fixture)。
- 横断修正は CLAUDE.md の規定どおり core / ooxml-common を含む1本の協調 PR にまとめてよい(commit は 1 関心 = 1 commit)。
- セキュリティ項目(RB 系)は「悪意入力の再現 fixture + 修正後に無害化されることのテスト」をセットで。
- 既知の継続スレッド: issue #698(SPACE_SHRINK_RATIO)、A2 残り 120 プリセット移行(参照更新承認とセット)は本計画と別トラック。

## 検証コマンド early reference

```bash
pnpm build:wasm                          # parser を触ったら必須
pnpm test                                # vitest(unit)
pnpm typecheck
pnpm vrt                                 # ローカルのみ・全パッケージ
pnpm --filter @silurus/ooxml-docx vrt    # 単一パッケージ
cargo fmt --all --check && cargo clippy --all-targets -- -D warnings
cargo test
node packages/node/src/bench-parse.mjs   # 境界込み parse ベンチ
node packages/node/src/bench-handle.mjs  # handle 反復 work ベンチ
```

---

## Phase 1 — 即効の正しさ・安全網・クイックウィン

### 正しさ(チャート)
- [x] CH1: 棒チャート負値 — 縦棒/横棒とも実 dataMin からゼロライン基準の両方向描画に一般化。stacked は正負分離スタック、percentStacked は符号保持(分母 Σ|v| と逆側スタックは spec が定めない Excel/PowerPoint 実挙動である旨をコメント化)。正値のみチャートは byte-stable(snapshot 不変で証明)
- [x] CH2: stackedLine / stackedLinePct — area の stackBase パターンを移植。**追加発見: 参照実装とされた stackedAreaPct 自体が正規化未実装** → 同 PR で修正(1c56f66)。bar/line/area の pct 分配規約(Σ|v|・符号保持)を 3 実装で統一
- [x] CH3: xlsx チャート色解決の私的コピー排除(bg2/tx2 + HSL) — 私的 3 関数を削除し ooxml_common::color::parse_color_node + 共有 XlsxSchemeResolver へ委譲。横断点検の結果 **pptx は委譲済みでバグ無し**(xlsx が唯一の外れ値だった)。注: alpha 付き solidFill は 8 桁 hex になる(shape 経路と整合、旧実装は alpha を落としていた)
- [x] CH4: c:dateAx 認識 — xlsx/pptx 両パーサーで catAx と同扱い(§21.2.2.39 の EG_AxShared 共有構造)。pptx は従来 None 固定だった cat_axis_format_code の配線も獲得。**レビューで発覚した穴: catAxisFormatCode の TS 側消費者が scatter のみで時系列 line/bar が生シリアル値のまま** → §21.2.2.71 の一般則として全カテゴリ軸ファミリーの目盛りラベル書式化(formatCategoryLabel)を core に追加実装(a39eca8)。baseTimeUnit / majorTimeUnit 等の日付単位制御は CH6 の範囲
- [x] CH11: チャート数値のロケール固定(§18.8.30 エンジンへ) — **再定義: 実箇所は scatter でなく waterfall のみ**(scatter は formatChartValWithCode 対応済みだった)。waterfall の軸目盛り/データラベル 2+1 箇所を formatChartValWithCode へ(valAxisFormatCode / dataLabelFormatCode 配線込み)。renderer 内の toLocaleString は 0 件に

> CH1–CH4/CH11 は PR fix/chart-correctness(9 commits)。検証 = vitest 1757 / typecheck / cargo 3 点 / VRT 163 全 green・参照画像差分ゼロ。独立 4 観点レビュー(spec 忠実性・正しさ・横断統一・テスト品質)→ REQUEST_CHANGES 3 件(MAJOR)を修正済み。別トラック送りの発見: pptx の cat_axis_crosses/crossesAt None 固定(CH6 圏)、percentStacked 分母計算の三重実装の共通化(コスメティック)、stacked の null セル= dispBlanksAs zero 既定(CH9 圏)

### 正しさ(xlsx / DrawingML / docx)
- [ ] XL1: date1904(+1900-02-29 バグ互換)
- [ ] XL2: General 書式 11 有効桁
- [ ] XF8: scrgbClr / hslClr / prstClr transforms
- [ ] XF11: docx 下線 17 種(pptx drawUnderline の core hoist 込み)
- [ ] XF12: cNvPr@hidden(3フォーマット)
- [ ] WD1: noBreakHyphen / cr / ptab / softHyphen
- [ ] WD8: kashida / thaiDistribute → both マップ(+真 kashida issue 起票)
- [ ] WD9: OMML phant / sPre / box / borderBox
- [ ] PP2: SmartArt relId バインディング
- [ ] XF1: **OOXML Strict 名前空間**(共通述語 + 3パーサー + Strict fixture)

### セキュリティ(S)
- [ ] RB3: sparkline レンジのセル数キャップ
- [ ] RB4: DIB 検証順序 + megapixel 予算
- [ ] RB11: zip reserve のキャップ
- [ ] RB12: CFB シグネチャ検知 → typed error
- [ ] RB8: cargo audit / pnpm audit ゲート(report-only 開始)

### worker / ライフサイクル
- [ ] SC17: Archive 所有権渡し(WASM 内二重コピー解消)
- [ ] SC18: extract 系の buffer 直 transfer
- [ ] SC20: load() の旧エンジン orphan 解消(5箇所)
- [ ] AR4: pptx worker init 失敗の永久 hang
- [ ] PD14: render エラーの運命統一(onError 契約)

### CI 網
- [ ] QA1: xlsx smoke
- [ ] QA2: worker-equivalence を CI へ
- [ ] QA3: webkit / firefox smoke
- [ ] QA5: VRT 数値 ratchet
- [ ] QA12: ast-grep ルール追加(unwrap / as-unknown-as / spec § コメント)
- [ ] PD13: dev-wasm base64 再混入ガード

### コンプライアンス / プロダクト即効
- [ ] QA14: MathJax 帰属 + THIRD_PARTY_NOTICES
- [ ] PD1: markdown 公開(toMarkdown API + パッケージ + CLI)
- [ ] PD7: サイト Why ページ(定量比較、出典付き)
- [ ] PD11: STABILITY.md
- [ ] XF13 註: README「EMF not yet rendered」是正

## Phase 2 — チャートエンジンの制覇

- [ ] CH5: チャート XML パース ooxml-common 一本化(enabling step、byte-stable oracle)
- [ ] CH6: 軸モデル完成(gridlines / units / log / orientation / rot / trendline)
- [ ] CH7: 第2軸の全ファミリー化
- [ ] CH8: pie / doughnut 完成
- [ ] CH9: line ファミリーのマーカー / エラーバー / ラベル / smooth
- [ ] CH10: チャートフォント(テーマ + c:txPr)
- [ ] CH12: docx チャート(参照画像承認とセット)
- [ ] CH13: 3D 平坦化 + stock + ofPie
- [ ] CH14: xlsx chartEx 認識
- [ ] CH15: chartEx レンダラー(funnel/histogram → treemap → sunburst → boxWhisker)

## Phase 3 — DrawingML 横断パリティと埋め込みフォント

- [ ] XF2: effectLst 共有化 + docx/xlsx 消費
- [ ] XF3: gradFill / pattFill 消費(xlsx/docx)
- [ ] XF4: fillRef 共有化
- [ ] XF5: rPr 共有化(xlsx/docx 図形テキストの underline/strike/hyperlink 獲得)
- [ ] XF6: gradient scaled / path / fillToRect
- [ ] XF7: 画像 recolor(duotone/grayscl/biLevel/lum、clrChange の docx 実装を共有化)
- [ ] XF10: pptx themeOverride
- [ ] XF13: EMF プレイヤー完成(window/viewport mapping + pptx 配線)
- [ ] XF14: 埋め込みフォント(docx ODTTF + pptx fntdata、node パス含む)
- [ ] XF9: 縦書き段階実装(docx textbox → docx tcPr → pptx 残り)

## Phase 4 — フォーマット固有フィデリティ深化

- [ ] WD2: pgNumType / フィールド書式
- [ ] WD3: w:object OLE + VML imagedata / textpath(watermark)
- [ ] WD4: run spacing / w / position / kern(VRT 承認前提)
- [ ] WD5: w:em 圏点
- [ ] WD6: pgBorders / lnNumType / セクション vAlign
- [ ] WD7: 隣接セル境界競合(§17.4.66)
- [ ] PP1: OLE プレビュー画像
- [ ] PP3: SmartArt フォールバック(S: 枠 → M: テキストリスト)
- [ ] PP4: WordArt prstTxWarp
- [ ] XL3: 数値書式文法完成(+ corpus テーブル)
- [ ] XL4: アウトライングルーピング
- [ ] XL5: ふりがな(rPh / phoneticPr)

## Phase 5 — インタラクション層とアクセシビリティ

- [ ] IX1: ハイパーリンク(3フォーマット、sanitizer 必須、内部アンカー収集込み)
- [ ] IX2: 文書内検索(findText API + ハイライト + next/prev)
- [ ] IX3: キーボード(host focusable + ナビゲーション + xlsx 矢印移動)
- [ ] IX4: a11y テキストレイヤー(PDF.js 型)
- [ ] IX5: alt text(docPr / cNvPr descr)パース(3フォーマット)
- [ ] IX6: worker モードのテキスト選択パリティ(docx / pptx)
- [ ] IX7: worker モードの OMML パリティ
- [ ] IX8: タッチピンチズーム(core/interaction 共有トラッカー)
- [ ] IX9: ズーム契約統一(setScale / fitWidth / fitPage × 4 viewer)+ PD12
- [ ] IX10: onProgress API
- [ ] IX11: エラーステート UI
- [ ] IX12: xlsx Ctrl+C のフォーカスゲート
- [ ] IX13: strings / theme token
- [ ] a11y 自動チェック(axe-core)を smoke へ

## Phase 6 — スケール性能

- [ ] SC1: xlsx ワイヤ正規化(CellValue::Shared + serde skip + colRef 削除)
- [ ] SC2: sheetData ストリーミングパース
- [ ] SC3: 窓配送(sheet_meta + sheet_rows)
- [ ] SC4: per-cell memo(書式 / wrap / font)
- [ ] SC5: チャート bitmap キャッシュ
- [ ] SC6: AxisMetrics 共有 + アンカー O(log n)
- [ ] XL6: C12 ScaledAxis 一本化(正しさ兼用)
- [ ] SC7: スクロール blit
- [ ] SC8: docx per-page 画像 decode
- [ ] SC9: docx incremental pagination(first paint 即時化)
- [ ] SC10: テキスト advance LRU
- [ ] SC11: ページ bitmap LRU(byte budget)
- [ ] SC12: scroll rAF 結合(docx/pptx)
- [ ] SC13: 画像 LRU の byte budget 化
- [ ] SC14: observeDpr
- [ ] SC15: メディア rAF dirty-driven 化
- [ ] SC16: B2-T2 セル段落 stamp
- [ ] SC19: per-sheet rels / drawing の re-parse 解消
- [ ] SC21: (opt-in)worker pool / SharedWorker

## Phase 7 — 堅牢性

- [ ] RB1: 画像ピクセル寸法キャップ(ヘッダ sniff)
- [ ] RB2: 再帰 DepthGuard(grpSp / tbl / OMML)
- [ ] RB5: clampCanvasSize
- [ ] RB6: WASM trap 後の self-poison + auto-respawn
- [ ] RB7: part 名付きエラー + 破損スライド degrade
- [ ] AR5: worker エラーの stack / name 保全
- [ ] RB9: CSP ドキュメント
- [ ] RB10: fontBaseUrl hook
- [ ] QA6: cargo-fuzz(3 parser + handle + rels)、24h クリーン目標
- [ ] QA7: proptest / fast-check

## Phase 8 — プロダクト展開

- [ ] PD2: node 公開 + docx/xlsx サムネイル完成
- [ ] PD3: PDF(S: ラスタ印刷 → L: ベクター記録コンテキスト)
- [ ] PD4: OoxmlError typed errors
- [ ] PD5: @silurus/ooxml-react
- [ ] PD6: Web Component + 静的ホスト
- [ ] PD8: 暗号化対応(Agile Encryption)
- [ ] PD9: fidelity スコアカード
- [ ] PD10: API リファレンス生成化

## 継続トラック

- [ ] AR1: docx parser.rs モジュール分割
- [ ] AR2: docx renderer.ts 物理分割
- [ ] AR3: noUncheckedIndexedAccess burn-down
- [ ] QA4: 差分 VRT(merge-base vs HEAD)
- [ ] QA8: カバレッジ計測
- [ ] QA9: perf 回帰 CI
- [ ] QA10: ground-truth ループ産業化
- [ ] QA11: 公開 conformance corpus
- [ ] QA13: ワイヤ契約 codegen / golden round-trip
