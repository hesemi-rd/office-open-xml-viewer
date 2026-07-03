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
- [x] XL1: date1904(+1900-02-29 バグ互換) — serial→date 変換を core `excel-date.ts` に集約(逆変換 `utcDateToExcelSerial` 込み、serial 60 は非生成)。**追加発見: 重複は 2 でなく 3 実装だった**(formula.ts の YEAR/MONTH/DAY/WEEKDAY 系も 1900 バグ未補正 → 委譲で解消)。**横断拡張: chart `c:date1904`(§21.2.2.38)も同じ穴** → ooxml-common `extract_chart_date1904` + xlsx/pptx 両パーサー + core renderer 全 19 箇所へ伝搬(chartEx は cx: スキーマに c:date1904 が無いため対象外)。chart spec の epoch 文言(Dec 31 基準)はセル定義 §18.17.4.1 と 1 日矛盾するため後者に統一(文言通りだと全日付が 1 日ズレることを数値検証)。TODAY/NOW volatile は「エンジンが 1900 系 serial を生成 → 1900 系で書式化」で自己一貫(formula エンジン自体の date1904 対応は別トラック)
- [x] XL2: General 書式 11 有効桁 — `formatGeneralNumber`(toPrecision(11) + trailing zero 除去 + 12 桁以上/1e-6 未満は Excel 式指数表記 E+NN)。小さい数の切替閾値(< 1e-5)は Microsoft 未文書だが Excel 実挙動(1E-06)と一致する一貫則としてコメント明示(spec レビューで許容判定)。列幅依存の桁数調整はスコープ外と明記。チャートの `formatChartVal`(toFixed(6))は同じ穴でないことを確認し不変

> XL1/XL2 は PR fix/xlsx-date-general(9 commits)。検証 = vitest 1806 / typecheck / cargo 3 点(304 tests)/ VRT 163 全 green・参照画像差分ゼロ。独立 4 観点レビュー → REQUEST_CHANGES 2 件(formula.ts の第 3 重複、TODAY/NOW×date1904 テスト欠落)+ LOW 群を全て同ブランチで解消。別トラック送り: formula エンジンへの date1904 伝搬(1904 ブックの TODAY() が 1904 系 serial を返すべき — 現設計は表示レベルで自己一貫)、mcp-server の cell_display は生値抽出面なので General 丸め対象外と判定
- [ ] XF8: scrgbClr / hslClr / prstClr transforms
- [ ] XF11: docx 下線 17 種(pptx drawUnderline の core hoist 込み)
- [ ] XF12: cNvPr@hidden(3フォーマット)
- [ ] WD1: noBreakHyphen / cr / ptab / softHyphen
- [ ] WD8: kashida / thaiDistribute → both マップ(+真 kashida issue 起票)
- [ ] WD9: OMML phant / sPre / box / borderBox
- [ ] PP2: SmartArt relId バインディング
- [x] XF1: **OOXML Strict 名前空間**(共通述語 + 3パーサー + Strict fixture) — **完了**(branch fix/strict-namespaces)。Office の「Strict 形式で保存」は Transitional(`http://schemas.openxmlformats.org/...`)でなく Strict(`http://purl.oclc.org/ooxml/...`)URI を使い、現パーサーは Transitional URI のみ照合するため 3 フォーマットとも真っ白になっていた。**Strict URI の確定は ECMA-376 5th ed. の Strict XML Schema `targetNamespace`(Part 1 `OfficeOpenXML-XMLSchema-Strict/*.xsd`)+ RELAX-NG Strict の 2 源で照合**(推測なし)。`ooxml_common::ns` に両クラスの `TRANSITIONAL`/`STRICT` 定数と `is_*_ns(Option<&str>) -> bool` 述語群(w/a/c/p/x/xdr/wp/pic/m/r)+ 属性 2-URI lookup `attr_ns` を集約。**実照合箇所は当初見積 154 → 実数約 131**(docx: xml_util ヘルパー集約 + parser/numbering の r:id・m:val・font-table、pptx: `attr_r` 1 箇所のみ〔他は local-name 照合で Strict 免疫〕、xlsx: chart.rs 82 + lib/styles/slicer/table/drawing、ooxml-common: blip `r:embed`・math `m:val`)。**触らなかったもの**: local-name 照合(方式 a)は Strict でも動くので不変、共有 `ooxml_common::rels` は Id→Target のみで Type URI を読まないので対象外だが、pptx ローカルの `find_rel_target_by_type` は Type を suffix 比較(`ends_with`)で読んでおり Strict の purl プレフィクスでも一致する(完全一致に変えると Strict で壊れる旨コメント注記済み)、メインパートパス(`word/document.xml` 等)はハードコードで Strict 共通、OPC package-relationships ns(`.../package/2006/relationships`)は ISO-29500 Part 2 で Strict/Transitional 不変ゆえ変更せず、MCE ns(markup-compatibility)は単一系統で対象外。Transitional 既存テスト全 green(回帰 oracle)を維持したまま、docx/pptx/ooxml-common に Strict fixture(合成)を追加(段落+run 書式/hyperlink r:id 解決/OMML m:val/shape+pic r:embed/blip・math)。実 Office の Strict 保存ファイルによる end-to-end 検証は QA11(公開 conformance corpus)の別トラック。

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
- ~~PD7: サイト Why ページ(定量比較、出典付き)~~ — **見送り(2026-07-03 ユーザー決定)**: ニーズが出ていない段階でのマーケ施策は行わない
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
- [ ] WD3: w:object OLE + VML imagedata / textpath(watermark) — **OLE 部は前倒しで完了**(PR feature/ole-preview、ユーザー指示 2026-07-03): `w:object` → `v:imagedata` を既存 ImageRun へ(寸法 = VML CSS style pt、fallback dxaOrig/dyaOrig ÷20)。残 = 素の `w:pict` 内 v:imagedata(非 OLE inline VML 画像)/ v:textpath watermark / CT_Object 第一子の modern `w:drawing` 委譲(§17.3.3.19)
- [ ] WD4: run spacing / w / position / kern(VRT 承認前提)
- [ ] WD5: w:em 圏点
- [ ] WD6: pgBorders / lnNumType / セクション vAlign
- [ ] WD7: 隣接セル境界競合(§17.4.66)
- [x] PP1: OLE プレビュー画像 — **前倒しで完了**(PR feature/ole-preview、ユーザー指示 2026-07-03)。graphicFrame の ole URI → oleObj の preview `p:pic` を parse_picture 再利用で emit(graphicFrame xfrm が権威 §19.3.2.4)。**レビュー CRITICAL 検出**: PowerPoint 正規出力は Choice=spid(pic と排他、Part 3 §B.1)/Fallback=pic なので「pic を持つ oleObj を選ぶ」能力述語に修正(MCE §9.3)。README の Not Planned 撤回。EMF/WMF プレビューの描画品質は共有メタファイルデコーダーの限界に従う(XF13 は pptx 固有でなく 3 フォーマット共通のデコーダー天井と再定義)。**xlsx 横断**: `<oleObjects>`(§18.3.1.60)パース + image MIME ガードまで実装。ただし Excel 実出力のプレビューは vmlDrawing パート(`oleObject@shapeId` ↔ `v:shape@id`)にあり、objectPr@r:id はデータ part(MS-OI29500 確認、当初実装の前提はレビューが反証)— **vmlDrawing 経路は PR feature/xlsx-ole-vml-preview で完結**(`<legacyDrawing r:id>` → vmlDrawing の `<v:imagedata o:relid>` を shapeId で解決、image-typed `objectPr` があれば優先、`<x:Anchor>` px→EMU ×9525 フォールバック。README を ✅ に反転)
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

## Phase 8 — プロダクト展開 【全項目見送り(2026-07-03 ユーザー決定)】

> ニーズが実証されていない段階でのプロダクト展開・エコシステム拡張は行わない。ユーザー需要が出た時点で個別に再検討する(PD7 も同判断で見送り)。Phase 1 の PD1(markdown 公開 — 既に全ユーザーの WASM に同梱済み機能の露出)と PD11(STABILITY.md — 運用ドキュメント)は対象外で存続。

- ~~PD2: node 公開 + docx/xlsx サムネイル完成~~
- ~~PD3: PDF(S: ラスタ印刷 → L: ベクター記録コンテキスト)~~
- ~~PD4: OoxmlError typed errors~~
- ~~PD5: @silurus/ooxml-react~~
- ~~PD6: Web Component + 静的ホスト~~
- ~~PD8: 暗号化対応(Agile Encryption)~~
- ~~PD9: fidelity スコアカード~~
- ~~PD10: API リファレンス生成化~~

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
