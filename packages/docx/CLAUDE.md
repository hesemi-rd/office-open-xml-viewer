# @silurus/ooxml-docx — DOCX セッション用

## あなたのロール
このディレクトリ（`packages/docx/`）の編集責任者。DOCX の parser・renderer を 0 から実装する。

## MVP スコープ（Phase 2）
- ZIP 解凍 + `word/document.xml` パース（Rust）
- パラグラフ + ラン（bold/italic/color/font-size）
- 明示改ページ `<w:br w:type="page"/>` のみ
- `DocxDocument.renderPage(target, pageIndex)` 公開 API
- 非対応（後続）: テーブル、画像、ヘッダー/フッター、自動改ページ、脚注

## Storybook

Storybook はルート一本化のため、パッケージ単体では起動しない。
ルートから `pnpm storybook` で全パッケージのストーリーが参照できる。

## 編集してよいもの
- `packages/docx/**` すべて
- `packages/docx/parser/src/**`（Rust）

## 絶対に編集してはいけないもの
- `packages/core/**` ← 共有コード。変更が必要なら main ブランチへ PR
- `packages/pptx/**` / `packages/xlsx/**` ← 他セッションの領域
- root の config 類 ← main で管理

## 参考にしてよいもの（読み取り専用）
- `src/renderer.ts` — Canvas 描画プリミティブの参考
- `pptx-parser/src/lib.rs` — wasm_bindgen + roxmltree の使い方
- `src/types.ts` — 型定義パターンの参考

## 参照画像
`packages/docx/tests/visual/references/` は Word export PNG のみ配置。自動更新禁止。

## ディレクトリ構成

```
packages/docx/
├── CLAUDE.md
├── package.json
├── tsconfig.json
├── vite.config.ts
├── public/              ← sample .docx ファイル置き場（git 管理対象外）
├── src/
│   ├── index.ts         ← 公開 API
│   ├── types.ts         ← Document モデル型（Rust JSON 出力と1:1）
│   ├── document.ts      ← DocxDocument クラス
│   ├── viewer.ts        ← DocxViewer クラス
│   ├── renderer.ts      ← renderPage 実装
│   └── wasm/            ← wasm-pack 出力（git 管理対象外）
├── parser/
│   ├── Cargo.toml
│   └── src/
│       ├── lib.rs       ← WASM エントリポイント
│       ├── types.rs     ← Rust 型定義
│       └── parser.rs    ← DOCX パーサー実装
└── tests/visual/
    ├── visual.spec.ts
    ├── fixture.html
    ├── references/      ← Word export PNG（git 管理対象外）
    ├── screenshots/
    └── diffs/
```

## WASM ビルド手順

```bash
cd packages/docx && npm run wasm
# または
cd packages/docx/parser && wasm-pack build --target web --out-dir ../src/wasm
```

## テスト実行

```bash
npx playwright test packages/docx/tests/visual
```

## ヒューリスティック禁止 (root CLAUDE.md と整合)

DOCX 描画では特に anchor / wgp / docGrid / ruby 周りでヒューリスティックを入れたくなる場面が多い。以下を遵守すること:

- **anchor 位置**: `positionH/V` の `relativeFrom` と `posOffset` / `align` / `pctPos` の意味は ECMA-376 §20.4.3.x で完全に定義されている。「sample-N に合わせるための定数 / 特殊分岐」は禁止。
- **wgp グループ変換**: `lIns/tIns/rIns/bIns` は page-space EMU（グループ変換後の寸法で測る）。`anchorXPt/anchorYPt` は子座標系の offset に sx/sy を掛けて page-relative にする。`widthPt/heightPt` も cx/cy × sx/sy。これら全てが一貫したスケーリングを受けていないと位置ズレが積み重なる。
- **フォント分類**: 名前パターンマッチではなく `word/fontTable.xml` の `<w:family w:val=…/>` (§17.8.3.10) を使う。パターンマッチはフォントテーブルにないケースへのフォールバックとしてのみ残す。
- **改ページヒント**: `<w:lastRenderedPageBreak/>` (§17.3.1.20) は Word のレイアウトキャッシュ。「自前計算が壊れている特定パスでだけ信用する」のは禁止。自前計算を直すか、常に無視するかのいずれかを選ぶ。混在させない。
- **仕様デフォルト値**: 属性省略時の `unwrap_or(0.0)` は多くの場合誤り。§21.1.2.1.1 の lIns=91440 EMU のように ECMA-376 にデフォルト値が定義されているものは必ずその値を使う。
