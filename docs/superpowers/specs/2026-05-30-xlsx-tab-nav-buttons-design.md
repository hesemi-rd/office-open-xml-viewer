# XLSX シートタブ送りボタン 設計

- 日付: 2026-05-30
- パッケージ: `packages/xlsx`
- 対象ファイル: `packages/xlsx/src/viewer.ts`

## 背景・課題

シートタブ帯は `overflow-x:auto` でスクロール可能だが、スクロールバーを
CSS で非表示にしている（`scrollbar-width:none` + `::-webkit-scrollbar{display:none}`）。
そのため:

- タブがコンテナ幅に入り切らなくても、スクロールできることが視覚的に伝わらない。
- プレーンなマウス（縦ホイールのみ）ユーザーには事実上スクロール手段がない。
  トラックパッド横スワイプ / Shift+ホイール / Magic Mouse 横スワイプを知らないと操作不能。

Excel 同様に明示的な送りボタンを設けて、この affordance 欠如を解消する。

## 方針

- タブ帯の**左端に ◀▶ の2ボタンを常時表示**する（Excel 準拠）。
- ボタンはタブ帯（スクロール領域）を**スクロールさせるだけ**で、アクティブシートは変更しない。
- オーバーフローしていない、またはスクロール端に達したら該当ボタンを **disabled（グレーアウト）** にする。
- 既存の `overflow-x:auto` は残し、トラックパッド横スワイプ / Shift+ホイールも併用可能なままにする。
- スクロールバー非表示は維持（ボタンが affordance を担うため）。

## DOM 構造の変更

現状は `this.tabBar` 自体がスクロールコンテナ（`overflow-x:auto`）で、タブボタンを
直接子に持つ。これを2層に分割する:

```
tabBar (flex, align-items:flex-end, height:TAB_BAR_H, flex-shrink:0, スクロールしない)
├── navGroup (flex-shrink:0)        ◀ ▶  ← 固定。左端
└── tabStrip (flex:1, overflow-x:auto, overflow-y:hidden, scrollbar 非表示)
        └── tab buttons ...         ← buildTabs() の append 先をここに変更
```

- `overflow-x:auto` 等のスクロール関連スタイルは `tabBar` から `tabStrip` へ移す。
- `tabBar` 自体はスクロールせず、navGroup と tabStrip を横並びにするだけ。
- `.xlsx-tab-bar::-webkit-scrollbar{display:none}` のセレクタは tabStrip に付けたクラスへ合わせる。

## コンポーネント / メソッド

### フィールド
- `private tabStrip: HTMLDivElement` — スクロール領域。タブボタンの親。
- `private navPrev: HTMLButtonElement` — ◀
- `private navNext: HTMLButtonElement` — ▶

### `buildTabs()`（変更）
タブボタンの append 先を `this.tabBar` → `this.tabStrip` に変更。生成後に
`this.updateNavButtons()` を呼ぶ。

### `updateTabActive(index)`（変更なしで動作）
`scrollIntoView` の対象タブは `tabStrip` の子のままなので従来どおり機能する。
スクロール後に `updateNavButtons()` が `scroll` イベント経由で呼ばれる。

### `scrollTabs(dir: -1 | 1)`（新設）
クリック方向で「見切れている次のタブ」が見える位置までスクロールする
（1タブずつ送るフィール）:

- `dir === -1`（◀）: tabStrip の左端より左に隠れているタブのうち最も右のものを探し、
  その左端が見えるよう `scrollTo({ left, behavior: 'smooth' })`。
- `dir === 1`（▶）: 右端より右に隠れているタブのうち最も左のものを探し、
  その右端が見えるよう `scrollTo`。
- 各タブの `offsetLeft` / `offsetWidth` と tabStrip の `scrollLeft` / `clientWidth` から境界を算出。

### `updateNavButtons()`（新設）
disabled 状態を再計算してボタンに反映:

- ◀ disabled: `scrollLeft <= 0`
- ▶ disabled: `scrollLeft + clientWidth >= scrollWidth - ε`（ε は 1px 程度）
- disabled 時は `opacity` を下げ、`pointer-events:none`、`cursor:default`。

### 呼び出し箇所
`updateNavButtons()` を以下から呼ぶ:

- `tabStrip` の `scroll` イベントリスナー
- 既存の `ResizeObserver`（canvasArea を観測中。タブ帯幅も連動して変わる）
- `buildTabs()` 直後
- ボタンクリック（`scrollTabs` 内のスクロール後 / scroll イベントでも更新される）

## スタイル

- ボタンはタブ帯の色調（`#f0f0f0` 系）に馴染む小型の矢印ボタン。
- 高さは `TAB_BAR_H` に揃え、`align-items` 基準で他タブと整合させる。
- 矢印グリフは `◀` `▶`（または `‹` `›`）。`font-size` で小さめに。
- `flex-shrink:0` で潰れない。`white-space:nowrap`。

## テスト

ロジックは `scrollWidth` / `clientWidth` 等のレイアウト計算に依存し、jsdom では
0 になり検証できない。よって **Playwright で実ブラウザのインタラクションテスト**を追加する:

1. 多数シート（横幅を確実に超える数）の workbook を読み込む。
2. 初期状態: ◀ は disabled、▶ は enabled。
3. ▶ をクリック → `tabStrip.scrollLeft` が増加することを確認。
4. 末尾まで送ると ▶ が disabled、◀ が enabled になることを確認。
5. ◀ で戻れることを確認。

VRT 参照画像（`tests/visual/references/`、Excel export PNG）は対象外。送りボタンは
viewer の chrome であり Excel export には存在しないため、参照画像は更新しない。

## 非対象 (YAGNI)

- ボタン長押しでの連続スクロール（Excel にはあるが MVP では見送り）。
- 先頭/末尾ジャンプ（⏮⏭）。
- 右クリックでのシート一覧ポップアップ。
