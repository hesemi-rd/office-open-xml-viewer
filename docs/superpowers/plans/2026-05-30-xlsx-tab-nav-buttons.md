# XLSX Sheet Tab Navigation Buttons Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Excel-style ◀▶ navigation buttons to the XLSX viewer's sheet tab bar so tabs that overflow the container width can be scrolled into view without relying on a hidden scrollbar.

**Architecture:** Split the existing single scroll-container tab bar into a fixed left-side nav-button group plus a scrollable tab strip. The buttons scroll the strip (they do not change the active sheet) and disable themselves at the ends / when there is no overflow. The strip keeps `overflow-x:auto` so trackpad / Shift+wheel still work.

**Tech Stack:** TypeScript, DOM API (no framework), Playwright (real-browser interaction test), Vite dev server.

---

## File Structure

- **Modify:** `packages/xlsx/src/viewer.ts` — all production changes live here (DOM construction in the constructor, `buildTabs`, new `makeNavButton` / `navButtonStyle` / `scrollTabs` / `updateNavButtons`, ResizeObserver wiring, new fields).
- **Create:** `packages/xlsx/tests/visual/viewer-fixture.html` — mounts a real `XlsxViewer` in a deliberately narrow (360px) container so the 5 long-named sheets in `demo/sample-1.xlsx` overflow.
- **Create:** `packages/xlsx/tests/visual/tab-nav.spec.ts` — Playwright interaction test driving the buttons.

`demo/sample-1.xlsx` has 5 sheets (`Dashboard`, `Forest Inventory`, `Species Analysis`, `Carbon & Growth`, `Biodiversity Index`). At up to `max-width:160px` per tab these always overflow a 360px container, giving a stable overflow condition without committing a new `.xlsx` fixture (committing xlsx files is forbidden by repo rules).

All paths below are relative to the repo root. Run commands from `packages/xlsx/`.

---

## Task 1: Failing interaction test (red)

**Files:**
- Create: `packages/xlsx/tests/visual/viewer-fixture.html`
- Create: `packages/xlsx/tests/visual/tab-nav.spec.ts`

- [ ] **Step 1: Create the viewer fixture page**

Create `packages/xlsx/tests/visual/viewer-fixture.html`:

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>XLSX Viewer – tab nav fixture</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { height: 100%; background: #fff; }
    /* Narrow host so the 5 sheet tabs in sample-1 overflow horizontally. */
    #host { width: 360px; height: 480px; }
  </style>
</head>
<body>
<div id="host"></div>
<script type="module">
  import { XlsxViewer } from '/src/index.ts';

  const params = new URLSearchParams(location.search);
  const file = params.get('file') ?? 'demo/sample-1.xlsx';
  const host = document.getElementById('host');

  const viewer = new XlsxViewer(host, {
    onReady: () => { document.body.dataset.status = 'ready'; },
    onError: (e) => {
      document.body.dataset.status = 'error';
      document.body.dataset.errorMessage = e instanceof Error ? e.message : String(e);
    },
  });

  await viewer.load('/' + file);
</script>
</body>
</html>
```

- [ ] **Step 2: Write the failing Playwright spec**

Create `packages/xlsx/tests/visual/tab-nav.spec.ts`:

```ts
import { test, expect, type Locator } from '@playwright/test';

test.describe('xlsx sheet tab navigation buttons', () => {
  test('prev/next scroll the tab strip and disable at the ends', async ({ page }) => {
    await page.goto('/tests/visual/viewer-fixture.html');
    await expect(page.locator('body')).toHaveAttribute('data-status', 'ready');

    const prev = page.locator('[data-xlsx-tab-nav="prev"]');
    const next = page.locator('[data-xlsx-tab-nav="next"]');
    await expect(prev).toBeVisible();
    await expect(next).toBeVisible();

    const scrollLeft = () =>
      page.evaluate(() => {
        const strip = document.querySelector('.xlsx-tab-strip');
        return strip ? (strip as HTMLElement).scrollLeft : -1;
      });

    // Disabled is encoded as pointer-events:none on the button's inline style.
    const isDisabled = (loc: Locator) =>
      loc.evaluate((el) => (el as HTMLElement).style.pointerEvents === 'none');

    // Initial state: scrolled to start, prev disabled, next enabled (5 long
    // tabs overflow the 360px host).
    await expect.poll(scrollLeft).toBe(0);
    await expect.poll(() => isDisabled(prev)).toBe(true);
    await expect.poll(() => isDisabled(next)).toBe(false);

    // One next click must increase scrollLeft.
    const before = await scrollLeft();
    await next.click();
    await expect.poll(scrollLeft).toBeGreaterThan(before);

    // Click next until it disables (reached the end).
    for (let i = 0; i < 12; i++) {
      if (await isDisabled(next)) break;
      await next.click();
    }
    await expect.poll(() => isDisabled(next)).toBe(true);
    await expect.poll(() => isDisabled(prev)).toBe(false);

    // Click prev back to the start.
    for (let i = 0; i < 12; i++) {
      if (await isDisabled(prev)) break;
      await prev.click();
    }
    await expect.poll(() => isDisabled(prev)).toBe(true);
    await expect.poll(scrollLeft).toBe(0);
  });
});
```

- [ ] **Step 3: Run the test to verify it fails**

Run: `pnpm --filter @silurus/ooxml-xlsx vrt -- tab-nav.spec.ts`

(Equivalent: from `packages/xlsx/`, `npx playwright test --config playwright.config.ts tab-nav.spec.ts`.)

Expected: FAIL. The locators `[data-xlsx-tab-nav="prev"]` / `"next"` do not exist yet, so `expect(prev).toBeVisible()` times out.

- [ ] **Step 4: Commit the failing test**

```bash
git add packages/xlsx/tests/visual/viewer-fixture.html packages/xlsx/tests/visual/tab-nav.spec.ts
git commit -m "test(xlsx): add failing interaction test for sheet tab nav buttons"
```

---

## Task 2: Restructure the tab bar DOM (nav group + scrollable strip)

**Files:**
- Modify: `packages/xlsx/src/viewer.ts`

- [ ] **Step 1: Add the new private fields**

In the field declaration block (currently around `packages/xlsx/src/viewer.ts:50-51`, after `private tabBar: HTMLDivElement;` and `private tabs: HTMLButtonElement[] = [];`), add three fields. Follow the existing pattern (no `!` definite-assignment — they are assigned in the constructor, like `canvas` / `canvasArea`):

```ts
  private tabBar: HTMLDivElement;
  private tabStrip: HTMLDivElement;
  private navPrev: HTMLButtonElement;
  private navNext: HTMLButtonElement;
  private tabs: HTMLButtonElement[] = [];
```

- [ ] **Step 2: Replace the tab-bar construction block**

In the constructor, replace the current block (`packages/xlsx/src/viewer.ts:100-108`):

```ts
    this.tabBar = document.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:#f0f0f0;border-top:1px solid #c8ccd0;` +
      `overflow-x:auto;overflow-y:hidden;padding:0 4px;gap:1px;scrollbar-width:none;`;
    const style = document.createElement('style');
    style.textContent = `.xlsx-tab-bar::-webkit-scrollbar{display:none}`;
    document.head.appendChild(style);
    this.tabBar.classList.add('xlsx-tab-bar');
```

with this (tabBar no longer scrolls; it holds the nav buttons + a scrollable strip):

```ts
    this.tabBar = document.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:#f0f0f0;border-top:1px solid #c8ccd0;padding:0 4px;gap:2px;`;

    // Excel-style scroll buttons. They scroll the tab strip; they do NOT change
    // the active sheet. Disabled (greyed) at the ends / when there is no overflow.
    this.navPrev = this.makeNavButton('◀', 'Scroll tabs left', () => this.scrollTabs(-1));
    this.navNext = this.makeNavButton('▶', 'Scroll tabs right', () => this.scrollTabs(1));
    this.navPrev.dataset.xlsxTabNav = 'prev';
    this.navNext.dataset.xlsxTabNav = 'next';
    this.navNext.style.marginRight = '4px';

    // The scrollable strip that actually holds the sheet tabs. position:relative
    // so each tab's offsetLeft is measured against the strip's scroll content.
    this.tabStrip = document.createElement('div');
    this.tabStrip.style.cssText =
      `position:relative;display:flex;align-items:flex-end;flex:1;min-width:0;height:100%;` +
      `overflow-x:auto;overflow-y:hidden;gap:1px;scrollbar-width:none;`;
    this.tabStrip.classList.add('xlsx-tab-strip');
    const style = document.createElement('style');
    style.textContent = `.xlsx-tab-strip::-webkit-scrollbar{display:none}`;
    document.head.appendChild(style);
    this.tabStrip.addEventListener('scroll', () => this.updateNavButtons());

    this.tabBar.appendChild(this.navPrev);
    this.tabBar.appendChild(this.navNext);
    this.tabBar.appendChild(this.tabStrip);
```

- [ ] **Step 3: Verify it compiles**

Run: `pnpm --filter @silurus/ooxml-xlsx typecheck`

Expected: FAIL — `makeNavButton`, `scrollTabs`, and `updateNavButtons` are referenced but not yet defined (errors like `Property 'makeNavButton' does not exist`). This confirms the wiring points at the methods added in Task 3. (Do not commit yet — the next task makes it compile.)

---

## Task 3: Implement nav-button helpers, scrolling, and disabled state

**Files:**
- Modify: `packages/xlsx/src/viewer.ts`

- [ ] **Step 1: Point `buildTabs` at the strip and refresh button state**

Replace `buildTabs` (`packages/xlsx/src/viewer.ts:682-694`):

```ts
  private buildTabs(): void {
    this.tabBar.innerHTML = '';
    this.tabs = [];
    this.wb.sheetNames.forEach((name, i) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabStyle(false);
      btn.addEventListener('click', () => this.showSheet(i));
      this.tabBar.appendChild(btn);
      this.tabs.push(btn);
    });
  }
```

with (append to `tabStrip`, not `tabBar`, and refresh nav state after building):

```ts
  private buildTabs(): void {
    this.tabStrip.innerHTML = '';
    this.tabs = [];
    this.wb.sheetNames.forEach((name, i) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabStyle(false);
      btn.addEventListener('click', () => this.showSheet(i));
      this.tabStrip.appendChild(btn);
      this.tabs.push(btn);
    });
    this.updateNavButtons();
  }
```

- [ ] **Step 2: Add the nav-button helpers and scroll/state logic**

Insert these four methods immediately after `buildTabs` (before `updateTabActive` at `packages/xlsx/src/viewer.ts:696`):

```ts
  private makeNavButton(glyph: string, label: string, onClick: () => void): HTMLButtonElement {
    const btn = document.createElement('button');
    btn.textContent = glyph;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.style.cssText = this.navButtonStyle(false);
    btn.addEventListener('click', onClick);
    return btn;
  }

  private navButtonStyle(disabled: boolean): string {
    // Sit a touch lower than the active tab (matches inactive tab height) and
    // share the tab palette so the buttons read as part of the tab strip.
    const base =
      `flex-shrink:0;width:22px;height:${TAB_BAR_H - 5}px;` +
      `display:flex;align-items:center;justify-content:center;` +
      `border:1px solid #c8ccd0;border-bottom:none;border-radius:3px 3px 0 0;` +
      `background:#e0e0e0;color:#555;font-size:9px;line-height:1;` +
      `box-sizing:border-box;outline:none;`;
    return disabled
      ? base + `opacity:0.35;cursor:default;pointer-events:none;`
      : base + `cursor:pointer;`;
  }

  private scrollTabs(dir: -1 | 1): void {
    const strip = this.tabStrip;
    const viewLeft = strip.scrollLeft;
    const viewRight = viewLeft + strip.clientWidth;
    let target: number | null = null;
    if (dir === 1) {
      // First tab clipped on the right; align its right edge to the viewport.
      for (const tab of this.tabs) {
        const right = tab.offsetLeft + tab.offsetWidth;
        if (right > viewRight + 1) {
          target = right - strip.clientWidth;
          break;
        }
      }
    } else {
      // Last tab clipped on the left; align its left edge to the viewport.
      for (let i = this.tabs.length - 1; i >= 0; i--) {
        const left = this.tabs[i].offsetLeft;
        if (left < viewLeft - 1) {
          target = left;
          break;
        }
      }
    }
    if (target !== null) {
      // Instant (not smooth) so the disabled state is consistent the moment the
      // click resolves — keeps the interaction deterministic to drive/test.
      strip.scrollLeft = Math.max(0, target);
    }
    this.updateNavButtons();
  }

  private updateNavButtons(): void {
    const strip = this.tabStrip;
    const atStart = strip.scrollLeft <= 0;
    const atEnd = strip.scrollLeft + strip.clientWidth >= strip.scrollWidth - 1;
    // No overflow => scrollWidth ≈ clientWidth => both ends true => both disabled.
    this.navPrev.style.cssText = this.navButtonStyle(atStart);
    this.navNext.style.cssText = this.navButtonStyle(atEnd);
  }
```

- [ ] **Step 3: Refresh nav state on resize**

The constructor already has a `ResizeObserver` observing `canvasArea` (currently `packages/xlsx/src/viewer.ts:120-124`):

```ts
    this.resizeObserver = new ResizeObserver(() => {
      this.renderCurrentSheet();
      this.updateSelectionOverlay();
    });
    this.resizeObserver.observe(this.canvasArea);
```

Add `this.updateNavButtons();` to the callback (the tab bar resizes together with `canvasArea`, so overflow state can change):

```ts
    this.resizeObserver = new ResizeObserver(() => {
      this.renderCurrentSheet();
      this.updateSelectionOverlay();
      this.updateNavButtons();
    });
    this.resizeObserver.observe(this.canvasArea);
```

- [ ] **Step 4: Verify it compiles**

Run: `pnpm --filter @silurus/ooxml-xlsx typecheck`

Expected: PASS (no errors).

---

## Task 4: Green + commit

**Files:** none (verification only)

- [ ] **Step 1: Run the interaction test**

Run: `pnpm --filter @silurus/ooxml-xlsx vrt -- tab-nav.spec.ts`

Expected: PASS (1 passed).

- [ ] **Step 2: Run the full xlsx VRT to confirm no regression**

First rebuild the parser per repo rule, then run the full suite:

Run: `pnpm --filter @silurus/ooxml-xlsx vrt`

Expected: PASS. The existing visual specs render via `XlsxWorkbook.renderViewport` straight to a canvas and never mount `XlsxViewer`, so the tab-bar change cannot affect the reference-image comparisons. (If `demo/sample-1.xlsx` is missing locally, the existing visual tests are skipped/fail for that reason, not because of this change — the new `tab-nav` test still runs against the same sample.)

- [ ] **Step 3: Commit**

```bash
git add packages/xlsx/src/viewer.ts
git commit -m "feat(xlsx): add Excel-style prev/next buttons to sheet tab bar

The tab strip hides its scrollbar, leaving plain-mouse users with no way to
reach overflowing sheet tabs. Add fixed ◀▶ buttons at the left of the tab bar
that scroll the strip one clipped tab per click and disable at the ends / when
there is no overflow. overflow-x:auto stays so trackpad / Shift+wheel still work."
```

---

## Self-Review Notes

- **Spec coverage:** left ◀▶ always-visible buttons (Task 2), scroll-the-strip semantics not active-sheet change (Task 3 `scrollTabs`), disabled at ends + no-overflow (Task 3 `updateNavButtons`), DOM split into navGroup + tabStrip (Task 2), `buildTabs` retargeted (Task 3), state hooks on scroll/resize/buildTabs (Tasks 2–3), kept `overflow-x:auto` + hidden scrollbar (Task 2), Playwright interaction test, VRT references untouched (Tasks 1 & 4). All covered.
- **Type consistency:** `tabStrip` / `navPrev` / `navNext` declared in Task 2 and used in Tasks 2–3; method names `makeNavButton` / `navButtonStyle` / `scrollTabs` / `updateNavButtons` consistent across wiring (Task 2) and definitions (Task 3); `.xlsx-tab-strip` class used in both product code and the test selector.
- **Deviation from spec (deliberate):** spec mentioned `scrollTo({ behavior:'smooth' })`; the plan uses an instant `scrollLeft` assignment so the disabled state is settled when the click resolves, keeping the test deterministic without arbitrary waits. UX impact is negligible for one-tab steps.
