# Auto Fit Page — Design Spec
**Date:** 2026-04-11  
**Project:** PPT Viewer (Electron)

---

## Overview

Add a persistent **Auto Fit Page** toggle that scales every slide so the entire slide (width *and* height) fits inside the viewport with no scrolling required. When the toggle is off, the viewer reverts to the existing fit-to-width behavior.

---

## Feature Behavior

### Toggle ON
- Fit scale is computed as `min(availableWidth / slideWidth, availableHeight / slideHeight)`.
- `availableWidth = mainView.clientWidth - 40` (20 px padding each side, same as today).
- `availableHeight = mainView.clientHeight - 20` (10 px padding top/bottom).
- This scale replaces the current `fitScale` used in `setupContinuousView`, `refreshView`, and `renderVisiblePages`.
- Every page in the continuous scroll view uses the same fit-page scale, so navigating pages requires no manual re-zoom.
- On window resize, `refreshView()` already fires — it will recalculate the fit-page scale automatically.

### Toggle OFF
- Reverts to normal: `fitScale = availableWidth / slideWidth` (existing behavior, unchanged).

### Interaction with manual zoom
- Pressing zoom in / zoom out (Ctrl+`+`, Ctrl+`-`, scroll wheel + Ctrl) while Auto Fit is ON automatically turns Auto Fit **OFF**, then applies the zoom increment. This matches the behavior of mainstream PDF viewers (Adobe, browser PDF viewer).
- The zoom-reset button ("Fit Width") turns Auto Fit **OFF** and resets `zoomLevel` to `1`.

---

## Persistence

- The toggle state is stored in `localStorage` under the key `pptviewer-auto-fit`.
- Loaded once on startup; applied immediately when the first file is opened.

---

## UI

### Toolbar button
- Placed between the **zoom controls** and the **rotate controls** (after the existing zoom-reset button, before the separator before rotate).
- Label: **"Fit Page"**
- Icon: an inward-arrow / fit-to-screen SVG (4 corner arrows pointing inward).
- Visually active (accent color highlight) when toggle is ON — same treatment as the "Dark Read" button.
- `title` attribute: `Auto Fit Page (F)`.

### Hamburger menu
- New entry under the **View** section, below "Document Dark Mode".
- Label: **"Auto Fit Page"**
- Shows a checkmark `✓` prefix when active.
- Keyboard shortcut hint: `F`.

### Keyboard shortcut
- Key: `F` (matches the single-key pattern already used: `I` for dark mode).
- Fires `toggleAutoFit()`.

---

## Code Changes

### `renderer/app.js`

1. **New state variable:**
   ```js
   var autoFitPage = localStorage.getItem('pptviewer-auto-fit') === 'true';
   ```

2. **New helper — `getFitScale()`:**
   Centralises the fit-scale calculation used in `setupContinuousView`, `refreshView`, and `renderVisiblePages`:
   ```js
   function getFitScale() {
     var mainView = document.getElementById('main-view');
     var availW = mainView.clientWidth - 40;
     if (!autoFitPage) return availW / pageBaseDims.width;
     var availH = mainView.clientHeight - 20;
     return Math.min(availW / pageBaseDims.width, availH / pageBaseDims.height);
   }
   ```
   Replace the three inline `fitScale = maxWidth / pageBaseDims.width` expressions with calls to `getFitScale()`.

3. **New function — `toggleAutoFit()`:**
   ```js
   function toggleAutoFit() {
     autoFitPage = !autoFitPage;
     localStorage.setItem('pptviewer-auto-fit', autoFitPage);
     updateAutoFitButton();
     refreshView();
   }
   ```

4. **New function — `updateAutoFitButton()`:**
   Sets the active/inactive visual state on the toolbar button and the checkmark in the hamburger menu item.

5. **Zoom interaction:** In `setZoom()`, add `autoFitPage = false; updateAutoFitButton();` before applying the zoom level.

6. **Event wiring:**
   - Toolbar button click → `toggleAutoFit()`.
   - Hamburger menu item click → `toggleAutoFit()`.
   - Keyboard `F` key in the existing `keydown` handler → `toggleAutoFit()`.

### `renderer/index.html`

1. Add toolbar button (between zoom-reset and the separator before rotate):
   ```html
   <button id="toolbar-auto-fit" class="toolbar-btn" title="Auto Fit Page (F)">
     <!-- fit-to-screen SVG -->
     Auto Fit
   </button>
   ```

2. Add hamburger menu item (below "Document Dark Mode" entry):
   ```html
   <button class="hmenu-item" id="hmenu-auto-fit">
     <span class="hmenu-icon"><!-- icon --></span>
     Auto Fit Page
     <span class="hmenu-kbd">F</span>
   </button>
   ```

### `renderer/style.css`

- No new rules needed. The active state reuses the `.active` or inline style pattern already used by the "Dark Read" button. Confirm exact pattern in `updateDocDarkButton()` and mirror it.

---

## Out of Scope

- Per-tab auto-fit state (toggle is global, applies to all tabs).
- Saving auto-fit state per file.
- Changes to the thumbnail sidebar scale.

---

## Success Criteria

1. With Auto Fit ON, opening any PPT/PDF shows each slide fully visible without vertical scrolling.
2. Resizing the window keeps the slide fully visible.
3. Manual zoom turns Auto Fit OFF and the button/menu item updates immediately.
4. The toggle state persists after closing and reopening the app.
5. The `F` keyboard shortcut toggles the mode.
6. The hamburger menu checkmark reflects the current state.
