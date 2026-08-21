# Tejas Dangodra — Watermark & Copy-Protection Style Guide

A reusable recipe for adding **my signature subtle watermark** plus light copy-protection to any HTML
study-guide / document. Copy the two snippets below into any HTML file.

The effect has two parts:

1. **A subtle diagonal watermark** tiled faintly across the whole page (also shows when printed).
2. **Light copy-protection** — copied text carries my credit, and casual copy/save/right-click is discouraged.

> ⚠️ **Honest note:** this is *client-side only*. It deters casual copying and makes my name travel with any
> copied text, but a determined person can still bypass it (e.g. via browser dev tools or by disabling
> JavaScript). For real protection you need a server-rendered document with proper DRM.

---

## 1 · The watermark (CSS + one div)

### Step A — add this CSS inside your `<style>` block

```css
/* ===== Subtle watermark ===== */
.wm-layer{
  position:fixed; inset:0; z-index:9999; pointer-events:none;
  background-repeat:repeat;
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='560' height='360'><text x='280' y='190' font-family='Segoe UI,Arial,sans-serif' font-size='14' font-weight='500' letter-spacing='1' fill='rgba(120,100,175,0.035)' text-anchor='middle' transform='rotate(-24 280 190)'>Tejas Dangodra</text></svg>");
}
@media print{
  .wm-layer{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}
```

### Step B — add this div immediately after `<body>`

```html
<div class="wm-layer" aria-hidden="true"></div>
```

That's it — the watermark now floats behind everything (it never blocks clicks because of
`pointer-events:none`).

---

## 2 · The dials — how to make it MORE or LESS visible

Everything is controlled inside the inline `<svg>` in the CSS above. Change these values:

| What you want | Setting to change | Notes |
|---|---|---|
| **Fainter / more subtle** | `fill='rgba(120,100,175, 0.035)'` → lower the last number | `0.035` = barely visible (my default) · `0.10` = clearly visible · `0.02` = almost invisible |
| **Bolder / more evident** | raise that same opacity number | e.g. `0.08`–`0.12` |
| **Fewer, more spread-out marks** | increase SVG `width` & `height` | `560 x 360` = sparse (my default). Bigger tile = fewer repeats |
| **Denser, tighter marks** | decrease SVG `width` & `height` | e.g. `340 x 200` |
| **Bigger / smaller text** | `font-size='14'` and `font-weight='500'` | keep it small & light for subtlety |
| **Different tilt** | `transform='rotate(-24 280 190)'` | change `-24` to any angle (the `280 190` is the pivot = half of width/height) |
| **Different colour** | the `rgb` part of `fill` | `120,100,175` = soft lavender. Match your theme |
| **Different text** | the words `Tejas Dangodra` | keep short so tiles stay clean |

> ✅ **My preferred "subtle" default:** opacity **0.035**, font **14px / weight 500**, tile **560×360**,
> rotation **−24°**, colour soft lavender `rgba(120,100,175,…)`, text **"Tejas Dangodra"**.

---

## 3 · Copy-protection (optional but recommended)

### Step A — add `user-select:none;` to your `body{}` rule

```css
body{ /* ...your existing styles... */ -webkit-user-select:none; -moz-user-select:none; user-select:none; }
```

### Step B — add this script just before `</body>`

```html
<script>
(function(){
  var CREDIT = "\n\n— © Made by Tejas Dangodra";

  // Append my credit to anything copied
  document.addEventListener('copy', function(e){
    var sel = (window.getSelection && window.getSelection().toString()) || '';
    if(sel.length){ e.clipboardData.setData('text/plain', sel + CREDIT); e.preventDefault(); }
  });

  // Discourage cut / right-click / drag
  ['contextmenu','dragstart','cut'].forEach(function(evt){
    document.addEventListener(evt, function(e){ e.preventDefault(); }, {passive:false});
  });

  // Discourage save / view-source / copy / print / devtools shortcuts
  document.addEventListener('keydown', function(e){
    var k = (e.key || '').toLowerCase();
    var block = ((e.ctrlKey || e.metaKey) && ['s','u','c','x','p'].indexOf(k) !== -1) ||
                k === 'f12' ||
                ((e.ctrlKey || e.metaKey) && e.shiftKey && ['i','j','c'].indexOf(k) !== -1);
    if(block){ e.preventDefault(); e.stopPropagation(); return false; }
  }, {passive:false});
})();
</script>
```

---

## 4 · Footer credit (always include)

Alongside the watermark, keep a **visible footer** on the page:

```html
<footer>
  <div>Your document title — Study Guide</div>
  <div><strong>Made by Tejas Dangodra</strong></div>
</footer>
```

---

## 5 · Quick checklist for any new HTML file

- [ ] `.wm-layer` CSS added to `<style>`
- [ ] `<div class="wm-layer" aria-hidden="true"></div>` right after `<body>`
- [ ] Opacity set to **0.035** (subtle) — tweak with the dials table if needed
- [ ] `user-select:none` on `body`
- [ ] Copy-protection `<script>` before `</body>`
- [ ] Visible **"Made by Tejas Dangodra"** footer

---

## 6 · See it working

A live, self-contained demo is in this repo: **`Watermark_Demo.html`** — open it in a browser, try to
select/copy text, and print-preview it to see the watermark behaviour.

---

*Reference maintained for Tejas Dangodra. Reuse this recipe across all study-guide HTML files.*
