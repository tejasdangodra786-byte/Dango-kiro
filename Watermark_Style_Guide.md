# Tejas Dangodra — Watermark & Copy-Attribution Style Guide

A reusable recipe for my signature look on any HTML study-guide / document. Copy the two snippets below.

The effect has two parts:

1. **A very sparse, subtle watermark** — only about **4 faint marks** reading "Made by Tejas Dangodra"
   spread across the whole page (also shows when printed).
2. **Copying is blocked** — text selection is disabled, right-click / drag / copy-save-view-source
   shortcuts are prevented, and if any text still reaches the clipboard it is replaced with a
   **"Made by Tejas Dangodra — All Rights Reserved"** notice.

> ⚠️ **Honest note:** this is *client-side only*. It marks copied text and prints the watermark, but a
> determined person can still bypass it. For real protection you need a server-rendered document with DRM.

---

## 1 · The watermark — only ~4 marks on the page

### Step A — add this CSS inside your `<style>` block

```css
/* Very sparse, subtle watermark: background-size 50vw 50vh forces a 2x2 grid = ~4 marks. */
.wm-layer{
  position:fixed; inset:0; z-index:9999; pointer-events:none;
  background-repeat:repeat;
  background-size:50vw 50vh;   /* <-- this is what keeps it to ~4 marks, on any screen size */
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='800' height='560'><text x='400' y='290' font-family='Segoe UI,Arial,sans-serif' font-size='13' font-weight='500' letter-spacing='1' fill='rgba(120,100,175,0.04)' text-anchor='middle' transform='rotate(-24 400 290)'>Made by Tejas Dangodra</text></svg>");
}
@media print{
  .wm-layer{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}
```

### Step B — add this div immediately after `<body>`

```html
<div class="wm-layer" aria-hidden="true"></div>
```

It floats behind everything and never blocks clicks (`pointer-events:none`).

---

## 2 · The dials — tune count & visibility

| What you want | Setting to change | Notes |
|---|---|---|
| **Fewer marks (default ~4)** | `background-size:50vw 50vh` | `50vw 50vh` = 2×2 grid = ~4 marks. Use `100vw 100vh` for **1** mark, `33vw 33vh` for ~9 |
| **More marks** | make the `background-size` smaller | e.g. `25vw 25vh` ≈ 16 marks |
| **Fainter / more subtle** | `fill='rgba(120,100,175, 0.04)'` → lower last number | `0.04` = barely visible (my default) · `0.02` = almost invisible · `0.10` = clearly visible |
| **Smaller / lighter text** | `font-size='13'`, `font-weight='500'` | keep small & light |
| **Different tilt** | `transform='rotate(-24 400 290)'` | change `-24`; keep the pivot at half of width/height |
| **Different colour** | the `rgb` part of `fill` | `120,100,175` = soft lavender — match your theme |
| **Different text** | the words `Tejas Dangodra` | keep it short |

> ✅ **My preferred default:** ~**4 marks** (`background-size:50vw 50vh`), opacity **0.04**, font
> **13px / weight 500**, rotation **−24°**, soft lavender, text **"Tejas Dangodra"**.

---

## 3 · Block copying — and replace the clipboard with a rights-reserved notice

### Step A — disable text selection: add to your `body{}` rule

```css
body{ /* ...existing... */ -webkit-user-select:none; -moz-user-select:none; user-select:none; }
```

### Step B — add this script just before `</body>`

```html
<script>
(function(){
  var NOTICE = "Made by Tejas Dangodra — All Rights Reserved.\n"
             + "© Tejas Dangodra. This content is protected. Unauthorised copying is not permitted.";

  // Any copy/cut is replaced with the rights-reserved notice
  ['copy','cut'].forEach(function(evt){
    document.addEventListener(evt, function(e){
      try{ e.clipboardData.setData('text/plain', NOTICE); e.preventDefault(); }catch(err){}
    });
  });

  // Discourage right-click and drag
  ['contextmenu','dragstart'].forEach(function(evt){
    document.addEventListener(evt, function(e){ e.preventDefault(); }, {passive:false});
  });

  // Block copy/cut/save/view-source/print/select-all/devtools shortcuts
  document.addEventListener('keydown', function(e){
    var k=(e.key||'').toLowerCase();
    var block = ((e.ctrlKey||e.metaKey) && ['c','x','s','u','p','a'].indexOf(k)!==-1) ||
                k==='f12' || ((e.ctrlKey||e.metaKey)&&e.shiftKey&&['i','j','c'].indexOf(k)!==-1);
    if(block){ e.preventDefault(); e.stopPropagation(); return false; }
  }, {passive:false});
})();
</script>
```

Edit `NOTICE` to change the message that lands on the clipboard.

---

## 4 · Footer credit (always include)

```html
<footer>
  <div>Your document title — Study Guide</div>
  <div><strong>Made by Tejas Dangodra</strong></div>
</footer>
```

---

## 5 · Quick checklist for any new HTML file

- [ ] `.wm-layer` CSS added, with `background-size:50vw 50vh` (≈4 marks)
- [ ] `<div class="wm-layer" aria-hidden="true"></div>` right after `<body>`
- [ ] Opacity ≈ **0.04** (subtle)
- [ ] `user-select:none` on `body` — copying is blocked
- [ ] Copy-blocking `<script>` (replaces clipboard with the rights-reserved notice) before `</body>`
- [ ] Visible **"Made by Tejas Dangodra"** footer

---

## 6 · See it working

A live demo is in this repo: **`Watermark_Demo.html`** — open it, adjust the sliders, try to select-and-copy
(you'll get the rights-reserved notice), and print-preview to see the ~4 marks.

---

*Reference maintained for Tejas Dangodra. Reuse across all study-guide HTML files.*
