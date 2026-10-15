# Tejas Dangodra — Watermark & Copy-Attribution Style Guide

A reusable recipe for my signature look on any HTML study-guide / document. Copy the two snippets below.

The effect has two parts:

1. **A clearly visible tiled watermark** — soft diagonal marks reading "Made by Tejas Dangodra" spread
   across the whole page in the background (also shows when printed).
2. **Copying is blocked** — text selection is disabled, right-click / drag / copy-save-view-source
   shortcuts are prevented, and if any text still reaches the clipboard it is replaced with a
   **"Made by Tejas Dangodra — All Rights Reserved"** notice.

> ⚠️ **Honest note:** this is *client-side only*. It marks copied text and prints the watermark, but a
> determined person can still bypass it. For real protection you need a server-rendered document with DRM.

---

## 1 · The watermark — clearly visible, tiled in the background

### Step A — add this CSS inside your `<style>` block

```css
/* Visible tiled watermark: background-size ~34vw 30vh gives a ~3-column diagonal grid. */
.wm-layer{
  position:fixed; inset:0; z-index:9999; pointer-events:none;
  background-repeat:repeat;
  background-size:34vw 30vh;   /* <-- controls how many marks show (smaller = more marks) */
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='560' height='360'><text x='280' y='190' font-family='Segoe UI,Arial,sans-serif' font-size='15' font-weight='600' letter-spacing='1' fill='rgba(120,100,175,0.10)' text-anchor='middle' transform='rotate(-24 280 190)'>Made by Tejas Dangodra</text></svg>");
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
| **More marks** | make `background-size` smaller | `34vw 30vh` ≈ 3 columns (my default) · `25vw 25vh` ≈ more · `20vw 18vh` = dense |
| **Fewer marks** | make `background-size` bigger | `50vw 50vh` = ~4 marks · `100vw 100vh` = 1 mark |
| **Stronger / more visible** | `fill='rgba(120,100,175, 0.10)'` → raise last number | `0.10` = clearly visible (my default) · `0.13` = stronger · `0.07` = softer · `0.04` = barely visible |
| **Bigger / bolder text** | `font-size='15'`, `font-weight='600'` | raise for more presence |
| **Different tilt** | `transform='rotate(-24 280 190)'` | change `-24`; keep the pivot at half of width/height |
| **Different colour** | the `rgb` part of `fill` | `120,100,175` = soft lavender — match your theme |
| **Different text** | the words `Made by Tejas Dangodra` | keep it short |

> ✅ **My preferred default:** clearly visible tiled marks (`background-size:34vw 30vh`), opacity **0.10**,
> font **15px / weight 600**, rotation **−24°**, soft lavender, text **"Made by Tejas Dangodra"**.

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
