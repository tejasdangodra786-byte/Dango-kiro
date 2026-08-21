# Tejas Dangodra — Watermark & Copy-Attribution Style Guide

A reusable recipe for my signature look on any HTML study-guide / document. Copy the two snippets below.

The effect has two parts:

1. **A very sparse, subtle watermark** — only about **4 faint marks** spread across the whole page (also
   shows when printed).
2. **Copy is allowed** — but when someone copies text, my credit **"Made by Tejas Dangodra" is woven all
   through** the copied text (a header, a tag every ~12 words, and a footer).

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
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='800' height='560'><text x='400' y='290' font-family='Segoe UI,Arial,sans-serif' font-size='13' font-weight='500' letter-spacing='1' fill='rgba(120,100,175,0.04)' text-anchor='middle' transform='rotate(-24 400 290)'>Tejas Dangodra</text></svg>");
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

## 3 · Copy is allowed — but credit is woven throughout

Do **not** disable text selection. Instead, add this script just before `</body>`:

```html
<script>
(function(){
  var TAG    = " [Made by Tejas Dangodra] ";
  var HEADER = "Made by Tejas Dangodra\n\n";
  var FOOTER = "\n\n— © Made by Tejas Dangodra";

  // Insert the tag after roughly every ~12 words so it appears all over.
  function weave(text){
    var words = text.split(/(\s+)/), out = [], n = 0;
    for(var i=0;i<words.length;i++){
      out.push(words[i]);
      if(/\S/.test(words[i])){ n++; if(n % 12 === 0){ out.push(TAG); } }
    }
    return HEADER + out.join('') + FOOTER;
  }

  document.addEventListener('copy', function(e){
    try{
      var sel = (window.getSelection && window.getSelection().toString()) || '';
      if(sel.length){ e.clipboardData.setData('text/plain', weave(sel)); e.preventDefault(); }
    }catch(err){ /* let normal copy proceed on failure */ }
  });
})();
</script>
```

Change the frequency by editing `n % 12` (smaller number = credit appears more often).

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
- [ ] **Do NOT** set `user-select:none` — copying stays allowed
- [ ] Copy `<script>` with the `weave()` credit-injection before `</body>`
- [ ] Visible **"Made by Tejas Dangodra"** footer

---

## 6 · See it working

A live demo is in this repo: **`Watermark_Demo.html`** — open it, adjust the sliders, select-and-copy some
text to see the woven credit, and print-preview to see the ~4 marks.

---

*Reference maintained for Tejas Dangodra. Reuse across all study-guide HTML files.*
