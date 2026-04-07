# revealjs-to-pptx

Converts [reveal.js](https://revealjs.com) HTML slide decks into editable PowerPoint (`.pptx`) files, preserving the **publictimes_CMYK / base_2022** design system:

- Black background, white text
- Cyan `#00CCFF` for bold text and list markers
- Magenta `#FF0099` for running-head pills and TOC underlines
- Yellow `#FFEE00` for links and slide counter
- Georgia as a substitute for the proprietary *Happy Times* serif
- Em-dash list style, cyan blockquote bars, background images

---

## Requirements

- Python 3.8 or newer
- The reveal.js project folder (HTML file + `img/` subfolder with images)

---

## Installation

**1. Clone or download this repository**

```bash
git clone https://github.com/YOUR_USERNAME/revealjs-to-pptx.git
cd revealjs-to-pptx
```

Or just download `revealjs_to_pptx.py` directly if you don't want the whole repo.

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
python revealjs_to_pptx.py path/to/Week2_Seminar.html
```

This creates `Week2_Seminar.pptx` in the same folder as the HTML file.

To specify a different output location:

```bash
python revealjs_to_pptx.py path/to/Week2_Seminar.html --out path/to/output.pptx
```

---

## Expected folder structure

The script expects the standard reveal.js export layout:

```
Week2_Seminar.html
img/
    balbi.jpg
    driscoll.jpg
    w2_seminar_title2.png
    ...
```

Images are resolved by filename from the `img/` folder next to the HTML file. Missing images are replaced with a grey placeholder so the conversion still completes.

---

## Slide types detected automatically

| Type | Detected by |
|------|-------------|
| Title | `h1` present on first slide, or `data-state="title"` |
| Table of contents | `class="toc"` |
| Section break / center | `class="center"` |
| Reading discussion | `.twothird` containing an `<img>` |
| Two-column exercise | `class="twocol"` |
| Standard | everything else |

Running-head labels are extracted automatically from the inline CSS `header:after { content: "…" }` rules in the HTML.

---

## Note on HTML structure

This script handles the **nested `<li>` pattern** used in this particular deck, where list items are nested inside each other rather than being siblings. If your colleague's slides use a different HTML structure, the list rendering may need adjustment.

---

## Limitations

- The proprietary *Happy Times* font is substituted with Georgia. If you install Happy Times on your system and rename the font constant in the script, it will use it instead.
- Slides with complex custom layouts not covered by the six detected types will fall back to the standard renderer and may need manual clean-up in PowerPoint.
- Animated or interactive reveal.js elements (fragments, transitions) are not converted.
