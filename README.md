<p align="center"><img width="200" alt="PPTAnim2PDF" src="https://github.com/user-attachments/assets/4c647bf0-299a-4cce-afc0-84c2f72257c6" /></p>

# PPTAnim2PDF

## Presentation

**PPTAnim2PDF** is a Python script that converts animated presentations into static, shareable PDFs by turning each slide into a sequence of slides showing successive animation states. This script is heavily inspired by [PPspliT](https://www.maxonthenet.altervista.org/ppsplit.php) by Massimo Rimondini (maxonthegit). Many thanks to him; his tool has been invaluable to me over the years.

The main difference between the two is that PPTAnim2PDF is a standalone script that directly parses and rewrites the `.pptx` OOXML package, using only open-source, cross-platform dependencies. PPspliT takes a different approach: it runs inside PowerPoint and relies on its VBA object model to interpret and transform slides. For standard desktop use, PPspliT remains the safer default, because it stays closer to PowerPoint’s native handling of the file. PPTAnim2PDF is designed for workflows where relying on PowerPoint is undesirable or impossible, such as batch or headless conversion. I created it to automate this conversion step so that I could integrate it into the CI/CD workflow I use to maintain the course materials in my Git repository (see my [nbworkshop repository](https://github.com/bcohenboulakia/nbworkshop)).

## Usage

```bash
pptanim2pdf.py [-h] [-o OUTPUT] [--report {none,summary,detail}] input
```

Convert an animated PPTX into a static PPTX with one slide per animation state.

- `input`: Input PPTX file.
- `-h`, `--help`: Show this help message and exit.
- `-o OUTPUT`, `--output OUTPUT`: Output PPTX file. If omitted, `_split` is appended to the input file name.
- `--report {none,summary,detail}`, `--report-level {none,summary,detail}`: Report detail level: `none` = no report, `summary` = concise report, `detail` = detailed report.

## Compatibility

PPTAnim2PDF is intended for `.pptx` files produced by Microsoft PowerPoint 2007 or later, with best-effort support for recent PowerPoint desktop versions on Windows and macOS.

PPTAnim2PDF directly reads and rewrites the `.pptx` Office Open XML package. The conversion process is developed against the PresentationML structures emitted by Microsoft PowerPoint, with tests currently focused on PowerPoint 2019. Different PowerPoint versions, platforms, or third-party exporters may encode the same animation differently. When such variants are detected but not understood, the converter reports them in the conversion summary instead of silently producing incorrect slides.

Legacy `.ppt` files are not supported. They should first be converted to `.pptx` using PowerPoint or another compatible tool.

Presentations exported by LibreOffice, Keynote, Google Slides, or other software may work if their animations are encoded using standard PresentationML structures recognized by the converter, but they are not guaranteed to be fully supported.

## Animations

### Converted animations

PPTAnim2PDF converts supported animations into explicit slide states. It flattens timing and trigger parameters into a linear sequence of static visual states.

- Detected entrance animations become object or paragraph appearances.
- Detected exit animations become object or paragraph disappearances.
- Detected motion-path animations move objects to their final position.

The following animation variants are supported during conversion:

- `On click` animations create a new slide state.
- `After previous` animations also create a new slide state; automatic timing is replaced with an explicit slide step.
- `With previous` animations without delay are merged into the current slide state.
- `With previous` animations with delay create a new slide state; the delay value itself is ignored.
- Dynamic entrance and exit effects are reduced to their final visibility state.
- Multiple motion paths applied to the same object are accumulated.
- Motion paths applied before an object appears are preserved, so the object appears directly at its updated position.
- Simultaneous entrance and motion-path animations are merged into the same resulting slide state.

### Ignored animations that cannot be converted into a distinct static state

Some animations cannot be converted faithfully because a PDF-compatible slide sequence can only represent stable visual states, not transient motion or time-based behavior. In such cases, the original PowerPoint file should be edited to express the effect as explicit static steps, for example by replacing a closed-loop motion path with separate duplicated objects or slides.

- Closed or zero-length motion paths.
- Redundant visibility changes.
- Animation steps whose final result is visually identical to the previous slide.

### Ignored animations outside the supported scope

These animations are not converted because supporting them would require much heavier slide analysis and transformation.

- Word-by-word text animations.
- Letter-by-letter or character-by-character text animations.
- Unsupported motion-path encodings.
- Media playback animations.
- Animations targeting internal parts of charts or SmartArt.
