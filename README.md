<p align="center"><img width="200" alt="PPTAnim2PDF" src="https://github.com/user-attachments/assets/a86ee40a-6d44-4d8c-9e51-456552d84a05" /></p>

# PPTAnim2PDF

## Presentation

**PPTAnim2PDF** is a Python script that converts animated presentations into static, shareable PDFs by turning each slide into a sequence of slides showing successive animation states. This script is heavily inspired by [PPspliT](https://www.maxonthenet.altervista.org/ppsplit.php) by Massimo Rimondini (maxonthegit). Many thanks to him; his tool has been invaluable to me over the years.

The main difference between the two is that PPTAnim2PDF is a standalone script that directly parses and rewrites the `.pptx` OOXML package, using only open-source, cross-platform dependencies. PPspliT takes a different approach: it runs inside PowerPoint and relies on its VBA object model to interpret and transform slides. For standard desktop use, PPspliT remains the safer default, because it stays closer to PowerPoint’s native handling of the file. PPTAnim2PDF is designed for workflows where relying on PowerPoint is undesirable or impossible, such as batch or headless conversion. I created it to automate this conversion step so that I could integrate it into the CI/CD workflow I use to maintain the course materials in my Git repository (see my [nbworkshop repository](https://github.com/bcohenboulakia/nbworkshop)).

PPTAnim2PDF requires Python 3 and `lxml`.

## Usage

```bash
pptanim2pdf.py [-h] [-o OUTPUT] [-f {pptx,pdf}] [--report {none,summary,detail}] [--no-text-style-width-compensation] input
```

Convert an animated PPTX into a static PPTX or PDF with one slide/page per animation state.

- `input`: Input PPTX file.
- `-h`, `--help`: Show this help message and exit.
- `-o OUTPUT`, `--output OUTPUT`: Output file. If omitted, `_split` is appended to the input file name for PPTX output, while PDF output reuses the input stem with the `.pdf` extension.
- `-f {pptx,pdf}`, `--format {pptx,pdf}`: Output format. `pptx` generates a static PPTX. `pdf` generates a PDF by first creating a temporary static PPTX and then exporting it with LibreOffice in headless mode.
- `--report {none,summary,detail}`, `--report-level {none,summary,detail}`: Controls the conversion report. `none` disables the report, `summary` prints aggregate counts, and `detail` also lists skipped animation steps with the original slide number and the reason why no separate static slide was generated.
- `--no-text-style-width-compensation`: Disables the default width compensation applied to text boxes when a Bold text-style animation may make the text wider. By default, PPTAnim2PDF tries to prevent unwanted line wrapping in LibreOffice/PDF rendering.

## Compatibility

PPTAnim2PDF is intended for `.pptx` files produced by Microsoft PowerPoint 2007 or later, with best-effort support for recent PowerPoint desktop versions on Windows and macOS. Malformed or inconsistent `.pptx` packages, such as files with missing slide parts or broken slide relationships, are rejected with an explicit error message.


PPTAnim2PDF directly reads and rewrites the `.pptx` Office Open XML package. The conversion process is developed against the PresentationML structures emitted by Microsoft PowerPoint, with tests currently focused on PowerPoint 2019. Different PowerPoint versions, platforms, or third-party exporters may encode the same animation differently. When such variants are detected but not understood, the converter reports them in the conversion summary instead of silently producing incorrect slides.

Legacy `.ppt` files are not supported. They should first be converted to `.pptx` using PowerPoint or another compatible tool.

Presentations exported by LibreOffice, Keynote, Google Slides, or other software may work if their animations are encoded using standard PresentationML structures recognized by the converter, but they are not guaranteed to be fully supported.

## Animations

### Converted animations

PPTAnim2PDF converts supported animations into explicit slide states. It flattens timing and trigger parameters into a linear sequence of static visual states.

- Detected entrance animations become object or paragraph appearances.
- Detected exit animations become object or paragraph disappearances.
- Detected motion-path animations move objects to their final position.
- Detected Grow/Shrink emphasis effects resize objects to their final size.
- Detected Spin emphasis effects rotate objects to their final angle.
- Detected Transparency effects apply the final opacity.
- Detected Fill Color and Line Color effects apply the final fill or line color.
- Detected Font Color effects apply the final text color.
- Detected Bold, Italic, Underline, and Strikethrough effects apply the final text style.
- For Bold text-style animations, PPTAnim2PDF also tries to compensate for the increased text width: it estimates the extra width, reduces the right text padding first, and then widens the object if the padding is not enough. This behavior can be disabled with `--no-text-style-width-compensation`.
- Some “After animation” effects are converted when they correspond to a static final state, such as dimming text or hiding an object after an animation.

The following animation variants are supported during conversion:

- `On click` animations create a new slide state.
- `After previous` animations create a new slide state, except when they occur before any click and have no delay, in which case they are merged into the initial slide state.
- `With previous` animations without delay are merged into the current slide state.
- `With previous` animations with delay create a new slide state; the delay value itself is ignored.
- Dynamic entrance and exit effects are reduced to their final visibility state.
- Multiple transformations applied to the same object are accumulated in timeline order.
- Transformations applied before an object appears are preserved, so the object appears directly at its updated position.
- Simultaneous entrance and transformation animations are merged into the same resulting slide state.

### Ignored animations that cannot be converted into a distinct static state

Some animations cannot be converted faithfully because a static slide sequence can only represent stable visual states, not transient motion or time-based behavior. In such cases, the original PowerPoint file should be edited to express the effect as explicit static steps, for example by replacing a closed-loop motion path with separate duplicated objects or slides.

- Closed or zero-length motion paths.
- Grow/Shrink effects whose final size is unchanged.
- Spin effects whose final angle is unchanged.
- Redundant visibility changes.
- Style or color effects that apply a state already present on the target.
- Transformations applied while the target object is not visible.
- Animation steps whose final result is visually identical to the previous slide.

### Ignored animations outside the supported scope

These animations are not converted because supporting them would require much heavier slide analysis and transformation.

- Word-by-word text animations.
- Letter-by-letter or character-by-character text animations.
- Other text animations targeting character ranges instead of full paragraphs.
- Media playback animations.
- Animations targeting internal parts of charts or SmartArt.
- Animations targeting embedded objects or other complex objects whose internal visual state is not represented as ordinary slide shapes.
- Emphasis effects other than the supported Grow/Shrink, Spin, Transparency, Fill Color, Line Color, Font Color, Bold, Italic, Underline, and Strikethrough effects.
