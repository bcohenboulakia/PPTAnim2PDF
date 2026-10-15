<p align="center"><img width="200" alt="PPTAnim2PDF" src="https://github.com/user-attachments/assets/a86ee40a-6d44-4d8c-9e51-456552d84a05" /></p>

# PPTAnim2PDF

## Presentation

**PPTAnim2PDF** is a tool for converting animated PowerPoint presentations into static PPTX files or shareable PDFs. Starting from a `.pptx` presentation, it generates a sequence of static slides showing the successive animation states, preserving the order in which objects and text appear, disappear or change.

The project combines a Python script that directly reads and rewrites the `.pptx` OOXML package with a GitHub Actions workflow that exports the resulting static presentation to PDF using [office2pdf](https://github.com/developer0hye/office2pdf). The workflow runs entirely on GitHub and does not require PowerPoint or any local installation. The script can also be used independently to generate static PPTX files or, with LibreOffice, PDFs.

PPTAnim2PDF is heavily inspired by [PPspliT](https://www.maxonthenet.altervista.org/ppsplit.php) by Massimo Rimondini (maxonthegit). Many thanks to him, his tool has been invaluable to me over the years. PPspliT runs inside PowerPoint and relies on its VBA object model to interpret and transform slides. For standard desktop use, PPspliT remains the safer default, because it stays closer to PowerPoint’s native handling of the file. PPTAnim2PDF takes a different approach, using open-source components that can run without PowerPoint, including in batch or headless workflows. I created it to automate this conversion step in the CI/CD workflow I use to maintain course materials in my Git repository (see my [nbworkshop repository](https://github.com/bcohenboulakia/nbworkshop)).

## Installation and prerequisites

PPTAnim2PDF can be used in two different ways:

- **Using the GitHub workflow:** No local installation is required. The workflow installs its dependencies on a GitHub Actions runner and performs both the animation conversion and PDF export. You only need a GitHub repository containing the script, the workflow and the presentation to convert.
- **Running the Python script directly:** Python 3 and `lxml` are required. Install the dependency with `python -m pip install lxml`. To generate PDFs directly from the script, LibreOffice must also be installed; it is not needed for static PPTX output.

The two components are located in the `.github` directory:

- `.github/pptanim2pdf/pptanim2pdf.py`: The Python script that converts the animations into static slide states.
- `.github/workflows/convert-pptx.yml`: The GitHub Actions workflow that runs the script and, for PDF output, converts the static PPTX with office2pdf.

## Quick start

To use the GitHub workflow:

1. Fork the repository, or copy the script and workflow into the corresponding `.github` paths in your own repository.
2. Add the PowerPoint presentation to the repository and push it to GitHub.
3. Open the **Actions** tab, select **Convert PPTX**, and click **Run workflow**.
4. Enter the path of the PPTX file in the repository, choose the output format (`pdf` or `pptx`), and run the workflow.

The generated file is committed to the branch on which the workflow was run. Conversion is currently triggered manually; pushing a presentation does not start it automatically.

Direct use of the script is described in the [Conversion script](#conversion-script) section below.

## GitHub workflow

The GitHub Actions workflow runs on Windows. It first uses the Python script to generate a temporary static PPTX, then uses [office2pdf](https://github.com/developer0hye/office2pdf) to export it to PDF. When `pptx` output is selected, only the Python conversion step is performed. Microsoft PowerPoint and LibreOffice are not required on the runner.

The workflow allows you to specify the input and output paths, output format, and report level. For PDF conversion, the office2pdf version can be selected from the latest `main` commit (the default), the latest release, a specific release tag, or a specific commit SHA. The corresponding source revision and executable hash are recorded in the workflow summary. The generated file is committed back to the selected branch.

PDF fidelity depends on office2pdf. The results are generally good, but some rendering differences with PowerPoint remain. office2pdf is under active development and improves rapidly.

## Conversion script

The Python script can be run independently to generate a static PPTX, with one slide per animation state. It can also export a PDF through LibreOffice in headless mode.

```bash
python .github/pptanim2pdf/pptanim2pdf.py [-h] [-o OUTPUT] [-f {pptx,pdf}] [--report {none,summary,detail}] [--text-style-width-compensation] input
```

- `input`: Input PPTX file.
- `-h`, `--help`: Show the help message and exit.
- `-o OUTPUT`, `--output OUTPUT`: Output file. If omitted, `_split` is appended to the input file name for PPTX output, while PDF output reuses the input stem with the `.pdf` extension.
- `-f {pptx,pdf}`, `--format {pptx,pdf}`: Output format. `pptx` (the default) generates a static PPTX. `pdf` first generates a temporary static PPTX, then exports it using LibreOffice.
- `--report {none,summary,detail}`, `--report-level {none,summary,detail}`: Controls the conversion report. `none` disables the report, `summary` prints aggregate counts, and `detail` also lists skipped animation steps with the original slide number and the reason why no separate static slide was generated.
- `--soffice`: Path to the LibreOffice executable, if it is not found automatically (PDF output only).
- `--pdf-timeout`: Maximum time allowed for LibreOffice PDF export, in seconds (default: 120).
- `--text-style-width-compensation`: Enables legacy width compensation for text boxes after Bold text-style animations, for renderers that need it. Disabled by default.

## Compatibility

PPTAnim2PDF is intended for `.pptx` files produced by Microsoft PowerPoint 2007 or later, with best-effort support for recent PowerPoint desktop versions on Windows and macOS. Malformed or inconsistent `.pptx` packages, such as files with missing slide parts or broken slide relationships, are rejected with an explicit error message.

Embedding the fonts used by the presentation in the source PPTX is essential for the closest possible match to the original. Missing fonts may be substituted, changing text metrics, line wrapping, and layout.

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
- For Bold text-style animations, optional width compensation estimates the extra text width, reduces the right text padding first, and then widens the object if needed. This legacy workaround is disabled by default and can be enabled with `--text-style-width-compensation`.
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
